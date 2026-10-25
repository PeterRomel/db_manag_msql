import os
from dotenv import load_dotenv

load_dotenv()

import os
import re
import time
import json
from enum import Enum
from typing import Optional
import pandas as pd
from pydantic import BaseModel, Field
from tqdm import tqdm
from google import genai
from google.genai import types
from google.genai.errors import ClientError

# ==============================================================================
# CONFIGURATION
# ==============================================================================
API_KEY = os.getenv("GCP_API_KEY")  # os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY_HERE")
INPUT_FILE = "used_items_extract.csv"
OUTPUT_USER_TARGET_FORMAT = "final_deduplicated_catalog_user_format.csv"
OUTPUT_DUPLICATES_ONLY = "duplicate_clusters_for_review.csv"
OUTPUT_ALL = "deduplicated_pharmacy_master.csv"
CHECKPOINT_FILE = "checkpoint_parsed_items.json"

# gemini-3.5-flash-lite provides 500 RPD on Google AI Studio Free Tier
MODEL_ID = "gemini-3.5-flash-lite"

# 100 items/batch = ~68 total requests for ~6,772 items
BATCH_SIZE = 100
REQUEST_DELAY_SECONDS = 10  # Keeps throughput safely under 15 RPM


# ==============================================================================
# 1. PRE-CLEANER (Strips Cashier Noise & ERP Artifacts)
# ==============================================================================
def pre_clean_erp_name(text: str) -> str:
    if not isinstance(text, str):
        return ""

    s = text.strip()
    s = re.sub(r"\b1000\s*mMG\b", "1000 MG", s, flags=re.I)
    s = re.sub(r"\bmMG\b", "MG", s, flags=re.I)
    s = re.sub(r"[@#*!?\[\]{}()]", " ", s)
    s = re.sub(r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP|جنيه)\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+LE\b", " ", s, flags=re.I)

    status_keywords = [
        r"\bNEW\s*PRICE\b",
        r"\bOLD\s*PRICE\b",
        r"\bN\.?P\.?\d*\b",
        r"\bN\/PRICE\b",
        r"\bOLD\/PRICE\b",
        r"\bNEU\s*\.?PRICE\b",
        r"\bNEW\b",
        r"\bOLD\b",
        r"\bPRES\b",
        r"\bEGY\b",
        r"سعر\s*جديد",
        r"سعر\s*قديم",
        r"بدون\s*خصم",
        r"موقوف",
        r"\bOFFER\b",
        r"\bFREE\b",
        r"\bGIFT\b",
    ]
    for pattern in status_keywords:
        s = re.sub(pattern, " ", s, flags=re.I)

    s = re.sub(r"\s+(m|mup|sedico|pharco|amryia|t3a|aug)\s*$", "", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


# ==============================================================================
# 2. ENUM & PYDANTIC STRUCTURED OUTPUT SCHEMA
# ==============================================================================
class DosageFormEnum(str, Enum):
    TABLET = "TABLET"
    CAPSULE = "CAPSULE"
    SYRUP = "SYRUP"
    SUSPENSION = "SUSPENSION"
    DROPS = "DROPS"
    CREAM = "CREAM"
    OINTMENT = "OINTMENT"
    GEL = "GEL"
    AMPOULE = "AMPOULE"
    VIAL = "VIAL"
    SACHET = "SACHET"
    SUPPOSITORY = "SUPPOSITORY"
    SHAMPOO = "SHAMPOO"
    SOAP = "SOAP"
    SPRAY = "SPRAY"
    INHALER = "INHALER"
    SOLUTION = "SOLUTION"
    LOTION = "LOTION"
    PATCH = "PATCH"
    MEDICAL_SUPPLY = "MEDICAL_SUPPLY"
    OTHER = "OTHER"


class ParsedProduct(BaseModel):
    item_code: str
    original_input: str
    canonical_brand: str = Field(
        description="Standardized English trade brand name only. Never substitute with generic molecule."
    )
    strength: Optional[str] = Field(
        default="",
        description="Active pharmaceutical strength (e.g. '500 MG', '40/10/25 MG').",
    )
    dosage_form: DosageFormEnum = Field(
        description="Strict standardized dosage form from the enum."
    )
    pack_size: Optional[str] = Field(
        default="",
        description="Package count or volume: '20 TAB', '100 ML', '15 GM', '3 AMP'. Infer if standard in Egypt.",
    )


class ProductBatchResponse(BaseModel):
    products: list[ParsedProduct]


SYSTEM_INSTRUCTION = """You are a senior pharmaceutical database auditor for Egyptian pharmacy catalogs.
Extract exact structured attributes from product strings.

RULES:
1. BRAND ISOLATION: Keep commercial trade brands distinct. PARAMOL != PANADOL != CETAL != ABIMOL.
2. STRENGTH PRECISION: Preserve combination ratios ('10/160 MG', '50/500 MG').
3. PACK SIZES: Normalize counts to standard short forms: '20 TAB', '30 CAP', '100 ML', '6 AMP', '15 GM'.
4. IMPLICIT PACK SIZES: If pack size is omitted but standard in Egypt (e.g. Bivatracin spray is 150 GM), infer it.
5. ARABIC TRANSLATION: Translate Arabic trade terms to English canonical equivalents (e.g., 'أقراص' -> TABLET).
"""


# ==============================================================================
# 3. EXTRACTION ENGINE WITH SELF-HEALING RATE-LIMIT RETRIES
# ==============================================================================
def parse_batch_with_retry(
    client: genai.Client, items_batch: list[dict], max_retries: int = 5
) -> list[dict]:
    prompt = "Parse the following Egyptian pharmacy records:\n\n"
    for item in items_batch:
        prompt += f"- Code: {item['code']}, Raw: '{item['cleaned_name']}'\n"

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_INSTRUCTION,
                    response_mime_type="application/json",
                    response_schema=ProductBatchResponse,
                    temperature=0.0,
                ),
            )
            data = ProductBatchResponse.model_validate_json(response.text)
            return [p.model_dump() for p in data.products]

        except ClientError as e:
            error_str = str(e)
            if "limit: 0" in error_str or "NOT_FOUND" in error_str:
                raise RuntimeError(f"Permanent configuration error: {error_str}")

            # Parse dynamic wait time provided by Google API
            retry_match = re.search(r"retryDelay':\s*'(\d+)s", error_str)
            if not retry_match:
                retry_match = re.search(r"retry in (\d+(\.\d+)?)s", error_str)

            wait_time = int(float(retry_match.group(1))) + 2 if retry_match else 20
            print(
                f"\n[!] Rate limit reached. Pausing {wait_time}s as requested by API (Attempt {attempt+1}/{max_retries})..."
            )
            time.sleep(wait_time)

        except Exception as e:
            print(f"\n[!] Unexpected error: {e}. Pausing 10s...")
            time.sleep(10)

    raise RuntimeError(f"Batch failed after {max_retries} attempts.")


# ==============================================================================
# 4. TYPE-SAFE SORTING KEY (Fixes the Pandas TypeError Bug)
# ==============================================================================
def safe_sort_key(code):
    """
    Guarantees no comparison between int and str.
    Sorts numeric codes numerically, followed by alphanumeric codes alphabetically.
    """
    s = str(code).strip()
    if s.isdigit():
        return (0, int(s), "")
    return (1, 0, s)


# ==============================================================================
# 5. EXECUTION & POST-PROCESSING PIPELINE
# ==============================================================================
def run_pipeline():
    if not API_KEY or API_KEY == "YOUR_GEMINI_API_KEY_HERE":
        raise ValueError("Please provide a valid GEMINI_API_KEY.")

    client = genai.Client(api_key=API_KEY)
    df = pd.read_csv(INPUT_FILE)
    print(f"[*] Loaded {len(df)} records from '{INPUT_FILE}'")

    items_to_process = [
        {
            "code": str(row["Item_Code"]),
            "raw_name": str(row["Item_English_Name"]),
            "cleaned_name": pre_clean_erp_name(str(row["Item_English_Name"])),
        }
        for _, row in df.iterrows()
    ]

    # 1. Load checkpoint if resuming an earlier run
    all_parsed = []
    start_index = 0
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                all_parsed = json.load(f)
            start_index = len(all_parsed)
            print(
                f"[*] Resuming from checkpoint: {start_index}/{len(items_to_process)} records already parsed."
            )
        except Exception:
            all_parsed = []

    # 2. Run API batches
    if start_index < len(items_to_process):
        print(f"[*] Target Model: {MODEL_ID} | Batch Size: {BATCH_SIZE}")
        for i in tqdm(range(start_index, len(items_to_process), BATCH_SIZE)):
            batch = items_to_process[i : i + BATCH_SIZE]
            results = parse_batch_with_retry(client, batch)
            all_parsed.extend(results)

            # Checkpoint after every batch
            with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
                json.dump(all_parsed, f, ensure_ascii=False)

            time.sleep(REQUEST_DELAY_SECONDS)

    print(
        "\n[*] All items parsed. Commencing Pandas deduplication and entity clustering..."
    )
    parsed_df = pd.DataFrame(all_parsed)

    # 3. Clean and normalize fields
    parsed_df["canonical_brand"] = (
        parsed_df["canonical_brand"]
        .fillna("UNKNOWN")
        .astype(str)
        .str.upper()
        .str.strip()
    )
    parsed_df["strength"] = (
        parsed_df["strength"].fillna("").astype(str).str.upper().str.strip()
    )
    parsed_df["dosage_form"] = (
        parsed_df["dosage_form"].fillna("OTHER").astype(str).str.upper().str.strip()
    )
    parsed_df["pack_size"] = (
        parsed_df["pack_size"].fillna("").astype(str).str.upper().str.strip()
    )
    parsed_df["item_code"] = parsed_df["item_code"].astype(str).str.strip()
    parsed_df["original_input"] = parsed_df["original_input"].astype(str).str.strip()

    # 4. Generate composite matching key (Rule 1-A strictly enforced)
    parsed_df["Canonical_SKU_Signature"] = (
        parsed_df["canonical_brand"]
        + " | "
        + parsed_df["strength"]
        + " | "
        + parsed_df["dosage_form"]
        + " | "
        + parsed_df["pack_size"]
    )

    # 5. Calculate duplicate metrics
    parsed_df["Duplicate_Count"] = (
        parsed_df.groupby("Canonical_SKU_Signature")["item_code"].transform("count") - 1
    )

    # 6. Collect sibling codes using the type-safe sort key
    def collect_codes(series):
        unique_codes = list(series.unique())
        sorted_codes = sorted(unique_codes, key=safe_sort_key)
        return "; ".join(sorted_codes)

    sibling_codes_map = (
        parsed_df.groupby("Canonical_SKU_Signature")["item_code"]
        .apply(collect_codes)
        .to_dict()
    )
    parsed_df["duplicate_codes"] = parsed_df["Canonical_SKU_Signature"].map(
        sibling_codes_map
    )

    # Determine Master code (oldest/lowest numerical code)
    def pick_master_code(series):
        unique_codes = list(series.unique())
        return sorted(unique_codes, key=safe_sort_key)[0]

    master_code_map = (
        parsed_df.groupby("Canonical_SKU_Signature")["item_code"]
        .apply(pick_master_code)
        .to_dict()
    )
    parsed_df["Canonical_Master_Item_Code"] = parsed_df["Canonical_SKU_Signature"].map(
        master_code_map
    )
    parsed_df["Is_Master_Record"] = (
        parsed_df["item_code"] == parsed_df["Canonical_Master_Item_Code"]
    )

    # 7. Build Canonical Master Title
    parsed_df["proposed_master_name"] = (
        (
            parsed_df["canonical_brand"]
            + " "
            + parsed_df["strength"]
            + " "
            + parsed_df["dosage_form"]
            + " "
            + parsed_df["pack_size"]
        )
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # 8. Build target 6-column user format
    target_df = parsed_df.rename(
        columns={
            "item_code": "code",
            "original_input": "name",
            "Duplicate_Count": "number of duplicates",
            "proposed_master_name": "proposed master name",
        }
    )

    # Sort alphabetically by proposed master name, then by code
    target_df["sort_code_helper"] = target_df["code"].apply(safe_sort_key)
    target_df = target_df.sort_values(
        by=["proposed master name", "sort_code_helper"]
    ).drop(columns=["sort_code_helper"])

    target_df["extra comments"] = target_df.apply(
        lambda r: (
            f"Duplicate record clustered into '{r['proposed master name']}'"
            if r["number of duplicates"] > 0
            else "Unique SKU"
        ),
        axis=1,
    )

    final_export = target_df[
        [
            "code",
            "name",
            "number of duplicates",
            "duplicate_codes",
            "proposed master name",
            "extra comments",
        ]
    ]

    # 9. Save all target CSV files
    final_export.to_csv(OUTPUT_USER_TARGET_FORMAT, index=False, encoding="utf-8-sig")
    parsed_df.to_csv(OUTPUT_ALL, index=False, encoding="utf-8-sig")

    duplicates_df = final_export[final_export["number of duplicates"] > 0]
    duplicates_df.to_csv(OUTPUT_DUPLICATES_ONLY, index=False, encoding="utf-8-sig")

    print("\n" + "=" * 70)
    print("DEDUPLICATION AUDIT COMPLETE")
    print("=" * 70)
    print(f"Total Rows Processed:                    {len(final_export)}")
    print(
        f"Unique Canonical Master SKUs:            {parsed_df['Canonical_SKU_Signature'].nunique()}"
    )
    print(f"Total Duplicate Rows Identified:         {len(duplicates_df)}")
    print(
        f"Total Unique/Single Rows:                {len(final_export) - len(duplicates_df)}"
    )
    print(f"\n[+] Master CSV (All Items, User Format): '{OUTPUT_USER_TARGET_FORMAT}'")
    print(f"[+] Duplicates Review File:              '{OUTPUT_DUPLICATES_ONLY}'")
    print(f"[+] Full Audit Ledger:                   '{OUTPUT_ALL}'")
    print("=" * 70)


if __name__ == "__main__":
    run_pipeline()
