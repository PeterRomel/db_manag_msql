import os
import re
import csv
import time
import json
import glob
import pandas as pd
from typing import List, Optional
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from google.genai.errors import ClientError
from dotenv import load_dotenv

load_dotenv()

# ==============================================================================
# 1. CONFIGURATION & LIMIT SAFEGUARDS
# ==============================================================================
API_KEY = os.getenv("GEMINI_API_KEY", os.getenv("GCP_API_KEY"))
# Using the standard flash model available on free tier.
# If you specifically have access to a 'lite' model, change this to 'gemini-2.0-flash-lite-preview-02-05'
MODEL_ID = "gemini-3.5-flash-lite"

BATCH_SIZE = 30  # 30 items per prompt is the sweet spot for Flash models
REQUEST_DELAY_SECONDS = 5  # Ensures we stay under 15 Requests Per Minute
MAX_RETRIES = 5

CHECKPOINT_FILE = "checkpoint_parsed_catalog.json"
OUTPUT_FILE = "final_deduplicated_catalog.csv"


# ==============================================================================
# 2. PYDANTIC SCHEMAS (Forces Gemini to return perfect JSON)
# ==============================================================================
class ParsedItem(BaseModel):
    item_code: str = Field(description="The original item code provided.")
    canonical_brand: str = Field(
        description="Standardized English trade brand name only (e.g., 'PANADOL', 'CONCOR')."
    )
    strength: str = Field(
        description="Active strength (e.g., '500 MG', '5/160 MG'). Leave blank if none."
    )
    dosage_form: str = Field(
        description="Standardized form: TAB, CAP, SYRUP, SUSP, DROPS, AMP, VIAL, CREAM, OINT, GEL, SACHET, SUPP, SPRAY, LOTION."
    )
    pack_size: str = Field(
        description="Package count or volume (e.g., '20 TAB', '100 ML', '15 GM')."
    )


class BatchResponse(BaseModel):
    items: list[ParsedItem]


# ==============================================================================
# 3. PYTHON DETERMINISTIC SANITIZER
# ==============================================================================
def clean_trade_name(text: str) -> str:
    """Strips prices, cashier noise, and fixes unit typos."""
    if not isinstance(text, str):
        return ""
    s = text.strip()

    # Normalize common Egyptian cashier typos
    s = re.sub(r"\b1000\s*mMG\b", "1000 MG", s, flags=re.I)
    s = re.sub(r"\bmMG\b", "MG", s, flags=re.I)
    s = re.sub(r"\bCABS(ULES)?\b", "CAP", s, flags=re.I)
    s = re.sub(r"\bTABLES\b", "TAB", s, flags=re.I)
    s = re.sub(r"\bSYRY?P\b", "SYRUP", s, flags=re.I)

    # Remove symbols
    s = re.sub(r"[@#*!?\[\]{}()]", " ", s)

    # Strip explicit currency patterns (e.g., 25LE, 30.5 EGP, 50 جنيه)
    s = re.sub(r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP|L\.E\.|جنيه)\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+LE\b", " ", s, flags=re.I)

    # Strip trailing naked prices (e.g. "curam 457mg 70ml 97" -> drops "97")
    s = re.sub(
        r"(\b(?:TAB|CAP|TABLETS|CAPSULES|ML|GM|G|SACH|SACHETS|AMP|VIAL|PIECES|P|DOSE|PUFF)\b\s*)\d{2,4}$",
        r"\1",
        s.strip(),
        flags=re.I,
    )
    s = re.sub(r"(\b(?:ML|GM|MG)\b\s*)\d{2,4}$", r"\1", s.strip(), flags=re.I)

    # Strip lifecycle and status noise
    noise_patterns = [
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
    for p in noise_patterns:
        s = re.sub(p, " ", s, flags=re.I)

    # Strip cashier and company suffixes at end of string
    s = re.sub(
        r"\s+\b(m|k|pres|aug|t3a|sedico|pharco|amryia|amrya|mup|eipico)\b\s*$",
        "",
        s,
        flags=re.I,
    )

    return re.sub(r"\s+", " ", s).strip().upper()


def get_brand_root(name: str) -> str:
    """Extracts brand root WITHOUT splitting hyphenated drug prefixes."""
    clean = re.sub(r"^[^a-zA-Z0-9\u0621-\u064A]+", "", name.strip())
    if not clean:
        return "OTHER"

    match_compound = re.match(
        r"^(CO|E|TRI|X|NO|DE|BIO|PRO|ULTRA)[-_ ]+([a-zA-Z0-9]+)", clean, flags=re.I
    )
    if match_compound:
        return f"{match_compound.group(1).upper()}-{match_compound.group(2).upper()}"

    tokens = re.split(r"[\s/]+", clean)
    return tokens[0].upper() if tokens else "OTHER"


def safe_sort_key(code):
    s = str(code).strip()
    return (0, int(s), "") if s.isdigit() else (1, 0, s)


# ==============================================================================
# 4. AI EXTRACTION ENGINE
# ==============================================================================
def extract_attributes_with_gemini(client: genai.Client, batch: list) -> list:
    prompt = "Extract the pharmaceutical attributes for the following Egyptian pharmacy items:\n\n"
    for item in batch:
        prompt += f"- Code: {item['code']} | Name: '{item['clean_name']}'\n"

    for attempt in range(MAX_RETRIES):
        try:
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=BatchResponse,
                    temperature=0.0,
                ),
            )
            # Parse the validated JSON response
            data = json.loads(response.text)
            return data.get("items", [])

        except ClientError as e:
            error_str = str(e)
            if "limit: 0" in error_str or "NOT_FOUND" in error_str:
                raise RuntimeError(f"Fatal API error: {error_str}")

            print(
                f"\n[!] Rate limit hit. Sleeping 15s (Attempt {attempt+1}/{MAX_RETRIES})..."
            )
            time.sleep(15)
        except Exception as e:
            print(f"\n[!] Network error: {e}. Retrying in 10s...")
            time.sleep(10)

    # Fallback if API fails completely
    return [
        {
            "item_code": item["code"],
            "canonical_brand": item["clean_name"],
            "strength": "",
            "dosage_form": "",
            "pack_size": "",
        }
        for item in batch
    ]


# ==============================================================================
# 5. MAIN PIPELINE
# ==============================================================================
def run():
    print("=================================================================")
    print("  GENIUS ERP PHARMACEUTICAL DEDUPLICATION ENGINE (HYBRID AI)     ")
    print("=================================================================")

    if not API_KEY:
        print("[ERROR] GEMINI_API_KEY environment variable not set.")
        return

    client = genai.Client(api_key=API_KEY)

    # 1. Merge all CSV parts automatically
    csv_files = glob.glob("used_items_extract_part_*.csv")
    if not csv_files:
        # Fallback to single file if parts don't exist
        csv_files = (
            ["used_items_extract.csv"]
            if os.path.exists("used_items_extract.csv")
            else []
        )

    if not csv_files:
        print("[ERROR] No CSV files found to process.")
        return

    print(f"[*] Found {len(csv_files)} CSV files. Merging and cleaning...")
    all_records = []
    for file in csv_files:
        df = pd.read_csv(file)
        for _, row in df.iterrows():
            code = str(row.iloc[0]).strip()
            raw_name = str(row.iloc[1]).strip()
            clean_name = clean_trade_name(raw_name)
            root = get_brand_root(clean_name)
            all_records.append(
                {
                    "code": code,
                    "raw_name": raw_name,
                    "clean_name": clean_name,
                    "root": root,
                }
            )

    print(f"[*] Total items loaded: {len(all_records)}")

    # 2. Group by Brand Root
    root_groups = {}
    for r in all_records:
        root_groups.setdefault(r["root"], []).append(r)

    # 3. Load Checkpoint
    parsed_results = {}  # code -> ParsedItem dict
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                parsed_results = json.load(f)
            print(
                f"[*] Resuming from checkpoint: {len(parsed_results)} items already parsed."
            )
        except Exception:
            pass

    # 4. Process Singletons (FREE - No API Calls)
    singletons_resolved = 0
    for root, items in root_groups.items():
        if len(items) == 1:
            code = items[0]["code"]
            if code not in parsed_results:
                parsed_results[code] = {
                    "item_code": code,
                    "canonical_brand": items[0]["clean_name"],
                    "strength": "",
                    "dosage_form": "",
                    "pack_size": "",
                }
                singletons_resolved += 1

    if singletons_resolved > 0:
        print(
            f"[*] Auto-resolved {singletons_resolved} unique singletons (0 API calls used)."
        )

    # 5. Batch Multi-Item Families for AI Extraction
    items_to_process = []
    for root, items in root_groups.items():
        if len(items) > 1:
            for item in items:
                if item["code"] not in parsed_results:
                    items_to_process.append(item)

    if items_to_process:
        print(
            f"[*] Sending {len(items_to_process)} ambiguous items to Gemini for attribute extraction..."
        )

        for i in range(0, len(items_to_process), BATCH_SIZE):
            batch = items_to_process[i : i + BATCH_SIZE]
            print(
                f"    Processing batch {i//BATCH_SIZE + 1}/{(len(items_to_process)//BATCH_SIZE)+1}..."
            )

            extracted_items = extract_attributes_with_gemini(client, batch)

            for ext in extracted_items:
                parsed_results[ext["item_code"]] = ext

            # Save checkpoint
            with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
                json.dump(parsed_results, f, ensure_ascii=False, indent=2)

            time.sleep(REQUEST_DELAY_SECONDS)

    # 6. Python Deterministic Clustering (The Judge)
    print("\n[*] AI Extraction complete. Executing deterministic clustering...")

    # Build Signatures
    for code, data in parsed_results.items():
        brand = str(data.get("canonical_brand", "")).strip().upper()
        strength = str(data.get("strength", "")).strip().upper()
        form = str(data.get("dosage_form", "")).strip().upper()
        pack = str(data.get("pack_size", "")).strip().upper()

        # Create the unbreakable signature
        signature = f"{brand} | {strength} | {form} | {pack}"
        data["signature"] = signature

        # Create proposed master name
        parts = [brand]
        if strength:
            parts.append(strength)
        if form and form not in pack:
            parts.append(form)
        if pack:
            parts.append(pack)
        data["master_name"] = " ".join(parts).strip()

    # Group by Signature
    signature_groups = {}
    for code, data in parsed_results.items():
        signature_groups.setdefault(data["signature"], []).append(code)

    # 7. Build Final CSV
    final_rows = []
    for r in all_records:
        code = r["code"]
        data = parsed_results.get(code)

        if not data:
            continue  # Should not happen

        sig = data["signature"]
        cluster_codes = sorted(signature_groups[sig], key=safe_sort_key)

        dup_count = len(cluster_codes) - 1
        dup_str = "; ".join(cluster_codes) if dup_count > 0 else ""
        comment = (
            f"Duplicate record clustered into '{data['master_name']}'"
            if dup_count > 0
            else "Unique SKU"
        )

        final_rows.append(
            {
                "code": code,
                "name": r["raw_name"],
                "number of duplicates": dup_count,
                "duplicate_codes": dup_str,
                "proposed master name": data["master_name"],
                "extra comments": comment,
                "sort_key": safe_sort_key(code),
            }
        )

    # Sort alphabetically by master name, then by code
    final_rows.sort(key=lambda x: (x["proposed master name"], x["sort_key"]))

    with open(OUTPUT_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "code",
                "name",
                "number of duplicates",
                "duplicate_codes",
                "proposed master name",
                "extra comments",
            ]
        )
        for row in final_rows:
            writer.writerow(
                [
                    row["code"],
                    row["name"],
                    row["number of duplicates"],
                    row["duplicate_codes"],
                    row["proposed master name"],
                    row["extra comments"],
                ]
            )

    print("\n" + "=" * 70)
    print(f"[✔] DEDUPLICATION FINISHED SUCCESSFULLY")
    print(f"[✔] Total Records Written:   {len(final_rows)}")
    dups = sum(1 for r in final_rows if r["number of duplicates"] > 0)
    print(f"[✔] Total Duplicate Records: {dups}")
    print(f"[✔] Total Unique SKUs:       {len(final_rows) - dups}")
    print(f"[✔] Output File Saved To:    '{os.path.abspath(OUTPUT_FILE)}'")
    print("=" * 70)


if __name__ == "__main__":
    run()
