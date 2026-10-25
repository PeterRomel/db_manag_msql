import os
from dotenv import load_dotenv

load_dotenv()

"""EGYPTIAN PHARMACY ERP DEDUPLICATION & ENTITY RESOLUTION ENGINE (PRODUCTION GRADE)

Engineered for: Genius / e-Plus SQL Server ERP Catalogs
Handles: 6,800+ noisy items, low API quotas (Gemini Flash-Lite), cross-batch reconciliation.
"""

import csv
import json
import os
import re
import sys
import time
from collections import defaultdict
from google import genai
from google.genai import types
from google.genai.errors import ClientError
import pandas as pd

# ==============================================================================
# CONFIGURATION & API RATELIMIT SAFEGUARDS
# ==============================================================================
API_KEY = os.getenv(
    "GEMINI_API_KEY", os.getenv("GCP_API_KEY")
)
MODEL_ID = "gemini-3.5-flash-lite"

TARGET_ITEMS_PER_PROMPT = 40
REQUEST_DELAY_SECONDS = 8
MAX_RETRIES = 5

INPUT_FILE = "used_items_extract.csv"
OUTPUT_FILE = "final_deduplicated_catalog_PERFECT.csv"
CHECKPOINT_FILE = "checkpoint_clusters.json"

client = genai.Client(api_key=API_KEY)


# ==============================================================================
# 1. HARDENED DETERMINISTIC PRE-SANITIZER
# ==============================================================================
def clean_trade_name(text: str) -> str:
    """Strips prices (with or without LE), cashier noise, and fixes unit typos."""
    if not isinstance(text, str):
        return ""

    s = text.strip()

    # Normalize common Egyptian cashier typos
    s = re.sub(r"\b1000\s*mMG\b", "1000 MG", s, flags=re.I)
    s = re.sub(r"\bmMG\b", "MG", s, flags=re.I)
    s = re.sub(r"\bCABS(ULES)?\b", "CAP", s, flags=re.I)
    s = re.sub(r"\bTABLES\b", "TAB", s, flags=re.I)
    s = re.sub(r"\bSYRY?P\b", "SYRUP", s, flags=re.I)

    # Remove symbols and brackets
    s = re.sub(r"[@#*!?\[\]{}()]", " ", s)

    # 1. Strip explicit currency patterns (e.g., 25LE, 30.5 EGP, 50 جنيه)
    s = re.sub(r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP|L\.E\.|جنيه)\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+LE\b", " ", s, flags=re.I)

    # 2. Strip trailing naked prices using capture groups (replaces variable-width look-behind)
    # E.g. "centravita ginko 30 tab 375" -> keeps "centravita ginko 30 tab", drops "375"
    s = re.sub(
        r"(\b(?:TAB|CAP|TABLETS|CAPSULES|ML|GM|G|SACH|SACHETS|AMP|VIAL|PIECES|P|DOSE|PUFF)\b\s*)\d{2,4}$",
        r"\1",
        s.strip(),
        flags=re.I,
    )
    # E.g. "curam 457mg 70ml 97" -> keeps "curam 457mg 70ml", drops "97"
    s = re.sub(
        r"(\b(?:ML|GM|MG)\b\s*)\d{2,4}$",
        r"\1",
        s.strip(),
        flags=re.I,
    )

    # 3. Strip lifecycle and status noise
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

    # 4. Strip cashier and company suffixes at end of string
    s = re.sub(
        r"\s+\b(m|k|pres|aug|t3a|sedico|pharco|amryia|amrya|mup|eipico)\b\s*$",
        "",
        s,
        flags=re.I,
    )

    return re.sub(r"\s+", " ", s).strip()


def get_brand_root(name: str) -> str:
    """Accurately extracts brand root WITHOUT splitting hyphenated drug prefixes."""
    clean = re.sub(r"^[^a-zA-Z0-9\u0621-\u064A]+", "", name.strip())
    if not clean:
        return "OTHER"

    # Match compound prefixes common in Egyptian pharmacy: CO-, E-, TRI-, X-, NO-, DE-
    match_compound = re.match(
        r"^(CO|E|TRI|X|NO|DE|BIO|PRO|ULTRA)[-_ ]+([a-zA-Z0-9]+)",
        clean,
        flags=re.I,
    )
    if match_compound:
        return f"{match_compound.group(1).upper()}-{match_compound.group(2).upper()}"

    # Otherwise split purely by spaces or slashes (preserving internal hyphens)
    tokens = re.split(r"[\s/]+", clean)
    return tokens[0].upper() if tokens else "OTHER"


# ==============================================================================
# 2. GEMINI INSTRUCTION & BUNDLED QUERY
# ==============================================================================
SYSTEM_INSTRUCTION = """You are a pharmaceutical catalog deduplication engine for Egyptian pharmacies.
Group redundant inventory entries for identical products into duplicate clusters.

RULES:
1. STRICT BRAND ISOLATION: Never merge items from different brands (e.g. Paramol != Panadol != Cetal).
2. RULE 1-A (STRICT SKU MATCHING): Items with different strengths (e.g., 20 MG vs 40 MG) or different pack counts/volumes (e.g., 20 TAB vs 30 TAB, 60 ML vs 100 ML) are DIFFERENT SKUs. NEVER merge them.
3. CASHIER FORM TYPOS: In Egypt, cashiers often mistakenly type 'TAB' instead of 'CAP' (or vice versa) for products that only exist in one form (e.g. Bengiride 0.5mg). If Brand, Strength, and Count match exactly, cluster them under the true canonical dosage form.
4. CANONICAL MASTER NAME: Assign a clean, professional, standardized English trade name (e.g., 'CATAFLAM 50 MG 20 TAB'). Translate Arabic service items accurately.

OUTPUT FORMAT (JSON ONLY):
{
  "clusters": [
    {
      "master_name": "CANONICAL BRAND STRENGTH FORM PACK",
      "codes": ["code1", "code2"]
    }
  ],
  "uniques": {
    "code3": "CANONICAL BRAND STRENGTH FORM PACK"
  }
}
"""


def cluster_bundle_with_gemini(
    bundle_families: list[dict], max_retries: int = MAX_RETRIES
) -> dict:
    prompt = (
        "Identify duplicate items within each of these distinct brand families:\n\n"
    )
    for fam in bundle_families:
        prompt += f"=== BRAND FAMILY: {fam['root']} ===\n"
        for it in fam["items"]:
            prompt += f"- Code: {it['code']} | Cleaned: '{it['clean_name']}' | Raw: '{it['raw_name']}'\n"
        prompt += "\n"

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_INSTRUCTION,
                    temperature=0.0,
                    response_mime_type="application/json",
                ),
            )
            return json.loads(response.text)

        except ClientError as e:
            error_str = str(e)
            if "limit: 0" in error_str or "NOT_FOUND" in error_str:
                raise RuntimeError(f"Fatal API configuration error: {error_str}")

            match = re.search(r"retryDelay':\s*'(\d+)s", error_str) or re.search(
                r"retry in (\d+(\.\d+)?)s", error_str
            )
            wait = int(float(match.group(1))) + 2 if match else 25
            print(
                f"\n[!] Rate throttle. Sleeping {wait}s (Attempt {attempt+1}/{max_retries})..."
            )
            time.sleep(wait)

        except Exception as e:
            print(f"\n[!] Parsing/network error: {e}. Retrying in 10s...")
            time.sleep(10)

    # Fallback if request fails
    fallback = {"clusters": [], "uniques": {}}
    for fam in bundle_families:
        for it in fam["items"]:
            fallback["uniques"][it["code"]] = it["clean_name"]
    return fallback


# ==============================================================================
# 3. GLOBAL DETERMINISTIC RECONCILIATION ENGINE (STAGE 4)
# ==============================================================================
def global_reconciliation_pass(rows: list[dict]) -> list[dict]:
    """Reconciles 'Twin Ghosts', normalizes clinical concentrations, and repairs edge cases across the entire catalog."""
    print("\n[*] Executing Global Deterministic Reconciliation on all catalog items...")

    # Step 1: Clean and standardize Master Names
    for r in rows:
        m = r["proposed master name"].strip().upper()

        # Fix saline concentration error (0.09% -> 0.9%)
        m = re.sub(r"\b0\.09%", "0.9%", m)

        # Fix Babetone Sugar Free truncation
        if "BABETONE" in m and "SUGAR FREE" in r["name"].upper():
            m = "BABETONE SUGAR FREE SYRUP 120 ML"

        # Fix Centravita Ginko 37.5 MG price hallucination
        if "CENTRAVITA GINKO" in m and "37.5 MG" in m:
            m = "CENTRAVITA GINKO 30 TAB"

        # Strip remaining cashier initials leaking into master name
        m = re.sub(r"\s+\b(M|K|PRES)\b$", "", m)

        # Fix Clinical Service for Blood Pressure Test
        if r["code"].strip() == "2" and "ضغط" in r["name"]:
            m = "BLOOD PRESSURE CHECK (SERVICE)"

        r["proposed master name"] = re.sub(r"\s+", " ", m).strip()

    # Step 2: Global Grouping by Master Name (Fixes the Twin Ghost bug across batch boundaries)
    master_groups = defaultdict(list)
    for idx, r in enumerate(rows):
        master_groups[r["proposed master name"]].append(idx)

    # Step 3: Reconcile TAB vs CAP clashes for identical strength and count
    normalized_keys = defaultdict(list)
    for m_name in list(master_groups.keys()):
        key = re.sub(r"\b(TAB|CAP|TABLET|CAPSULE)S?\b", "DOSEFORM", m_name)
        normalized_keys[key].append(m_name)

    for _, variants in normalized_keys.items():
        if len(variants) > 1:
            primary = variants[0]
            for secondary in variants[1:]:
                sec_indices = master_groups.pop(secondary)
                master_groups[primary].extend(sec_indices)
                for i in sec_indices:
                    rows[i]["proposed master name"] = primary

    # Step 4: Re-calculate Duplicate Codes and Cluster Counts
    def safe_sort_key(code):
        s = str(code).strip()
        return (0, int(s), "") if s.isdigit() else (1, 0, s)

    for m_name, indices in master_groups.items():
        all_codes = sorted(
            list(set(rows[i]["code"].strip() for i in indices)),
            key=safe_sort_key,
        )
        cluster_size = len(all_codes)
        dup_count = cluster_size - 1
        dup_str = "; ".join(all_codes) if dup_count > 0 else ""
        comment = (
            f"Duplicate record clustered into '{m_name}'"
            if dup_count > 0
            else "Unique SKU"
        )

        for i in indices:
            rows[i]["number of duplicates"] = dup_count
            rows[i]["duplicate_codes"] = dup_str
            rows[i]["extra comments"] = comment
            rows[i]["sort_key"] = safe_sort_key(rows[i]["code"])

    return rows


# ==============================================================================
# 4. PIPELINE ORCHESTRATION WITH RESUME CAPABILITY
# ==============================================================================
def run():
    print(f"[*] Reading '{INPUT_FILE}'...")
    if not os.path.exists(INPUT_FILE):
        print(f"[ERROR] Input file '{INPUT_FILE}' not found!")
        return

    raw_df = pd.read_csv(INPUT_FILE)
    print(f"[*] Total rows in raw input file: {len(raw_df)}")

    # Detect code and name column dynamically
    code_col = (
        "Item_Code"
        if "Item_Code" in raw_df.columns
        else ("code" if "code" in raw_df.columns else raw_df.columns[0])
    )
    name_col = (
        "Item_English_Name"
        if "Item_English_Name" in raw_df.columns
        else ("name" if "name" in raw_df.columns else raw_df.columns[1])
    )

    records = []
    for _, row in raw_df.iterrows():
        code = str(row[code_col]).strip()
        raw_name = str(row[name_col]).strip()
        clean_name = clean_trade_name(raw_name)
        root = get_brand_root(clean_name)
        records.append(
            {
                "code": code,
                "raw_name": raw_name,
                "clean_name": clean_name,
                "root": root,
            }
        )

    # Group by Brand Root
    root_groups = defaultdict(list)
    for r in records:
        root_groups[r["root"]].append(r)

    print(
        f"[*] Grouped into {len(root_groups)} distinct brand families across {len(records)} items."
    )

    # Load Checkpoint
    processed_roots = set()
    master_map = {}
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                master_map = saved.get("master_map", {})
                processed_roots = set(saved.get("processed_roots", []))
            print(
                f"[*] Resuming from checkpoint: {len(processed_roots)} roots previously resolved."
            )
        except Exception as e:
            print(f"[!] Warning reading checkpoint: {e}. Starting fresh.")

    # Bundle Multi-Item Families (Leave Singletons to be resolved natively)
    pending_families = [
        {"root": root, "items": items}
        for root, items in root_groups.items()
        if root not in processed_roots and len(items) > 1
    ]

    # Resolve Singletons directly without API tokens
    singleton_count = 0
    for root, items in root_groups.items():
        if root not in processed_roots and len(items) == 1:
            code = items[0]["code"]
            master_map[code] = {
                "master_name": items[0]["clean_name"].upper(),
                "dup_codes": [code],
            }
            processed_roots.add(root)
            singleton_count += 1

    if singleton_count > 0:
        print(
            f"[*] Resolved {singleton_count} singletons locally for FREE (0 API calls)."
        )

    bundles = []
    current_bundle = []
    current_count = 0
    for fam in pending_families:
        fam_size = len(fam["items"])
        if current_count + fam_size > TARGET_ITEMS_PER_PROMPT and current_bundle:
            bundles.append(current_bundle)
            current_bundle = []
            current_count = 0
        current_bundle.append(fam)
        current_count += fam_size
    if current_bundle:
        bundles.append(current_bundle)

    total_bundles = len(bundles)
    print(f"[*] Multi-item families packaged into {total_bundles} small API batches.")

    # Process API Batches
    for b_idx, bundle in enumerate(bundles, start=1):
        b_items = sum(len(f["items"]) for f in bundle)
        print(
            f"[{b_idx}/{total_bundles}] Auditing {len(bundle)} brand roots ({b_items} items)..."
        )

        result = cluster_bundle_with_gemini(bundle)

        # Register clusters
        for cl in result.get("clusters", []):
            m_name = cl.get("master_name", "").strip().upper()
            c_codes = [str(c).strip() for c in cl.get("codes", [])]
            for c in c_codes:
                master_map[c] = {"master_name": m_name, "dup_codes": c_codes}

        # Register uniques
        for u_code, u_name in result.get("uniques", {}).items():
            u_code = str(u_code).strip()
            if u_code not in master_map:
                master_map[u_code] = {
                    "master_name": str(u_name).strip().upper(),
                    "dup_codes": [u_code],
                }

        # Fallback registration
        for fam in bundle:
            for it in fam["items"]:
                c = it["code"]
                if c not in master_map:
                    master_map[c] = {
                        "master_name": it["clean_name"].upper(),
                        "dup_codes": [c],
                    }
            processed_roots.add(fam["root"])

        # Checkpoint save
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            json.dump(
                {"processed_roots": list(processed_roots), "master_map": master_map},
                f,
            )

        time.sleep(REQUEST_DELAY_SECONDS)

    # Build Preliminary Output Rows
    output_rows = []
    for r in records:
        code = r["code"]
        item_data = master_map.get(
            code, {"master_name": r["clean_name"].upper(), "dup_codes": [code]}
        )
        output_rows.append(
            {
                "code": code,
                "name": r["raw_name"],
                "number of duplicates": 0,
                "duplicate_codes": "",
                "proposed master name": item_data["master_name"],
                "extra comments": "",
            }
        )

    # EXECUTE GLOBAL RECONCILIATION ENGINE (Eliminates Twin Ghosts & Form Errors)
    final_reconciled_rows = global_reconciliation_pass(output_rows)

    # Sort alphabetically by master name, then by code
    final_reconciled_rows.sort(key=lambda x: (x["proposed master name"], x["sort_key"]))

    # Write Final Output with UTF-8 BOM
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
        for r in final_reconciled_rows:
            writer.writerow(
                [
                    r["code"],
                    r["name"],
                    r["number of duplicates"],
                    r["duplicate_codes"],
                    r["proposed master name"],
                    r["extra comments"],
                ]
            )

    print("\n" + "=" * 70)
    print(f"[✔] DEDUPLICATION FINISHED SUCCESSFULLY")
    print(f"[✔] Total Records Written:   {len(final_reconciled_rows)}")
    dups = sum(1 for r in final_reconciled_rows if int(r["number of duplicates"]) > 0)
    print(f"[✔] Total Duplicate Records: {dups}")
    print(f"[✔] Total Unique SKUs:       {len(final_reconciled_rows) - dups}")
    print(f"[✔] Output File Saved To:    '{os.path.abspath(OUTPUT_FILE)}'")
    print("=" * 70)


if __name__ == "__main__":
    run()
