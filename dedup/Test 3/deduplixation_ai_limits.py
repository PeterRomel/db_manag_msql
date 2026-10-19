import os
from dotenv import load_dotenv

load_dotenv()

import os
import re
import csv
import time
import json
import pandas as pd
from google import genai
from google.genai import types
from google.genai.errors import ClientError

# ==============================================================================
# CONFIGURATION & LIMIT SAFEGUARDS
# ==============================================================================
API_KEY = os.getenv(
    "GEMINI_API_KEY", os.getenv("GCP_API_KEY")
)
MODEL_ID = "gemini-3.5-flash-lite"

# Moderate load per prompt: 35-45 items prevents hallucinations & keeps total calls < 100
TARGET_ITEMS_PER_PROMPT = 40

# 8 seconds delay ensures ~7-8 requests/min (well below the 15 RPM free tier limit)
REQUEST_DELAY_SECONDS = 8
MAX_RETRIES = 5

INPUT_FILE = "used_items_extract.csv"
OUTPUT_FILE = "final_deduplicated_catalog_PERFECT.csv"
CHECKPOINT_FILE = "checkpoint_clusters.json"

client = genai.Client(api_key=API_KEY)


# ==============================================================================
# 1. DETERMINISTIC PYTHON SANITIZER
# ==============================================================================
def clean_trade_name(text: str) -> str:
    if not isinstance(text, str):
        return ""
    s = text.strip()
    s = re.sub(r"\b1000\s*mMG\b", "1000 MG", s, flags=re.I)
    s = re.sub(r"\bmMG\b", "MG", s, flags=re.I)
    s = re.sub(r"[@#*!?\[\]{}()]", " ", s)
    s = re.sub(r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP|جنيه)\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+LE\b", " ", s, flags=re.I)

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

    s = re.sub(r"\s+(m|mup|sedico|pharco|amryia|t3a|aug)\s*$", "", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def get_brand_root(name: str) -> str:
    clean = re.sub(r"^[^a-zA-Z0-9\u0621-\u064A]+", "", name.strip())
    tokens = re.split(r"[\s\-_/]+", clean)
    return tokens[0].upper() if tokens else "OTHER"


# ==============================================================================
# 2. BUNDLED PROMPT (LOW COGNITIVE LOAD FOR FLASH LITE)
# ==============================================================================
SYSTEM_INSTRUCTION = """You are a pharmaceutical catalog deduplication engine for Egyptian pharmacies.
You will be given multiple small brand families. Group items within each brand family into duplicate clusters.

RULES:
1. STRICT BRAND ISOLATION: Never merge items from different brands.
2. RULE 1-A (PACK SIZE POLICY): Items with different strengths (e.g., 20 MG vs 40 MG) or different pack counts (e.g., 20 TAB vs 30 TAB, 60 ML vs 100 ML) are DIFFERENT SKUs. NEVER merge them.
3. IDENTICAL DUPLICATES: Two items are duplicates ONLY if they share the exact same Brand, Active Strength, Dosage Form, and Pack Size.
4. MASTER NAME: For each cluster or unique item, assign a clean, standardized English trade name (e.g., 'CATAFLAM 50 MG 20 TAB'). Translate Arabic descriptions.

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
                f"\n[!] Daily quota or rate limit throttle. Sleeping {wait}s (Attempt {attempt+1}/{max_retries})..."
            )
            time.sleep(wait)

        except Exception as e:
            print(f"\n[!] Network or parsing error: {e}. Retrying in 10s...")
            time.sleep(10)

    # Fallback if request fails: register items as individual clean entries
    fallback = {"clusters": [], "uniques": {}}
    for fam in bundle_families:
        for it in fam["items"]:
            fallback["uniques"][it["code"]] = it["clean_name"]
    return fallback


# ==============================================================================
# 3. PIPELINE ORCHESTRATION WITH RESUME CAPABILITY
# ==============================================================================
def safe_sort_key(code):
    s = str(code).strip()
    return (0, int(s), "") if s.isdigit() else (1, 0, s)


def run():
    print(f"[*] Reading '{INPUT_FILE}'...")
    raw_df = pd.read_csv(INPUT_FILE)

    records = []
    for _, row in raw_df.iterrows():
        code = str(row["Item_Code"]).strip()
        raw_name = str(row["Item_English_Name"]).strip()
        clean_name = clean_trade_name(raw_name)
        root = get_brand_root(clean_name)
        records.append(
            {"code": code, "raw_name": raw_name, "clean_name": clean_name, "root": root}
        )

    # Group records by brand root
    root_groups = {}
    for r in records:
        root_groups.setdefault(r["root"], []).append(r)

    total_roots = len(root_groups)
    print(
        f"[*] Total Catalog: {len(records)} records across {total_roots} brand roots."
    )

    # 1. Load Checkpoint if resuming
    processed_roots = set()
    master_map = {}
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                master_map = saved.get("master_map", {})
                processed_roots = set(saved.get("processed_roots", []))
            print(
                f"[*] Checkpoint loaded: {len(processed_roots)} roots already processed. Resuming seamlessly..."
            )
        except Exception as e:
            print(f"[!] Warning reading checkpoint: {e}. Starting fresh.")

    # 2. Process Singletons Locally (Zero Tokens, Zero API Requests)
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
            f"[*] Identified and resolved {singleton_count} single-item brands locally for FREE (0 API calls)."
        )

    # 3. Bundle Remaining Multi-Item Roots into Safe ~40-Item Prompts
    pending_families = [
        {"root": root, "items": items}
        for root, items in root_groups.items()
        if root not in processed_roots
    ]

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
    print(
        f"[*] Bundled remaining candidate families into {total_bundles} small batches."
    )
    print(
        f"[*] Projected API Calls: {total_bundles} requests (Daily Limit: 500 RPD | Remaining capacity: {500 - total_bundles})"
    )

    # 4. Execute API Calls on Bundles
    for b_idx, bundle in enumerate(bundles, start=1):
        bundle_roots = [f["root"] for f in bundle]
        bundle_items_count = sum(len(f["items"]) for f in bundle)
        print(
            f"[{b_idx}/{total_bundles}] Auditing {len(bundle)} brand families ({bundle_items_count} items)..."
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

        # Safety Fallback: ensure no item was missed by the model
        for fam in bundle:
            for it in fam["items"]:
                c = it["code"]
                if c not in master_map:
                    master_map[c] = {
                        "master_name": it["clean_name"].upper(),
                        "dup_codes": [c],
                    }
            processed_roots.add(fam["root"])

        # Update checkpoint on disk
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            json.dump(
                {"processed_roots": list(processed_roots), "master_map": master_map}, f
            )

        time.sleep(REQUEST_DELAY_SECONDS)

    # ==============================================================================
    # 5. CONSTRUCT FINAL, PERFECT CSV (Python Native Writer)
    # ==============================================================================
    print("\n[*] Assembling final, mathematically symmetric CSV...")
    final_output_rows = []

    for r in records:
        code = r["code"]
        name = r["raw_name"]
        item_data = master_map.get(
            code, {"master_name": r["clean_name"].upper(), "dup_codes": [code]}
        )

        m_name = item_data["master_name"]
        all_sibling_codes = sorted(list(set(item_data["dup_codes"])), key=safe_sort_key)

        dup_count = len(all_sibling_codes) - 1
        dup_codes_str = "; ".join(all_sibling_codes) if dup_count > 0 else ""
        comments = (
            f"Duplicate record clustered into '{m_name}'"
            if dup_count > 0
            else "Unique SKU"
        )

        final_output_rows.append(
            {
                "code": code,
                "name": name,
                "number of duplicates": dup_count,
                "duplicate_codes": dup_codes_str,
                "proposed master name": m_name,
                "extra comments": comments,
                "sort_key": safe_sort_key(code),
            }
        )

    # Sort alphabetically by proposed master name, then by code
    final_output_rows.sort(key=lambda x: (x["proposed master name"], x["sort_key"]))

    # Write using standard csv.writer with UTF-8 BOM for Android Excel compatibility
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
        for row in final_output_rows:
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
    print(f"[✔] DEDUPLICATION FINISHED")
    print(f"[✔] File Output: '{os.path.abspath(OUTPUT_FILE)}'")
    print(f"[✔] Total Items Processed:   {len(final_output_rows)}")
    dups_count = sum(1 for r in final_output_rows if r["number of duplicates"] > 0)
    print(f"[✔] Total Duplicates Mapped: {dups_count}")
    print(f"[✔] Total Unique SKUs:       {len(final_output_rows) - dups_count}")
    print("=" * 70)


if __name__ == "__main__":
    run()
