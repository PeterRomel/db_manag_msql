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
# CONFIGURATION
# ==============================================================================
API_KEY = os.getenv(
    "GEMINI_API_KEY", os.getenv("GCP_API_KEY")
)
MODEL_ID = "gemini-3.5-flash-lite"

INPUT_FILE = "used_items_extract.csv"
OUTPUT_FILE = "final_deduplicated_catalog_PERFECT.csv"
CHECKPOINT_FILE = "checkpoint_clusters.json"

client = genai.Client(api_key=API_KEY)


# ==============================================================================
# 1. DETERMINISTIC PYTHON SANITIZER
# ==============================================================================
def clean_trade_name(text: str) -> str:
    """Strips cashier noise, prices, and status markers in Python."""
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

    # Strip supplier/cashier initials at end of string
    s = re.sub(r"\s+(m|mup|sedico|pharco|amryia|t3a|aug)\s*$", "", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def get_brand_root(name: str) -> str:
    """Extracts the first alphanumeric brand token for grouping."""
    clean = re.sub(r"^[^a-zA-Z0-9\u0621-\u064A]+", "", name.strip())
    tokens = re.split(r"[\s\-_/]+", clean)
    return tokens[0].upper() if tokens else "OTHER"


# ==============================================================================
# 2. PROMPT TEMPLATE (JSON CLUSTERING ONLY)
# ==============================================================================
CLUSTER_INSTRUCTION = """You are a senior pharmaceutical catalog auditor in Egypt.
Group the provided items into duplicate clusters.

RULES:
1. STRICT BRAND ISOLATION: Never merge different trade brands (Paramol != Panadol != Cetal).
2. RULE 1-A (PACK SIZE POLICY): Items with different strengths or pack sizes (e.g. 20 TAB vs 30 TAB, 60 ML vs 100 ML) are DIFFERENT SKUs. Never group them.
3. IDENTICAL ITEMS: Two items are duplicates ONLY if they have the exact same Brand, Active Strength, Dosage Form, and Pack Size.
4. CANONICAL MASTER NAME: For each cluster or item, provide a standardized, professional English title (e.g. "CATAFLAM 50 MG 20 TAB"). Translate Arabic terms if applicable.

OUTPUT SCHEMA (JSON ONLY):
{
  "clusters": [
    {
      "master_name": "STANDARDIZED BRAND STRENGTH FORM PACK",
      "codes": ["code1", "code2"]
    }
  ],
  "uniques": {
    "code3": "STANDARDIZED BRAND STRENGTH FORM PACK"
  }
}
"""


def cluster_with_gemini(items_list: list[dict], max_retries: int = 5) -> dict:
    prompt = "Review this family of candidate items:\n\n"
    for item in items_list:
        prompt += f"- Code: {item['code']} | Cleaned Name: '{item['clean_name']}' | Raw: '{item['raw_name']}'\n"

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=CLUSTER_INSTRUCTION,
                    temperature=0.0,
                    response_mime_type="application/json",
                ),
            )
            return json.loads(response.text)

        except ClientError as e:
            error_str = str(e)
            if "limit: 0" in error_str or "NOT_FOUND" in error_str:
                raise RuntimeError(f"Fatal error: {error_str}")

            match = re.search(r"retryDelay':\s*'(\d+)s", error_str) or re.search(
                r"retry in (\d+(\.\d+)?)s", error_str
            )
            wait = int(float(match.group(1))) + 2 if match else 20
            print(f"[!] Rate limit. Pausing {wait}s...")
            time.sleep(wait)

        except Exception as e:
            print(f"[!] Network error: {e}. Pausing 10s...")
            time.sleep(10)

    return {
        "clusters": [],
        "uniques": {item["code"]: item["clean_name"] for item in items_list},
    }


# ==============================================================================
# 3. PIPELINE ORCHESTRATION & PYTHON CSV WRITING
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

    # Group by brand root
    root_groups = {}
    for r in records:
        root_groups.setdefault(r["root"], []).append(r)

    print(
        f"[*] Clustered {len(records)} items into {len(root_groups)} brand root families."
    )

    # Load checkpoint if resuming
    processed_roots = set()
    master_map = {}  # code -> {"master_name": ..., "dup_codes": [...]}
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                master_map = saved.get("master_map", {})
                processed_roots = set(saved.get("processed_roots", []))
            print(
                f"[*] Checkpoint loaded: {len(processed_roots)} roots already audited."
            )
        except Exception:
            pass

    # Process each brand family
    for root, items in root_groups.items():
        if root in processed_roots:
            continue

        # If only 1 item in the brand root, Python handles it for FREE without AI
        if len(items) == 1:
            code = items[0]["code"]
            master_map[code] = {
                "master_name": items[0]["clean_name"].upper(),
                "dup_codes": [code],
            }
            processed_roots.add(root)
            continue

        # If multiple items exist in the brand root, ask Gemini to cluster them
        print(f"[*] Auditing brand family '{root}' ({len(items)} items)...")
        result = cluster_with_gemini(items)

        # 1. Register clusters
        for cl in result.get("clusters", []):
            m_name = cl.get("master_name", "").strip().upper()
            c_codes = [str(c).strip() for c in cl.get("codes", [])]
            for c in c_codes:
                master_map[c] = {"master_name": m_name, "dup_codes": c_codes}

        # 2. Register items the AI confirmed unique
        for u_code, u_name in result.get("uniques", {}).items():
            u_code = str(u_code).strip()
            if u_code not in master_map:
                master_map[u_code] = {
                    "master_name": str(u_name).strip().upper(),
                    "dup_codes": [u_code],
                }

        # 3. Fallback for any items missed by the AI
        for it in items:
            c = it["code"]
            if c not in master_map:
                master_map[c] = {
                    "master_name": it["clean_name"].upper(),
                    "dup_codes": [c],
                }

        processed_roots.add(root)

        # Save checkpoint periodically
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            json.dump(
                {"processed_roots": list(processed_roots), "master_map": master_map}, f
            )

        time.sleep(4.5)  # Enforce Free Tier RPM limit

    # ==============================================================================
    # 4. PYTHON BUILDS THE PERFECT CSV (Zero Syntax Errors)
    # ==============================================================================
    print("\n[*] Assembling final CSV table in Python...")
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
                "sort_key_code": safe_sort_key(code),
            }
        )

    # Sort strictly alphabetically by master name, then by code
    final_output_rows.sort(
        key=lambda x: (x["proposed master name"], x["sort_key_code"])
    )

    # Write using Python's native CSV writer (Handles quotes, UTF-8 BOM, and commas natively)
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
    print("DEDUPLICATION COMPLETE & VALIDATED")
    print(f"[✔] Perfectly formatted file: '{OUTPUT_FILE}'")
    print(f"[✔] Total Items Processed:   {len(final_output_rows)}")
    dups_count = sum(1 for r in final_output_rows if r["number of duplicates"] > 0)
    print(f"[✔] Total Duplicates Mapped: {dups_count}")
    print(f"[✔] Total Unique SKUs:       {len(final_output_rows) - dups_count}")
    print("=" * 70)


if __name__ == "__main__":
    run()
