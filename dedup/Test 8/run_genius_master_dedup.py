#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
========================================================================================
GENIUS ERP CATALOG DEDUPLICATION & CLINICAL RESOLUTION ENGINE
Target Platform: Egyptian Pharmacy ERP (Genius / e-Plus on Microsoft SQL Server)
Model: Gemini 3.5 Flash-Lite (Google AI Studio Free Tier Optimized)
Output: RFC-4180 CSV with UTF-8 BOM + Production T-SQL Migration Script
========================================================================================
"""

import os
import sys
import re
import csv
import glob
import time
import json
from collections import defaultdict
from typing import Any, Dict, List, Tuple
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# ==============================================================================
# 1. API & EXECUTION CONFIGURATION
# ==============================================================================
API_KEY = os.getenv("GEMINI_API_KEY", os.getenv("GCP_API_KEY"))
if not API_KEY:
    print("[ERROR] GEMINI_API_KEY environment variable not detected.")
    print("Please set it in your environment or .env file.")
    sys.exit(1)

from google import genai
from google.genai import types
from google.genai.errors import ClientError

client = genai.Client(api_key=API_KEY)
MODEL_ID = "gemini-3.5-flash-lite"

# Safe pacing: 6.5 seconds = ~9.2 RPM (well below 15 RPM limit)
TARGET_ITEMS_PER_PROMPT = 40
REQUEST_DELAY_SECONDS = 6.5
MAX_RETRIES = 5

CHECKPOINT_FILE = "genius_dedup_checkpoint.json"
OUTPUT_CSV_FILE = "final_deduplicated_catalog_PERFECT.csv"
OUTPUT_SQL_FILE = "execute_genius_deduplication.sql"

# Common Egyptian pharmacy protected prefixes (never split from brand stem)
PROTECTED_PREFIXES = {
    "CO",
    "TRI",
    "E",
    "X",
    "NO",
    "DE",
    "BIO",
    "PRO",
    "ULTRA",
    "MICRO",
    "PAN",
    "MAX",
    "NEO",
    "NEW",
    "MULTI",
    "UNI",
}

# Clinical consumables & service lines that must remain unmerged singletons
SERVICE_OR_SUPPLY_CODES = {
    "2",
    "IN",
    "G5",
    "G10",
    "S3",
    "S5",
    "S10",
    "WI",
    "MA",
    "pl",
    "ss5",
    "cat",
}


# ==============================================================================
# 2. DETERMINISTIC PRE-SANITIZER
# ==============================================================================
def clean_trade_name(raw_name: str) -> str:
    """Strips cashier noise, embedded prices, and unit typos deterministically."""
    if not isinstance(raw_name, str) or not raw_name.strip():
        return ""

    s = raw_name.strip()

    # Step 1: Normalize common cashier typos
    s = re.sub(r"\b1000\s*mMG\b", "1000 MG", s, flags=re.I)
    s = re.sub(r"\bmMG\b", "MG", s, flags=re.I)
    s = re.sub(r"\bCABS(ULES)?\b", "CAP", s, flags=re.I)
    s = re.sub(r"\bTABLES\b", "TAB", s, flags=re.I)
    s = re.sub(r"\bSYRY?P\b", "SYRUP", s, flags=re.I)
    s = re.sub(r"[@#*!?\[\]{}()~^%_]", " ", s)

    # Step 2: Strip currency prices (e.g., 25LE, 30.5 EGP, 50 جنيه)
    s = re.sub(r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP|L\.E\.|جنيه)\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+LE\b", " ", s, flags=re.I)

    # Step 3: Strip trailing naked prices (prevents AI confusing price with strength)
    # E.g., 'centravita ginko 30 tab 375' -> 'centravita ginko 30 tab'
    s = re.sub(
        r"(\b(?:TAB|CAP|TABLETS?|CAPSULES?|ML|GM|G|SACH|SACHETS?|AMP|VIALS?|PIECES|P|DOSE|PUFF)\b\s*)\d{2,4}$",
        r"\1",
        s.strip(),
        flags=re.I,
    )
    # E.g., 'curam 457mg 70ml 97' -> 'curam 457mg 70ml'
    s = re.sub(r"(\b(?:ML|GM|MG)\b\s*)\d{2,4}$", r"\1", s.strip(), flags=re.I)

    # Step 4: Strip lifecycle status tags
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
        r"\bNOW\s*PRICE\b",
    ]
    for p in noise_patterns:
        s = re.sub(p, " ", s, flags=re.I)

    # Step 5: Strip trailing manufacturer/cashier initials
    s = re.sub(
        r"\s+\b(m|k|pres|aug|t3a|sedico|pharco|amryia|amrya|mup|eipico|sed|alex)\b\s*$",
        "",
        s,
        flags=re.I,
    )

    return re.sub(r"\s+", " ", s).strip()


def get_brand_root(name: str) -> str:
    """Extracts the brand family root preserving compound drug prefixes."""
    clean = re.sub(r"^[^a-zA-Z0-9\u0621-\u064A]+", "", name.strip())
    if not clean:
        return "OTHER"

    # Match compound prefixes common in Egyptian trade names
    match_compound = re.match(r"^([a-zA-Z0-9]+)[-_ ]+([a-zA-Z0-9]+)", clean, flags=re.I)
    if match_compound:
        prefix = match_compound.group(1).upper()
        if prefix in PROTECTED_PREFIXES:
            return f"{prefix}-{match_compound.group(2).upper()}"

    tokens = re.split(r"[\s/]+", clean)
    return tokens[0].upper() if tokens else "OTHER"


def safe_sort_key(code: Any) -> Tuple[int, int, str]:
    """Safe key avoiding TypeError between integers and string codes."""
    s = str(code).strip()
    return (0, int(s), "") if s.isdigit() else (1, 0, s.upper())


# ==============================================================================
# 3. AI STRUCTURED REASONING PROMPT
# ==============================================================================
SYSTEM_INSTRUCTION = """You are an expert Egyptian Pharmacy ERP catalog entity-resolution auditor.
Your job is to group duplicate items representing the EXACT same product into clusters.

CLINICAL & ERP DEDUPLICATION RULES:
1. STRICT BRAND ISOLATION: Never merge different trade brands (e.g., Paramol != Panadol != Cetal != Abimol).
2. RULE 1-A (STRICT SKU MATCHING): Items with different active strengths (e.g., 20 MG vs 40 MG) OR different pack counts / volumes (e.g., 20 TAB vs 30 TAB, 60 ML vs 100 ML, 5 AMP vs 6 AMP) are DIFFERENT commercial SKUs. NEVER merge them.
3. COMBINATION INTEGRITY: Monotherapy drugs must NEVER merge with combination counterparts (e.g., CONCOR 5 MG != CONCOR PLUS 5 MG != CONCOR AMLO 5/5 MG).
4. CASHIER FORM TYPOS: If brand, strength, and count match exactly but one record says 'TAB' and another says 'CAP' for a product that only exists in one form in Egypt (e.g., Bengiride 0.5mg), cluster them under the true canonical form.
5. STANDARDIZED MASTER NAME: For each cluster or unique item, assign a clean, professional canonical English trade name:
   Format: [BRAND] [STRENGTH] [DOSAGE FORM] [PACK SIZE]
   Example: 'CATAFLAM 50 MG 20 TAB', 'AUGMENTIN 1 GM 14 TAB', 'BRUFEN 400 MG 30 TAB'.
   Translate Arabic service/supply descriptions to clean English.

OUTPUT JSON FORMAT (ONLY VALID JSON, NO MARKDOWN):
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
    bundle_families: List[Dict], max_retries: int = MAX_RETRIES
) -> Dict:
    prompt = "Audit and identify duplicate items within these brand families:\n\n"
    for fam in bundle_families:
        prompt += f"=== BRAND ROOT: {fam['root']} ===\n"
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
            wait = int(float(match.group(1))) + 2 if match else 20
            print(
                f"\n[!] Rate throttle. Pausing {wait}s (Attempt {attempt+1}/{max_retries})..."
            )
            time.sleep(wait)

        except Exception as e:
            print(f"\n[!] Network or parsing error: {e}. Retrying in 8s...")
            time.sleep(8)

    # Fallback: Treat all items in batch as unique clean entries
    fallback = {"clusters": [], "uniques": {}}
    for fam in bundle_families:
        for it in fam["items"]:
            fallback["uniques"][it["code"]] = it["clean_name"].upper()
    return fallback


# ==============================================================================
# 4. CLINICAL SAFETY GUARDRAILS (POST-PROCESSING)
# ==============================================================================
def extract_strength_signature(text: str) -> str:
    m = re.search(
        r"\b(\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)*)\s*(MG|MCG|GM|G|IU|%|ML)\b",
        text,
        flags=re.I,
    )
    return m.group(0).upper().replace(" ", "") if m else ""


def extract_pack_signature(text: str) -> str:
    m = re.search(
        r"\b(\d+)\s*(TAB|CAP|AMP|VIAL|SACH|ML|GM|DOSE|SUPP|PATCH)S?\b", text, flags=re.I
    )
    return m.group(0).upper().replace(" ", "") if m else ""


def extract_form_category(text: str) -> str:
    s = text.upper()
    if any(t in s for t in ["CREAM", "OINT", "GEL", "LOTION"]):
        return "TOPICAL"
    if any(t in s for t in ["TAB", "CAP", "TABLET", "CAPSULE", "LOZ"]):
        return "ORAL_SOLID"
    if any(t in s for t in ["SYRUP", "SUSP", "DROPS", "ORAL SOL"]):
        return "ORAL_LIQUID"
    if any(t in s for t in ["AMP", "VIAL", "INJECTION"]):
        return "INJECTABLE"
    return "OTHER"


def apply_clinical_guardrails(
    master_map: Dict[str, Dict], records: List[Dict]
) -> Dict[str, Dict]:
    """Inspects AI clusters and splits any illegal merges that violate Rule 1-A."""
    print("[*] Running Clinical Safety Guardrails on duplicate clusters...")

    # Invert mapping: master_name -> list of item records
    clusters = defaultdict(list)
    for r in records:
        code = r["code"]
        m_data = master_map.get(
            code, {"master_name": r["clean_name"].upper(), "dup_codes": [code]}
        )
        clusters[m_data["master_name"]].append(r)

    repaired_clusters = 0
    split_items_count = 0

    for m_name, items in list(clusters.items()):
        if len(items) <= 1:
            continue

        base_item = items[0]
        base_str = extract_strength_signature(base_item["clean_name"])
        base_pack = extract_pack_signature(base_item["clean_name"])
        base_cat = extract_form_category(base_item["clean_name"])
        base_is_combo = bool(
            re.search(
                r"\b(PLUS|CO|COMP|TRIO|AMLO|HCT|FORTE|EXTRA)\b",
                base_item["clean_name"],
                re.I,
            )
        )

        valid_items = [base_item]
        invalid_items = []

        for item in items[1:]:
            curr_str = extract_strength_signature(item["clean_name"])
            curr_pack = extract_pack_signature(item["clean_name"])
            curr_cat = extract_form_category(item["clean_name"])
            curr_is_combo = bool(
                re.search(
                    r"\b(PLUS|CO|COMP|TRIO|AMLO|HCT|FORTE|EXTRA)\b",
                    item["clean_name"],
                    re.I,
                )
            )

            # Breach 1: Form Category mismatch (e.g. Gel vs Tablet)
            if base_cat != "OTHER" and curr_cat != "OTHER" and base_cat != curr_cat:
                invalid_items.append(item)
                continue
            # Breach 2: Active strength conflict
            if base_str and curr_str and base_str != curr_str:
                invalid_items.append(item)
                continue
            # Breach 3: Pack count mismatch (20 TAB vs 30 TAB)
            if base_pack and curr_pack and base_pack != curr_pack:
                invalid_items.append(item)
                continue
            # Breach 4: Monotherapy vs Combination clash
            if base_is_combo != curr_is_combo:
                invalid_items.append(item)
                continue

            valid_items.append(item)

        if invalid_items:
            repaired_clusters += 1
            split_items_count += len(invalid_items)

            # Re-link valid siblings
            valid_codes = sorted([it["code"] for it in valid_items], key=safe_sort_key)
            for it in valid_items:
                master_map[it["code"]] = {
                    "master_name": m_name,
                    "dup_codes": valid_codes,
                }

            # Unmerge invalid siblings into separate clean entries
            for it in invalid_items:
                master_map[it["code"]] = {
                    "master_name": it["clean_name"].upper(),
                    "dup_codes": [it["code"]],
                }

    print(
        f"[✓] Clinical Guardrails Verified: {repaired_clusters} over-merged clusters corrected ({split_items_count} items safeguarded)."
    )
    return master_map


# ==============================================================================
# 5. DATA INGESTION ENGINE (AUTO-MERGES PARTS 1 TO 9)
# ==============================================================================
def load_all_catalog_records() -> List[Dict]:
    """Detects either a single catalog file or automatically concatenates parts 1 to 9."""
    records = []

    # Check for part files first
    part_files = sorted(glob.glob("used_items_extract_part_*.csv"))
    if part_files:
        print(
            f"[*] Detected {len(part_files)} catalog part files: {', '.join(part_files)}"
        )
        for pf in part_files:
            df_part = pd.read_csv(pf, encoding="utf-8-sig")
            code_col = df_part.columns[0]
            name_col = df_part.columns[1]
            for _, row in df_part.iterrows():
                code = str(row[code_col]).strip()
                raw_name = str(row[name_col]).strip()
                if code and raw_name and code != "nan":
                    records.append({"code": code, "raw_name": raw_name})
    else:
        # Fallback to single extract file
        candidate_files = ["used_items_extract.csv", "items.csv", "catalog.csv"]
        single_file = next((f for f in candidate_files if os.path.exists(f)), None)
        if not single_file:
            print(
                "[ERROR] Could not find 'used_items_extract_part_*.csv' or 'used_items_extract.csv'."
            )
            sys.exit(1)
        print(f"[*] Loading unified extract: {single_file}")
        df = pd.read_csv(single_file, encoding="utf-8-sig")
        code_col = df.columns[0]
        name_col = df.columns[1]
        for _, row in df.iterrows():
            code = str(row[code_col]).strip()
            raw_name = str(row[name_col]).strip()
            if code and raw_name and code != "nan":
                records.append({"code": code, "raw_name": raw_name})

    # Deduplicate ingestion list by code
    unique_records = {}
    for r in records:
        if r["code"] not in unique_records:
            clean = clean_trade_name(r["raw_name"])
            root = get_brand_root(clean)
            unique_records[r["code"]] = {
                "code": r["code"],
                "raw_name": r["raw_name"],
                "clean_name": clean,
                "root": root,
            }

    final_list = list(unique_records.values())
    print(f"[✓] Successfully ingested {len(final_list):,} unique active catalog SKUs.")
    return final_list


# ==============================================================================
# 6. PIPELINE ORCHESTRATION WITH RESUME CAPABILITY
# ==============================================================================
def run_deduplication_pipeline():
    print("=" * 80)
    print(
        "  GENIUS ERP MASTER CATALOG DEDUPLICATION PIPELINE (GEMINI 3.5 FLASH-LITE)  "
    )
    print("=" * 80)

    records = load_all_catalog_records()

    # Group records by brand root
    root_groups = defaultdict(list)
    for r in records:
        root_groups[r["root"]].append(r)

    print(f"[*] Clustered catalog into {len(root_groups):,} brand root families.")

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
                f"[*] Checkpoint loaded: {len(processed_roots)} roots previously resolved. Resuming..."
            )
        except Exception as e:
            print(f"[!] Warning reading checkpoint: {e}. Starting fresh.")

    # 2. Resolve Singletons Locally for Free (Zero API Tokens)
    singleton_count = 0
    for root, items in root_groups.items():
        if root not in processed_roots and (
            len(items) == 1 or items[0]["code"] in SERVICE_OR_SUPPLY_CODES
        ):
            for it in items:
                master_map[it["code"]] = {
                    "master_name": it["clean_name"].upper(),
                    "dup_codes": [it["code"]],
                }
            processed_roots.add(root)
            singleton_count += len(items)

    if singleton_count > 0:
        print(
            f"[✓] Resolved {singleton_count} singletons locally for FREE (0 API calls)."
        )

    # 3. Bundle Multi-Item Families into Micro-Batches of ~40 items
    pending_families = [
        {"root": root, "items": items}
        for root, items in root_groups.items()
        if root not in processed_roots and len(items) > 1
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
        f"[*] Packaged remaining multi-item families into {total_bundles} API requests."
    )
    print(
        f"[*] Daily Quota Budget: {total_bundles} / 500 RPD (Est. execution time: ~{round(total_bundles * REQUEST_DELAY_SECONDS / 60, 1)} mins)."
    )

    # 4. Execute AI Bundles
    for b_idx, bundle in enumerate(bundles, start=1):
        b_items_count = sum(len(f["items"]) for f in bundle)
        roots_preview = ", ".join(f["root"] for f in bundle[:3])
        print(
            f"[{b_idx}/{total_bundles}] Auditing {len(bundle)} brand roots ({b_items_count} items: {roots_preview}...)..."
        )

        result = cluster_bundle_with_gemini(bundle)

        # Register AI duplicate clusters
        for cl in result.get("clusters", []):
            m_name = cl.get("master_name", "").strip().upper()
            c_codes = [str(c).strip() for c in cl.get("codes", [])]
            for c in c_codes:
                master_map[c] = {"master_name": m_name, "dup_codes": c_codes}

        # Register AI confirmed uniques
        for u_code, u_name in result.get("uniques", {}).items():
            u_code = str(u_code).strip()
            if u_code not in master_map:
                master_map[u_code] = {
                    "master_name": str(u_name).strip().upper(),
                    "dup_codes": [u_code],
                }

        # Safety Fallback: Ensure no item is orphaned
        for fam in bundle:
            for it in fam["items"]:
                c = it["code"]
                if c not in master_map:
                    master_map[c] = {
                        "master_name": it["clean_name"].upper(),
                        "dup_codes": [c],
                    }
            processed_roots.add(fam["root"])

        # Checkpoint to disk after every batch
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            json.dump(
                {"processed_roots": list(processed_roots), "master_map": master_map}, f
            )

        time.sleep(REQUEST_DELAY_SECONDS)

    # 5. Apply Deterministic Clinical Guardrails (Eliminate False Merges)
    master_map = apply_clinical_guardrails(master_map, records)

    # ==============================================================================
    # 7. EXPORT STANDARDIZED CSV & T-SQL MIGRATION SCRIPT
    # ==============================================================================
    print("\n[*] Assembling final audited catalog rows...")
    final_output_rows = []

    for r in records:
        code = r["code"]
        raw_name = r["raw_name"]
        item_data = master_map.get(
            code, {"master_name": r["clean_name"].upper(), "dup_codes": [code]}
        )

        m_name = item_data["master_name"]
        all_siblings = sorted(list(set(item_data["dup_codes"])), key=safe_sort_key)
        dup_count = len(all_siblings) - 1
        dup_str = "; ".join(all_siblings) if dup_count > 0 else ""

        if dup_count == 0:
            comment = "Unique SKU"
        elif code == all_siblings[0]:
            comment = f"MASTER SKU (Consolidates {dup_count} duplicate records)"
        else:
            comment = f"Duplicate record to merge into Master Code {all_siblings[0]}"

        final_output_rows.append(
            {
                "code": code,
                "name": raw_name,
                "number of duplicates": dup_count,
                "duplicate_codes": dup_str,
                "proposed master name": m_name,
                "extra comments": comment,
                "sort_key_code": safe_sort_key(code),
            }
        )

    # Sort alphabetically by master name, then numerically by code
    final_output_rows.sort(
        key=lambda x: (x["proposed master name"], x["sort_key_code"])
    )

    # Write Output CSV with UTF-8 BOM
    with open(OUTPUT_CSV_FILE, "w", encoding="utf-8-sig", newline="") as f:
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

    # Generate Safe T-SQL Migration Script
    generate_tsql_script(final_output_rows, OUTPUT_SQL_FILE)

    print("\n" + "=" * 80)
    print("DEDUPLICATION PIPELINE COMPLETED SUCCESSFULLY")
    print(f"[✓] Audited Catalog CSV:  '{os.path.abspath(OUTPUT_CSV_FILE)}'")
    print(f"[✓] T-SQL Migration Script: '{os.path.abspath(OUTPUT_SQL_FILE)}'")
    print(f"[✓] Total Catalog Items:   {len(final_output_rows):,}")
    dups_total = sum(1 for r in final_output_rows if int(r["number of duplicates"]) > 0)
    print(f"[✓] Total Redundant SKUs:  {dups_total:,}")
    print(f"[✓] Canonical Master SKUs: {len(final_output_rows) - dups_total:,}")
    print("=" * 80)


def generate_tsql_script(rows: List[Dict], sql_filepath: str):
    """Generates a non-destructive, transaction-safe T-SQL migration script for Genius."""
    # Collect duplicate pairs: (Master_Code, Duplicate_Code)
    duplicate_pairs: List[Tuple[str, str]] = []
    seen_pairs = set()

    for r in rows:
        dup_str = r["duplicate_codes"].strip()
        if dup_str and int(r["number of duplicates"]) > 0:
            codes = [c.strip() for c in dup_str.split(";")]
            master_code = codes[0]
            for dup_code in codes[1:]:
                pair = (master_code, dup_code)
                if pair not in seen_pairs:
                    seen_pairs.add(pair)
                    duplicate_pairs.append(pair)

    with open(sql_filepath, "w", encoding="utf-8") as f:
        f.write(
            """-- =============================================================================
-- GENIUS (E-PLUS) ERP: CATALOG DEDUPLICATION MIGRATION SCRIPT
-- Target Database: Genius (Microsoft SQL Server)
-- Safety Directives:
--   1. ZERO HARD DELETES: Preserves referential integrity in sales_trans_d & pur_trans_d.
--   2. Secondary barcodes repointed to Master SKU in Item_Objects.
--   3. Active branch inventory balances consolidated in Item_Class_Store.
--   4. Duplicate items soft-deactivated in Item_Catalog (itm_active = '0', itm_stop_sell = 1).
-- =============================================================================

USE [Genius];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRANSACTION;
BEGIN TRY
    -- 1. Create Staging Table for Verified Deduplication Pairs
    IF OBJECT_ID('tempdb..#VerifiedDuplicates') IS NOT NULL
        DROP TABLE #VerifiedDuplicates;

    CREATE TABLE #VerifiedDuplicates (
        Master_Code VARCHAR(50) NOT NULL,
        Duplicate_Code VARCHAR(50) NOT NULL,
        PRIMARY KEY CLUSTERED (Master_Code, Duplicate_Code)
    );

    PRINT 'Loading verified duplicate clusters into staging table...';
"""
        )
        for m_code, d_code in duplicate_pairs:
            f.write(
                f"    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) VALUES ('{m_code}', '{d_code}');\n"
            )

        f.write(f"""
    PRINT 'Loaded {len(duplicate_pairs)} duplicate relationships.';

    -- 2. Repoint Secondary Barcodes in Item_Objects to Master SKU
    PRINT 'Step 1: Repointing secondary barcodes in Item_Objects...';
    UPDATE obj
    SET obj.io_itm_id = m.itm_id
    FROM Item_Objects obj
    INNER JOIN Item_Catalog dup ON obj.io_itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Catalog m ON m.itm_code = map.Master_Code
    WHERE obj.io_itm_id <> m.itm_id;

    -- 3. Consolidate Active Branch Stock Balances in Item_Class_Store
    PRINT 'Step 2: Consolidating branch inventory in Item_Class_Store...';

    -- A. Transfer quantities from duplicate batch records to master batch records where same store exists
    UPDATE master_ics
    SET master_ics.itm_qty = master_ics.itm_qty + dup_ics.itm_qty
    FROM Item_Class_Store master_ics
    INNER JOIN Item_Catalog m ON master_ics.itm_id = m.itm_id
    INNER JOIN #VerifiedDuplicates map ON m.itm_code = map.Master_Code
    INNER JOIN Item_Catalog dup ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Class_Store dup_ics ON dup_ics.itm_id = dup.itm_id 
        AND dup_ics.sto_id = master_ics.sto_id;

    -- B. Zero out inventory on duplicate records
    UPDATE dup_ics
    SET dup_ics.itm_qty = 0
    FROM Item_Class_Store dup_ics
    INNER JOIN Item_Catalog dup ON dup_ics.itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code;

    -- 4. Soft-Deactivate Redundant Catalog Rows
    PRINT 'Step 3: Soft-deactivating redundant items in Item_Catalog...';
    UPDATE cat
    SET cat.itm_active = '0',
        cat.itm_stop_sell = 1,
        cat.itm_stop_pur = 1,
        cat.itm_notes = ISNULL(cat.itm_notes, '') + ' [MERGED INTO MASTER CODE ' + map.Master_Code + ']'
    FROM Item_Catalog cat
    INNER JOIN #VerifiedDuplicates map ON cat.itm_code = map.Duplicate_Code;

    COMMIT TRANSACTION;
    PRINT 'SUCCESS: Catalog deduplication migration completed and committed.';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;
    PRINT 'ERROR OCCURRED: Transaction rolled back successfully.';
    PRINT ERROR_MESSAGE();
END CATCH;
GO
""")


if __name__ == "__main__":
    run_deduplication_pipeline()
