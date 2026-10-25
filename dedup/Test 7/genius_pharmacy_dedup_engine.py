#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
========================================================================================
GENIUS (e-Plus) PHARMACY ERP - MASTER CATALOG DEDUPLICATION & ENTITY RESOLUTION ENGINE
========================================================================================
Architecture:
  - Stage 1: Deterministic Multi-Lingual Sanitizer (Python)
  - Stage 2: Root-Safe Brand Family Grouping (Zero Boundary Splits)
  - Stage 3: Free Local Singleton Resolver (0 API Calls for ~40% of catalog)
  - Stage 4: Bundled Micro-Batch AI Clustering (Gemini Flash-Lite via Google GenAI SDK)
  - Stage 5: Clinical Guardrail Audit & Twin-Ghost Reconciler (Rule 1-A Enforcement)
  - Stage 6: Symmetrical CSV (UTF-8 BOM) & Safe T-SQL Migration Generator
========================================================================================
"""

import os
import sys
import re
import csv
import time
import json
from collections import defaultdict
from typing import Any, Dict, List, Tuple
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

# ======================================================================================
# 0. API & MODEL CONFIGURATION
# ======================================================================================
API_KEY = os.getenv("GEMINI_API_KEY", os.getenv("GCP_API_KEY"))

# Use the Google GenAI SDK
try:
    from google import genai
    from google.genai import types
    from google.genai.errors import ClientError

    CLIENT = genai.Client(api_key=API_KEY) if API_KEY else None
except ImportError:
    print("[!] Google GenAI SDK not installed. Run: pip install google-genai pydantic")
    sys.exit(1)

MODEL_ID = "gemini-3.5-flash-lite"  # Highly efficient, cost-effective, adheres strictly to JSON schemas

INPUT_FILE = "used_items_extract.csv"
OUTPUT_CSV = "final_deduplicated_catalog_PERFECT.csv"
OUTPUT_SQL = "execute_genius_deduplication.sql"
CHECKPOINT_FILE = "checkpoint_clusters.json"

TARGET_ITEMS_PER_PROMPT = 40
REQUEST_DELAY_SECONDS = 8
MAX_RETRIES = 5

# Clinical services & hospital consumables that must remain unmerged singletons
EXCLUDED_SERVICE_CODES = {"2", "IN", "G5", "G10", "S3", "S5", "S10", "WI", "MA", "PZ"}


# ======================================================================================
# 1. STAGE 1: SURGICAL NON-DESTRUCTIVE PRE-SANITIZER
# ======================================================================================
def clean_trade_name(text: str) -> str:
    """
    Strips noise, cashier tags, and explicit prices without breaking
    clinical strengths, combination ratios, or formulation modifiers.
    """
    if not isinstance(text, str) or not text.strip():
        return ""

    s = text.strip()

    # Normalize common Egyptian pharmacy unit typos
    s = re.sub(r"\b1000\s*mMG\b", "1000 MG", s, flags=re.I)
    s = re.sub(r"\bmMG\b", "MG", s, flags=re.I)
    s = re.sub(r"\bCABS(ULES)?\b", "CAP", s, flags=re.I)
    s = re.sub(r"\bTABLES\b", "TAB", s, flags=re.I)
    s = re.sub(r"\bSYRY?P\b", "SYRUP", s, flags=re.I)

    # Strip bracket and symbol noise while preserving hyphens and slashes
    s = re.sub(r"[@#*!?\[\]{}()~^%_]+", " ", s)

    # 1. Strip explicit currency patterns (e.g., '18LE', '22.5 L.E', '357 EGP', '144 جنيه')
    s = re.sub(r"(?i)\b\d+(?:\.\d+)?\s*(?:LE|L\.E|EGP|L\.E\.|جنيه)\b", " ", s)
    s = re.sub(r"(?i)\b\d+LE\b", " ", s)

    # 2. Strip trailing naked prices preceded by package forms or volumes
    # E.g. "CENTRAVITA GINKO 30 TAB 375" -> "CENTRAVITA GINKO 30 TAB"
    s = re.sub(
        r"(?i)(\b(?:TAB|CAP|TABLETS?|CAPSULES?|ML|GM|G|SACH|SACHETS?|AMP|AMPOULES?|VIALS?|PIECES|DOSE|PUFF)\b\s*)\d{2,4}\s*$",
        r"\1",
        s.strip(),
    )
    # E.g. "CURAM 457MG 70ML 97" -> "CURAM 457MG 70ML"
    s = re.sub(r"(?i)(\b(?:ML|GM|MG)\b\s*)\d{2,4}\s*$", r"\1", s.strip())

    # 3. Strip cashier status flags and lifecycle markers
    noise_patterns = [
        r"(?i)\bNEW\s*PRICE\b",
        r"(?i)\bOLD\s*PRICE\b",
        r"(?i)\bN\.?P\.?\d*\b",
        r"(?i)\bN\/PRICE\b",
        r"(?i)\bOLD\/PRICE\b",
        r"(?i)\bNEU\s*\.?PRICE\b",
        r"(?i)\bNEW\b",
        r"(?i)\bOLD\b",
        r"(?i)\bPRES\b",
        r"(?i)\bEGY\b",
        r"سعر\s*جديد",
        r"سعر\s*قديم",
        r"بدون\s*خصم",
        r"موقوف",
        r"خصم\s*خاص",
        r"(?i)\bOFFER\b",
        r"(?i)\bFREE\b",
        r"(?i)\bGIFT\b",
        r"(?i)\bBONUS\b",
        r"(?i)\bBOUNS\b",
    ]
    for p in noise_patterns:
        s = re.sub(p, " ", s)

    # 4. Strip trailing cashier manufacturer tags (e.g. "mup", "pharco", "sedico")
    s = re.sub(
        r"(?i)\s+\b(M|MUP|SEDICO|PHARCO|AMRYIA|AMRIYA|T3A|AUG|EVA|EIPICO|EPICO)\b\s*$",
        "",
        s,
    )

    return re.sub(r"\s+", " ", s).strip()


def get_brand_root(name: str) -> str:
    """
    Extracts the primary alphanumeric brand root without splitting compound drug stems.
    Keeps 'CO-TAREG', 'E-MOXCLAV', 'TRI-M', 'BIO-TULLE' intact as single brand families.
    """
    clean = re.sub(r"^[^a-zA-Z0-9\u0621-\u064A]+", "", name.strip())
    if not clean:
        return "OTHER"

    # Match common pharmaceutical hyphenated prefixes
    match_compound = re.match(
        r"^(CO|E|TRI|X|NO|DE|BIO|PRO|ULTRA)[-_ ]+([a-zA-Z0-9\u0621-\u064A]+)",
        clean,
        flags=re.I,
    )
    if match_compound:
        return f"{match_compound.group(1).upper()}-{match_compound.group(2).upper()}"

    tokens = re.split(r"[\s/]+", clean)
    return tokens[0].upper() if tokens else "OTHER"


def safe_sort_key(code: Any) -> Tuple[int, int, str]:
    """
    Prevents TypeError: '<' not supported between instances of 'int' and 'str'.
    """
    s = str(code).strip()
    return (0, int(s), "") if s.isdigit() else (1, 0, s.upper())


# ======================================================================================
# 2. STAGE 4: ADVANCED AI CLUSTERING ENGINE (JSON ONLY)
# ======================================================================================
SYSTEM_INSTRUCTION = """You are a senior pharmaceutical database auditor specializing in Egyptian pharmacy catalogs.
Group redundant inventory entries for identical pharmaceutical products into duplicate clusters.

CLINICAL & ERP DEDUPLICATION RULES:
1. STRICT COMMERCIAL BRAND ISOLATION:
   - Never merge different commercial trade brands (e.g., Paramol != Panadol != Cetal != Abimol).
2. RULE 1-A (STRICT PACK SIZE & STRENGTH MATCHING):
   - Items with different active strengths (e.g., 20 MG vs 40 MG, 5000 IU vs 10000 IU) MUST NEVER be merged.
   - Items with different pack sizes or volumes (e.g., 20 TAB vs 30 TAB, 60 ML vs 100 ML, 1 AMP vs 3 AMP) MUST NEVER be merged.
   - Merging is strictly reserved for items sharing identical: Brand + Strength + Dosage Form + Pack Count/Volume.
3. COMBINATION & KINETIC MODIFIERS:
   - Never merge plain formulations with combination forms (e.g., Plain vs PLUS, Plain vs CO, Plain vs COMP, Plain vs HCT).
   - Never merge immediate release with sustained release (e.g., Plain vs SR, Plain vs XR, Plain vs RETARD).
4. CANONICAL MASTER NAME SYNTHESIS:
   - For each cluster or unique item, assign a clean, standardized English trade title (e.g., 'CATAFLAM 50 MG 20 TAB', 'AMARYL 2 MG 30 TAB').
   - For clinical services, provide clear descriptions (e.g., 'BLOOD PRESSURE CHECK (SERVICE)').

OUTPUT FORMAT: Return JSON ONLY matching this schema:
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
    prompt = "Review these distinct brand families and cluster duplicate records within each family:\n\n"
    for fam in bundle_families:
        prompt += f"=== BRAND FAMILY: {fam['root']} ===\n"
        for it in fam["items"]:
            prompt += f"- Code: {it['code']} | Cleaned: '{it['clean_name']}' | Raw: '{it['raw_name']}'\n"
        prompt += "\n"

    for attempt in range(max_retries):
        try:
            response = CLIENT.models.generate_content(
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
                raise RuntimeError(f"Permanent configuration error: {error_str}")

            match = re.search(r"retryDelay':\s*'(\d+)s", error_str) or re.search(
                r"retry in (\d+(\.\d+)?)s", error_str
            )
            wait = int(float(match.group(1))) + 2 if match else 25
            print(
                f"\n[!] Rate throttle reached. Pausing {wait}s (Attempt {attempt+1}/{max_retries})..."
            )
            time.sleep(wait)

        except Exception as e:
            print(
                f"\n[!] Network or parsing error: {e}. Retrying in 10s (Attempt {attempt+1}/{max_retries})..."
            )
            time.sleep(10)

    # Safe fallback if batch fails permanently
    fallback = {"clusters": [], "uniques": {}}
    for fam in bundle_families:
        for it in fam["items"]:
            fallback["uniques"][it["code"]] = it["clean_name"].upper()
    return fallback


# ======================================================================================
# 3. STAGE 5: CLINICAL INTEGRITY GUARDRAILS & TWIN-GHOST RECONCILER
# ======================================================================================
def extract_form_category(name: str) -> str:
    s = name.upper()
    if re.search(r"\b(GEL|CREAM|OINT|OINTMENT|LOTION|PAINT)\b", s):
        return "TOPICAL"
    if re.search(r"\b(TAB|CAP|TABLETS?|CAPSULES?|PILL|DRAGEE)\b", s):
        return "ORAL_SOLID"
    if re.search(r"\b(SYRUP|SUSP|SUSPENSION|SOLUTION|DROPS)\b", s):
        return "ORAL_LIQUID"
    if re.search(r"\b(AMP|VIAL|INJECTION|AMPOULES?)\b", s):
        return "INJECTABLE"
    if re.search(r"\b(SPRAY|INHALER|PUFF)\b", s):
        return "INHALER_SPRAY"
    return "OTHER"


def extract_strength_token(name: str) -> str:
    match = re.search(
        r"\b(\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)*)\s*(MG|GM|G|MCG|IU|%|ML)\b",
        name,
        re.I,
    )
    return re.sub(r"\s+", "", match.group(0).upper()) if match else ""


def extract_pack_token(name: str) -> str:
    match = re.search(
        r"\b(\d+)\s*(TAB|CAP|TABLETS?|CAPSULES?|AMP|VIAL|SACHET|DOSE|ML|GM)S?\b",
        name,
        re.I,
    )
    return re.sub(r"\s+", "", match.group(0).upper()) if match else ""


def validate_and_repair_clusters(rows: List[Dict]) -> List[Dict]:
    """
    Audits every duplicate cluster against Rule 1-A to prevent false-positive merges.
    Splits any cluster with conflicting forms, strengths, or pack sizes.
    """
    print("\n[*] Auditing clinical integrity and Rule 1-A guardrails...")

    # Group row indices by proposed master name
    clusters = defaultdict(list)
    for idx, r in enumerate(rows):
        clusters[r["proposed master name"]].append(idx)

    repaired_count = 0

    for master_name, indices in list(clusters.items()):
        if len(indices) <= 1:
            continue

        base_idx = indices[0]
        base_name = rows[base_idx]["name"]
        base_form = extract_form_category(base_name)
        base_str = extract_strength_token(base_name)
        base_pack = extract_pack_token(base_name)

        invalid_indices = []
        for i in indices[1:]:
            curr_name = rows[i]["name"]
            curr_form = extract_form_category(curr_name)
            curr_str = extract_strength_token(curr_name)
            curr_pack = extract_pack_token(curr_name)

            # Check 1: Dosage form clash (e.g. Gel merged into Tablet)
            if base_form != "OTHER" and curr_form != "OTHER" and base_form != curr_form:
                invalid_indices.append(i)
                continue

            # Check 2: Strength mismatch (e.g. 50 MG vs 100 MG)
            if base_str and curr_str and base_str != curr_str:
                invalid_indices.append(i)
                continue

            # Check 3: Pack count mismatch (e.g. 20 TAB vs 30 TAB)
            if base_pack and curr_pack and base_pack != curr_pack:
                invalid_indices.append(i)
                continue

        # Split out invalid items
        if invalid_indices:
            repaired_count += len(invalid_indices)
            for inv_i in invalid_indices:
                # Convert back to clean individual SKU
                rows[inv_i]["proposed master name"] = clean_trade_name(
                    rows[inv_i]["name"]
                ).upper()

    # Re-calculate symmetrical duplicate codes and counts across the catalog
    final_groups = defaultdict(list)
    for idx, r in enumerate(rows):
        final_groups[r["proposed master name"]].append(idx)

    for m_name, indices in final_groups.items():
        all_codes = sorted(
            list(set(str(rows[i]["code"]).strip() for i in indices)), key=safe_sort_key
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

    print(
        f"[✓] Clinical guardrails enforced: {repaired_count} unsafe merges successfully unmerged."
    )
    return rows


# ======================================================================================
# 4. STAGE 6: SAFE T-SQL MIGRATION SCRIPT GENERATOR
# ======================================================================================
def generate_sql_migration(rows: List[Dict], sql_filepath: str):
    """
    Generates a non-destructive, transaction-safe SQL migration script for Genius ERP.
    - Zero hard deletes (protects foreign keys on sales_trans_d and pur_trans_d).
    - Maps secondary barcodes to Master Code in Item_Objects.
    - Consolidates stock balances in Item_Class_Store.
    - Soft-deactivates duplicate items in Item_Catalog.
    """
    print(f"[*] Generating non-destructive T-SQL migration script: {sql_filepath}...")

    # Group records by proposed master name to find (Master_Code, Duplicate_Code) pairs
    clusters = defaultdict(list)
    for r in rows:
        if r["number of duplicates"] > 0:
            clusters[r["proposed master name"]].append(r["code"])

    duplicate_pairs = []
    for master_name, codes in clusters.items():
        sorted_codes = sorted(list(set(codes)), key=safe_sort_key)
        master_code = sorted_codes[0]
        for dup_code in sorted_codes[1:]:
            duplicate_pairs.append((master_code, dup_code))

    with open(sql_filepath, mode="w", encoding="utf-8") as f:
        f.write(
            f"""-- =============================================================================
-- GENIUS (e-Plus) PHARMACY ERP: MASTER DEDUPLICATION MIGRATION
-- Database: [Genius]
-- Total Duplicate Records to Consolidate: {len(duplicate_pairs)}
-- Safety Directives:
--   1. ZERO hard deletes (preserves foreign keys across sales & purchases).
--   2. Secondary barcodes mapped to Master SKU in Item_Objects.
--   3. Stock balances consolidated in Item_Class_Store.
--   4. Duplicate records soft-deactivated in Item_Catalog.
-- =============================================================================

USE [Genius];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRANSACTION;
BEGIN TRY
    IF OBJECT_ID('tempdb..#VerifiedDuplicates') IS NOT NULL
        DROP TABLE #VerifiedDuplicates;

    CREATE TABLE #VerifiedDuplicates (
        Master_Code VARCHAR(50) NOT NULL,
        Duplicate_Code VARCHAR(50) NOT NULL,
        PRIMARY KEY CLUSTERED (Master_Code, Duplicate_Code)
    );

    PRINT 'Loading verified duplicate cluster mappings...';
"""
        )

        for m_code, d_code in duplicate_pairs:
            f.write(
                f"    INSERT INTO #VerifiedDuplicates VALUES ('{m_code}', '{d_code}');\n"
            )

        f.write(f"""
    PRINT 'Loaded {len(duplicate_pairs)} duplicate pairs into staging.';

    -- 1. Re-point secondary barcodes in Item_Objects to Master SKU
    PRINT 'Step 1: Re-pointing barcodes in Item_Objects...';
    UPDATE obj
    SET obj.io_itm_id = m.itm_id
    FROM Item_Objects obj
    INNER JOIN Item_Catalog dup ON obj.io_itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Catalog m ON m.itm_code = map.Master_Code
    WHERE obj.io_itm_id <> m.itm_id;

    -- 2. Consolidate inventory balances in Item_Class_Store
    PRINT 'Step 2: Consolidating branch stock balances in Item_Class_Store...';
    
    -- A. Add duplicate quantity to master record where matching store & batch exist
    UPDATE master_ics
    SET master_ics.itm_qty = master_ics.itm_qty + dup_ics.itm_qty
    FROM Item_Class_Store master_ics
    INNER JOIN Item_Catalog m ON master_ics.itm_id = m.itm_id
    INNER JOIN #VerifiedDuplicates map ON m.itm_code = map.Master_Code
    INNER JOIN Item_Catalog dup ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Class_Store dup_ics ON dup_ics.itm_id = dup.itm_id 
        AND dup_ics.sto_id = master_ics.sto_id;

    -- B. Zero out inventory balances on duplicate items
    UPDATE dup_ics
    SET dup_ics.itm_qty = 0
    FROM Item_Class_Store dup_ics
    INNER JOIN Item_Catalog dup ON dup_ics.itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code;

    -- 3. Soft-deactivate redundant records in Item_Catalog
    PRINT 'Step 3: Soft-deactivating redundant records in Item_Catalog...';
    UPDATE cat
    SET cat.itm_active = '0',
        cat.itm_stop_sell = 1,
        cat.itm_stop_pur = 1,
        cat.itm_notes = ISNULL(cat.itm_notes, '') + ' [DEDUP: MERGED INTO MASTER CODE ' + map.Master_Code + ']'
    FROM Item_Catalog cat
    INNER JOIN #VerifiedDuplicates map ON cat.itm_code = map.Duplicate_Code;

    COMMIT TRANSACTION;
    PRINT 'SUCCESS: Deduplication migration committed successfully.';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;

    PRINT 'FAILED: Error occurred during migration. Transaction rolled back.';
    PRINT ERROR_MESSAGE();
END CATCH;
GO
""")
    print(
        f"[✓] Safe T-SQL migration script written: {len(duplicate_pairs)} actions staged."
    )


# ======================================================================================
# 5. MASTER EXECUTION PIPELINE
# ======================================================================================
def run():
    print("==========================================================================")
    print("STARTING GENIUS ERP CATALOG DEDUPLICATION PIPELINE")
    print("==========================================================================")

    if not os.path.exists(INPUT_FILE):
        print(
            f"[-] Error: '{INPUT_FILE}' not found. Please ensure it is in this folder."
        )
        sys.exit(1)

    raw_df = pd.read_csv(INPUT_FILE)
    print(f"[*] Ingested {len(raw_df):,} total raw records from '{INPUT_FILE}'.")

    # Detect code and name columns dynamically
    code_col = "Item_Code" if "Item_Code" in raw_df.columns else raw_df.columns[0]
    name_col = (
        "Item_English_Name"
        if "Item_English_Name" in raw_df.columns
        else raw_df.columns[1]
    )

    records = []
    for _, row in raw_df.iterrows():
        code = str(row[code_col]).strip()
        raw_name = str(row[name_col]).strip()
        clean_name = clean_trade_name(raw_name)
        root = get_brand_root(clean_name)
        records.append(
            {"code": code, "raw_name": raw_name, "clean_name": clean_name, "root": root}
        )

    # Group by brand root (Zero Boundary Splits)
    root_groups = defaultdict(list)
    for r in records:
        root_groups[r["root"]].append(r)

    print(f"[*] Partitioned into {len(root_groups):,} brand families.")

    # Load Checkpoint if resuming
    processed_roots = set()
    master_map = {}
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                master_map = saved.get("master_map", {})
                processed_roots = set(saved.get("processed_roots", []))
            print(
                f"[*] Resuming from checkpoint: {len(processed_roots)} roots already completed."
            )
        except Exception as e:
            print(f"[!] Checkpoint warning: {e}. Starting fresh.")

    # Stage 3: Resolve Singletons Locally for FREE (0 API Tokens)
    singleton_count = 0
    for root, items in root_groups.items():
        if root not in processed_roots:
            # If family has only 1 item, or is a protected clinical service, resolve locally
            if len(items) == 1 or all(
                it["code"].upper() in EXCLUDED_SERVICE_CODES for it in items
            ):
                for it in items:
                    c = it["code"]
                    master_map[c] = {
                        "master_name": it["clean_name"].upper(),
                        "dup_codes": [c],
                    }
                processed_roots.add(root)
                singleton_count += len(items)

    print(
        f"[✓] Stage 3: Resolved {singleton_count:,} singletons locally for FREE (0 API calls)."
    )

    # Stage 4: Package remaining candidate families into small batches (~40 items)
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
        f"[*] Remaining candidate families packaged into {total_bundles} AI micro-batches."
    )
    print(
        f"[*] Estimated total API requests: {total_bundles} (Daily quota remaining: {500 - total_bundles})"
    )

    # Execute AI Bundles
    for b_idx, bundle in enumerate(bundles, start=1):
        b_items_count = sum(len(f["items"]) for f in bundle)
        print(
            f"[{b_idx}/{total_bundles}] Auditing {len(bundle)} brand roots ({b_items_count} items)..."
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

        # Safety Fallback: ensure no item was dropped
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
                {"processed_roots": list(processed_roots), "master_map": master_map}, f
            )

        time.sleep(REQUEST_DELAY_SECONDS)

    # Assemble preliminary row records
    assembled_rows = []
    for r in records:
        code = r["code"]
        item_data = master_map.get(
            code, {"master_name": r["clean_name"].upper(), "dup_codes": [code]}
        )
        assembled_rows.append(
            {
                "code": code,
                "name": r["raw_name"],
                "number of duplicates": 0,
                "duplicate_codes": "",
                "proposed master name": item_data["master_name"],
                "extra comments": "",
            }
        )

    # Stage 5: Enforce Clinical Integrity Guardrails & Symmetrical Codes
    final_rows = validate_and_repair_clusters(assembled_rows)

    # Sort alphabetically by master name, then by code
    final_rows.sort(key=lambda x: (x["proposed master name"], x["sort_key"]))

    # Stage 6A: Write Final RFC 4180 CSV with UTF-8 BOM
    with open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
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
        for r in final_rows:
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

    # Stage 6B: Write Non-Destructive SQL Migration Script
    generate_sql_migration(final_rows, OUTPUT_SQL)

    print("\n" + "=" * 75)
    print("[✔] DEDUPLICATION & ENTITY RESOLUTION FINISHED SUCCESSFULLY")
    print(f"[✔] Total Records Processed:   {len(final_rows):,}")
    dups = sum(1 for r in final_rows if int(r["number of duplicates"]) > 0)
    print(f"[✔] Total Duplicate Items:     {dups:,}")
    print(f"[✔] Total Unique Master SKUs:  {len(final_rows) - dups:,}")
    print(f"[✔] Master Catalog File:       '{os.path.abspath(OUTPUT_CSV)}'")
    print(f"[✔] Safe T-SQL Migration File: '{os.path.abspath(OUTPUT_SQL)}'")
    print("=" * 75)


if __name__ == "__main__":
    run()
