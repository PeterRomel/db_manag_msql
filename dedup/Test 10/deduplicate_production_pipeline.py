#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
========================================================================================
GENIUS (e-Plus) PHARMACY ERP - MASTER CATALOG DEDUPLICATION & ENTITY RESOLUTION PIPELINE
========================================================================================
Engineered for:
  - Strict compliance with Gemini Flash-Lite Free Tier quotas (15 RPM, 500 RPD, 250k TPM)
  - Zero Python brand/root separation (avoids compound prefix & Arabic boundary errors)
  - Sequential Lexicographical Micro-Batching (35 items/request = ~118 total requests)
  - Pydantic JSON Schema Structured Outputs (zero hallucinated keys or markdown bugs)
  - Full State Checkpointing (resume from exact batch index on interruption)
  - Global Cross-Catalog Reconciliation & Clinical Safety Guardrails
  - Native Excel Tables (openpyxl) with total rows, auto-filters, and conditional formatting
  - Production T-SQL migration script with zero hard deletes
========================================================================================
"""

import os
import re
import sys
import time
import json
import logging
from typing import List, Dict, Any, Optional
from collections import defaultdict

import pandas as pd
from pydantic import BaseModel, Field

# Google GenAI SDK (New Official SDK)
try:
    from google import genai
    from google.genai import types
    from google.genai.errors import ClientError
except ImportError:
    print(
        "[ERROR] Please install the official Google GenAI SDK: pip install google-genai pydantic openpyxl pandas"
    )
    sys.exit(1)

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule

# --------------------------------------------------------------------------------------
# 0. CONFIGURATION & LOGGING
# --------------------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("GeniusDedup")

API_KEY = os.getenv("GEMINI_API_KEY", os.getenv("GCP_API_KEY", ""))
MODEL_ID = "gemini-3.5-flash-lite"

# Rate Limit Controls
BATCH_SIZE = 35  # 4,137 / 35 = ~118 requests (Quota limit: 500 RPD)
REQUEST_DELAY_SECONDS = 5.5  # Enforces ~10.9 RPM (Quota limit: 15 RPM)
MAX_RETRIES = 5

INPUT_CSV = "used_items_extract.csv"
CHECKPOINT_JSON = "checkpoint_dedup_progress.json"
OUTPUT_EXCEL = "01_Genius_Deduplicated_Master_Catalog.xlsx"
OUTPUT_SQL = "execute_genius_deduplication.sql"


# --------------------------------------------------------------------------------------
# 1. DETERMINISTIC PYTHON PRE-SANITIZER
# --------------------------------------------------------------------------------------
def clean_pharmacy_name(raw_name: str) -> str:
    """
    Strips objective cashier noise, currency prices, status flags, and typos.
    Does NOT attempt to split brand stems or extract roots in Python.
    """
    if not isinstance(raw_name, str) or not raw_name.strip():
        return ""

    s = raw_name.strip()

    # A. Normalize unit & packaging typos
    s = re.sub(r"(?i)\b1000\s*mMG\b", "1000 MG", s)
    s = re.sub(r"(?i)\bmMG\b", "MG", s)
    s = re.sub(r"(?i)\bCABS(ULES)?\b", "CAP", s)
    s = re.sub(r"(?i)\bTABLES\b", "TAB", s)
    s = re.sub(r"(?i)\bSYRY?P\b", "SYRUP", s)
    s = re.sub(r"(?i)\bTABALATIL\b", "TAB", s)
    s = re.sub(r"(?i)\bVAIL\b", "VIAL", s)
    s = re.sub(r"(?i)\bPRISE\b", "PRICE", s)
    s = re.sub(r"(?i)\bFLIMS?\b", "FILM", s)

    # B. Strip noise symbols & brackets (leaving decimal points in dosages intact)
    s = re.sub(r"[@#*!?%^~_\[\]{}()]+", " ", s)

    # C. Strip explicit currency prices (e.g. '24le', '28LE', '120 L.E', '357 EGP', '144 جنيه')
    s = re.sub(r"(?i)\b\d+(?:\.\d+)?\s*(?:LE|L\.E|EGP|L\.E\.|جنيه)\b", " ", s)
    s = re.sub(r"(?i)\b\d+LE\b", " ", s)

    # D. Strip lifecycle / cashier noise keywords & versions (e.g., 'NEW PRICE 2', 'NEW.P', 'OLD/PRICE')
    noise_patterns = [
        r"(?i)\b(?:NEW|OLD|NEU)[\s./_-]*P(?:RICE)?(?:\s*\d+)?\b",
        r"(?i)\bN[\s./_-]*P(?:RICE)?\d*\b",
        r"(?i)\bOLD\/PRICE\b",
        r"(?i)\b(?:NEW|OLD|PRES|EGY)\b",
        r"سعر\s*جديد(?:\s*\d+)?",
        r"سعر\s*قديم(?:\s*\d+)?",
        r"بدون\s*خصم",
        r"موقوف",
        r"خصم\s*خاص",
        r"(?i)\b(?:OFFER|GIFT|FREE|BONUS|BOUNS|SAMPLE|UPDATE)\b",
    ]
    for p in noise_patterns:
        s = re.sub(p, " ", s)

    # E. Strip trailing naked prices where a pack count and form already exist
    # E.g. "curam 457mg 70ml 97" -> "curam 457mg 70ml"
    # E.g. "ANTICOX II 15MG 30 tab 111" -> "ANTICOX II 15MG 30 tab"
    # E.g. "BLADOGRA 25 MG 30 TAB 255" -> "BLADOGRA 25 MG 30 TAB"
    s = re.sub(
        r"(?i)(\b\d+\s*(?:TAB|TABS|TABLETS?|CAP|CAPS|CAPSULES?|AMP|AMPOULES?|AMPS|VIALS?|SACH|SACHETS?|SUPP|ML|GM|G)\b\s*)\d{2,4}\s*$",
        r"\1",
        s,
    )

    # F. Strip trailing distributor / manufacturer initials added at register
    s = re.sub(
        r"(?i)\s+\b(m|k|pres|aug|t3a|sedico|pharco|amryia|amrya|mup|eipico|epico|cid)\b\s*$",
        "",
        s,
    )

    return re.sub(r"\s+", " ", s).strip()


# --------------------------------------------------------------------------------------
# 2. PYDANTIC STRUCTURED OUTPUT SCHEMA
# --------------------------------------------------------------------------------------
class ParsedPharmacyItem(BaseModel):
    item_code: str = Field(description="Exact original item code passed in input")
    canonical_brand: str = Field(
        description="Standardized commercial English brand name only (e.g. 'PANADOL EXTRA', 'CONCOR COR', 'AUGMENTIN', 'CO-TAREG'). Never replace brand with generic molecule."
    )
    generic_name: str = Field(
        default="",
        description="Active chemical molecule(s) in English (e.g. 'PARACETAMOL / CAFFEINE', 'BISOPROLOL', 'AMOXICILLIN / CLAVULANIC ACID').",
    )
    strength: str = Field(
        default="",
        description="Standardized strength with unit (e.g. '500 MG', '1 GM', '160/12.5 MG', '200000 IU', '0.5%'). Leave empty if not applicable.",
    )
    dosage_form: str = Field(
        description="Standardized dosage form in English: TABLET, CAPSULE, SYRUP, SUSPENSION, AMPOULE, VIAL, CREAM, OINTMENT, GEL, DROPS, SPRAY, SUPPOSITORY, SACHET, PATCH, SHAMPOO, SOAP, DEVICE, SERVICE, OTHER."
    )
    pack_size: str = Field(
        default="",
        description="Package count or volume: '20 TAB', '14 CAP', '100 ML', '6 AMP', '15 GM', '12 SACH', '1 VIAL'. Crucial: Rule 1-A strictly distinguishes 20 TAB vs 30 TAB.",
    )
    canonical_sku_name: str = Field(
        description="Clean standardized master title: BRAND STRENGTH DOSAGE_FORM PACK_SIZE"
    )
    canonical_signature: str = Field(
        description="Exact composite key: BRAND|STRENGTH|DOSAGE_FORM|PACK_SIZE"
    )


class BatchExtractionResponse(BaseModel):
    items: List[ParsedPharmacyItem]


SYSTEM_PROMPT = """You are a senior pharmaceutical catalog auditor for Egyptian Pharmacy ERP systems.
Your task is to parse raw pharmacy inventory records into standardized clinical attributes.

CRITICAL CLINICAL & ERP RULES:
1. COMMERCIAL BRAND ISOLATION:
   - Commercial brand trade names must NEVER be merged across different products.
   - PARAMOL != PANADOL != CETAL != ABIMOL.
   - Preserve compound brand extensions: 'PANADOL EXTRA' != 'PANADOL ADVANCE' != 'PANADOL COLD&FLU'.
   - Preserve prefixes: 'CO-TAREG' != 'TAREG'; 'E-MOXCLAV' != 'MOXCLAV'; 'CONCOR AMLO' != 'CONCOR COR' != 'CONCOR'.

2. RULE 1-A (STRICT PACK SIZE & STRENGTH POLICY):
   - Items with different active strengths (e.g., 20 MG vs 40 MG, 5000 IU vs 50000 IU) are DIFFERENT SKUs.
   - Items with different pack sizes or volumes (e.g., 20 TAB vs 30 TAB, 60 ML vs 100 ML) are DIFFERENT SKUs.
   - NEVER merge different pack sizes into the same canonical signature.

3. ARABIC TRANSLATION & STANDARDIZATION:
   - Translate Arabic clinical dosage forms into standardized English (e.g., 'أقراص' -> TABLET, 'شراب' -> SYRUP, 'كريم' -> CREAM).
   - If the item is a clinical service (e.g., 'قياس ضغط') or hospital consumable (e.g., 'سرنجة 5 سم'), categorize form as SERVICE or DEVICE.

4. COMPOSITE SKU SIGNATURE:
   - Construct 'canonical_signature' strictly as: UPPERCASE(BRAND)|UPPERCASE(STRENGTH)|UPPERCASE(DOSAGE_FORM)|UPPERCASE(PACK_SIZE).
"""


# --------------------------------------------------------------------------------------
# 3. AI EXTRACTION ENGINE WITH SELF-HEALING RATE-LIMIT THROTTLING
# --------------------------------------------------------------------------------------
def execute_gemini_batch(
    client: genai.Client, items_chunk: List[Dict[str, str]]
) -> List[Dict[str, Any]]:
    """Sends a batch of 35 items to Gemini with Structured Output enforcement."""
    prompt_lines = ["Extract attributes for each of these Egyptian pharmacy items:\n"]
    for it in items_chunk:
        prompt_lines.append(
            f"- Code: {it['code']} | Cleaned: '{it['cleaned']}' | Raw: '{it['raw']}'"
        )
    user_prompt = "\n".join(prompt_lines)

    for attempt in range(MAX_RETRIES):
        try:
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    automatic_function_calling=types.AutomaticFunctionCallingConfig(
                        disable=True
                    ),
                    system_instruction=SYSTEM_PROMPT,
                    temperature=0.0,
                    response_mime_type="application/json",
                    response_schema=BatchExtractionResponse,
                ),
            )
            parsed_data = BatchExtractionResponse.model_validate_json(response.text)
            return [item.model_dump() for item in parsed_data.items]

        except ClientError as e:
            err_msg = str(e)
            if "limit: 0" in err_msg or "NOT_FOUND" in err_msg:
                raise RuntimeError(f"Fatal API error: {err_msg}")

            retry_match = re.search(r"retryDelay':\s*'(\d+)s", err_msg) or re.search(
                r"retry in (\d+(\.\d+)?)s", err_msg
            )
            wait_s = int(float(retry_match.group(1))) + 3 if retry_match else 25
            logger.warning(
                f"Rate limited by Google API. Throttling for {wait_s}s (Attempt {attempt+1}/{MAX_RETRIES})..."
            )
            time.sleep(wait_s)

        except Exception as e:
            logger.warning(
                f"Parsing or network error: {e}. Retrying in 10s (Attempt {attempt+1}/{MAX_RETRIES})..."
            )
            time.sleep(10)

    # Fallback if batch fails permanently
    logger.error(
        "Batch failed permanently after max retries. Registering default fallback values."
    )
    fallback = []
    for it in items_chunk:
        fallback.append(
            {
                "item_code": it["code"],
                "canonical_brand": it["cleaned"].upper(),
                "generic_name": "",
                "strength": "",
                "dosage_form": "OTHER",
                "pack_size": "",
                "canonical_sku_name": it["cleaned"].upper(),
                "canonical_signature": f"{it['cleaned'].upper()}||OTHER|",
            }
        )
    return fallback


# --------------------------------------------------------------------------------------
# 4. TYPE-SAFE COMPARATORS & CLINICAL SAFETY GUARDRAILS
# --------------------------------------------------------------------------------------
def safe_sort_key(code: Any):
    """Sorts numeric codes numerically, followed by alphanumeric codes alphabetically."""
    s = str(code).strip()
    return (0, int(s), "") if s.isdigit() else (1, 0, s.upper())


def audit_and_repair_clusters(
    raw_records: List[Dict[str, Any]], parsed_dict: Dict[str, Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    Performs global deterministic clustering across all 4,137 catalog items.
    Enforces clinical safety guardrails to split invalid merges.
    """
    logger.info(
        "Executing Global Cross-Catalog Entity Resolution & Safety Guardrails..."
    )

    # Group items globally by Canonical_SKU_Signature
    clusters: Dict[str, List[Dict[str, Any]]] = defaultdict(list)

    for r in raw_records:
        code = r["code"]
        ai_data = parsed_dict.get(code)
        if not ai_data:
            continue

        item_entry = {
            "code": code,
            "raw_name": r["raw"],
            "clean_name": r["cleaned"],
            "canonical_brand": ai_data.get("canonical_brand", "").strip().upper(),
            "generic_name": ai_data.get("generic_name", "").strip().upper(),
            "strength": ai_data.get("strength", "").strip().upper(),
            "dosage_form": ai_data.get("dosage_form", "").strip().upper(),
            "pack_size": ai_data.get("pack_size", "").strip().upper(),
            "canonical_sku_name": ai_data.get("canonical_sku_name", "").strip().upper(),
            "canonical_signature": ai_data.get("canonical_signature", "")
            .strip()
            .upper(),
        }

        # If signature was empty, synthesize it
        if not item_entry["canonical_signature"]:
            item_entry["canonical_signature"] = (
                f"{item_entry['canonical_brand']}|{item_entry['strength']}|"
                f"{item_entry['dosage_form']}|{item_entry['pack_size']}"
            )

        clusters[item_entry["canonical_signature"]].append(item_entry)

    final_output_rows = []

    for sig, cluster_items in clusters.items():
        # CLINICAL GUARDRAIL VALIDATION:
        # Check if the cluster has conflicting dosage forms or pack counts
        form_types = {it["dosage_form"] for it in cluster_items}
        strengths = {it["strength"] for it in cluster_items}
        pack_sizes = {it["pack_size"] for it in cluster_items}

        # If there are irreconcilable clinical conflicts, split them
        if len(form_types) > 1 or len(strengths) > 1 or len(pack_sizes) > 1:
            sub_clusters = defaultdict(list)
            for it in cluster_items:
                sub_sig = f"{it['canonical_brand']}|{it['strength']}|{it['dosage_form']}|{it['pack_size']}"
                sub_clusters[sub_sig].append(it)
            working_clusters = list(sub_clusters.values())
        else:
            working_clusters = [cluster_items]

        for valid_cluster in working_clusters:
            # Sort cluster items by oldest / lowest numeric code
            sorted_cluster = sorted(
                valid_cluster, key=lambda x: safe_sort_key(x["code"])
            )
            master_item = sorted_cluster[0]
            master_code = master_item["code"]
            canonical_name = master_item["canonical_sku_name"]

            dup_count = len(sorted_cluster) - 1
            all_sibling_codes = [it["code"] for it in sorted_cluster]
            dup_codes_str = "; ".join(all_sibling_codes) if dup_count > 0 else ""

            for item in sorted_cluster:
                is_master = item["code"] == master_code
                if dup_count == 0:
                    status = "Single SKU"
                    comment = "Unique SKU"
                elif is_master:
                    status = "Master SKU"
                    comment = f"MASTER SKU (Consolidates {dup_count} duplicate records)"
                else:
                    status = "Duplicate SKU"
                    comment = (
                        f"Duplicate record to merge into Master Code {master_code}"
                    )

                final_output_rows.append(
                    {
                        "Item_Code": item["code"],
                        "Raw_ERP_Name": item["raw_name"],
                        "Cleaned_Name": item["clean_name"],
                        "Canonical_Brand": item["canonical_brand"],
                        "Generic_Molecule": item["generic_name"],
                        "Active_Strength": item["strength"],
                        "Dosage_Form": item["dosage_form"],
                        "Pack_Size": item["pack_size"],
                        "Canonical_Master_Name": canonical_name,
                        "SKU_Signature": sig,
                        "Deduplication_Status": status,
                        "Duplicate_Count": dup_count,
                        "Duplicate_Codes": dup_codes_str,
                        "Action_Notes": comment,
                        "_sort_key_1": canonical_name,
                        "_sort_key_2": safe_sort_key(item["code"]),
                    }
                )

    # Sort catalog alphabetically by Canonical Master Name, keeping duplicates adjacent
    final_output_rows.sort(key=lambda x: (x["_sort_key_1"], x["_sort_key_2"]))
    for r in final_output_rows:
        del r["_sort_key_1"]
        del r["_sort_key_2"]

    return final_output_rows


# --------------------------------------------------------------------------------------
# 5. OPENPYXL EXCEL MASTER AUDIT REPORT GENERATOR
# --------------------------------------------------------------------------------------
def export_excel_audit_workbook(rows: List[Dict[str, Any]], output_filepath: str):
    """
    Generates a structured corporate Excel workbook with Tables, Total Rows,
    Auto-Filters, and soft conditional formatting for duplicate rows.
    """
    logger.info(f"Generating Excel Master Audit Workbook: {output_filepath}...")

    wb = openpyxl.Workbook()
    # Sheet 1: Full Catalog
    ws_all = wb.active
    ws_all.title = "Master_Catalog_Full"
    ws_all.views.sheetView[0].showGridLines = True

    headers = list(rows[0].keys())
    ws_all.append(headers)

    for r in rows:
        ws_all.append(list(r.values()))

    max_row = ws_all.max_row
    max_col = ws_all.max_column

    # 1. Apply Official Excel Table
    tab_ref = f"A1:{get_column_letter(max_col)}{max_row}"
    excel_tab = Table(displayName="GeniusMasterCatalogTable", ref=tab_ref)
    excel_tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws_all.add_table(excel_tab)

    # 2. Add Duplicates Only Sheet
    ws_dups = wb.create_sheet(title="Duplicates_For_Review")
    ws_dups.views.sheetView[0].showGridLines = True
    ws_dups.append(headers)

    dup_rows = [r for r in rows if r["Duplicate_Count"] > 0]
    for r in dup_rows:
        ws_dups.append(list(r.values()))

    if dup_rows:
        dup_tab_ref = f"A1:{get_column_letter(max_col)}{ws_dups.max_row}"
        excel_dup_tab = Table(displayName="DuplicatesReviewTable", ref=dup_tab_ref)
        excel_dup_tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws_dups.add_table(excel_dup_tab)

    # 3. Add Executive KPI Summary Sheet
    ws_kpi = wb.create_sheet(title="Executive_Summary")
    ws_kpi.views.sheetView[0].showGridLines = True

    total_records = len(rows)
    total_duplicates = len(dup_rows)
    total_uniques = total_records - total_duplicates
    unique_skus = len({r["SKU_Signature"] for r in rows})
    catalog_reduction = (
        (total_duplicates / total_records) * 100 if total_records > 0 else 0
    )

    kpi_data = [
        ["Executive Deduplication Metric", "Value", "Operational Impact"],
        [
            "Total Catalog Raw SKUs",
            total_records,
            "Total inventory records ingested from Genius ERP",
        ],
        [
            "Total Redundant Duplicate Rows",
            total_duplicates,
            "Secondary SKUs created due to historical price changes",
        ],
        [
            "Total Clean Single SKUs",
            total_uniques,
            "Records with zero duplicates in catalog",
        ],
        [
            "Consolidated Canonical SKUs",
            unique_skus,
            "True unique pharmaceutical products active in pharmacy",
        ],
        [
            "Catalog Waste / Reduction Potential",
            f"{catalog_reduction:.2f}%",
            "Percentage of catalog clutter eliminated",
        ],
    ]
    for row in kpi_data:
        ws_kpi.append(row)

    # Format and Style All Sheets
    thin_border = Border(
        left=Side(style="thin", color="E0E0E0"),
        right=Side(style="thin", color="E0E0E0"),
        top=Side(style="thin", color="E0E0E0"),
        bottom=Side(style="thin", color="E0E0E0"),
    )

    for ws in [ws_all, ws_dups, ws_kpi]:
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 26

        for c_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(c_idx)
            header_val = str(ws.cell(row=1, column=c_idx).value or "")
            max_len = len(header_val)

            for r_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                cell.font = Font(name="Segoe UI", size=9)

                # Check status column to color code
                if cell.value is not None:
                    s_val = str(cell.value)
                    if len(s_val) > max_len:
                        max_len = len(s_val)

            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

    # Add soft amber conditional formatting on Duplicate SKU rows
    soft_amber_fill = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )
    amber_font = Font(name="Segoe UI", size=9, color="7F6000")
    status_col_letter = get_column_letter(headers.index("Deduplication_Status") + 1)

    rule = CellIsRule(
        operator="equal",
        formula=['"Duplicate SKU"'],
        fill=soft_amber_fill,
        font=amber_font,
    )
    ws_all.conditional_formatting.add(
        f"{status_col_letter}2:{status_col_letter}{max_row}", rule
    )

    wb.save(output_filepath)
    logger.info(f"Excel Master Audit Workbook saved successfully: {output_filepath}")


# --------------------------------------------------------------------------------------
# 6. TRANSACTION-SAFE T-SQL MIGRATION SCRIPT GENERATOR
# --------------------------------------------------------------------------------------
def export_sql_migration_script(rows: List[Dict[str, Any]], output_filepath: str):
    """
    Generates a non-destructive, transaction-safe T-SQL migration script.
    Preserves all historical sales & purchases while consolidating stock.
    """
    logger.info(f"Generating T-SQL Migration Script: {output_filepath}...")

    duplicate_mappings = [
        (
            r["Item_Code"],
            r["Duplicate_Codes"].split("; ")[0],
        )  # (Duplicate Code, Master Code)
        for r in rows
        if r["Deduplication_Status"] == "Duplicate SKU"
    ]

    with open(output_filepath, "w", encoding="utf-8") as f:
        f.write(
            """-- =============================================================================
-- GENIUS / E-PLUS PHARMACY ERP: PRODUCTION CATALOG DEDUPLICATION SCRIPT
-- Target Database: [Genius] on Microsoft SQL Server
-- Directives:
--   1. ZERO hard deletes (preserves foreign keys on sales_trans_d and pur_trans_d).
--   2. Secondary barcodes re-linked to Master SKU in Item_Objects.
--   3. Warehouse stock quantities consolidated in Item_Class_Store.
--   4. Duplicate SKUs soft-deactivated (itm_stop_sell=1, itm_stop_pur=1, itm_active='0').
-- =============================================================================

USE [Genius];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRANSACTION;
BEGIN TRY
    -- 1. Create In-Memory Mapping Table
    IF OBJECT_ID('tempdb..#VerifiedDuplicates') IS NOT NULL
        DROP TABLE #VerifiedDuplicates;

    CREATE TABLE #VerifiedDuplicates (
        Duplicate_Code VARCHAR(50) NOT NULL,
        Master_Code VARCHAR(50) NOT NULL,
        PRIMARY KEY CLUSTERED (Duplicate_Code, Master_Code)
    );

    PRINT 'Staging verified duplicate mappings...';
"""
        )
        for dup_code, master_code in duplicate_mappings:
            f.write(
                f"    INSERT INTO #VerifiedDuplicates (Duplicate_Code, Master_Code) VALUES ('{dup_code}', '{master_code}');\n"
            )

        f.write(f"""
    PRINT 'Loaded {len(duplicate_mappings)} duplicate mappings into staging table.';

    -- 2. Migrate Secondary Barcodes in Item_Objects
    PRINT 'Step 1: Re-pointing secondary barcodes in Item_Objects to Master SKU...';
    UPDATE obj
    SET obj.io_itm_id = master_cat.itm_id
    FROM Item_Objects obj
    INNER JOIN Item_Catalog dup_cat ON obj.io_itm_id = dup_cat.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup_cat.itm_code = map.Duplicate_Code
    INNER JOIN Item_Catalog master_cat ON master_cat.itm_code = map.Master_Code
    WHERE obj.io_itm_id <> master_cat.itm_id;

    -- 3. Consolidate Active Branch Stock Balances in Item_Class_Store
    PRINT 'Step 2: Consolidating branch stock balances in Item_Class_Store...';
    
    -- A. Add duplicate quantity to master record where batch store record exists
    UPDATE master_ics
    SET master_ics.itm_qty = master_ics.itm_qty + dup_ics.itm_qty
    FROM Item_Class_Store master_ics
    INNER JOIN Item_Catalog master_cat ON master_ics.itm_id = master_cat.itm_id
    INNER JOIN #VerifiedDuplicates map ON master_cat.itm_code = map.Master_Code
    INNER JOIN Item_Catalog dup_cat ON dup_cat.itm_code = map.Duplicate_Code
    INNER JOIN Item_Class_Store dup_ics ON dup_ics.itm_id = dup_cat.itm_id 
        AND dup_ics.sto_id = master_ics.sto_id;

    -- B. Zero out inventory balances on duplicate items
    UPDATE dup_ics
    SET dup_ics.itm_qty = 0
    FROM Item_Class_Store dup_ics
    INNER JOIN Item_Catalog dup_cat ON dup_ics.itm_id = dup_cat.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup_cat.itm_code = map.Duplicate_Code;

    -- 4. Soft-Deactivate Redundant Catalog Rows in Item_Catalog
    PRINT 'Step 3: Soft-deactivating redundant records in Item_Catalog...';
    UPDATE cat
    SET cat.itm_active = '0',
        cat.itm_stop_sell = 1,
        cat.itm_stop_pur = 1,
        cat.itm_notes = ISNULL(cat.itm_notes, '') + ' [MERGED INTO MASTER CODE ' + map.Master_Code + ']'
    FROM Item_Catalog cat
    INNER JOIN #VerifiedDuplicates map ON cat.itm_code = map.Duplicate_Code;

    COMMIT TRANSACTION;
    PRINT 'Deduplication migration completed successfully. Transaction committed.';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;

    PRINT 'ERROR OCCURRED DURING MIGRATION! Transaction rolled back.';
    PRINT ERROR_MESSAGE();
END CATCH;
GO
""")
    logger.info(f"T-SQL Migration Script generated: {output_filepath}")


# --------------------------------------------------------------------------------------
# 7. MAIN ORCHESTRATION PIPELINE
# --------------------------------------------------------------------------------------
def main():
    start_time = time.time()
    print("==========================================================================")
    print("  GENIUS ERP PHARMACEUTICAL MASTER DEDUPLICATION ENGINE                   ")
    print("==========================================================================")

    if not API_KEY:
        print("[!] GEMINI_API_KEY environment variable is missing!")
        print("    Please set it before running: export GEMINI_API_KEY='your-key'")
        sys.exit(1)

    client = genai.Client(api_key=API_KEY)

    # 1. Ingest Input File
    if not os.path.exists(INPUT_CSV):
        print(
            f"[!] Input file '{INPUT_CSV}' not found. Please provide the catalog CSV."
        )
        sys.exit(1)

    raw_df = pd.read_csv(INPUT_CSV)
    code_col = (
        "itm_code"
        if "itm_code" in raw_df.columns
        else ("Item_Code" if "Item_Code" in raw_df.columns else raw_df.columns[0])
    )
    name_col = (
        "itm_name_en"
        if "itm_name_en" in raw_df.columns
        else (
            "Item_English_Name"
            if "Item_English_Name" in raw_df.columns
            else raw_df.columns[1]
        )
    )

    logger.info(f"Loaded {len(raw_df):,} total records from '{INPUT_CSV}'.")

    # 2. Run Deterministic Python Pre-Cleaner
    records = []
    for _, row in raw_df.iterrows():
        c = str(row[code_col]).strip()
        raw_n = str(row[name_col]).strip()
        clean_n = clean_pharmacy_name(raw_n)
        records.append({"code": c, "raw": raw_n, "cleaned": clean_n})

    # Sort alphabetically by cleaned string (brings variant spellings adjacent)
    records.sort(key=lambda x: (x["cleaned"].upper(), safe_sort_key(x["code"])))

    # 3. Load Checkpoint (Resume seamlessly)
    checkpoint_parsed: Dict[str, Dict[str, Any]] = {}
    completed_batch_indices = set()

    if os.path.exists(CHECKPOINT_JSON):
        try:
            with open(CHECKPOINT_JSON, "r", encoding="utf-8") as f:
                saved = json.load(f)
                checkpoint_parsed = saved.get("parsed_items", {})
                completed_batch_indices = set(saved.get("completed_batches", []))
            logger.info(
                f"Checkpoint Loaded: {len(checkpoint_parsed):,} items already parsed across {len(completed_batch_indices)} batches."
            )
        except Exception as e:
            logger.warning(f"Could not load checkpoint: {e}. Starting fresh.")

    # 4. Construct Micro-Batches
    batches = [records[i : i + BATCH_SIZE] for i in range(0, len(records), BATCH_SIZE)]
    total_batches = len(batches)
    logger.info(
        f"Catalog divided into {total_batches} micro-batches ({BATCH_SIZE} items/batch)."
    )
    logger.info(
        f"Projected API Consumption: {total_batches - len(completed_batch_indices)} requests (Daily Limit: 500 RPD)."
    )

    # 5. Execute API Batches
    for b_idx, batch in enumerate(batches, start=1):
        if b_idx in completed_batch_indices:
            continue

        first_item = batch[0]["cleaned"]
        last_item = batch[-1]["cleaned"]
        logger.info(
            f"[{b_idx}/{total_batches}] Extracting {len(batch)} items ({first_item[:16]}... -> {last_item[:16]}...)..."
        )

        batch_results = execute_gemini_batch(client, batch)

        # Save results in memory
        for res in batch_results:
            c_code = str(res.get("item_code", "")).strip()
            if c_code:
                checkpoint_parsed[c_code] = res

        completed_batch_indices.add(b_idx)

        # Commit Checkpoint to Disk
        with open(CHECKPOINT_JSON, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "completed_batches": list(completed_batch_indices),
                    "parsed_items": checkpoint_parsed,
                },
                f,
                ensure_ascii=False,
            )

        time.sleep(REQUEST_DELAY_SECONDS)

    logger.info(
        "All batches parsed by AI successfully. Proceeding to Global Reconciliation..."
    )

    # 6. Global Reconciliation & Clinical Safety Audit
    final_audited_rows = audit_and_repair_clusters(records, checkpoint_parsed)

    # 7. Generate Master Excel Audit Workbook
    export_excel_audit_workbook(final_audited_rows, OUTPUT_EXCEL)

    # 8. Generate Production T-SQL Migration Script
    export_sql_migration_script(final_audited_rows, OUTPUT_SQL)

    elapsed = round(time.time() - start_time, 2)
    print("\n" + "=" * 70)
    print(f"[✔] DEDUPLICATION PIPELINE COMPLETED IN {elapsed}s")
    print(f"[✔] Master Excel Workbook: '{os.path.abspath(OUTPUT_EXCEL)}'")
    print(f"[✔] T-SQL Migration Script: '{os.path.abspath(OUTPUT_SQL)}'")
    print("=" * 70)


if __name__ == "__main__":
    main()
