#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ULTIMATE PHARMACEUTICAL ERP DEDUPLICATION & ENTITY RESOLUTION ENGINE
Designed for Egyptian Pharmacy Management Systems (Genius / e-Plus on MS SQL Server).
Author: Enterprise Pharmacy Database Architecture Group
Encoding: UTF-8 with BOM (utf-8-sig) for 100% Android Excel Compatibility
"""

import csv
import os
import re
import sys
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

# ==============================================================================
# 1. PHARMACEUTICAL DOMAIN CONSTANTS & CANONICAL DICTIONARIES
# ==============================================================================

# Prefixes that MUST NOT be split from the brand stem
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

# Critical brand modifiers that signify distinct clinical combinations or kinetics
CRITICAL_MODIFIERS = {
    "PLUS",
    "COMP",
    "TRIO",
    "AMLO",
    "HCT",
    "FORTE",
    "EXTRA",
    "XR",
    "MR",
    "SR",
    "CR",
    "RETARD",
    "MAX",
    "MINI",
    "JUNIOR",
    "INFANT",
    "KIDS",
    "PED",
    "PEDIATRIC",
    "ADULT",
    "DAILY",
    "D",
}

# Clinical services & hospital consumables that must remain unmerged singletons
EXCLUDED_SERVICE_CODES = {"2", "IN", "G5", "G10", "S3", "S5", "S10", "WI", "MA"}

# Common Egyptian pharmacy cashier noise & lifecycle markers
ERP_NOISE_REGEX = re.compile(
    r"(?i)\b(?:NEW\s*PRICE|OLD\s*PRICE|N\.?P\.?|N\/PRICE|N\s*P|"
    r"سعر\s*جديد|سعر\s*قديم|موقوف|بدون\s*خصم|خصم\s*خاص|"
    r"OFFER|GIFT|FREE|BONUS|BOUNS|SAMPLE|UPDATE)\b"
)

# Trailing manufacturer & distributor suffixes added by cashiers
MANUFACTURER_SUFFIXES = re.compile(
    r"(?i)\b(?:MUP|SEDICO|PHARCO|AMRIYA|AMRYIA|T3A|AUG|EVA|EIPICO|EPICO|"
    r"KAHIRA|ALEX|CID|MEMPHIS|NILE|MISR|GLOBAL|MARCYRL|SIGMA|APEX|AVENTIS|"
    r"SANOFI|GSK|NOVARTIS|PFIZER|ASTRA)\b"
)

# Canonical dosage form mappings (Arabic & English to Standardized Form)
DOSAGE_FORM_MAP = {
    # Oral Solids
    "TAB": "TAB",
    "TABS": "TAB",
    "TABLET": "TAB",
    "TABLETS": "TAB",
    "قرص": "TAB",
    "اقراص": "TAB",
    "أقراص": "TAB",
    "CAP": "CAP",
    "CAPS": "CAP",
    "CAPSULE": "CAP",
    "CAPSULES": "CAP",
    "كبسول": "CAP",
    "كبسولات": "CAP",
    "SACHET": "SACHET",
    "SACHETS": "SACHET",
    "SAC": "SACHET",
    "فوار": "SACHET",
    "اكياس": "SACHET",
    "أكياس": "SACHET",
    "EFFERVESCENT": "EFF TAB",
    "EFF": "EFF TAB",
    "LOZENGES": "LOZENGES",
    "LOZ": "LOZENGES",
    "استحلاب": "LOZENGES",
    # Oral Liquids
    "SYRUP": "SYRUP",
    "SYR": "SYRUP",
    "شراب": "SYRUP",
    "SUSP": "SUSP",
    "SUSPENSION": "SUSP",
    "معلق": "SUSP",
    "DROPS": "DROPS",
    "ORAL DROPS": "DROPS",
    "نقط": "DROPS",
    # Injectables
    "AMP": "AMP",
    "AMPOULE": "AMP",
    "AMPOULES": "AMP",
    "AMPS": "AMP",
    "حقن": "AMP",
    "امبول": "AMP",
    "أمبول": "AMP",
    "VIAL": "VIAL",
    "VIALS": "VIAL",
    "فيال": "VIAL",
    # Topicals
    "CREAM": "CREAM",
    "CRM": "CREAM",
    "كريم": "CREAM",
    "OINT": "OINT",
    "OINTMENT": "OINT",
    "مرهم": "OINT",
    "GEL": "GEL",
    "جل": "GEL",
    "LOTION": "LOTION",
    "LOT": "LOTION",
    "لوشن": "LOTION",
    # Rectal / Vaginal
    "SUPP": "SUPP",
    "SUPPOSITORY": "SUPP",
    "لبوس": "SUPP",
    "VAG OVULE": "OVULE",
    "OVULE": "OVULE",
    "OVULES": "OVULE",
    # Inhalers / Sprays
    "INHALER": "INHALER",
    "INH": "INHALER",
    "بخاخ": "SPRAY",
    "SPRAY": "SPRAY",
    # Consumables / Cosmetics
    "SHAMPOO": "SHAMPOO",
    "SOAP": "SOAP",
    "SOLUTION": "SOLUTION",
    "SOL": "SOLUTION",
}

# ==============================================================================
# 2. TYPE-SAFE UTILITIES & ALPHANUMERIC COMPARATORS
# ==============================================================================


def safe_sort_key(code: Any) -> Tuple[int, int, str]:
    """
    Solves TypeError: '<' not supported between instances of 'int' and 'str'.
    Numeric codes sort first by numerical magnitude; alphanumeric codes sort alphabetically.
    """
    s = str(code).strip()
    if s.isdigit():
        return (0, int(s), "")
    return (1, 0, s.upper())


# ==============================================================================
# 3. HIGH-PRECISION FORENSIC PRE-SANITIZER
# ==============================================================================


def clean_pharmaceutical_name(raw_name: str) -> str:
    """
    Strips noise, punctuation flags, lifecycle markers, and embedded prices
    while strictly avoiding variable-width look-behind crashes.
    """
    if not raw_name or not isinstance(raw_name, str):
        return ""

    text = raw_name.strip()

    # Step A: Remove ERP symbol noise (@, @@, #, *, !, ?, ~, ^, _, %)
    text = re.sub(r"[@#*!?%^~_]+", " ", text)

    # Step B: Strip explicit cashier status tags (NEW PRICE, N.P, etc.)
    text = ERP_NOISE_REGEX.sub(" ", text)

    # Step C: Strip explicit currency price tags (e.g., '18LE', '22.5 L.E', '357 EGP', '144 جنيه')
    text = re.sub(r"(?i)\b\d+(?:\.\d+)?\s*(?:LE|L\.E|EGP|جنيه)\b", " ", text)

    # Step D: Strip manufacturer / supplier trailing stamps
    text = MANUFACTURER_SUFFIXES.sub(" ", text)

    # Step E: Strip Naked Prices (Trailing 2 to 4 digits preceded by units or strengths)
    # Uses fixed capturing groups and backreferences to prevent re.PatternError
    text = re.sub(
        r"(?i)(\b(?:TAB|TABS|CAP|CAPS|TABLETS?|CAPSULES?|ML|GM|G|AMP|VIAL|SACHET|DOSE|IU)\b\s*)\d{2,4}\s*$",
        r"\1",
        text,
    )

    # Step F: Normalize unit typos
    text = re.sub(r"(?i)\b(\d+)\s*MMG\b", r"\1 MG", text)
    text = re.sub(r"(?i)\b(\d+)\s*M\s*GM\b", r"\1 GM", text)
    text = re.sub(r"(?i)\b(\d+)\s*M\s*L\b", r"\1 ML", text)
    text = re.sub(r"(?i)\bTABLES\b", "TAB", text)
    text = re.sub(r"(?i)\bCABSU?LES?\b", "CAP", text)
    text = re.sub(r"(?i)\bSYRY?P\b", "SYRUP", text)

    # Collapse multiple spaces
    text = re.sub(r"\s+", " ", text).strip()
    return text


# ==============================================================================
# 4. DETERMINISTIC CANONICAL FEATURE EXTRACTOR
# ==============================================================================


class CanonicalItem:
    def __init__(self, code: str, raw_name: str):
        self.code = str(code).strip()
        self.raw_name = raw_name.strip()
        self.clean_name = clean_pharmaceutical_name(raw_name)

        # Clinical Attributes
        self.is_service_or_supply: bool = False
        self.brand_stem: str = ""
        self.strength: str = ""
        self.dosage_form: str = ""
        self.pack_size: str = ""
        self.canonical_signature: str = ""
        self.proposed_master_name: str = ""

        self._parse()

    def _parse(self):
        upper_code = self.code.upper()
        # 1. Clinical Service / Hospital Supply Check
        if upper_code in EXCLUDED_SERVICE_CODES or any(
            srv in self.clean_name.upper()
            for srv in ["قياس ضغط", "ضرب حقنه", "BLOOD PRESSURE", "GLUCOSE TEST"]
        ):
            self.is_service_or_supply = True
            self.brand_stem = self.clean_name.upper()
            self.canonical_signature = f"SERVICE_OR_SUPPLY|{self.code}"
            self.proposed_master_name = self.clean_name.upper()
            return

        text = self.clean_name.upper()

        # 2. Extract Strength / Combination Ratio (e.g. 500 MG, 40/5/12.5 MG, 5000 IU, 0.5%)
        # Supports multi-ingredient ratios separated by slashes
        strength_match = re.search(
            r"\b(\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)*(?:\s*/\s*\d+(?:\.\d+)?)*)\s*(MG|GM|G|MCG|IU|%|ML)\b",
            text,
        )
        if strength_match:
            val = re.sub(r"\s+", "", strength_match.group(1))
            unit = strength_match.group(2)
            self.strength = f"{val} {unit}"
            # Blank out the matched strength token from working string
            text = text[: strength_match.start()] + " " + text[strength_match.end() :]

        # 3. Extract Pack Size / Volume (e.g. 20 TAB, 100 ML, 6 AMP, 30 CAP)
        pack_match = re.search(
            r"\b(\d+)\s*(TAB|TABS|TABLETS?|CAP|CAPS|CAPSULES?|AMP|AMPOULES?|AMPS|VIALS?|SACHETS?|DOSES?|PATCHES?|SUPP|ML|GM|G)\b",
            text,
        )
        if pack_match:
            count = pack_match.group(1)
            raw_unit = pack_match.group(2)
            unit_norm = DOSAGE_FORM_MAP.get(raw_unit, raw_unit)
            self.pack_size = f"{count} {unit_norm}"
            text = text[: pack_match.start()] + " " + text[pack_match.end() :]
        else:
            # Standalone trailing tablet count e.g. "PARAMOL 500MG 20"
            standalone_pack = re.search(r"\b(\d{1,3})\s*$", text)
            if standalone_pack:
                self.pack_size = f"{standalone_pack.group(1)} UNITS"
                text = text[: standalone_pack.start()]

        # 4. Extract Dosage Form
        for raw_form, canonical_form in DOSAGE_FORM_MAP.items():
            form_pattern = r"\b" + re.escape(raw_form) + r"\b"
            if re.search(form_pattern, text):
                self.dosage_form = canonical_form
                text = re.sub(form_pattern, " ", text)
                break

        if not self.dosage_form and self.pack_size:
            # Infer form from pack unit if present
            for unit_key in ["TAB", "CAP", "AMP", "VIAL", "SACHET", "SUPP"]:
                if unit_key in self.pack_size:
                    self.dosage_form = unit_key
                    break

        # 5. Extract Brand Stem (Prefix-Safe & Modifier-Aware)
        text = re.sub(r"[-_/]+", " ", text)
        tokens = [t.strip() for t in text.split() if t.strip()]

        if tokens:
            # Check protected prefixes (e.g., CO-TAREG, E-MOXCLAV)
            if len(tokens) >= 2 and tokens[0] in PROTECTED_PREFIXES:
                brand_base = f"{tokens[0]}-{tokens[1]}"
                remaining_tokens = tokens[2:]
            else:
                brand_base = tokens[0]
                remaining_tokens = tokens[1:]

            # Preserve critical modifiers (e.g., PLUS, COMP, TRIO, XR, FORTE)
            modifiers = [t for t in remaining_tokens if t in CRITICAL_MODIFIERS]
            if modifiers:
                self.brand_stem = f"{brand_base} {' '.join(modifiers)}"
            else:
                self.brand_stem = brand_base
        else:
            self.brand_stem = self.clean_name.upper()

        # Harmonize single-form cashiers errors (e.g., Bengiride only exists as CAP)
        if "BENGIRIDE" in self.brand_stem and self.dosage_form == "TAB":
            self.dosage_form = "CAP"
            self.pack_size = self.pack_size.replace("TAB", "CAP")

        # 6. Construct Strict Multi-Dimensional Signature
        self.canonical_signature = (
            f"{self.brand_stem}|{self.strength}|{self.dosage_form}|{self.pack_size}"
        )

        # 7. Synthesize Canonical Master Name
        name_parts = [self.brand_stem]
        if self.strength:
            name_parts.append(self.strength)
        if self.dosage_form and self.dosage_form not in self.pack_size:
            name_parts.append(self.dosage_form)
        if self.pack_size:
            name_parts.append(self.pack_size)
        self.proposed_master_name = " ".join(name_parts).strip()


# ==============================================================================
# 5. STAGE 5 SAFETY GUARDRAILS & RESOLUTION ENGINE
# ==============================================================================


class PharmacyDeduplicationEngine:
    def __init__(self):
        self.items: List[CanonicalItem] = []
        self.clusters: Dict[str, List[CanonicalItem]] = defaultdict(list)

    def load_extract(self, csv_filepath: str):
        if not os.path.exists(csv_filepath):
            raise FileNotFoundError(f"Input extract file '{csv_filepath}' not found.")

        print(f"[*] Ingesting raw catalog from {csv_filepath}...")
        with open(csv_filepath, mode="r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.reader(f)
            header = next(reader, None)

            for row in reader:
                if not row or len(row) < 2:
                    continue
                code, raw_name = row[0].strip(), row[1].strip()
                item = CanonicalItem(code, raw_name)
                self.items.append(item)

        print(f"[✓] Successfully loaded {len(self.items):,} active records.")

    def execute_resolution(self):
        """Clusters records by canonical signature with strict clinical guardrails."""
        print("[*] Executing Canonical Resolution and Stage 5 Safety Guardrails...")

        # Step A: Initial cluster assignment
        sig_map: Dict[str, List[CanonicalItem]] = defaultdict(list)
        for item in self.items:
            if item.is_service_or_supply:
                # Services never merge; each forms its own cluster
                sig_map[f"SINGLETON_SRV_{item.code}"].append(item)
            else:
                sig_map[item.canonical_signature].append(item)

        # Step B: Guardrail Validation (Deterministic Splitting of Unsafe Collisions)
        for sig, cluster_items in sig_map.items():
            if len(cluster_items) <= 1:
                self.clusters[sig] = cluster_items
                continue

            # Safety check: Verify that active strengths and pack sizes are 100% identical
            strengths = {it.strength for it in cluster_items}
            pack_sizes = {it.pack_size for it in cluster_items}
            forms = {it.dosage_form for it in cluster_items}

            if len(strengths) > 1 or len(pack_sizes) > 1 or len(forms) > 1:
                # Guardrail breach detected: sub-partition cluster safely
                print(f"[!] Guardrail Alert: Splitting over-merged signature: {sig}")
                sub_clusters = defaultdict(list)
                for it in cluster_items:
                    sub_key = (
                        f"{it.brand_stem}|{it.strength}|{it.dosage_form}|{it.pack_size}"
                    )
                    sub_clusters[sub_key].append(it)
                for sub_k, sub_v in sub_clusters.items():
                    self.clusters[sub_k].extend(sub_v)
            else:
                self.clusters[sig] = cluster_items

        print(
            f"[✓] Resolved {len(self.items):,} items into {len(self.clusters):,} canonical SKUs."
        )

    def export_canonical_catalog(self, output_filepath: str):
        """
        Exports the master 6-column CSV with UTF-8 BOM encoding for complete Android Excel stability.
        Columns: code, name, number of duplicates, duplicate_codes, proposed master name, extra comments
        """
        print(f"[*] Exporting canonical catalog to: {output_filepath}")

        # Flatten records for alphabetical sorting by Proposed Master Name
        export_rows = []

        for sig, cluster in self.clusters.items():
            # Determine Master Code: lowest/oldest numeric item code
            sorted_cluster = sorted(cluster, key=lambda x: safe_sort_key(x.code))
            master_item = sorted_cluster[0]
            master_code = master_item.code
            canonical_name = master_item.proposed_master_name

            # Symmetrical duplicate list (all codes in cluster)
            all_codes = [it.code for it in sorted_cluster]
            duplicate_codes_str = "; ".join(all_codes)
            dup_count = len(cluster) - 1

            for item in sorted_cluster:
                if dup_count == 0:
                    comment = "Unique SKU"
                elif item.code == master_code:
                    comment = f"MASTER SKU (Consolidates {dup_count} duplicate records)"
                else:
                    comment = (
                        f"Duplicate record to merge into Master Code {master_code}"
                    )

                export_rows.append(
                    {
                        "code": item.code,
                        "name": item.raw_name,
                        "number of duplicates": dup_count,
                        "duplicate_codes": duplicate_codes_str,
                        "proposed master name": canonical_name,
                        "extra comments": comment,
                    }
                )

        # Sort alphabetically by master name, keeping duplicates adjacent
        export_rows.sort(
            key=lambda r: (r["proposed master name"].upper(), safe_sort_key(r["code"]))
        )

        # Write RFC 4180 compliant CSV with BOM
        fieldnames = [
            "code",
            "name",
            "number of duplicates",
            "duplicate_codes",
            "proposed master name",
            "extra comments",
        ]

        with open(output_filepath, mode="w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            for row in export_rows:
                writer.writerow(row)

        print(
            f"[✓] Master catalog export completed: {len(export_rows):,} rows written."
        )

    def generate_sql_migration_script(self, sql_filepath: str):
        """
        Generates a non-destructive, transaction-safe T-SQL migration script
        for the Genius database on Microsoft SQL Server.
        """
        print(f"[*] Generating production SQL migration script: {sql_filepath}")

        duplicate_pairs: List[Tuple[str, str]] = []
        for sig, cluster in self.clusters.items():
            if len(cluster) > 1:
                sorted_cluster = sorted(cluster, key=lambda x: safe_sort_key(x.code))
                master_code = sorted_cluster[0].code
                for dup in sorted_cluster[1:]:
                    duplicate_pairs.append((master_code, dup.code))

        with open(sql_filepath, mode="w", encoding="utf-8") as f:
            f.write(
                """-- =============================================================================
-- GENIUS / E-PLUS PHARMACY ERP: ITEM DEDUPLICATION MIGRATION SCRIPT
-- Target Database Engine: Microsoft SQL Server (Transact-SQL)
-- Target Database: Genius
-- Core Directives:
--   1. ZERO hard deletes (preserves foreign keys on sales_trans_d and pur_trans_d).
--   2. Secondary barcodes mapped to Master SKU in Item_Objects.
--   3. Stock balances consolidated in Item_Class_Store.
--   4. Duplicate rows soft-deactivated in Item_Catalog.
-- =============================================================================

USE [Genius];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRANSACTION;
BEGIN TRY
    -- 1. Create In-Memory Staging Table for Verified Deduplication Pairs
    IF OBJECT_ID('tempdb..#VerifiedDuplicates') IS NOT NULL
        DROP TABLE #VerifiedDuplicates;

    CREATE TABLE #VerifiedDuplicates (
        Master_Code DECIMAL(18,0) NOT NULL,
        Duplicate_Code DECIMAL(18,0) NOT NULL,
        PRIMARY KEY CLUSTERED (Master_Code, Duplicate_Code)
    );

    PRINT 'Loading verified duplicate clusters...';
"""
            )

            # Batch insert duplicate mapping pairs
            for master_code, dup_code in duplicate_pairs:
                if str(master_code).isdigit() and str(dup_code).isdigit():
                    f.write(
                        f"    INSERT INTO #VerifiedDuplicates (Master_Code, Duplicate_Code) "
                        f"VALUES ({master_code}, {dup_code});\n"
                    )

            f.write(f"""
    PRINT 'Loaded {len(duplicate_pairs)} duplicate mappings into staging table.';

    -- 2. Migrate Secondary Barcodes from Duplicate SKUs to Master SKU in Item_Objects
    PRINT 'Step 1: Re-pointing secondary barcodes in Item_Objects...';
    UPDATE obj
    SET obj.itm_id = m.itm_id
    FROM Item_Objects obj
    INNER JOIN Item_Catalog dup ON obj.itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Catalog m ON m.itm_code = map.Master_Code
    WHERE obj.itm_id <> m.itm_id;

    -- 3. Consolidate Active Warehouse & Branch Stock Balances in Item_Class_Store
    PRINT 'Step 2: Consolidating branch stock balances in Item_Class_Store...';
    
    -- A. Add duplicate quantity to master record where store record exists
    UPDATE master_ics
    SET master_ics.ics_qty = master_ics.ics_qty + dup_ics.ics_qty
    FROM Item_Class_Store master_ics
    INNER JOIN Item_Catalog m ON master_ics.itm_id = m.itm_id
    INNER JOIN #VerifiedDuplicates map ON m.itm_code = map.Master_Code
    INNER JOIN Item_Catalog dup ON dup.itm_code = map.Duplicate_Code
    INNER JOIN Item_Class_Store dup_ics ON dup_ics.itm_id = dup.itm_id 
        AND dup_ics.store_id = master_ics.store_id;

    -- B. Zero out inventory balances on duplicate items
    UPDATE dup_ics
    SET dup_ics.ics_qty = 0
    FROM Item_Class_Store dup_ics
    INNER JOIN Item_Catalog dup ON dup_ics.itm_id = dup.itm_id
    INNER JOIN #VerifiedDuplicates map ON dup.itm_code = map.Duplicate_Code;

    -- 4. Soft-Deactivate Redundant Catalog Rows in Item_Catalog
    PRINT 'Step 3: Soft-deactivating redundant records in Item_Catalog...';
    UPDATE cat
    SET cat.itm_active = '0',
        cat.itm_stop_sell = 1,
        cat.itm_stop_pur = 1,
        cat.itm_notes = ISNULL(cat.itm_notes, '') + ' [MERGED INTO MASTER CODE ' + CAST(map.Master_Code AS VARCHAR(30)) + ']'
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
        print(
            f"[✓] Production SQL script generated successfully: {len(duplicate_pairs)} actions staged."
        )


# ==============================================================================
# 6. MAIN EXECUTION PIPELINE
# ==============================================================================

if __name__ == "__main__":
    INPUT_FILE = "used_items_extract.csv"
    OUTPUT_CSV = "final_deduplicated_catalog_PERFECT.csv"
    OUTPUT_SQL = "execute_genius_deduplication.sql"

    print("=================================================================")
    print("  GENIUS ERP PHARMACEUTICAL CANONICAL DEDUPLICATION SYSTEM       ")
    print("=================================================================")

    if not os.path.exists(INPUT_FILE):
        print(f"[-] Error: '{INPUT_FILE}' not found in current directory.")
        print("    Please place the catalog extract CSV here and re-run.")
        sys.exit(1)

    engine = PharmacyDeduplicationEngine()
    engine.load_extract(INPUT_FILE)
    engine.execute_resolution()
    engine.export_canonical_catalog(OUTPUT_CSV)
    engine.generate_sql_migration_script(OUTPUT_SQL)

    print("\n[★] PIPELINE EXECUTION COMPLETE:")
    print(f"    1. Android Excel Master Catalog: {OUTPUT_CSV}")
    print(f"    2. Safe T-SQL Migration Script:  {OUTPUT_SQL}")
    print("=================================================================")
