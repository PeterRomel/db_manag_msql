#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
========================================================================================
GENIUS (e-Plus) PHARMACY ERP - MASTER USED ITEMS CATALOG EXTRACTION ENGINE
========================================================================================
Outputs:
  1. 01_Genius_Master_Used_Items_Catalog.xlsx -> Complete catalog with all DB columns,
                                                 enriched attributes, last action & stores.
  2. By_Letter/Items_Letter_<X>.csv          -> Segmented CSVs per first letter (Code + EN Name).
  3. All_Used_Items.csv                      -> Master CSV of all used items (Code + EN Name).
========================================================================================
"""

import os
import sys
import time
import logging
import urllib.parse
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine, text
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------------------
# 0. LOGGING & SYSTEM SETUP
# --------------------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("GeniusCatalogExtractor")

# --------------------------------------------------------------------------------------
# 1. DATABASE CONNECTION SETUP
# --------------------------------------------------------------------------------------
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

conn_str = (
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USERNAME};"
    f"PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
params = urllib.parse.quote_plus(conn_str)
ENGINE = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", pool_pre_ping=True)

# --------------------------------------------------------------------------------------
# 2. PROPRIETARY BINARY DECRYPTION ENGINE
# --------------------------------------------------------------------------------------
def decode_legacy_bidi_bytes(raw_bytes):
    """
    Decodes obfuscated varbinary medicine trade names in Item_Catalog.
    Splits binary in half (N // 2), reverses each chunk independently,
    concatenates back together, and decodes with windows-1256.
    """
    if pd.isna(raw_bytes) or not raw_bytes:
        return ""
    try:
        raw_b = bytes(raw_bytes)
        N = len(raw_b)
        L = N // 2
        
        left_chunk = raw_b[:L]
        right_chunk = raw_b[L:]
        
        corrected_bytes = left_chunk[::-1] + right_chunk[::-1]
        return corrected_bytes.decode('windows-1256', errors='replace').strip()
    except Exception:
        return ""

# --------------------------------------------------------------------------------------
# 3. EXCEL STYLING HELPER
# --------------------------------------------------------------------------------------
THIN_BORDER = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'),
    bottom=Side(style='thin', color='D3D3D3')
)

def style_master_catalog_sheet(ws, header_fill_hex="1B365D"):
    """
    Applies executive formatting:
      - Custom Header Navy Fill & White Bold Font
      - Freeze top row
      - Standardized borders & alignments
      - Accounting & Date format detection
      - Auto-fitted column widths based on a representative sample
    """
    header_fill = PatternFill(start_color=header_fill_hex, end_color=header_fill_hex, fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "C2"  # Freeze code and name columns

    # Header Row
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 32

    max_row = ws.max_row
    max_col = ws.max_column

    col_meta = {}
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(ws.cell(row=1, column=col_idx).value or "")
        h_lower = header_val.lower()
        
        is_money = any(t in h_lower for t in ["price", "cost", "val", "margin", "tax"])
        is_perc = any(t in h_lower for t in ["perc", "rate", "%"])
        is_qty = any(t in h_lower for t in ["qty", "stock", "count", "factor", "limit"])
        
        col_meta[col_idx] = {
            'letter': col_letter,
            'max_len': min(max(len(header_val), 12), 45),
            'is_money': is_money,
            'is_perc': is_perc,
            'is_qty': is_qty
        }

    # Data Rows
    for row in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            col_idx = cell.column
            meta = col_meta[col_idx]
            val = cell.value
            cell.font = data_font
            cell.border = THIN_BORDER

            if val is not None:
                str_val = str(val)
                if len(str_val) > meta['max_len'] and cell.row <= 500:
                    meta['max_len'] = min(len(str_val), 50)

                if isinstance(val, (int, float)):
                    if meta['is_money']:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif meta['is_perc']:
                        cell.number_format = "0.00%"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif meta['is_qty']:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                elif isinstance(val, (datetime, pd.Timestamp)):
                    cell.number_format = "YYYY-MM-DD HH:MM"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

    for col_idx in range(1, max_col + 1):
        meta = col_meta[col_idx]
        ws.column_dimensions[meta['letter']].width = meta['max_len'] + 3

# --------------------------------------------------------------------------------------
# 4. DATA EXTRACTION & AGGREGATION PIPELINE
# --------------------------------------------------------------------------------------
def extract_master_used_catalog():
    start_time = time.time()
    logger.info("==========================================================================")
    logger.info("PHASE 1: IDENTIFYING ALL HISTORICALLY USED ITEMS ACROSS THE DATABASE")
    logger.info("==========================================================================")

    # 1. Identify all used items
    q_used_ids = """
    SELECT DISTINCT itm_id FROM (
        SELECT itm_id FROM sales_trans_d WITH (NOLOCK) WHERE itm_id IS NOT NULL AND itm_id > 0
        UNION
        SELECT itm_id FROM pur_trans_d WITH (NOLOCK) WHERE itm_id IS NOT NULL AND itm_id > 0
        UNION
        SELECT itm_id FROM Item_Class_Store WITH (NOLOCK) WHERE itm_id IS NOT NULL AND itm_id > 0 AND itm_qty <> 0
        UNION
        SELECT itm_id FROM Stock_Open_d WITH (NOLOCK) WHERE itm_id IS NOT NULL AND itm_id > 0
        UNION
        SELECT st_itm_id AS itm_id FROM Store_Trans WITH (NOLOCK) WHERE st_itm_id IS NOT NULL AND st_itm_id > 0
        UNION
        SELECT invd_itm_id AS itm_id FROM Inventory_d WITH (NOLOCK) WHERE invd_itm_id IS NOT NULL AND invd_itm_id > 0
        UNION
        SELECT iqc_itm_id AS itm_id FROM Item_Quantity_Changes WITH (NOLOCK) WHERE iqc_itm_id IS NOT NULL AND iqc_itm_id > 0
        UNION
        SELECT ibd_itm_id AS itm_id FROM Item_back_d WITH (NOLOCK) WHERE ibd_itm_id IS NOT NULL AND ibd_itm_id > 0
    ) u;
    """
    df_used_ids = pd.read_sql(q_used_ids, ENGINE)
    used_ids_set = set(df_used_ids['itm_id'].dropna().astype(float).unique())
    logger.info(f"Identified {len(used_ids_set):,} items with transactional/inventory history.")

    # 2. Extract Base Item Catalog
    logger.info("Extracting core catalog attributes and decrypted trade names...")
    q_catalog = """
    SELECT 
        ic.itm_id,
        ic.itm_code,
        ic.itm_code2,
        ic.itm_int_code,
        ic.itm_usr_code,
        ic.ucp_code,
        ic.itm_name_ar,
        ic.itm_name_en,
        ic.itm_name_ar_encrypt,
        ic.itm_name_en_encrypt,
        c.com_name_ar,
        c.com_name_en,
        itf.itf_name_ar,
        sg.sg_name_ar,
        ic.itm_scientific_n1,
        ic.itm_scientific_n2,
        ic.itm_effictive,
        ic.itm_effictive_perc,
        ium.ium_name_ar,
        g.g_name_ar,
        io.io_text_ar,
        CASE ic.itm_ismedicine WHEN 1 THEN 'Medicine (دواء)' ELSE 'Non-Medicine (مستلزمات/تجميل)' END AS itm_ismedicine_desc,
        CASE ic.itm_freez WHEN 1 THEN 'Yes - Refrigerator (ثلاجة)' ELSE 'No (عادي)' END AS itm_freez_desc,
        CASE ic.itm_isprev WHEN 1 THEN 'Yes - Prescription (روشتة)' ELSE 'No - OTC (بدون روشتة)' END AS itm_isprev_desc,
        ic.itm_def_sell_price,
        ic.itm_def_pharm_price,
        (ic.itm_def_sell_price - ic.itm_def_pharm_price) AS unit_profit_margin,
        CASE WHEN ic.itm_def_sell_price > 0 
             THEN ((ic.itm_def_sell_price - ic.itm_def_pharm_price) / ic.itm_def_sell_price) 
             ELSE 0 END AS profit_margin_perc,
        ic.itm_def_tax,
        ic.Tax_Category,
        CASE ic.itm_allow_discount WHEN 1 THEN 'Yes (مسموح)' ELSE 'No (ممنوع)' END AS itm_allow_discount_desc,
        ic.itm_max_disc_per,
        ic.itm_max_disc_val,
        u1.u_name_ar AS u1_name,
        u2.u_name_ar AS u2_name,
        u3.u_name_ar AS u3_name,
        ic.itm_unit1_unit2,
        ic.itm_unit1_unit3,
        CASE ic.itm_purchase_unit WHEN 3 THEN 'Unit 3 (الصغرى)' WHEN 2 THEN 'Unit 2 (المتوسطة)' ELSE 'Unit 1 (الكبرى)' END AS itm_purchase_unit_desc,
        CASE ic.itm_sell_unit WHEN 3 THEN 'Unit 3 (الصغرى)' WHEN 2 THEN 'Unit 2 (المتوسطة)' ELSE 'Unit 1 (الكبرى)' END AS itm_sell_unit_desc,
        ic.itm_min_limit,
        ic.itm_max_limit,
        ic.itm_request_limit,
        ic.itm_default_limit,
        CASE ic.itm_has_expire WHEN 1 THEN 'Yes (متابع الصلاحية)' ELSE 'No (بدون صلاحية)' END AS itm_has_expire_desc,
        CASE ic.itm_isShortage WHEN 1 THEN 'Yes - Shortage (ناقص)' ELSE 'Available (متوفر)' END AS itm_isShortage_desc,
        CASE ic.itm_frac_qty WHEN 1 THEN 'Yes (مسموح)' ELSE 'No (ممنوع)' END AS itm_frac_qty_desc,
        CASE ic.itm_favourite WHEN 1 THEN 'Yes (مميز)' ELSE 'Standard' END AS itm_favourite_desc,
        CASE ic.itm_active WHEN '1' THEN 'Active (نشط)' ELSE 'Inactive (غير نشط)' END AS itm_active_desc,
        CASE ic.itm_stop_sell WHEN 1 THEN 'Blocked (موقوف)' ELSE 'Permitted (متاح)' END AS itm_stop_sell_desc,
        CASE ic.itm_stop_pur WHEN 1 THEN 'Blocked (موقوف)' ELSE 'Permitted (متاح)' END AS itm_stop_pur_desc,
        ic.itm_notes,
        ic.sec_insert_date,
        ic.sec_update_date
    FROM Item_Catalog ic WITH (NOLOCK)
    LEFT JOIN Company c WITH (NOLOCK) ON c.com_id = ic.com_id
    LEFT JOIN Item_Formats itf WITH (NOLOCK) ON itf.itf_id = ic.itm_itf_id
    LEFT JOIN Scientific_Groups sg WITH (NOLOCK) ON sg.sg_id = ic.itm_scientific_group_id
    LEFT JOIN item_usage_manner ium WITH (NOLOCK) ON ium.ium_id = ic.itm_usage_manner_id
    LEFT JOIN Groups g WITH (NOLOCK) ON g.g_id = ic.itm_g1
    LEFT JOIN Item_Origins io WITH (NOLOCK) ON io.io_id = ic.itm_origin
    LEFT JOIN Units u1 WITH (NOLOCK) ON u1.u_id = ic.itm_unit1
    LEFT JOIN Units u2 WITH (NOLOCK) ON u2.u_id = ic.itm_unit2
    LEFT JOIN Units u3 WITH (NOLOCK) ON u3.u_id = ic.itm_unit3;
    """
    df_cat = pd.read_sql(q_catalog, ENGINE)
    
    # Filter only used items
    df_cat = df_cat[df_cat['itm_id'].isin(used_ids_set)].copy()
    logger.info(f"Loaded {len(df_cat):,} used item master records.")

    # 3. Decrypt Trade Names
    logger.info("Executing binary decryption on Arabic and English names...")
    def resolve_ar(row):
        if row['itm_name_ar_encrypt']:
            d = decode_legacy_bidi_bytes(row['itm_name_ar_encrypt'])
            if d:
                return d
        return row['itm_name_ar'] or ""

    def resolve_en(row):
        if row['itm_name_en_encrypt']:
            d = decode_legacy_bidi_bytes(row['itm_name_en_encrypt'])
            if d:
                return d
        return row['itm_name_en'] or ""

    df_cat['itm_name_ar'] = df_cat.apply(resolve_ar, axis=1)
    df_cat['itm_name_en'] = df_cat.apply(resolve_en, axis=1)

    # Clean Fallbacks
    df_cat['itm_name_en'] = df_cat['itm_name_en'].fillna("").apply(lambda x: str(x).strip())
    df_cat['itm_name_ar'] = df_cat['itm_name_ar'].fillna("").apply(lambda x: str(x).strip())
    
    # In case English name is completely empty, use Arabic name or Code
    mask_empty = df_cat['itm_name_en'] == ""
    df_cat.loc[mask_empty, 'itm_name_en'] = df_cat.loc[mask_empty, 'itm_name_ar'].combine_first(df_cat.loc[mask_empty, 'itm_code'])

    # Drop encrypted binary blobs
    df_cat.drop(columns=['itm_name_ar_encrypt', 'itm_name_en_encrypt'], inplace=True)

    # 4. Supplementary Barcodes (Item_Objects)
    logger.info("Pulling additional barcodes from Item_Objects...")
    q_objects = "SELECT io_itm_id AS itm_id, io_itm_int AS extra_code FROM Item_Objects WITH (NOLOCK) WHERE io_itm_int IS NOT NULL;"
    df_obj = pd.read_sql(q_objects, ENGINE)
    df_obj_grouped = df_obj.groupby('itm_id')['extra_code'].apply(lambda x: ", ".join(sorted(set(str(v).strip() for v in x)))).reset_index()
    df_obj_grouped.rename(columns={'extra_code': 'extra_barcodes'}, inplace=True)
    df_cat = df_cat.merge(df_obj_grouped, on='itm_id', how='left')

    # 5. Indications (Item_Usage_Causes)
    logger.info("Pulling clinical indications from Item_Usage_Causes...")
    q_causes = """
    SELECT iuc.iuc_itm_id AS itm_id, uc.uc_name_ar AS cause_name
    FROM Item_Usage_Causes iuc WITH (NOLOCK)
    JOIN Usage_Causes uc WITH (NOLOCK) ON uc.uc_id = iuc.iuc_uc_id;
    """
    df_causes = pd.read_sql(q_causes, ENGINE)
    df_causes_grp = df_causes.groupby('itm_id')['cause_name'].apply(lambda x: ", ".join(sorted(set(str(v).strip() for v in x)))).reset_index()
    df_causes_grp.rename(columns={'cause_name': 'usage_indications'}, inplace=True)
    df_cat = df_cat.merge(df_causes_grp, on='itm_id', how='left')

    # 6. Preferred/Last Vendor Pricing
    logger.info("Pulling vendor pricing context from Item_Vendor...")
    q_vendor = """
    SELECT 
        iv.itm_id,
        v.ven_name_ar AS preferred_vendor_name,
        iv.itm_pharm_price AS last_vendor_price,
        iv.itm_disc1 AS last_vendor_discount,
        iv.itm_pricing_date AS last_vendor_date
    FROM Item_Vendor iv WITH (NOLOCK)
    JOIN Vendor v WITH (NOLOCK) ON v.ven_id = iv.ven_id
    WHERE iv.itm_pricing_date IS NOT NULL;
    """
    df_ven = pd.read_sql(q_vendor, ENGINE)
    # Deduplicate keeping latest pricing date
    df_ven = df_ven.sort_values(by=['itm_id', 'last_vendor_date']).groupby('itm_id').last().reset_index()
    df_cat = df_cat.merge(df_ven, on='itm_id', how='left')

    # 7. Current Live Stock Balances & Locations (Item_Class_Store)
    logger.info("Calculating on-hand store stock and active batches...")
    q_stock = """
    SELECT 
        ics.itm_id,
        ics.sto_id,
        st.sto_name_ar AS Store_Name,
        ics.itm_qty,
        ics.pharm_price,
        ics.sell_price,
        ics.itm_expiry_date
    FROM Item_Class_Store ics WITH (NOLOCK)
    JOIN Store st WITH (NOLOCK) ON st.sto_id = ics.sto_id;
    """
    df_stock = pd.read_sql(q_stock, ENGINE)
    df_stock_used = df_stock[df_stock['itm_id'].isin(used_ids_set)].copy()

    # Aggregate active stock
    active_stock_grp = df_stock_used[df_stock_used['itm_qty'] > 0].groupby('itm_id').agg(
        current_total_stock=('itm_qty', 'sum'),
        active_batches_count=('itm_qty', 'count'),
        earliest_expiry_date=('itm_expiry_date', 'min'),
        total_cost_valuation=('itm_qty', lambda s: (s * df_stock_used.loc[s.index, 'pharm_price']).sum()),
        total_retail_valuation=('itm_qty', lambda s: (s * df_stock_used.loc[s.index, 'sell_price']).sum()),
        stores_currently_stocked=('Store_Name', lambda x: ", ".join(sorted(set(str(v) for v in x))))
    ).reset_index()
    df_cat = df_cat.merge(active_stock_grp, on='itm_id', how='left')

    # 8. All Historical Store Locations
    logger.info("Tracing all historical branch locations per item...")
    q_stores_hist = """
    SELECT DISTINCT itm_id, st.sto_name_ar AS Store_Name
    FROM (
        SELECT itm_id, sto_id FROM Item_Class_Store WITH (NOLOCK)
        UNION
        SELECT std.itm_id, h.sto_id FROM sales_trans_d std WITH (NOLOCK) JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id = std.sth_id
        UNION
        SELECT ptd.itm_id, pth.sto_id FROM pur_trans_d ptd WITH (NOLOCK) JOIN pur_trans_h pth WITH (NOLOCK) ON pth.pth_id = ptd.pth_id
        UNION
        SELECT st_itm_id AS itm_id, st_from_store AS sto_id FROM Store_Trans WITH (NOLOCK)
        UNION
        SELECT st_itm_id AS itm_id, st_to_store AS sto_id FROM Store_Trans WITH (NOLOCK)
    ) sh
    JOIN Store st WITH (NOLOCK) ON st.sto_id = sh.sto_id
    WHERE sh.itm_id IS NOT NULL;
    """
    df_stores_hist = pd.read_sql(q_stores_hist, ENGINE)
    df_stores_hist_grp = df_stores_hist.groupby('itm_id')['Store_Name'].apply(lambda x: ", ".join(sorted(set(str(v) for v in x)))).reset_index()
    df_stores_hist_grp.rename(columns={'Store_Name': 'stores_historical_locations'}, inplace=True)
    df_cat = df_cat.merge(df_stores_hist_grp, on='itm_id', how='left')

    # 9. Forensics: Determine Date, Type and Store of the Very Last Action Performed
    logger.info("Computing the latest operational action across all transactions...")
    q_last_actions = """
    SELECT itm_id, act_date, act_type, sto_name_ar AS act_store FROM (
        -- 1. Sales
        SELECT std.itm_id, h.sec_insert_date AS act_date, 'Sale (فاتورة بيع)' AS act_type, st.sto_name_ar
        FROM sales_trans_d std WITH (NOLOCK)
        JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id = std.sth_id
        JOIN Store st WITH (NOLOCK) ON st.sto_id = h.sto_id
        WHERE h.sec_insert_date IS NOT NULL

        UNION ALL
        -- 2. Purchases
        SELECT ptd.itm_id, pth.sec_insert_date AS act_date, 'Purchase (فاتورة شراء)' AS act_type, st.sto_name_ar
        FROM pur_trans_d ptd WITH (NOLOCK)
        JOIN pur_trans_h pth WITH (NOLOCK) ON pth.pth_id = ptd.pth_id
        JOIN Store st WITH (NOLOCK) ON st.sto_id = pth.sto_id
        WHERE pth.sec_insert_date IS NOT NULL

        UNION ALL
        -- 3. Stock Transfers
        SELECT st.st_itm_id AS itm_id, st.sec_insert_date AS act_date, 'Stock Transfer (تحويل مخزني)' AS act_type, s.sto_name_ar
        FROM Store_Trans st WITH (NOLOCK)
        JOIN Store s WITH (NOLOCK) ON s.sto_id = st.st_to_store
        WHERE st.sec_insert_date IS NOT NULL

        UNION ALL
        -- 4. Inventory Audits
        SELECT id.invd_itm_id AS itm_id, ih.sec_insert_date AS act_date, 'Stock Audit (تسوية جرد)' AS act_type, st.sto_name_ar
        FROM Inventory_d id WITH (NOLOCK)
        JOIN Inventory_h ih WITH (NOLOCK) ON ih.inv_id = id.invd_inv_id
        JOIN Store st WITH (NOLOCK) ON st.sto_id = ih.inv_sto_id
        WHERE ih.sec_insert_date IS NOT NULL

        UNION ALL
        -- 5. Quantity Adjustments / Breakages
        SELECT iqc.iqc_itm_id AS itm_id, iqc.sec_insert_date AS act_date, 'Qty Adjustment (تسوية هالك/تالف)' AS act_type, st.sto_name_ar
        FROM Item_Quantity_Changes iqc WITH (NOLOCK)
        JOIN Store st WITH (NOLOCK) ON st.sto_id = iqc.iqc_store_id
        WHERE iqc.sec_insert_date IS NOT NULL

        UNION ALL
        -- 6. Vendor General Returns
        SELECT ibd.ibd_itm_id AS itm_id, ibh.sec_insert_date AS act_date, 'Vendor Return (مرتجع مورد)' AS act_type, st.sto_name_ar
        FROM Item_back_d ibd WITH (NOLOCK)
        JOIN Item_back_h ibh WITH (NOLOCK) ON ibh.ibh_id = ibd.ibd_ibh_id
        JOIN Store st WITH (NOLOCK) ON st.sto_id = ibh.ibh_sto_id
        WHERE ibh.sec_insert_date IS NOT NULL

        UNION ALL
        -- 7. Price Changes
        SELECT ich.ich_itm_id AS itm_id, ich.sec_insert_date AS act_date, 'Price Update (تعديل سعر)' AS act_type, 'System (المنظومة)' AS sto_name_ar
        FROM Item_Changes ich WITH (NOLOCK)
        WHERE ich.sec_insert_date IS NOT NULL
    ) acts;
    """
    df_actions = pd.read_sql(q_last_actions, ENGINE)
    df_actions = df_actions[df_actions['itm_id'].isin(used_ids_set)].copy()
    
    # Sort and take the absolute latest action per item
    df_latest_action = df_actions.sort_values(by=['itm_id', 'act_date']).groupby('itm_id').last().reset_index()
    df_latest_action.rename(columns={
        'act_date': 'last_action_date',
        'act_type': 'last_action_type',
        'act_store': 'last_action_store'
    }, inplace=True)
    df_cat = df_cat.merge(df_latest_action, on='itm_id', how='left')

    # Fill default values
    df_cat['current_total_stock'] = df_cat['current_total_stock'].fillna(0.0)
    df_cat['active_batches_count'] = df_cat['active_batches_count'].fillna(0)
    df_cat['total_cost_valuation'] = df_cat['total_cost_valuation'].fillna(0.0)
    df_cat['total_retail_valuation'] = df_cat['total_retail_valuation'].fillna(0.0)
    df_cat['stores_currently_stocked'] = df_cat['stores_currently_stocked'].fillna("None (لا يوجد رصيد حالي)")
    df_cat['stores_historical_locations'] = df_cat['stores_historical_locations'].fillna("Unassigned")
    df_cat['extra_barcodes'] = df_cat['extra_barcodes'].fillna("")
    df_cat['usage_indications'] = df_cat['usage_indications'].fillna("")
    df_cat['preferred_vendor_name'] = df_cat['preferred_vendor_name'].fillna("None Recorded")

    # 10. Sort Alphabetically using English Name
    logger.info("Sorting all extracted records alphabetically by item English name...")
    df_cat['sort_key'] = df_cat['itm_name_en'].astype(str).str.upper().str.strip()
    df_cat.sort_values(by='sort_key', ascending=True, inplace=True)
    df_cat.drop(columns=['sort_key'], inplace=True)

    # 11. Rename Columns to Contain Actual DB Column Headers + English/Arabic Explanations
    column_mapping = {
        'itm_id': 'itm_id (Item Unique ID - كود الصنف الداخلي)',
        'itm_code': 'itm_code (Primary Barcode - الباركود الرئيسي)',
        'itm_name_en': 'itm_name_en (English Trade Name - الاسم التجاري بالإنجليزي)',
        'itm_name_ar': 'itm_name_ar (Arabic Trade Name - الاسم التجاري بالعربي)',
        'itm_code2': 'itm_code2 (Secondary Barcode - كود إضافي)',
        'itm_int_code': 'itm_int_code (International Barcode - الكود الدولي)',
        'itm_usr_code': 'itm_usr_code (User Custom Code - كود مخصص)',
        'ucp_code': 'ucp_code (UCP Distributor Code - كود الشركة المتحدة)',
        'extra_barcodes': 'extra_barcodes (Additional Linked Barcodes - باركودات إضافية مرتبطة)',
        'com_name_ar': 'com_name_ar (Manufacturer Name Ar - الشركة المنتجة بالعربي)',
        'com_name_en': 'com_name_en (Manufacturer Name En - الشركة المنتجة بالإنجليزي)',
        'itf_name_ar': 'itf_name_ar (Dosage Form / Format - الشكل الصيدلي)',
        'sg_name_ar': 'sg_name_ar (Pharmacological Group - المجموعة العلمية)',
        'itm_scientific_n1': 'itm_scientific_n1 (Generic Name 1 - الاسم العلمي الأول)',
        'itm_scientific_n2': 'itm_scientific_n2 (Generic Name 2 - الاسم العلمي الثاني)',
        'itm_effictive': 'itm_effictive (Active Ingredient - المادة الفعالة)',
        'itm_effictive_perc': 'itm_effictive_perc (Ingredient Strength - تركيز المادة الفعالة)',
        'ium_name_ar': 'ium_name_ar (Route of Administration - طريقة الاستخدام)',
        'g_name_ar': 'g_name_ar (Therapeutic Group - تصنيف الصنف)',
        'usage_indications': 'usage_indications (Clinical Indications - دواعي الاستعمال)',
        'io_text_ar': 'io_text_ar (Product Origin - بلد المنشأ)',
        'itm_ismedicine_desc': 'itm_ismedicine (Medicine Classification - طبيعة الصنف)',
        'itm_freez_desc': 'itm_freez (Cold Chain Storage - حفظ ثلاجة)',
        'itm_isprev_desc': 'itm_isprev (Prescription Required - صرف بروشتة)',
        'itm_def_sell_price': 'itm_def_sell_price (Current Retail Price - سعر بيع الجمهور)',
        'itm_def_pharm_price': 'itm_def_pharm_price (Pharmacy Cost Price - سعر شراء الصيدلية)',
        'unit_profit_margin': 'unit_profit_margin (Unit Profit Margin L.E - هامش الربح بالجنيه)',
        'profit_margin_perc': 'profit_margin_perc (Profit Margin Ratio - نسبة هامش الربح)',
        'itm_def_tax': 'itm_def_tax (Sales Tax Amount - قيمة الضريبة)',
        'Tax_Category': 'Tax_Category (Tax Authority Category - تصنيف الضرائب)',
        'itm_allow_discount_desc': 'itm_allow_discount (Discount Eligibility - خضوع للخصم)',
        'itm_max_disc_per': 'itm_max_disc_per (Max Discount % - أقصى نسبة خصم)',
        'itm_max_disc_val': 'itm_max_disc_val (Max Discount Value - أقصى قيمة خصم)',
        'u1_name': 'u1_name (Major Package Unit - الوحدة الكبرى)',
        'u2_name': 'u2_name (Intermediate Unit - الوحدة المتوسطة)',
        'u3_name': 'u3_name (Dosage Unit - الوحدة الصغرى)',
        'itm_unit1_unit2': 'itm_unit1_unit2 (Unit 1 to 2 Factor - معامل التجزئة 1 إلى 2)',
        'itm_unit1_unit3': 'itm_unit1_unit3 (Unit 1 to 3 Factor - معامل التجزئة 1 إلى 3)',
        'itm_purchase_unit_desc': 'itm_purchase_unit (Default Purchase Unit - وحدة الشراء الافتراضية)',
        'itm_sell_unit_desc': 'itm_sell_unit (Default Sell Unit - وحدة البيع الافتراضية)',
        'itm_min_limit': 'itm_min_limit (Minimum Stock Level - الحد الأدنى للمخزون)',
        'itm_max_limit': 'itm_max_limit (Maximum Stock Level - الحد الأقصى للمخزون)',
        'itm_request_limit': 'itm_request_limit (Reorder Trigger - حد إعادة الطلب)',
        'itm_default_limit': 'itm_default_limit (Default Order Limit - الكمية الافتراضية للطلب)',
        'itm_has_expire_desc': 'itm_has_expire (Expiry Tracking - متابعة الصلاحية)',
        'itm_isShortage_desc': 'itm_isShortage (Market Shortage Status - حالة النواقص)',
        'itm_frac_qty_desc': 'itm_frac_qty (Allow Fractional Selling - بيع أجزاء ووحدات)',
        'itm_favourite_desc': 'itm_favourite (Favourite Fast Moving - صنف مميز ومفضل)',
        'current_total_stock': 'current_total_stock (Total On-Hand Stock - إجمالي الرصيد الحالي الفعلي)',
        'active_batches_count': 'active_batches_count (Active Batches Count - عدد التشغيلات المتوفرة بالمخزون)',
        'earliest_expiry_date': 'earliest_expiry_date (Earliest Expiry Date - أقرب تاريخ انتهاء صلاحية)',
        'total_cost_valuation': 'total_cost_valuation (Stock Valuation at Cost - إجمالي قيمة المخزون بالتكلفة)',
        'total_retail_valuation': 'total_retail_valuation (Stock Valuation at Retail - إجمالي قيمة المخزون بسعر الجمهور)',
        'stores_currently_stocked': 'stores_currently_stocked (Stores with Positive Stock - الفروع المتوفر بها رصيد حالياً)',
        'stores_historical_locations': 'stores_historical_locations (All Historical Stores - جميع الفروع والمخازن التي تواجد بها الصنف تاريخياً)',
        'preferred_vendor_name': 'preferred_vendor_name (Preferred/Last Vendor - المورد الأخير/المفضل)',
        'last_vendor_price': 'last_vendor_price (Last Vendor Cost - آخر سعر شراء معلن من المورد)',
        'last_vendor_discount': 'last_vendor_discount (Last Vendor Discount % - آخر نسبة خصم مورد)',
        'last_vendor_date': 'last_vendor_date (Vendor Price Date - تاريخ تسعير المورد)',
        'sec_insert_date': 'sec_insert_date (System Creation Date - تاريخ إضافة الصنف بالمنظومة)',
        'sec_update_date': 'sec_update_date (Catalog Last Update - تاريخ آخر تعديل لبطاقة الصنف)',
        'last_action_date': 'last_action_date (Last Movement Date - تاريخ آخر حركة فعلية على الصنف)',
        'last_action_type': 'last_action_type (Last Movement Type - نوع آخر حركة: بيع/شراء/تحويل/جرد/سعر)',
        'last_action_store': 'last_action_store (Last Movement Store - الفرع الذي تمت به آخر حركة)',
        'itm_active_desc': 'itm_active (Catalog Active Status - حالة الصنف: نشط/غير نشط)',
        'itm_stop_sell_desc': 'itm_stop_sell (Sale Prohibition - موقوف بيعه)',
        'itm_stop_pur_desc': 'itm_stop_pur (Purchase Prohibition - موقوف شراؤه)',
        'itm_notes': 'itm_notes (Item Remarks - ملاحظات الصنف)'
    }
    
    df_master = df_cat.rename(columns=column_mapping)
    logger.info(f"Master DataFrame assembled: {df_master.shape[0]:,} rows, {df_master.shape[1]:,} columns.")
    return df_cat, df_master

# --------------------------------------------------------------------------------------
# 5. DELIVERABLES EXPORTER
# --------------------------------------------------------------------------------------
def export_all_deliverables(df_cat, df_master, base_dir):
    os.makedirs(base_dir, exist_ok=True)
    by_letter_dir = os.path.join(base_dir, "By_Letter")
    os.makedirs(by_letter_dir, exist_ok=True)

    # ==================================================================================
    # DELIVERABLE 1: Master Formatted Excel Workbook
    # ==================================================================================
    excel_path = os.path.join(base_dir, "01_Genius_Master_Used_Items_Catalog.xlsx")
    logger.info("==========================================================================")
    logger.info(f"GENERATING DELIVERABLE 1: {excel_path}...")
    logger.info("==========================================================================")
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_master.to_excel(writer, sheet_name="Used_Items_Catalog", index=False)

    logger.info("Styling Master Catalog Excel Workbook with OpenPyXL...")
    wb = openpyxl.load_workbook(excel_path)
    style_master_catalog_sheet(wb["Used_Items_Catalog"], header_fill_hex="1B365D")  # Deep Corporate Navy
    wb.save(excel_path)
    logger.info(f"Deliverable 1 Complete -> {excel_path}")

    # ==================================================================================
    # DELIVERABLE 2: Separate CSV for Each First Letter of English Name
    # ==================================================================================
    logger.info("==========================================================================")
    logger.info("GENERATING DELIVERABLE 2: Segmented Letter-by-Letter CSV Files...")
    logger.info("==========================================================================")
    
    # Prepare clean 2-column DataFrame
    df_clean_2col = df_cat[['itm_code', 'itm_name_en']].copy()
    df_clean_2col['itm_code'] = df_clean_2col['itm_code'].fillna("").astype(str).str.strip()
    df_clean_2col['itm_name_en'] = df_clean_2col['itm_name_en'].fillna("").astype(str).str.strip()
    
    # Sort alphabetically by English name
    df_clean_2col['sort_key'] = df_clean_2col['itm_name_en'].str.upper()
    df_clean_2col.sort_values(by='sort_key', ascending=True, inplace=True)
    df_clean_2col.drop(columns=['sort_key'], inplace=True)

    # Determine first letter
    def get_initial_bucket(name):
        if not name:
            return "Other"
        first_char = name[0].upper()
        if 'A' <= first_char <= 'Z':
            return first_char
        return "Other"

    df_clean_2col['Letter'] = df_clean_2col['itm_name_en'].apply(get_initial_bucket)

    letter_groups = df_clean_2col.groupby('Letter')
    created_letter_files = 0
    for letter, group in letter_groups:
        filename = f"Items_Letter_{letter}.csv"
        target_csv = os.path.join(by_letter_dir, filename)
        # Export ONLY itm_code and itm_name_en
        group[['itm_code', 'itm_name_en']].to_csv(target_csv, index=False, encoding='utf-8-sig')
        created_letter_files += 1

    logger.info(f"Deliverable 2 Complete -> {created_letter_files} CSV files written to: {by_letter_dir}")

    # ==================================================================================
    # DELIVERABLE 3: Single Master CSV File of All Used Items
    # ==================================================================================
    logger.info("==========================================================================")
    logger.info("GENERATING DELIVERABLE 3: Master Consolidated Used Items CSV...")
    logger.info("==========================================================================")
    
    master_csv_path = os.path.join(base_dir, "All_Used_Items.csv")
    # Export ONLY itm_code and itm_name_en sorted alphabetically
    df_clean_2col[['itm_code', 'itm_name_en']].to_csv(master_csv_path, index=False, encoding='utf-8-sig')
    logger.info(f"Deliverable 3 Complete -> {master_csv_path} ({len(df_clean_2col):,} items)")

# --------------------------------------------------------------------------------------
# 6. MAIN EXECUTION ORCHESTRATOR
# --------------------------------------------------------------------------------------
def main():
    total_start = time.time()
    logger.info("==========================================================================")
    logger.info("STARTING GENIUS ERP COMPREHENSIVE USED ITEMS CATALOG EXTRACTION")
    logger.info("==========================================================================")

    output_dir = os.path.join(os.getcwd(), "Exported_Items_Catalog")
    
    # Connectivity Test
    try:
        with ENGINE.connect() as conn:
            conn.execute(text("SELECT 1"))
        logger.info("SQL Server Connection Verified Successfully.")
    except Exception as e:
        logger.error(f"Cannot connect to database: {e}")
        sys.exit(1)

    try:
        df_cat, df_master = extract_master_used_catalog()
        export_all_deliverables(df_cat, df_master, output_dir)
    except Exception as e:
        logger.error(f"Execution failed due to an exception: {e}", exc_info=True)
        sys.exit(1)

    elapsed = round(time.time() - total_start, 2)
    logger.info("==========================================================================")
    logger.info(f"PIPELINE COMPLETED SUCCESSFULLY IN {elapsed} SECONDS.")
    logger.info(f"Output Master Folder: {output_dir}")
    logger.info("==========================================================================")

if __name__ == "__main__":
    main()