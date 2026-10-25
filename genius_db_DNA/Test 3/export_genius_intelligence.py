#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
========================================================================================
GENIUS (e-Plus) PHARMACY ERP - MASTER DATA CONSOLIDATION & INTELLIGENCE ENGINE
========================================================================================
Architecture Decoded Modules:
  1. Inventory & Batch Control Engine       -> 01_Genius_Inventory_Master.xlsx
  2. Sales & Revenue Operations             -> 02_Genius_Sales_Operations.xlsx
  3. Procurement & Vendor Supply Chain      -> 03_Genius_Purchasing_Operations.xlsx
  4. Treasury & General Financial Ledger    -> 04_Genius_Financial_Ledger.xlsx
  5. POS Cashier & Shift Audit Trail        -> 05_Genius_POS_Shift_Operations.xlsx
  6. Executive Business Intelligence        -> 06_Genius_Executive_Intelligence_Cockpit.xlsx
========================================================================================
"""

import os
import sys
import time
import logging
import urllib.parse
import pandas as pd
from datetime import datetime
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
logger = logging.getLogger("GeniusEngine")

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
    The binary buffer is split in half (N // 2), each half is reversed
    independently, concatenated back together, and decoded with windows-1256.
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

# Global In-Memory Decryption Cache
_ITEM_LOOKUP_CACHE = None

def get_item_name_cache(engine):
    """
    Builds a high-speed in-memory lookup table of decrypted item names.
    Executed once upon startup to decrypt 43,000+ catalog SKUs in ~0.2 seconds.
    """
    global _ITEM_LOOKUP_CACHE
    if _ITEM_LOOKUP_CACHE is None:
        logger.info("Initializing Item Decryption Engine & Loading Catalog Cache...")
        q = """
        SELECT 
            itm_id, 
            itm_name_ar, 
            itm_name_en, 
            itm_name_ar_encrypt, 
            itm_name_en_encrypt 
        FROM Item_Catalog WITH (NOLOCK)
        """
        try:
            df_cat = pd.read_sql(q, engine)
            cache_ar = {}
            cache_en = {}
            for _, row in df_cat.iterrows():
                i_id = row['itm_id']
                
                # Resolve Arabic Name
                raw_ar = row['itm_name_ar_encrypt']
                dec_ar = decode_legacy_bidi_bytes(raw_ar) if raw_ar else ""
                cache_ar[i_id] = dec_ar if dec_ar else (row['itm_name_ar'] or "")
                
                # Resolve English Name
                raw_en = row['itm_name_en_encrypt']
                dec_en = decode_legacy_bidi_bytes(raw_en) if raw_en else ""
                cache_en[i_id] = dec_en if dec_en else (row['itm_name_en'] or "")

            _ITEM_LOOKUP_CACHE = {'ar': cache_ar, 'en': cache_en}
            logger.info(f"Decryption Cache Ready: {len(cache_ar):,} item titles decoded.")
        except Exception as e:
            logger.error(f"Failed to build Item Decryption Cache: {e}")
            _ITEM_LOOKUP_CACHE = {'ar': {}, 'en': {}}
    return _ITEM_LOOKUP_CACHE

def decrypt_item_dataframe(df, engine):
    """
    Automatically detects item ID and name columns in any extracted DataFrame
    and populates them with the decrypted Arabic and English trade names.
    """
    if df is None or df.empty:
        return df

    # Detect Item ID Column
    id_col = None
    for candidate in ['Item_ID', 'itm_id', 'di_itm_id', 'st_itm_id']:
        if candidate in df.columns:
            id_col = candidate
            break

    if not id_col:
        return df

    lookup = get_item_name_cache(engine)

    # Decrypt Arabic Name columns
    for ar_col in ['Item_Name_Ar', 'Item_Name']:
        if ar_col in df.columns:
            decrypted_ar = df[id_col].map(lookup['ar'])
            df[ar_col] = decrypted_ar.combine_first(df[ar_col])

    # Decrypt English Name columns
    if 'Item_Name_En' in df.columns:
        decrypted_en = df[id_col].map(lookup['en'])
        df['Item_Name_En'] = decrypted_en.combine_first(df['Item_Name_En'])

    return df

# --------------------------------------------------------------------------------------
# 3. OPENPYXL STYLING & FORMATTING HELPERS
# --------------------------------------------------------------------------------------
THIN_BORDER = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'),
    bottom=Side(style='thin', color='D3D3D3')
)

def style_worksheet(ws, header_fill_hex="1F4E79", is_kpi=False):
    """
    Applies professional corporate layout:
      - Custom Header Colors & White Bold Text
      - Freeze top header row
      - Auto-fitted column widths
      - Standardized gridlines & cell alignments
      - Accounting & Date format detection
      - High-performance row-wise iteration
    """
    header_fill = PatternFill(start_color=header_fill_hex, end_color=header_fill_hex, fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    kpi_bold_font = Font(name="Segoe UI", size=9, bold=True)
    
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A2"

    # Style Header Row
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 28

    max_row = ws.max_row
    max_col = ws.max_column

    # Pre-calculate column header types & lengths
    col_meta = {}
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(ws.cell(row=1, column=col_idx).value or "")
        h_lower = header_val.lower()
        
        is_money = any(t in h_lower for t in ["price", "cost", "val", "total", "net", "credit", "debit", "balance", "revenue", "cogs", "profit", "cash"])
        is_perc = any(t in h_lower for t in ["perc", "rate", "ratio", "%"])
        is_qty = any(t in h_lower for t in ["qty", "count", "items", "no_"])
        
        col_meta[col_idx] = {
            'letter': col_letter,
            'header_val': header_val,
            'max_len': len(header_val),
            'is_money': is_money,
            'is_perc': is_perc,
            'is_qty': is_qty
        }

    # Style Data Cells
    for row in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            col_idx = cell.column
            meta = col_meta[col_idx]
            val = cell.value
            cell.font = kpi_bold_font if (is_kpi and col_idx == 1) else data_font
            cell.border = THIN_BORDER

            if val is not None:
                str_val = str(val)
                if len(str_val) > meta['max_len'] and cell.row <= 1000:
                    meta['max_len'] = len(str_val)

                # Format Numbers & Dates
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

    # Set Column Widths
    for col_idx in range(1, max_col + 1):
        meta = col_meta[col_idx]
        ws.column_dimensions[meta['letter']].width = min(max(meta['max_len'] + 4, 12), 48)

# ======================================================================================
# MODULE 1: CONSOLIDATING INVENTORY DETAILED DATA
# ======================================================================================
def generate_inventory_workbook(output_dir):
    filename = os.path.join(output_dir, "01_Genius_Inventory_Master.xlsx")
    logger.info(f"Generating Module 1: {filename}...")

    # 1. Active Store Batches
    q_batches = """
    SELECT 
        ics.sto_id AS Store_ID,
        st.sto_name_ar AS Store_Name,
        ics.itm_id AS Item_ID,
        ic.itm_code AS Barcode,
        ic.itm_name_ar AS Item_Name_Ar,
        ic.itm_name_en AS Item_Name_En,
        ics.c_id AS Batch_Class_ID,
        ics.itm_qty AS Batch_Quantity,
        u.u_name_ar AS Base_Unit,
        ics.pharm_price AS Unit_Cost_Price,
        ics.sell_price AS Unit_Sell_Price,
        ics.sell_tax AS Unit_Tax,
        CAST(ics.itm_qty * ics.pharm_price AS DECIMAL(18,2)) AS Total_Batch_Cost_Valuation,
        CAST(ics.itm_qty * ics.sell_price AS DECIMAL(18,2)) AS Total_Batch_Retail_Valuation,
        CONVERT(VARCHAR(10), ics.itm_expiry_date, 120) AS Expiry_Date,
        v.ven_name_ar AS Preferred_Vendor,
        ics.itm_location AS Rack_Location
    FROM Item_Class_Store ics WITH (NOLOCK)
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = ics.itm_id
    JOIN Store st WITH (NOLOCK) ON st.sto_id = ics.sto_id
    LEFT JOIN Units u WITH (NOLOCK) ON u.u_id = ic.itm_unit1
    LEFT JOIN Vendor v WITH (NOLOCK) ON v.ven_id = ics.ven_id
    ORDER BY ics.sto_id, ic.itm_name_ar;
    """

    # 2. Consolidated Stock Valuation per Item
    q_consolidated = """
    SELECT 
        ic.itm_id AS Item_ID,
        ic.itm_code AS Barcode,
        ic.itm_name_ar AS Item_Name_Ar,
        c.com_name_ar AS Manufacturer,
        COUNT(DISTINCT ics.c_id) AS Active_Batch_Count,
        SUM(ics.itm_qty) AS Total_Stock_Units,
        AVG(ics.pharm_price) AS Weighted_Avg_Cost,
        MAX(ics.sell_price) AS Current_Retail_Price,
        SUM(CAST(ics.itm_qty * ics.pharm_price AS DECIMAL(18,2))) AS Consolidated_Cost_Value,
        SUM(CAST(ics.itm_qty * ics.sell_price AS DECIMAL(18,2))) AS Consolidated_Retail_Value,
        MIN(ics.itm_expiry_date) AS Nearest_Expiry
    FROM Item_Class_Store ics WITH (NOLOCK)
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = ics.itm_id
    LEFT JOIN Company c WITH (NOLOCK) ON c.com_id = ic.com_id
    GROUP BY ic.itm_id, ic.itm_code, ic.itm_name_ar, c.com_name_ar
    HAVING SUM(ics.itm_qty) <> 0
    ORDER BY Consolidated_Cost_Value DESC;
    """

    # 3. Inter-Store Transfers
    q_transfers = """
    SELECT 
        sth.stnh_id AS Transfer_ID,
        sth.stnh_flag AS Status_Flag,
        fs.sto_name_ar AS From_Store,
        ts.sto_name_ar AS To_Store,
        sth.stnh_no_items AS Total_Item_Lines,
        sth.stnh_tot_cost_value AS Total_Cost_Transfer,
        sth.stnh_tot_sell_value AS Total_Retail_Transfer,
        st.st_itm_id AS Item_ID,
        ic.itm_name_ar AS Item_Name,
        st.st_itm_quantity AS Transferred_Qty,
        st.st_itm_cost AS Line_Cost,
        st.st_itm_sell AS Line_Sell,
        sth.sec_insert_date AS Transfer_Date,
        emp.e_Name AS Operator
    FROM Store_Trans_h sth WITH (NOLOCK)
    JOIN Store_Trans st WITH (NOLOCK) ON st.stnh_id = sth.stnh_id
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = st.st_itm_id
    JOIN Store fs WITH (NOLOCK) ON fs.sto_id = sth.stnh_f_Sto_id
    JOIN Store ts WITH (NOLOCK) ON ts.sto_id = sth.stnh_t_Sto_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = sth.sec_insert_uid
    ORDER BY sth.stnh_id DESC;
    """

    # 4. Physical Stock-Take Audits
    q_audits = """
    SELECT 
        ih.inv_id AS Audit_Doc_ID,
        st.sto_name_ar AS Store_Name,
        ih.inv_begin_time AS Audit_Timestamp,
        ih.inv_flag AS Doc_State,
        id.invd_itm_id AS Item_ID,
        ic.itm_name_ar AS Item_Name,
        id.invd_c_id AS Batch_ID,
        id.invd_old_qty AS System_Expected_Qty,
        id.invd_new_qty AS Physical_Counted_Qty,
        (id.invd_new_qty - id.invd_old_qty) AS Variance_Qty,
        id.invd_pharm_price AS Unit_Cost,
        CAST((id.invd_new_qty - id.invd_old_qty) * id.invd_pharm_price AS DECIMAL(18,2)) AS Financial_Cost_Impact,
        emp.e_Name AS Audited_By
    FROM Inventory_h ih WITH (NOLOCK)
    JOIN Inventory_d id WITH (NOLOCK) ON id.invd_inv_id = ih.inv_id
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = id.invd_itm_id
    JOIN Store st WITH (NOLOCK) ON st.sto_id = ih.inv_sto_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = ih.sec_insert_uid
    ORDER BY ih.inv_id DESC;
    """

    # 5. Quantity Adjustments & Breakages
    q_adjustments = """
    SELECT 
        iqc.iqc_id AS Adj_ID,
        st.sto_name_ar AS Store_Name,
        iqc.iqc_itm_id AS Item_ID,
        ic.itm_name_ar AS Item_Name,
        dc.dc_name_ar AS Cause_Description,
        iqc.iqc_old_quantity AS Old_Quantity,
        iqc.iqc_new_quantity AS New_Quantity,
        (iqc.iqc_new_quantity - iqc.iqc_old_quantity) AS Quantity_Delta,
        iqc.iqc_pharm_price AS Cost_Price,
        CAST((iqc.iqc_new_quantity - iqc.iqc_old_quantity) * iqc.iqc_pharm_price AS DECIMAL(18,2)) AS Adjustment_Cost_Value,
        iqc.iqc_notes AS Notes,
        iqc.sec_insert_date AS Action_Timestamp
    FROM Item_Quantity_Changes iqc WITH (NOLOCK)
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = iqc.iqc_itm_id
    LEFT JOIN Store st WITH (NOLOCK) ON st.sto_id = iqc.iqc_store_id
    LEFT JOIN destroy_causes dc WITH (NOLOCK) ON dc.dc_id = iqc.iqc_dc_id
    ORDER BY iqc.iqc_id DESC;
    """

    # 6. Opening Balances
    q_opening = """
    SELECT 
        so.so_id AS Opening_Doc_ID,
        st.sto_name_ar AS Store_Name,
        sod.itm_id AS Item_ID,
        ic.itm_name_ar AS Item_Name,
        sod.c_id AS Batch_ID,
        sod.itm_qty AS Opening_Quantity,
        sod.pharm_price AS Cost_Price,
        sod.sell_price AS Retail_Price,
        sod.itm_expiry_date AS Expiry_Date,
        so.sec_insert_date AS Setup_Date
    FROM Stock_Open so WITH (NOLOCK)
    JOIN Stock_Open_d sod WITH (NOLOCK) ON sod.so_id = so.so_id
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = sod.itm_id
    JOIN Store st WITH (NOLOCK) ON st.sto_id = so.so_sto_id
    ORDER BY so.so_id DESC;
    """

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        decrypt_item_dataframe(pd.read_sql(q_batches, ENGINE), ENGINE).to_excel(writer, sheet_name="Current_Store_Batches", index=False)
        decrypt_item_dataframe(pd.read_sql(q_consolidated, ENGINE), ENGINE).to_excel(writer, sheet_name="Consolidated_Inventory", index=False)
        decrypt_item_dataframe(pd.read_sql(q_transfers, ENGINE), ENGINE).to_excel(writer, sheet_name="Stock_Transfers", index=False)
        decrypt_item_dataframe(pd.read_sql(q_audits, ENGINE), ENGINE).to_excel(writer, sheet_name="Physical_Audits_Variance", index=False)
        decrypt_item_dataframe(pd.read_sql(q_adjustments, ENGINE), ENGINE).to_excel(writer, sheet_name="Quantity_Adjustments", index=False)
        decrypt_item_dataframe(pd.read_sql(q_opening, ENGINE), ENGINE).to_excel(writer, sheet_name="Opening_Stock", index=False)

    wb = openpyxl.load_workbook(filename)
    for ws_name in wb.sheetnames:
        style_worksheet(wb[ws_name], header_fill_hex="1B4D3E")  # Emerald theme
    wb.save(filename)
    logger.info("Module 1 Complete.")

# ======================================================================================
# MODULE 2: CONSOLIDATING SALES OPERATIONS DETAILED DATA
# ======================================================================================
def generate_sales_workbook(output_dir):
    filename = os.path.join(output_dir, "02_Genius_Sales_Operations.xlsx")
    logger.info(f"Generating Module 2: {filename}...")

    # 1. Sales Headers
    q_sales_h = """
    SELECT 
        h.sth_id AS Bill_ID,
        st.sto_name_ar AS Store_Name,
        h.sec_insert_date AS Invoice_Date,
        CASE h.bill_typ 
            WHEN 1 THEN 'Cash (نقدي)' 
            WHEN 2 THEN 'Credit (آجل)' 
            WHEN 3 THEN 'Delivery (توصيل)' 
            WHEN 4 THEN 'Visa (فيزا)' 
            ELSE 'Other' 
        END AS Bill_Type,
        ISNULL(c.cust_name_ar, 'عميل نقدي (Cash Walk-in)') AS Customer_Name,
        h.no_of_items AS Total_Lines,
        h.total_bill AS Gross_Amount,
        h.total_des_mon AS Cash_Discount,
        h.total_dis_per AS Percentage_Discount,
        h.total_bill_after_disc AS Net_Invoice_Amount,
        h.total_tax AS Tax_Amount,
        h.sth_cash AS Paid_Cash,
        h.sth_rest AS Customer_Remaining_Debt,
        h.sth_extra_expenses AS Delivery_Service_Fee,
        emp.e_Name AS Cashier_Name,
        h.sth_pc_name AS Station_Hostname
    FROM sales_trans_h h WITH (NOLOCK)
    JOIN Store st WITH (NOLOCK) ON st.sto_id = h.sto_id
    LEFT JOIN Customer c WITH (NOLOCK) ON c.cust_id = h.cust_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = h.emp_id
    ORDER BY h.sth_id DESC;
    """

    # 2. Sales Detailed Lines
    q_sales_d = """
    SELECT 
        d.sth_id AS Bill_ID,
        d.std_id AS Line_No,
        h.sec_insert_date AS Sale_Date,
        d.itm_id AS Item_ID,
        ic.itm_code AS Barcode,
        ic.itm_name_ar AS Item_Name,
        d.c_id AS Batch_Class_ID,
        d.exp_date AS Batch_Expiry,
        d.qnty AS Sold_Quantity,
        d.itm_unit AS Unit_Level,
        d.itm_sell AS Unit_Sale_Price,
        d.itm_cost AS Unit_Cost_COGS,
        d.itm_dis_per AS Item_Discount_Perc,
        d.itm_dis_mon AS Item_Discount_Cash,
        CAST((d.qnty * d.itm_sell * (1.0 - (d.itm_dis_per / 100.0))) - d.itm_dis_mon AS DECIMAL(18,2)) AS Net_Line_Revenue,
        CAST(d.qnty * d.itm_cost AS DECIMAL(18,2)) AS Net_Line_COGS,
        CAST(((d.qnty * d.itm_sell * (1.0 - (d.itm_dis_per / 100.0))) - d.itm_dis_mon) - (d.qnty * d.itm_cost) AS DECIMAL(18,2)) AS Gross_Profit
    FROM sales_trans_d d WITH (NOLOCK)
    JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id = d.sth_id
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = d.itm_id
    WHERE h.sth_flag = 'C'
    ORDER BY d.sth_id DESC, d.std_id ASC;
    """

    # 3. Payment Methods Breakdown
    q_payments = """
    SELECT 
        stp.stp_sth_id AS Bill_ID,
        st.sto_name_ar AS Store_Name,
        fcs.fcs_name_ar AS Cash_Drawer_Name,
        stp.stp_value AS Tendered_Value,
        stp.stp_insert_date AS Payment_Timestamp,
        emp.e_Name AS Processed_By,
        stp.stp_pc_name AS Station
    FROM sales_trans_payment stp WITH (NOLOCK)
    JOIN Store st WITH (NOLOCK) ON st.sto_id = stp.stp_sto_id
    JOIN F_Cash_Store fcs WITH (NOLOCK) ON fcs.fcs_id = stp.stp_fcs_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = stp.stp_insert_uid
    ORDER BY stp.stp_sth_id DESC;
    """

    # 4. Sales Returns Operations
    q_returns = """
    SELECT 
        sr.sr_id AS Return_Doc_ID,
        sr.sth_id AS Original_Bill_ID,
        h.sec_insert_date AS Original_Sale_Date,
        sr.returned_items_no AS Items_Count_Returned,
        sr.returned_items_value AS Gross_Refund_Value,
        sr.total_back_tax AS Tax_Refunded,
        srp.srp_value AS Actual_Cash_Refunded,
        fcs.fcs_name_ar AS Refunded_From_Drawer,
        sr.sec_insert_date AS Return_Timestamp,
        emp.e_Name AS Handled_By
    FROM sales_return sr WITH (NOLOCK)
    JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id = sr.sth_id
    LEFT JOIN sales_return_payment srp WITH (NOLOCK) ON srp.srp_sr_id = sr.sr_id
    LEFT JOIN F_Cash_Store fcs WITH (NOLOCK) ON fcs.fcs_id = srp.srp_fcs_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = sr.sec_insert_uid
    ORDER BY sr.sr_id DESC;
    """

    # 5. Visa Transactions
    q_visa = """
    SELECT 
        svi.sth_id AS Bill_ID,
        b.bnk_name_ar AS Terminal_Bank,
        svi.visa_name AS Cardholder_Name,
        svi.visa_no AS Card_Masked_Number,
        svi.exp_date AS Card_Expiry,
        svi.visa_value AS Transaction_Amount,
        svi.get_fees_perc AS Bank_Fee_Percentage,
        svi.sec_insert_date AS Transaction_Date
    FROM sales_visa_info svi WITH (NOLOCK)
    LEFT JOIN Bank b WITH (NOLOCK) ON b.bnk_id = svi.visa_bank_id
    ORDER BY svi.sth_id DESC;
    """

    # 6. Deleted Invoices Audit Log
    q_del_bills = """
    SELECT 
        del.sth_id AS Voided_Bill_ID,
        st.sto_name_ar AS Store_Name,
        del.total_bill_net AS Voided_Amount,
        del.no_of_items AS Voided_Lines,
        del.sec_insert_date AS Creation_Time,
        del.sth_delTime AS Deletion_Time,
        del.sth_del_pc_name AS Deletion_Station,
        emp.e_Name AS Original_Cashier
    FROM sales_trans_h_del del WITH (NOLOCK)
    JOIN Store st WITH (NOLOCK) ON st.sto_id = del.sto_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = del.sec_insert_uid
    ORDER BY del.sth_delTime DESC;
    """

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        pd.read_sql(q_sales_h, ENGINE).to_excel(writer, sheet_name="Sales_Invoices", index=False)
        decrypt_item_dataframe(pd.read_sql(q_sales_d, ENGINE), ENGINE).to_excel(writer, sheet_name="Sales_Line_Items", index=False)
        pd.read_sql(q_payments, ENGINE).to_excel(writer, sheet_name="Drawer_Payments", index=False)
        pd.read_sql(q_returns, ENGINE).to_excel(writer, sheet_name="Sales_Returns", index=False)
        pd.read_sql(q_visa, ENGINE).to_excel(writer, sheet_name="Card_Settlements", index=False)
        pd.read_sql(q_del_bills, ENGINE).to_excel(writer, sheet_name="Deleted_Bills_Audit", index=False)

    wb = openpyxl.load_workbook(filename)
    for ws_name in wb.sheetnames:
        style_worksheet(wb[ws_name], header_fill_hex="1F4E79")  # Sapphire Blue
    wb.save(filename)
    logger.info("Module 2 Complete.")

# ======================================================================================
# MODULE 3: CONSOLIDATING PURCHASES RELATED OPERATIONS
# ======================================================================================
def generate_purchases_workbook(output_dir):
    filename = os.path.join(output_dir, "03_Genius_Purchasing_Operations.xlsx")
    logger.info(f"Generating Module 3: {filename}...")

    # 1. Purchase Invoices Header
    q_pur_h = """
    SELECT 
        pth.pth_id AS Purchase_Bill_ID,
        st.sto_name_ar AS Receiving_Store,
        v.ven_name_ar AS Vendor_Name,
        pth.ven_bill_no AS Vendor_Invoice_Ref,
        pth.ven_bill_date AS Vendor_Bill_Date,
        pth.no_of_items AS Total_Lines,
        pth.total_bill AS Gross_Purchase_Total,
        pth.total_dis_per AS Header_Discount_Perc,
        pth.total_des_mon AS Header_Discount_Cash,
        pth.total_after_back AS Net_Purchase_Total,
        pth.pth_paid AS Paid_Amount,
        (pth.total_after_back - pth.pth_paid) AS Remaining_Vendor_Liability,
        pth.pth_source_tax AS Withholding_Tax_Rate,
        emp.e_Name AS Created_By,
        pth.sec_insert_date AS System_Entry_Date
    FROM pur_trans_h pth WITH (NOLOCK)
    JOIN Store st WITH (NOLOCK) ON st.sto_id = pth.sto_id
    JOIN Vendor v WITH (NOLOCK) ON v.ven_id = pth.ven_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = pth.sec_insert_uid
    WHERE pth.bill_status <> 1
    ORDER BY pth.pth_id DESC;
    """

    # 2. Purchase Invoices Details
    q_pur_d = """
    SELECT 
        ptd.pth_id AS Purchase_Bill_ID,
        ptd.ptd_id AS Line_No,
        pth.sec_insert_date AS Purchase_Date,
        v.ven_name_ar AS Vendor_Name,
        ptd.itm_id AS Item_ID,
        ic.itm_code AS Barcode,
        ic.itm_name_ar AS Item_Name,
        ptd.c_id AS Assigned_Batch_ID,
        CONVERT(VARCHAR(10), ptd.exp_date, 120) AS Expiry_Date,
        ptd.qnty AS Purchased_Qty,
        ptd.bonus AS Bonus_Free_Qty,
        ptd.itm_pur_price AS Unit_Wholesale_Price,
        ptd.itm_extra_dis AS Line_Discount_Perc,
        ptd.itm_tax_price AS Unit_Tax_Amount,
        ptd.itm_cost AS Effective_Unit_Cost,
        ptd.itm_sell AS Expected_Retail_Price,
        CAST((ptd.qnty * (ptd.itm_pur_price * (1.0 - (ISNULL(ptd.itm_extra_dis,0)/100.0)) + ptd.itm_tax_price)) AS DECIMAL(18,2)) AS Net_Line_Purchase_Cost,
        ptd.itm_back_qty AS Returned_Qty
    FROM pur_trans_d ptd WITH (NOLOCK)
    JOIN pur_trans_h pth WITH (NOLOCK) ON pth.pth_id = ptd.pth_id
    JOIN Vendor v WITH (NOLOCK) ON v.ven_id = pth.ven_id
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = ptd.itm_id
    WHERE pth.bill_status <> 1
    ORDER BY ptd.pth_id DESC, ptd.ptd_id ASC;
    """

    # 3. Vendor General Returns
    q_pur_returns = """
    SELECT 
        ibh.ibh_id AS Return_Doc_ID,
        st.sto_name_ar AS Store_Name,
        v.ven_name_ar AS Vendor_Name,
        ibh.ibh_total_value AS Total_Return_Value,
        ibd.ibd_itm_id AS Item_ID,
        ic.itm_name_ar AS Item_Name,
        ibd.ibd_qty AS Returned_Quantity,
        ibd.ibd_pharm_price AS Refund_Unit_Cost,
        ibd.ibd_tax_price AS Refund_Unit_Tax,
        CAST(ibd.ibd_qty * (ibd.ibd_pharm_price + ibd.ibd_tax_price) AS DECIMAL(18,2)) AS Total_Refund_Credit,
        CONVERT(VARCHAR(10), ibd.ibd_exp_date, 120) AS Expiry_Date,
        ibh.sec_insert_date AS Return_Date,
        emp.e_Name AS Returned_By
    FROM Item_back_h ibh WITH (NOLOCK)
    JOIN Item_back_d ibd WITH (NOLOCK) ON ibd.ibd_ibh_id = ibh.ibh_id
    JOIN Vendor v WITH (NOLOCK) ON v.ven_id = ibh.ibh_ven_id
    JOIN Store st WITH (NOLOCK) ON st.sto_id = ibh.ibh_sto_id
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = ibd.ibd_itm_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = ibh.sec_insert_uid
    ORDER BY ibh.ibh_id DESC;
    """

    # 4. Vendor Pricing History & Catalogs
    q_vendor_pricing = """
    SELECT 
        ivh.itm_id AS Item_ID,
        ic.itm_name_ar AS Item_Name,
        v.ven_name_ar AS Vendor_Name,
        ivh.itm_pharm_price AS Quoted_Pharm_Price,
        ivh.itm_sell_price AS Retail_Price,
        ivh.itm_disc1 AS Primary_Discount,
        ivh.itm_disc2 AS Secondary_Discount,
        ivh.itm_disc3 AS Special_Discount,
        ivh.itm_pricing_date AS Price_Recorded_Date
    FROM Item_Vendor_history ivh WITH (NOLOCK)
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = ivh.itm_id
    JOIN Vendor v WITH (NOLOCK) ON v.ven_id = ivh.ven_id
    ORDER BY ivh.itm_pricing_date DESC;
    """

    # 5. Vendor Summary Ledger
    q_vendor_ledger = """
    SELECT 
        v.ven_id AS Vendor_ID,
        v.ven_code AS Vendor_Code,
        v.ven_name_ar AS Vendor_Name_Ar,
        v.ven_tel AS Phone_Contact,
        v.ven_max_credit AS Credit_Limit,
        v.ven_start_debit AS Initial_Debit,
        v.ven_start_credit AS Initial_Credit,
        v.ven_current_credit AS Current_Outstanding_Balance,
        CASE WHEN v.ven_current_credit < 0 THEN 'We Owe Vendor (دائن)' ELSE 'Vendor Owes Us (مدين)' END AS Balance_Status,
        v.sec_update_date AS Last_Transaction_Date
    FROM Vendor v WITH (NOLOCK)
    WHERE v.ven_active = '1'
    ORDER BY ABS(v.ven_current_credit) DESC;
    """

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        pd.read_sql(q_pur_h, ENGINE).to_excel(writer, sheet_name="Purchase_Bills", index=False)
        decrypt_item_dataframe(pd.read_sql(q_pur_d, ENGINE), ENGINE).to_excel(writer, sheet_name="Purchase_Line_Items", index=False)
        decrypt_item_dataframe(pd.read_sql(q_pur_returns, ENGINE), ENGINE).to_excel(writer, sheet_name="Vendor_Returns", index=False)
        decrypt_item_dataframe(pd.read_sql(q_vendor_pricing, ENGINE), ENGINE).to_excel(writer, sheet_name="Vendor_Pricing_History", index=False)
        pd.read_sql(q_vendor_ledger, ENGINE).to_excel(writer, sheet_name="Vendor_Ledger_Balances", index=False)

    wb = openpyxl.load_workbook(filename)
    for ws_name in wb.sheetnames:
        style_worksheet(wb[ws_name], header_fill_hex="4A154B")  # Deep Burgundy
    wb.save(filename)
    logger.info("Module 3 Complete.")

# ======================================================================================
# MODULE 4: CONSOLIDATING FINANCIAL & MONEY LEDGER OPERATIONS
# ======================================================================================
def generate_financial_workbook(output_dir):
    filename = os.path.join(output_dir, "04_Genius_Financial_Ledger.xlsx")
    logger.info(f"Generating Module 4: {filename}...")

    # 1. Cash Safes, Drawers & Banks
    q_cash_stores = """
    SELECT 
        fcs.fcs_id AS Fund_ID,
        fcs.fcs_code AS Code,
        fcs.fcs_name_ar AS Account_Title,
        CASE fcs.fcs_type 
            WHEN '1' THEN 'POS Cashier Drawer (درج كاشير)' 
            WHEN '2' THEN 'Main Treasury/Vault (خزينة رئيسية)' 
            WHEN '3' THEN 'Bank Account (حساب بنكي)' 
            ELSE 'Other Fund' 
        END AS Account_Type,
        fcs.fcs_current_balance AS Current_Live_Balance,
        fcs.fcs_account_number AS Bank_Account_No,
        b.bnk_name_ar AS Associated_Bank,
        fcs.fcs_computer_name AS Bound_Workstation,
        fcs.sec_update_date AS Last_Balance_Update
    FROM F_Cash_Store fcs WITH (NOLOCK)
    LEFT JOIN Bank b WITH (NOLOCK) ON b.bnk_id = fcs.fcs_bank_id
    WHERE fcs.fcs_active = 1
    ORDER BY fcs.fcs_type, fcs.fcs_id;
    """

    # 2. Daily Treasury Transactions
    q_treasury_trans = """
    SELECT 
        fh.fh_id AS Entry_ID,
        fh.fh_actual_date AS Transaction_Date,
        CASE fh.fh_trans_type 
            WHEN '1' THEN 'Cash Movement (نقدية)' 
            WHEN '2' THEN 'Cheque/Commercial Paper (أوراق مالية)' 
            ELSE 'Other' 
        END AS Instrument_Type,
        fh.fh_value AS Amount,
        CASE fh.fh_from_type 
            WHEN '1' THEN 'Customer (عميل)' 
            WHEN '2' THEN 'Vendor (مورد)' 
            WHEN '3' THEN 'Cash Safe/Drawer (خزينة/درج)' 
            ELSE 'Account/Other' 
        END AS Source_Entity_Type,
        fh.fh_from_id AS Source_Entity_ID,
        CASE fh.fh_to_type 
            WHEN '1' THEN 'Customer (عميل)' 
            WHEN '2' THEN 'Vendor (مورد)' 
            WHEN '3' THEN 'Cash Safe/Drawer (خزينة/درج)' 
            ELSE 'Account/Other' 
        END AS Destination_Entity_Type,
        fh.fh_to_id AS Destination_Entity_ID,
        fh.fh_notes AS Operational_Narrative,
        fh.fh_computer AS Executing_Machine,
        emp.e_Name AS Operator_Name
    FROM F_Transaction_Header fh WITH (NOLOCK)
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = fh.sec_insert_uid
    ORDER BY fh.fh_id DESC;
    """

    # 3. Commercial Papers & Cheques Register
    q_cheques = """
    SELECT 
        fhe.fhe_id AS Paper_Registry_ID,
        fhe.fh_id AS Treasury_Header_ID,
        fp.fp_name_ar AS Paper_Type,
        fhe.fhe_paper_number AS Cheque_Paper_Number,
        fhe.fhe_paper_create_date AS Issue_Date,
        fhe.fhe_paper_valid_date AS Due_Maturity_Date,
        b.bnk_name_ar AS Deposit_Bank,
        CASE fhe.fhe_paper_status 
            WHEN '0' THEN 'Pending In-Portfolio (تحت التحصيل)' 
            WHEN '1' THEN 'Deposited/Collected (تم التحصيل)' 
            WHEN '2' THEN 'Rejected/Bounced (مرفوض/مرتد)' 
            WHEN '3' THEN 'Cancelled (ملغي)' 
            ELSE 'Unknown' 
        END AS Cheque_Status,
        fh.fh_value AS Face_Value,
        fhe.fhe_expenses AS Bank_Clearance_Expenses
    FROM F_Transaction_Header_Ext fhe WITH (NOLOCK)
    JOIN F_Transaction_Header fh WITH (NOLOCK) ON fh.fh_id = fhe.fh_id
    LEFT JOIN F_Papers fp WITH (NOLOCK) ON fp.fp_id = fhe.fhe_paper_type
    LEFT JOIN Bank b WITH (NOLOCK) ON b.bnk_id = fhe.fhe_paper_deposit_bank_id
    ORDER BY fhe.fhe_id DESC;
    """

    # 4. Automated GL Document Engine
    q_gl_engine = """
    SELECT 
        fah.fah_id AS Auto_Doc_ID,
        st.sto_name_ar AS Branch_Name,
        fah.fah_type AS System_Trigger,
        fah.fah_form_text AS Originating_Form,
        fah.fah_description AS Journal_Narrative,
        fah.fah_total_value AS Total_Document_Value,
        fad.fad_group_id AS Leg_Group,
        fad.fad_D_Account_name AS Debit_Account,
        fad.fad_D_object AS Debit_Target_Type,
        fad.fad_C_Account_name AS Credit_Account,
        fad.fad_C_object AS Credit_Target_Type,
        fad.fad_value AS Leg_Amount,
        fah.sec_insert_date AS Posting_Timestamp,
        fah.fah_computer_name AS Machine
    FROM F_Auto_Doc_h fah WITH (NOLOCK)
    JOIN F_Auto_Doc_d fad WITH (NOLOCK) ON fad.fad_fah_id = fah.fah_id
    JOIN Store st WITH (NOLOCK) ON st.sto_id = fah.fah_sto_id
    ORDER BY fah.fah_id DESC, fad.fad_group_id ASC;
    """

    # 5. Debit / Credit Adjustment Notices
    q_notices = """
    SELECT 
        fn.fn_id AS Notice_ID,
        CASE fn.fn_type 
            WHEN '1' THEN 'Addition Notice - Credit (إشعار إضافة)' 
            WHEN '2' THEN 'Discount Notice - Debit (إشعار خصم)' 
            ELSE 'Other' 
        END AS Notice_Type,
        CASE fn.fn_who_type 
            WHEN '1' THEN 'Vendor (مورد)' 
            WHEN '2' THEN 'Customer (عميل)' 
            ELSE 'Party' 
        END AS Party_Type,
        fn.fn_who_id AS Party_ID,
        fnc.fnc_name_ar AS Reason_Category,
        fn.fn_value AS Notice_Amount,
        fn.fn_notes AS Explanation,
        fn.sec_insert_date AS Issued_Timestamp,
        emp.e_Name AS Officer_Name
    FROM F_Notices fn WITH (NOLOCK)
    LEFT JOIN F_Notices_Causes fnc WITH (NOLOCK) ON fnc.fnc_id = fn.fn_cause_id
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = fn.sec_insert_uid
    ORDER BY fn.fn_id DESC;
    """

    # 6. Customer Credit & Debt Roster
    q_cust_debt = """
    SELECT 
        c.cust_id AS Customer_ID,
        c.cust_code AS Customer_Code,
        c.cust_name_ar AS Customer_Name,
        c.cust_mobile AS Phone_Number,
        c.cust_max_credit AS Authorized_Credit_Limit,
        c.cust_start_debit AS Starting_Debit,
        c.cust_start_credit AS Starting_Credit,
        c.cust_current_credit AS Current_Balance,
        CASE 
            WHEN c.cust_current_credit > 0 THEN 'Customer Owes Pharmacy (مدين)' 
            WHEN c.cust_current_credit < 0 THEN 'Pharmacy Owes Customer (دائن)' 
            ELSE 'Settled (خالص)' 
        END AS Position_Status,
        c.sec_update_date AS Last_Credit_Change
    FROM Customer c WITH (NOLOCK)
    WHERE c.cust_active = '1'
    ORDER BY c.cust_current_credit DESC;
    """

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        pd.read_sql(q_cash_stores, ENGINE).to_excel(writer, sheet_name="Cash_Safes_Banks", index=False)
        pd.read_sql(q_treasury_trans, ENGINE).to_excel(writer, sheet_name="Treasury_Ledger", index=False)
        pd.read_sql(q_cheques, ENGINE).to_excel(writer, sheet_name="Cheques_Portfolio", index=False)
        pd.read_sql(q_gl_engine, ENGINE).to_excel(writer, sheet_name="Automated_GL_Journals", index=False)
        pd.read_sql(q_notices, ENGINE).to_excel(writer, sheet_name="Debit_Credit_Notices", index=False)
        pd.read_sql(q_cust_debt, ENGINE).to_excel(writer, sheet_name="Customer_Credit_Balances", index=False)

    wb = openpyxl.load_workbook(filename)
    for ws_name in wb.sheetnames:
        style_worksheet(wb[ws_name], header_fill_hex="7D5A00")  # Dark Gold / Bronze
    wb.save(filename)
    logger.info("Module 4 Complete.")

# ======================================================================================
# MODULE 5: CONSOLIDATING SHIFTS & POS OPERATIONS
# ======================================================================================
def generate_shifts_workbook(output_dir):
    filename = os.path.join(output_dir, "05_Genius_POS_Shift_Operations.xlsx")
    logger.info(f"Generating Module 5: {filename}...")

    # 1. POS Cashier Shifts Closure Summary
    q_shifts = """
    SELECT 
        sc.sc_id AS Shift_Log_ID,
        emp.e_Name AS Cashier_Name,
        fcs.fcs_name_ar AS Cash_Drawer,
        sc.sc_shift_start_time AS Shift_Started_At,
        sc.sec_insert_date AS Shift_Closed_At,
        sc.sc_start_cash AS Starting_Float_Cash,
        sc.sc_curr_cash AS System_Calculated_Cash,
        sc.sc_act_cash AS Cashier_Actual_Counted_Cash,
        CAST(sc.sc_act_cash - sc.sc_curr_cash AS DECIMAL(18,2)) AS Cash_Variance_Over_Short,
        CASE 
            WHEN (sc.sc_act_cash - sc.sc_curr_cash) > 0 THEN 'Surplus (زيادة)' 
            WHEN (sc.sc_act_cash - sc.sc_curr_cash) < 0 THEN 'Deficit/Shortage (عجز)' 
            ELSE 'Balanced (مطابق)' 
        END AS Reconciliation_Status,
        sc.sc_trans_value AS Amount_Transferred_To_Vault,
        next_emp.e_Name AS Transferred_To_Employee,
        sc.sc_curr_credit AS Visa_System_Total,
        sc.sc_act_credit AS Visa_Actual_Slip_Total,
        sc.sc_notes AS Shift_Notes
    FROM Shift_Changes sc WITH (NOLOCK)
    JOIN Employee emp WITH (NOLOCK) ON emp.e_id = sc.sc_cashier_id
    JOIN F_Cash_Store fcs WITH (NOLOCK) ON fcs.fcs_id = sc.sc_fcs_id
    LEFT JOIN Employee next_emp WITH (NOLOCK) ON next_emp.e_id = sc.sc_to_employee
    ORDER BY sc.sc_id DESC;
    """

    # 2. Shift Operations Breakdown Details
    q_shifts_d = """
    SELECT 
        scd.scd_sc_id AS Shift_Log_ID,
        st.sto_name_ar AS Store_Name,
        fcs.fcs_name_ar AS Cash_Drawer,
        scd.scd_start AS Shift_Open,
        scd.scd_end AS Shift_Close,
        scd.scd_d_startBalance AS Opening_Balance,
        scd.scd_d_sales AS Cash_Sales_In,
        scd.scd_d_salesfromreturn AS Sales_From_Returns,
        scd.scd_d_visacollect AS Visa_Collections,
        scd.scd_d_cashIn AS Manual_Cash_In,
        scd.scd_c_purchases AS Outbound_Cash_Purchases,
        scd.scd_c_salesReturn AS Outbound_Sales_Refunds,
        scd.scd_c_cashOut AS Manual_Cash_Out_Expenses,
        scd.scd_o_salesVisa AS Visa_Card_Sales,
        scd.scd_o_salesCredit AS Credit_Account_Sales,
        scd.scd_o_salesPending AS Suspended_Bills_Value
    FROM Shift_Changes_d scd WITH (NOLOCK)
    JOIN Store st WITH (NOLOCK) ON st.sto_id = scd.scd_sto_id
    JOIN F_Cash_Store fcs WITH (NOLOCK) ON fcs.fcs_id = scd.scd_fcs_id
    ORDER BY scd.scd_sc_id DESC;
    """

    # 3. User Login / Logout Station Sessions
    q_sessions = """
    SELECT 
        lu.lg_id AS Session_ID,
        usr.usr_name AS System_Username,
        emp.e_Name AS Employee_Full_Name,
        fcs.fcs_name_ar AS POS_Drawer_Assigned,
        lu.computer_nm AS Client_Workstation,
        lu.prog_ver AS ePlus_Software_Version,
        lu.lg_begin_tm AS Login_Timestamp,
        lu.lg_close_tm AS Logout_Timestamp,
        DATEDIFF(MINUTE, lu.lg_begin_tm, lu.lg_close_tm) AS Session_Duration_Minutes
    FROM log_users lu WITH (NOLOCK)
    LEFT JOIN Users usr WITH (NOLOCK) ON usr.usr_e_id = CAST(lu.lg_usr_id AS DECIMAL(18,0))
    LEFT JOIN Employee emp WITH (NOLOCK) ON emp.e_id = usr.usr_e_id
    LEFT JOIN F_Cash_Store fcs WITH (NOLOCK) ON fcs.fcs_id = lu.f_cash_store_id
    ORDER BY lu.lg_id DESC;
    """

    # 4. Staff Shift Schedule & Privileges
    q_staff = """
    SELECT 
        e.e_id AS Employee_ID,
        e.e_code AS Employee_Code,
        e.e_Name AS Staff_Name,
        j.j_name_ar AS Job_Title,
        ews.ews_name AS Assigned_Shift_Scheme,
        ews.ews_start AS Shift_Default_Start,
        ews.ews_end AS Shift_Default_End,
        e.e_bill_disc_perc AS Max_Allowed_Bill_Discount,
        CASE e.e_select_cashier WHEN 1 THEN 'Yes' ELSE 'No' END AS Can_Select_Cashier,
        CASE e.e_change_sales_prices WHEN 1 THEN 'Yes' ELSE 'No' END AS Can_Override_Prices,
        CASE e.e_can_pend WHEN 1 THEN 'Yes' ELSE 'No' END AS Can_Suspend_Bills,
        e.sec_update_date AS Profile_Last_Updated
    FROM Employee e WITH (NOLOCK)
    LEFT JOIN Jobs j WITH (NOLOCK) ON j.j_id = e.e_job_id
    LEFT JOIN Emp_Work_Shifts ews WITH (NOLOCK) ON ews.ews_id = e.e_ews_id
    WHERE e.e_active = 1
    ORDER BY e.e_id ASC;
    """

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        pd.read_sql(q_shifts, ENGINE).to_excel(writer, sheet_name="Shift_Reconciliation", index=False)
        pd.read_sql(q_shifts_d, ENGINE).to_excel(writer, sheet_name="Shift_Operations_Breakdown", index=False)
        pd.read_sql(q_sessions, ENGINE).to_excel(writer, sheet_name="Station_Login_Sessions", index=False)
        pd.read_sql(q_staff, ENGINE).to_excel(writer, sheet_name="Cashier_Staff_Roster", index=False)

    wb = openpyxl.load_workbook(filename)
    for ws_name in wb.sheetnames:
        style_worksheet(wb[ws_name], header_fill_hex="1A5276")  # Slate Teal
    wb.save(filename)
    logger.info("Module 5 Complete.")

# ======================================================================================
# MODULE 6: EXECUTIVE BUSINESS INTELLIGENCE & ANALYTICAL COCKPIT
# ======================================================================================
def generate_executive_bi_workbook(output_dir):
    filename = os.path.join(output_dir, "06_Genius_Executive_Intelligence_Cockpit.xlsx")
    logger.info(f"Generating Module 6: {filename}...")

    # 1. Executive Master KPI Scorecard
    q_kpi_scorecard = """
    SELECT '1. Total Gross Sales Revenue' AS Metric, 
           CAST(ISNULL(SUM(total_bill), 0) AS DECIMAL(18,2)) AS Value_L_E, 
           'Total sales before any discounts across all history' AS Notes FROM sales_trans_h WITH (NOLOCK) WHERE sth_flag='C'
    UNION ALL
    SELECT '2. Total Discounts Conceded', 
           CAST(ISNULL(SUM(total_bill - total_bill_after_disc), 0) AS DECIMAL(18,2)), 
           'Customer discounts granted at header level' FROM sales_trans_h WITH (NOLOCK) WHERE sth_flag='C'
    UNION ALL
    SELECT '3. Total Net Sales Realized', 
           CAST(ISNULL(SUM(total_bill_after_disc), 0) AS DECIMAL(18,2)), 
           'Actual top-line revenue generated' FROM sales_trans_h WITH (NOLOCK) WHERE sth_flag='C'
    UNION ALL
    SELECT '4. Total Cost of Goods Sold (COGS)', 
           CAST(ISNULL(SUM(d.qnty * d.itm_cost), 0) AS DECIMAL(18,2)), 
           'Direct pharmacy acquisition cost of sold merchandise' FROM sales_trans_d d WITH (NOLOCK) JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id=d.sth_id WHERE h.sth_flag='C'
    UNION ALL
    SELECT '5. Cumulative Gross Profit', 
           CAST(ISNULL(SUM((d.qnty * d.itm_sell * (1.0 - (d.itm_dis_per/100.0))) - d.itm_dis_mon) - SUM(d.qnty * d.itm_cost), 0) AS DECIMAL(18,2)), 
           'Net Revenue minus Total Product Cost' FROM sales_trans_d d WITH (NOLOCK) JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id=d.sth_id WHERE h.sth_flag='C'
    UNION ALL
    SELECT '6. Total Current Inventory Valuation (At Cost)', 
           CAST(ISNULL(SUM(itm_qty * pharm_price), 0) AS DECIMAL(18,2)), 
           'Total capital currently tied up in warehouse & branches' FROM Item_Class_Store WITH (NOLOCK) WHERE itm_qty > 0
    UNION ALL
    SELECT '7. Total Current Inventory Valuation (At Retail)', 
           CAST(ISNULL(SUM(itm_qty * sell_price), 0) AS DECIMAL(18,2)), 
           'Potential retail revenue of on-shelf stock' FROM Item_Class_Store WITH (NOLOCK) WHERE itm_qty > 0
    UNION ALL
    SELECT '8. Unrealized Potential Profit on Shelves', 
           CAST(ISNULL(SUM(itm_qty * (sell_price - pharm_price)), 0) AS DECIMAL(18,2)), 
           'Unrealized gross margin sitting in inventory' FROM Item_Class_Store WITH (NOLOCK) WHERE itm_qty > 0
    UNION ALL
    SELECT '9. Total Customer Outstanding Debt (Accounts Receivable)', 
           CAST(ISNULL(SUM(cust_current_credit), 0) AS DECIMAL(18,2)), 
           'Total money owed to pharmacy by credit customers' FROM Customer WITH (NOLOCK) WHERE cust_current_credit > 0
    UNION ALL
    SELECT '10. Total Vendor Liabilities (Accounts Payable)', 
           CAST(ISNULL(ABS(SUM(ven_current_credit)), 0) AS DECIMAL(18,2)), 
           'Total unpaid invoices owed to pharmaceutical distributors' FROM Vendor WITH (NOLOCK) WHERE ven_current_credit < 0
    UNION ALL
    SELECT '11. Total Liquidity in Safes & Banks', 
           CAST(ISNULL(SUM(fcs_current_balance), 0) AS DECIMAL(18,2)), 
           'Total live cash available across all drawers, vaults and bank accounts' FROM F_Cash_Store WITH (NOLOCK) WHERE fcs_active = 1
    UNION ALL
    SELECT '12. Net POS Cash Drawer Discrepancies (Shortage/Surplus)', 
           CAST(ISNULL(SUM(sc_act_cash - sc_curr_cash), 0) AS DECIMAL(18,2)), 
           'Cumulative historical cashier cash drawer variance' FROM Shift_Changes WITH (NOLOCK);
    """

    # 2. ABC Inventory & Revenue Pareto Analysis
    q_abc_analysis = """
    WITH ItemSales AS (
        SELECT 
            ic.itm_id,
            ic.itm_code AS Barcode,
            ic.itm_name_ar AS Item_Name,
            c.com_name_ar AS Manufacturer,
            SUM(d.qnty) AS Units_Sold,
            CAST(SUM((d.qnty * d.itm_sell * (1.0 - (d.itm_dis_per/100.0))) - d.itm_dis_mon) AS DECIMAL(18,2)) AS Net_Revenue,
            CAST(SUM(d.qnty * d.itm_cost) AS DECIMAL(18,2)) AS Total_COGS,
            CAST(SUM((d.qnty * d.itm_sell * (1.0 - (d.itm_dis_per/100.0))) - d.itm_dis_mon) - SUM(d.qnty * d.itm_cost) AS DECIMAL(18,2)) AS Gross_Profit
        FROM sales_trans_d d WITH (NOLOCK)
        JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id = d.sth_id
        JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = d.itm_id
        LEFT JOIN Company c WITH (NOLOCK) ON c.com_id = ic.com_id
        WHERE h.sth_flag = 'C'
        GROUP BY ic.itm_id, ic.itm_code, ic.itm_name_ar, c.com_name_ar
    ),
    RankedSales AS (
        SELECT *,
            SUM(Net_Revenue) OVER () AS Total_Market_Revenue,
            SUM(Net_Revenue) OVER (ORDER BY Net_Revenue DESC ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS Running_Revenue
        FROM ItemSales
    )
    SELECT 
        itm_id AS Item_ID,
        Barcode,
        Item_Name,
        Manufacturer,
        Units_Sold,
        Net_Revenue,
        Total_COGS,
        Gross_Profit,
        CAST((Gross_Profit / NULLIF(Net_Revenue, 0)) AS DECIMAL(5,4)) AS Gross_Margin_Perc,
        CAST((Running_Revenue / NULLIF(Total_Market_Revenue, 0)) AS DECIMAL(5,4)) AS Cumulative_Revenue_Share,
        CASE 
            WHEN (Running_Revenue / NULLIF(Total_Market_Revenue, 0)) <= 0.80 THEN 'Class A (Top 80% Rev)'
            WHEN (Running_Revenue / NULLIF(Total_Market_Revenue, 0)) <= 0.95 THEN 'Class B (Next 15% Rev)'
            ELSE 'Class C (Bottom 5% Rev)'
        END AS ABC_Classification
    FROM RankedSales
    ORDER BY Net_Revenue DESC;
    """

    # 3. Expiry Risk & Dead Stock Radar
    q_dead_stock_radar = """
    SELECT 
        ics.sto_id AS Store_ID,
        st.sto_name_ar AS Store_Name,
        ic.itm_id AS Item_ID,
        ic.itm_code AS Barcode,
        ic.itm_name_ar AS Item_Name,
        ics.c_id AS Batch_ID,
        ics.itm_qty AS Stock_On_Hand,
        ics.pharm_price AS Unit_Cost,
        CAST(ics.itm_qty * ics.pharm_price AS DECIMAL(18,2)) AS Capital_At_Risk,
        CONVERT(VARCHAR(10), ics.itm_expiry_date, 120) AS Expiry_Date,
        DATEDIFF(DAY, GETDATE(), ics.itm_expiry_date) AS Days_Until_Expiration,
        CASE 
            WHEN ics.itm_expiry_date < GETDATE() THEN 'EXPIRED (منتهي الصلاحية)'
            WHEN DATEDIFF(DAY, GETDATE(), ics.itm_expiry_date) <= 90 THEN 'CRITICAL (< 90 Days)'
            WHEN DATEDIFF(DAY, GETDATE(), ics.itm_expiry_date) <= 180 THEN 'WARNING (< 180 Days)'
            ELSE 'STABLE'
        END AS Expiry_Risk_Category,
        ISNULL((
            SELECT MAX(h.sec_insert_date) 
            FROM sales_trans_d d WITH (NOLOCK) 
            JOIN sales_trans_h h WITH (NOLOCK) ON h.sth_id = d.sth_id 
            WHERE d.itm_id = ics.itm_id
        ), '2000-01-01') AS Last_Sold_Date
    FROM Item_Class_Store ics WITH (NOLOCK)
    JOIN Item_Catalog ic WITH (NOLOCK) ON ic.itm_id = ics.itm_id
    JOIN Store st WITH (NOLOCK) ON st.sto_id = ics.sto_id
    WHERE ics.itm_qty > 0 
      AND (ics.itm_expiry_date <= DATEADD(MONTH, 6, GETDATE()) OR ics.itm_expiry_date IS NULL)
    ORDER BY Days_Until_Expiration ASC;
    """

    # 4. Hourly Sales Distribution & Traffic Velocity
    q_hourly_traffic = """
    SELECT 
        DATEPART(HOUR, sec_insert_date) AS Hour_Of_Day_24H,
        COUNT(sth_id) AS Total_Bills_Count,
        CAST(SUM(total_bill_after_disc) AS DECIMAL(18,2)) AS Total_Revenue_Generated,
        CAST(AVG(total_bill_after_disc) AS DECIMAL(18,2)) AS Average_Basket_Size,
        CAST(SUM(total_bill_after_disc) / NULLIF((SELECT SUM(total_bill_after_disc) FROM sales_trans_h WHERE sth_flag='C'), 0) AS DECIMAL(5,4)) AS Share_Of_Total_Sales
    FROM sales_trans_h WITH (NOLOCK)
    WHERE sth_flag = 'C'
    GROUP BY DATEPART(HOUR, sec_insert_date)
    ORDER BY Hour_Of_Day_24H ASC;
    """

    # 5. Distributor / Vendor Supply Performance Matrix
    q_vendor_performance = """
    SELECT 
        v.ven_id AS Vendor_ID,
        v.ven_name_ar AS Vendor_Name,
        COUNT(DISTINCT pth.pth_id) AS Total_Purchase_Bills,
        CAST(SUM(pth.total_bill) AS DECIMAL(18,2)) AS Total_Purchased_Gross,
        CAST(SUM(pth.total_after_back) AS DECIMAL(18,2)) AS Total_Purchased_Net,
        CAST(AVG(pth.total_dis_per) AS DECIMAL(5,2)) AS Avg_Negotiated_Discount_Perc,
        ISNULL((
            SELECT CAST(SUM(ibh_total_value) AS DECIMAL(18,2)) 
            FROM Item_back_h WITH (NOLOCK) 
            WHERE ibh_ven_id = v.ven_id
        ), 0.00) AS Total_Returned_Value,
        v.ven_current_credit AS Current_Outstanding_Debt
    FROM Vendor v WITH (NOLOCK)
    JOIN pur_trans_h pth WITH (NOLOCK) ON pth.ven_id = v.ven_id
    WHERE pth.bill_status <> 1
    GROUP BY v.ven_id, v.ven_name_ar, v.ven_current_credit
    ORDER BY Total_Purchased_Net DESC;
    """

    # 6. Branch / Store Operational Matrix
    q_store_matrix = """
    SELECT 
        st.sto_id AS Store_ID,
        st.sto_name_ar AS Store_Name,
        ISNULL((SELECT COUNT(*) FROM Item_Class_Store WHERE sto_id=st.sto_id AND itm_qty > 0), 0) AS Active_SKUs_In_Stock,
        ISNULL((SELECT CAST(SUM(itm_qty * pharm_price) AS DECIMAL(18,2)) FROM Item_Class_Store WHERE sto_id=st.sto_id AND itm_qty > 0), 0) AS Total_Stock_Cost_Valuation,
        ISNULL((SELECT CAST(SUM(itm_qty * sell_price) AS DECIMAL(18,2)) FROM Item_Class_Store WHERE sto_id=st.sto_id AND itm_qty > 0), 0) AS Total_Stock_Retail_Valuation,
        ISNULL((SELECT COUNT(*) FROM sales_trans_h WHERE sto_id=st.sto_id AND sth_flag='C'), 0) AS Total_Sales_Transactions,
        ISNULL((SELECT CAST(SUM(total_bill_after_disc) AS DECIMAL(18,2)) FROM sales_trans_h WHERE sto_id=st.sto_id AND sth_flag='C'), 0) AS Total_Sales_Revenue,
        ISNULL((SELECT CAST(SUM(sc_act_cash - sc_curr_cash) AS DECIMAL(18,2)) FROM Shift_Changes sc JOIN F_Cash_Store fcs ON fcs.fcs_id=sc.sc_fcs_id WHERE fcs.fcs_code=st.sto_code), 0) AS Cumulative_Shift_Variance
    FROM Store st WITH (NOLOCK)
    WHERE st.activated = 1;
    """

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        pd.read_sql(q_kpi_scorecard, ENGINE).to_excel(writer, sheet_name="Executive_KPI_Summary", index=False)
        decrypt_item_dataframe(pd.read_sql(q_abc_analysis, ENGINE), ENGINE).to_excel(writer, sheet_name="ABC_Revenue_Pareto", index=False)
        decrypt_item_dataframe(pd.read_sql(q_dead_stock_radar, ENGINE), ENGINE).to_excel(writer, sheet_name="Expiry_&_DeadStock_Radar", index=False)
        pd.read_sql(q_hourly_traffic, ENGINE).to_excel(writer, sheet_name="Hourly_Traffic_Patterns", index=False)
        pd.read_sql(q_vendor_performance, ENGINE).to_excel(writer, sheet_name="Vendor_Supply_Matrix", index=False)
        pd.read_sql(q_store_matrix, ENGINE).to_excel(writer, sheet_name="MultiBranch_Benchmark", index=False)

    wb = openpyxl.load_workbook(filename)
    # Executive Dark Navy Theme
    for ws_name in wb.sheetnames:
        style_worksheet(wb[ws_name], header_fill_hex="0B132B", is_kpi=(ws_name == "Executive_KPI_Summary"))
    wb.save(filename)
    logger.info("Module 6 Complete.")

# ======================================================================================
# MASTER WORKFLOW ORCHESTRATOR
# ======================================================================================
def main():
    start_time = time.time()
    logger.info("==========================================================================")
    logger.info("STARTING GENIUS ERP COMPREHENSIVE DATA CONSOLIDATION PIPELINE")
    logger.info("==========================================================================")

    output_dir = os.path.join(os.getcwd(), "Genius_Consolidated_Reports")
    os.makedirs(output_dir, exist_ok=True)
    logger.info(f"Target Output Directory: {output_dir}")

    # Test Database Connectivity
    try:
        with ENGINE.connect() as conn:
            conn.execute(text("SELECT 1"))
        logger.info("Database Connection Test: SUCCESSFUL.")
    except Exception as e:
        logger.error(f"Failed to connect to Microsoft SQL Server: {e}")
        sys.exit(1)

    # Execute All 6 Modules Sequentially
    try:
        generate_inventory_workbook(output_dir)
        generate_sales_workbook(output_dir)
        generate_purchases_workbook(output_dir)
        generate_financial_workbook(output_dir)
        generate_shifts_workbook(output_dir)
        generate_executive_bi_workbook(output_dir)
    except Exception as e:
        logger.error(f"Extraction halted due to an unexpected exception: {e}", exc_info=True)
        sys.exit(1)

    elapsed = round(time.time() - start_time, 2)
    logger.info("==========================================================================")
    logger.info(f"ALL 6 MASTER WORKBOOKS SUCCESSFULLY GENERATED IN {elapsed} SECONDS.")
    logger.info(f"Location: {output_dir}")
    logger.info("==========================================================================")

if __name__ == "__main__":
    main()