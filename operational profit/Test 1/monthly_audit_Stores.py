import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

# ==============================================================================
# 1. DATABASE CONNECTION & CONFIGURATION
# ==============================================================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

START_DATE = "2025-01-01 00:00:00"
END_DATE = "2026-09-30 23:59:59"

params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USERNAME};"
    f"PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

# ==============================================================================
# 2. RAW SQL EXTRACTION QUERIES (RESOLVING STORE 0 AND ANOMALIES)
# ==============================================================================

# A. Raw Sales Lines (Store specific)
SQL_SALES = text("""
    SELECT 
        h.sto_id,
        h.sec_insert_date AS sale_date,
        d.qnty,
        ISNULL(d.itm_back, 0) AS itm_back,
        CAST(d.itm_sell AS float) AS itm_sell,
        CAST(ISNULL(d.itm_cost, 0) AS float) AS itm_cost,
        CAST(ISNULL(d.itm_dis_per, 0) AS float) AS itm_dis_per,
        CAST(ISNULL(d.itm_dis_mon, 0) AS float) AS itm_dis_mon,
        CAST(ISNULL(h.total_dis_per, 0) AS float) AS bill_dis_per
    FROM sales_trans_d d
    INNER JOIN sales_trans_h h ON h.sth_id = d.sth_id
    WHERE h.sto_id = :sto_id
      AND h.sth_flag = 'C'
      AND h.sec_insert_date BETWEEN :start_date AND :end_date
""")

# B. Real Cash & Visa Collected (With 8M typo filter)
SQL_COLLECTIONS = text("""
    SELECT 
        stp.stp_sto_id AS sto_id,
        stp.stp_insert_date AS trans_date,
        stp.stp_pt_id, -- 1 = Cash, 2 = Visa
        CAST(stp.stp_value AS float) AS amount
    FROM sales_trans_payment stp
    WHERE stp.stp_sto_id = :sto_id
      AND stp.stp_value < 1000000 -- Excludes the 8,010,871.00 cashier typo from Notice #89
      AND stp.stp_insert_date BETWEEN :start_date AND :end_date
""")

# C. Client Service Inflows (Sugar, Pressure, Compounding, Direct Cash-In)
SQL_SERVICE_CASH_IN = text("""
    SELECT 
        ISNULL(fh.fh_sto_id, 47) AS sto_id,
        fh.fh_actual_date AS trans_date,
        CAST(fh.fh_value AS float) AS cash_in_amount
    FROM F_Transaction_Header fh
    WHERE (fh.fh_sto_id = :sto_id OR (fh.fh_sto_id = 0 AND :sto_id = 47) OR fh.fh_sto_id IS NULL)
      AND fh.fh_trans_type = '1' -- Cash
      AND fh.fh_from_type = '1'  -- Customer Service Profile
      AND fh.fh_to_type = '3'    -- To POS Drawer
      AND fh.fh_notes NOT LIKE '%فاتورة بيع%' -- Non-invoice service cash
      AND fh.fh_actual_date BETWEEN :start_date AND :end_date
""")

# D. New Purchases at Wholesale Cost
SQL_PURCHASES = text("""
    SELECT 
        pth.sto_id,
        ISNULL(pth.ven_bill_date, pth.sec_insert_date) AS pur_date,
        ptd.qnty,
        ISNULL(ptd.bonus, 0) AS bonus,
        CAST(ptd.itm_pur_price AS float) AS itm_pur_price,
        CAST(ISNULL(ptd.itm_tax_price, 0) AS float) AS itm_tax_price,
        CAST(ISNULL(ptd.itm_extra_dis, 0) AS float) AS itm_extra_dis,
        CAST(ISNULL(pth.total_dis_per, 0) AS float) AS bill_dis_per
    FROM pur_trans_d ptd
    INNER JOIN pur_trans_h pth ON pth.pth_id = ptd.pth_id
    WHERE pth.sto_id = :sto_id
      AND pth.bill_status <> 1
      AND ISNULL(pth.ven_bill_date, pth.sec_insert_date) BETWEEN :start_date AND :end_date
""")

# E. Operating Expenses (From Shift Ledgers + General Ledger)
SQL_EXPENSES = text("""
    SELECT 
        scd.scd_sto_id AS sto_id,
        scd.scd_start AS expense_date,
        CAST(ISNULL(scd.scd_c_cashOut, 0) AS float) AS expense_amount
    FROM Shift_Changes_d scd
    WHERE scd.scd_sto_id = :sto_id
      AND scd.scd_start BETWEEN :start_date AND :end_date
      AND scd.scd_c_cashOut > 0
    UNION ALL
    SELECT 
        ISNULL(fh.fh_sto_id, 47) AS sto_id,
        fh.fh_actual_date AS expense_date,
        CAST(fh.fh_value AS float) AS expense_amount
    FROM F_Transaction_Header fh
    WHERE (fh.fh_sto_id = :sto_id OR (fh.fh_sto_id = 0 AND :sto_id = 47) OR fh.fh_sto_id IS NULL)
      AND fh.fh_trans_type = '1'
      AND fh.fh_from_type = '3'
      AND fh.fh_to_type IN ('5', '0')
      AND fh.fh_value < 1000000 -- Excludes accidental typos
      AND fh.fh_notes NOT LIKE '%فاتورة بيع%'
      AND fh.fh_actual_date BETWEEN :start_date AND :end_date
""")

# F. Expired Stock Write-Offs & Shelf Losses
SQL_EXPIRED_STOCK = text("""
    SELECT 
        ISNULL(iqc.iqc_store_id, 47) AS sto_id,
        iqc.sec_insert_date AS writeoff_date,
        CAST((iqc.iqc_old_quantity - iqc.iqc_new_quantity) * iqc.iqc_pharm_price AS float) AS expired_cost
    FROM Item_Quantity_Changes iqc
    WHERE (iqc.iqc_store_id = :sto_id OR (iqc.iqc_store_id = 0 AND :sto_id = 47))
      AND iqc.iqc_dc_id IN (1, 2) -- 2 = Expired, 1 = Damaged
      AND (iqc.iqc_old_quantity - iqc.iqc_new_quantity) > 0
      AND iqc.sec_insert_date BETWEEN :start_date AND :end_date
""")

# G. Baseline Inventory Valuation
SQL_CURRENT_STOCK = text("""
    SELECT 
        sto_id, 
        CAST(SUM(itm_qty * pharm_price) AS float) AS live_inventory_cost 
    FROM Item_Class_Store 
    WHERE itm_qty > 0 
    GROUP BY sto_id
""")

# ==============================================================================
# 3. REPORT GENERATION ENGINE
# ==============================================================================
def generate_reports():
    print("[1/4] Loading live inventory balances...")
    with engine.connect() as conn:
        df_live_stock = pd.read_sql(SQL_CURRENT_STOCK, conn).set_index("sto_id")

    months = pd.period_range(start="2025-01", end="2026-09", freq="M").astype(str)

    # --------------------------------------------------------------------------
    # PROCESS STORE 1 (MAIN PHARMACY) & STORE 9 (TAHT EL-TASREEF)
    # --------------------------------------------------------------------------
    for store in [{"sto_id": 47, "code": "1"}, {"sto_id": 48, "code": "9"}]:
        sto_id = store["sto_id"]
        sto_code = store["code"]
        print(f"\n[2/4] Processing Store {sto_code} (Database sto_id: {sto_id})...")

        params_dict = {"sto_id": sto_id, "start_date": START_DATE, "end_date": END_DATE}

        with engine.connect() as conn:
            df_sales = pd.read_sql(SQL_SALES, conn, params=params_dict)
            df_col = pd.read_sql(SQL_COLLECTIONS, conn, params=params_dict)
            df_srv = pd.read_sql(SQL_SERVICE_CASH_IN, conn, params=params_dict)
            df_pur = pd.read_sql(SQL_PURCHASES, conn, params=params_dict)
            df_exp = pd.read_sql(SQL_EXPENSES, conn, params=params_dict)
            df_exp_stock = pd.read_sql(SQL_EXPIRED_STOCK, conn, params=params_dict)

        m_df = pd.DataFrame(index=months)

        # 1. Sales & Two-Tier Discounts
        if not df_sales.empty:
            df_sales["Net_Qty"] = df_sales["qnty"] - df_sales["itm_back"]
            df_sales = df_sales[df_sales["Net_Qty"] > 0].copy()

            df_sales["Gross"] = df_sales["Net_Qty"] * df_sales["itm_sell"]
            adj_mon = (df_sales["itm_dis_mon"] * (df_sales["Net_Qty"] / df_sales["qnty"])).fillna(0)
            line_net = (df_sales["Gross"] * ((100.0 - df_sales["itm_dis_per"]) / 100.0)) - adj_mon
            df_sales["Net_Sales"] = (line_net.clip(lower=0.0) * ((100.0 - df_sales["bill_dis_per"]) / 100.0)).clip(lower=0.0)
            df_sales["Discounts"] = df_sales["Gross"] - df_sales["Net_Sales"]
            df_sales["COGS"] = df_sales["Net_Qty"] * df_sales["itm_cost"]

            df_sales["Month"] = pd.to_datetime(df_sales["sale_date"]).dt.to_period("M").astype(str)
            s_grp = df_sales.groupby("Month").agg(
                Gross_Sales=("Gross", "sum"),
                Discounts_Given=("Discounts", "sum"),
                Net_Sales=("Net_Sales", "sum"),
                COGS=("COGS", "sum")
            )
            m_df = m_df.join(s_grp)
        else:
            m_df["Gross_Sales"] = m_df["Discounts_Given"] = m_df["Net_Sales"] = m_df["COGS"] = 0.0

        m_df.fillna(0.0, inplace=True)

        # 2. Collections (Cash & Visa)
        if not df_col.empty:
            df_col["Month"] = pd.to_datetime(df_col["trans_date"]).dt.to_period("M").astype(str)
            cash = df_col[df_col["stp_pt_id"] == 1].groupby("Month")["amount"].sum().rename("Real_Cash")
            visa = df_col[df_col["stp_pt_id"] == 2].groupby("Month")["amount"].sum().rename("Real_Visa")
            m_df = m_df.join(cash).join(visa)
        else:
            m_df["Real_Cash"] = m_df["Real_Visa"] = 0.0

        m_df.fillna(0.0, inplace=True)

        # 3. Service Cash Inflows
        if not df_srv.empty:
            df_srv["Month"] = pd.to_datetime(df_srv["trans_date"]).dt.to_period("M").astype(str)
            m_df = m_df.join(df_srv.groupby("Month")["cash_in_amount"].sum().rename("Service_Cash_In"))
        else:
            m_df["Service_Cash_In"] = 0.0

        m_df.fillna(0.0, inplace=True)

        # 4. Purchases
        if not df_pur.empty:
            base_p = df_pur["itm_pur_price"] * ((100.0 - df_pur["itm_extra_dis"]) / 100.0)
            net_u = base_p * ((100.0 - df_pur["bill_dis_per"]) / 100.0)
            df_pur["Pur_Cost"] = (df_pur["qnty"] * (net_u + df_pur["itm_tax_price"])) + (df_pur["bonus"] * df_pur["itm_tax_price"])
            df_pur["Month"] = pd.to_datetime(df_pur["pur_date"]).dt.to_period("M").astype(str)
            m_df = m_df.join(df_pur.groupby("Month")["Pur_Cost"].sum().rename("Purchases_Cost"))
        else:
            m_df["Purchases_Cost"] = 0.0

        m_df.fillna(0.0, inplace=True)

        # 5. Operating Expenses
        if not df_exp.empty:
            df_exp["Month"] = pd.to_datetime(df_exp["expense_date"]).dt.to_period("M").astype(str)
            m_df = m_df.join(df_exp.groupby("Month")["expense_amount"].sum().rename("Expenses"))
        else:
            m_df["Expenses"] = 0.0

        m_df.fillna(0.0, inplace=True)

        # 6. Expired Stock Losses
        if not df_exp_stock.empty:
            df_exp_stock["Month"] = pd.to_datetime(df_exp_stock["writeoff_date"]).dt.to_period("M").astype(str)
            m_df = m_df.join(df_exp_stock.groupby("Month")["expired_cost"].sum().rename("Expired_Loss"))
        else:
            m_df["Expired_Loss"] = 0.0

        m_df.fillna(0.0, inplace=True)

        # 7. Inventory Trajectory (Perpetual Inventory Equation)
        live_val = df_live_stock.loc[sto_id, "live_inventory_cost"] if sto_id in df_live_stock.index else 500000.0
        m_df["Inventory_Delta"] = m_df["Purchases_Cost"] - m_df["COGS"] - m_df["Expired_Loss"]
        
        # Back-calculate monthly inventory balances leading up to the live valuation
        cum_deltas = m_df["Inventory_Delta"].cumsum()
        final_cum = cum_deltas.iloc[-1]
        base_inventory = live_val - final_cum
        m_df["Total_Stock_Cost"] = base_inventory + cum_deltas

        # 8. Derived KPIs
        m_df["POS_Gross_Profit"] = m_df["Net_Sales"] - m_df["COGS"]
        m_df["Gross_Margin_Pct"] = ((m_df["POS_Gross_Profit"] / m_df["Net_Sales"]) * 100.0).fillna(0.0).round(2)
        m_df["Procurement_Ratio"] = ((m_df["Purchases_Cost"] / m_df["COGS"])).replace([float('inf'), -float('inf')], 0.0).fillna(0.0).round(2)
        m_df["Cash_Tied_In_Inventory"] = m_df["Purchases_Cost"] - m_df["COGS"]
        m_df["Total_Real_Collected"] = m_df["Real_Cash"] + m_df["Real_Visa"] + m_df["Service_Cash_In"]
        m_df["Net_Operational_Profit"] = m_df["POS_Gross_Profit"] - m_df["Expenses"] + m_df["Service_Cash_In"]
        m_df["Net_Cash_Flow_Impact"] = m_df["Net_Operational_Profit"] - m_df["Cash_Tied_In_Inventory"]
        m_df["Expiry_Loss_Rate_Pct"] = ((m_df["Expired_Loss"] / m_df["Net_Sales"]) * 100.0).fillna(0.0).round(2)

        def flag_proc(r):
            if r["Net_Sales"] == 0: return "No Sales Activity"
            if r["Procurement_Ratio"] > 1.10: return "OVER-PURCHASING (Draining Liquid Cash)"
            elif r["Procurement_Ratio"] < 0.90: return "DESTOCKING (Freeing Cash / Watch Stockouts)"
            return "BALANCED (Procurement Matches Dispensing)"

        m_df["Purchasing_Status"] = m_df.apply(flag_proc, axis=1)

        # ----------------------------------------------------------------------
        # OUTPUT EXPORTS
        # ----------------------------------------------------------------------
        if sto_code == "9":
            out_s9 = pd.DataFrame({
                "Month": m_df.index,
                "Gross Sales": m_df["Gross_Sales"].round(2),
                "Total Discounts Given": m_df["Discounts_Given"].round(2),
                "Net Sales Revenue": m_df["Net_Sales"].round(2),
                "Cost of Goods Sold (COGS)": m_df["COGS"].round(2),
                "POS Gross Profit": m_df["POS_Gross_Profit"].round(2),
                "Profit % (Margin)": m_df["Gross_Margin_Pct"],
                "Total Inventory Cost": m_df["Total_Stock_Cost"].round(2),
                "Total Expiry Cost": m_df["Expired_Loss"].round(2)
            })
            file_s9 = "monthly_audit_Store_9_Taht_ElTasreef.csv"
            out_s9.to_csv(file_s9, index=False, encoding="utf-8-sig")
            print(f"      [OK] Exported Store 9 (8 columns) -> {file_s9}")

        elif sto_code == "1":
            out_s1 = pd.DataFrame({
                "Month": m_df.index,
                "(A) Gross Sales": m_df["Gross_Sales"].round(2),
                "(B) Total Discounts Given": m_df["Discounts_Given"].round(2),
                "(C) Net Sales Revenue": m_df["Net_Sales"].round(2),
                "(D) Cost of Goods Sold (COGS)": m_df["COGS"].round(2),
                "(E) POS Gross Profit": m_df["POS_Gross_Profit"].round(2),
                "Gross Margin %": m_df["Gross_Margin_Pct"],
                "(F) New Purchases at Cost": m_df["Purchases_Cost"].round(2),
                "Procurement Ratio (F / D)": m_df["Procurement_Ratio"],
                "Purchasing Audit Status": m_df["Purchasing_Status"],
                "Cash Tied Up in Extra Stock (F - D)": m_df["Cash_Tied_In_Inventory"].round(2),
                "(G) Operating Expenses": m_df["Expenses"].round(2),
                "(H) Service & Direct Cash Inflows": m_df["Service_Cash_In"].round(2),
                "Real Cash Collected (POS)": m_df["Real_Cash"].round(2),
                "Real Visa Collected (POS)": m_df["Real_Visa"].round(2),
                "Total Real Collected Money": m_df["Total_Real_Collected"].round(2),
                "(I) Net Operational Profit (E - G + H)": m_df["Net_Operational_Profit"].round(2),
                "(J) Net Cash Flow Impact (I - Cash Tied)": m_df["Net_Cash_Flow_Impact"].round(2),
                "[Audit] Total Stock Cost": m_df["Total_Stock_Cost"].round(2),
                "[Audit] Expired Stock Cost Loss": m_df["Expired_Loss"].round(2),
                "[Audit] Expiry Loss Rate %": m_df["Expiry_Loss_Rate_Pct"]
            })
            file_s1 = "monthly_audit_Store_1_Main_Pharmacy.csv"
            out_s1.to_csv(file_s1, index=False, encoding="utf-8-sig")
            print(f"      [OK] Exported Store 1 (21 columns) -> {file_s1}")

    print("\n[4/4] Reports successfully refreshed and reconciled.")

if __name__ == "__main__":
    generate_reports()