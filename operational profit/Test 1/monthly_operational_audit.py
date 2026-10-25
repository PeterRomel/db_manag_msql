import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

# ==========================================
# 1. DATABASE CONNECTION
# ==========================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"
START_DATE = "2025-01-01 00:00:00"

params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USERNAME};"
    f"PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

# ==========================================
# 2. RAW DATA SQL EXTRACTION QUERIES (>= 2025)
# ==========================================

# A. Raw Sales Logs
SQL_SALES = text("""
    SELECT 
        d.std_id,
        d.sth_id,
        d.itm_id,
        h.sec_insert_date AS trans_date,
        d.qnty,
        ISNULL(d.itm_back, 0) AS itm_back,
        CAST(d.itm_sell AS float) AS itm_sell,
        CAST(ISNULL(d.itm_cost, 0) AS float) AS itm_cost,
        CAST(ISNULL(d.itm_dis_per, 0) AS float) AS itm_dis_per,
        CAST(ISNULL(d.itm_dis_mon, 0) AS float) AS itm_dis_mon,
        CAST(ISNULL(h.total_dis_per, 0) AS float) AS bill_dis_per
    FROM sales_trans_d d
    INNER JOIN sales_trans_h h ON h.sth_id = d.sth_id
    WHERE h.sth_flag = 'C' 
      AND h.sec_insert_date >= :start_date
""")

# B. Direct Cash Deposits from Service Customer Profiles
SQL_SERVICE_CASH = text("""
    SELECT 
        fh.fh_id,
        fh.fh_actual_date AS trans_date,
        CAST(fh.fh_value AS float) AS service_revenue,
        c.cust_name_ar,
        fh.fh_notes
    FROM F_Transaction_Header fh
    INNER JOIN Customer c ON c.cust_id = fh.fh_from_id
    WHERE fh.fh_trans_type = '1' -- Cash transaction
      AND fh.fh_from_type = '1'  -- From Customer profile
      AND fh.fh_to_type = '3'    -- To Register Cash Drawer
      AND fh.fh_actual_date >= :start_date
      AND (
          c.cust_name_ar LIKE '%سكر%'
          OR c.cust_name_ar LIKE '%ضغط%'
          OR c.cust_name_ar LIKE '%تركيب%'
          OR c.cust_name_ar LIKE '%خدم%'
          OR c.cust_name_ar LIKE '%زيادة%'
          OR c.cust_name_ar LIKE '%نقود%'
          OR fh.fh_notes LIKE '%خدم%'
          OR fh.fh_notes LIKE '%جلس%'
      )
""")

# C. New Purchases (Incoming Stock at Wholesale Cost)
SQL_PURCHASES = text("""
    SELECT 
        ptd.ptd_id,
        pth.pth_id,
        pth.sec_insert_date AS trans_date,
        ptd.qnty,
        ISNULL(ptd.bonus, 0) AS bonus,
        ISNULL(ptd.itm_back_qty, 0) AS itm_back_qty,
        CAST(ptd.itm_pur_price AS float) AS itm_pur_price,
        CAST(ISNULL(ptd.itm_extra_dis, 0) AS float) AS itm_extra_dis,
        CAST(ISNULL(ptd.itm_dis_mon, 0) AS float) AS itm_dis_mon,
        CAST(ISNULL(pth.total_dis_per, 0) AS float) AS total_dis_per
    FROM pur_trans_d ptd
    INNER JOIN pur_trans_h pth ON pth.pth_id = ptd.pth_id
    WHERE pth.bill_status <> 1  -- Exclude voided/cancelled bills
      AND pth.pth_sell_back = 0 -- Exclude sell-backs
      AND pth.sec_insert_date >= :start_date
""")

# D. Expired & Damaged Medications Written Off
SQL_EXPIRED_STOCK = text("""
    SELECT 
        iqc.iqc_id,
        iqc.sec_insert_date AS trans_date,
        CAST(iqc.iqc_old_quantity AS float) AS old_qty,
        CAST(iqc.iqc_new_quantity AS float) AS new_qty,
        CAST(ISNULL(iqc.iqc_pharm_price, 0) AS float) AS unit_cost,
        iqc.iqc_dc_id,
        dc.dc_name_ar AS reason
    FROM Item_Quantity_Changes iqc
    LEFT JOIN destroy_causes dc ON dc.dc_id = iqc.iqc_dc_id
    WHERE iqc.sec_insert_date >= :start_date
      AND iqc.iqc_old_quantity > iqc.iqc_new_quantity -- Quantity reduced
      AND (iqc.iqc_dc_id IN (1, 2, 8) OR dc.dc_name_ar LIKE '%صلاحية%' OR dc.dc_name_ar LIKE '%تالف%')
""")

# E. Operating Expenses from General Ledger
SQL_EXPENSES = text("""
    SELECT 
        fh.fh_id,
        fh.fh_actual_date AS trans_date,
        CAST(fh.fh_value AS float) AS expense_amount,
        fh.fh_notes,
        ISNULL(oe.oe_name_ar, ISNULL(ac.ac_name_ar, 'مصروفات تشغيلية عامة')) AS category
    FROM F_Transaction_Header fh
    LEFT JOIN Other_Expenses oe ON oe.oe_id = fh.fh_to_id AND fh.fh_to_type = '5'
    LEFT JOIN F_Account_Chart ac ON ac.ac_id = fh.fh_to_id AND fh.fh_to_type = '5'
    WHERE fh.fh_trans_type = '1' -- Cash Out
      AND fh.fh_from_type = '3'  -- From POS Cash Drawer
      AND fh.fh_to_type = '5'    -- To Operating Expenses
      AND fh.fh_actual_date >= :start_date
""")


# ==========================================
# 3. PURE PYTHON PROCESSING PIPELINE
# ==========================================
def run_financial_audit():
    print(f"Loading raw transaction logs starting from {START_DATE}...")
    with engine.connect() as conn:
        df_sales = pd.read_sql(SQL_SALES, conn, params={"start_date": START_DATE})
        df_services = pd.read_sql(
            SQL_SERVICE_CASH, conn, params={"start_date": START_DATE}
        )
        df_purchases = pd.read_sql(
            SQL_PURCHASES, conn, params={"start_date": START_DATE}
        )
        df_expired = pd.read_sql(
            SQL_EXPIRED_STOCK, conn, params={"start_date": START_DATE}
        )
        df_expenses = pd.read_sql(SQL_EXPENSES, conn, params={"start_date": START_DATE})

    print(f"Extracted:")
    print(f" - Sales records: {len(df_sales):,}")
    print(f" - Direct service entries: {len(df_services):,}")
    print(f" - Purchase records: {len(df_purchases):,}")
    print(f" - Expired/damaged entries: {len(df_expired):,}")
    print(f" - Expense entries: {len(df_expenses):,}")

    # --- 1. Process Sales & COGS ---
    df_sales["Net_Qty"] = df_sales["qnty"] - df_sales["itm_back"]
    df_sales = df_sales[df_sales["Net_Qty"] > 0].copy()

    df_sales["Gross_Sales"] = df_sales["Net_Qty"] * df_sales["itm_sell"]
    adj_line_mon_dis = (
        df_sales["itm_dis_mon"] * (df_sales["Net_Qty"] / df_sales["qnty"])
    ).fillna(0)
    line_net = (
        df_sales["Gross_Sales"] * ((100.0 - df_sales["itm_dis_per"]) / 100.0)
    ) - adj_line_mon_dis
    df_sales["Net_Revenue"] = (
        line_net.clip(lower=0.0) * ((100.0 - df_sales["bill_dis_per"]) / 100.0)
    ).clip(lower=0.0)
    df_sales["COGS"] = df_sales["Net_Qty"] * df_sales["itm_cost"]

    df_sales["trans_date"] = pd.to_datetime(df_sales["trans_date"])
    df_sales["Month"] = df_sales["trans_date"].dt.to_period("M").astype(str)

    monthly_sales = (
        df_sales.groupby("Month")
        .agg(POS_Sales_Revenue=("Net_Revenue", "sum"), COGS=("COGS", "sum"))
        .reset_index()
    )

    # --- 2. Process Service Cash Inflows ---
    if not df_services.empty:
        df_services["trans_date"] = pd.to_datetime(df_services["trans_date"])
        df_services["Month"] = df_services["trans_date"].dt.to_period("M").astype(str)
        monthly_services = (
            df_services.groupby("Month")["service_revenue"].sum().reset_index()
        )
        monthly_services.rename(
            columns={"service_revenue": "Service_Revenue"}, inplace=True
        )
    else:
        monthly_services = pd.DataFrame(columns=["Month", "Service_Revenue"])

    # --- 3. Process New Purchases at Wholesale Cost ---
    if not df_purchases.empty:
        df_purchases["Net_Qty"] = (
            df_purchases["qnty"] + df_purchases["bonus"] - df_purchases["itm_back_qty"]
        )
        df_purchases = df_purchases[df_purchases["Net_Qty"] > 0].copy()

        # Calculate line net wholesale cost after discounts
        pur_line_net = (
            df_purchases["Net_Qty"]
            * df_purchases["itm_pur_price"]
            * ((100.0 - df_purchases["itm_extra_dis"]) / 100.0)
        ) - df_purchases["itm_dis_mon"]
        df_purchases["Final_Purchase_Cost"] = (
            pur_line_net.clip(lower=0.0)
            * ((100.0 - df_purchases["total_dis_per"]) / 100.0)
        ).clip(lower=0.0)

        df_purchases["trans_date"] = pd.to_datetime(df_purchases["trans_date"])
        df_purchases["Month"] = df_purchases["trans_date"].dt.to_period("M").astype(str)
        monthly_purchases = (
            df_purchases.groupby("Month")["Final_Purchase_Cost"].sum().reset_index()
        )
        monthly_purchases.rename(
            columns={"Final_Purchase_Cost": "New_Purchases"}, inplace=True
        )
    else:
        monthly_purchases = pd.DataFrame(columns=["Month", "New_Purchases"])

    # --- 4. Process Expired & Damaged Stock Write-offs ---
    if not df_expired.empty:
        df_expired["Loss_Qty"] = df_expired["old_qty"] - df_expired["new_qty"]
        df_expired["Loss_Value"] = df_expired["Loss_Qty"] * df_expired["unit_cost"]
        df_expired["trans_date"] = pd.to_datetime(df_expired["trans_date"])
        df_expired["Month"] = df_expired["trans_date"].dt.to_period("M").astype(str)
        monthly_expired = df_expired.groupby("Month")["Loss_Value"].sum().reset_index()
        monthly_expired.rename(
            columns={"Loss_Value": "Expired_Stock_Loss"}, inplace=True
        )
    else:
        monthly_expired = pd.DataFrame(columns=["Month", "Expired_Stock_Loss"])

    # --- 5. Process Operating Expenses ---
    if not df_expenses.empty:
        df_expenses["trans_date"] = pd.to_datetime(df_expenses["trans_date"])
        df_expenses["Month"] = df_expenses["trans_date"].dt.to_period("M").astype(str)
        monthly_expenses = (
            df_expenses.groupby("Month")["expense_amount"].sum().reset_index()
        )
        monthly_expenses.rename(
            columns={"expense_amount": "Operating_Expenses"}, inplace=True
        )
    else:
        monthly_expenses = pd.DataFrame(columns=["Month", "Operating_Expenses"])

    # ==========================================
    # 4. MASTER FINANCIAL MERGE & KPI SYNTHESIS
    # ==========================================
    df_merged = (
        monthly_sales.merge(monthly_services, on="Month", how="outer")
        .merge(monthly_purchases, on="Month", how="outer")
        .merge(monthly_expenses, on="Month", how="outer")
        .merge(monthly_expired, on="Month", how="outer")
        .fillna(0.0)
        .sort_values(by="Month")
    )

    # Financial Formulas
    df_merged["Total_Operating_Revenue"] = (
        df_merged["POS_Sales_Revenue"] + df_merged["Service_Revenue"]
    )
    df_merged["Gross_Operating_Profit"] = (
        df_merged["Total_Operating_Revenue"] - df_merged["COGS"]
    )

    # True Net Operational Profit
    df_merged["True_Net_Operational_Profit"] = (
        df_merged["Gross_Operating_Profit"]
        - df_merged["Operating_Expenses"]
        - df_merged["Expired_Stock_Loss"]
    )

    # Cash Tied Up in Excess Stock (Purchases - COGS)
    df_merged["Inventory_Cash_Drag"] = df_merged["New_Purchases"] - df_merged["COGS"]

    # True Net Cash Flow from Operations
    df_merged["Net_Cash_Flow"] = (
        df_merged["True_Net_Operational_Profit"] - df_merged["Inventory_Cash_Drag"]
    )

    # Core Management KPIs
    df_merged["Procurement_Ratio"] = df_merged.apply(
        lambda r: round(r["New_Purchases"] / r["COGS"], 2) if r["COGS"] > 0 else 0.0,
        axis=1,
    )

    df_merged["Expiry_Loss_Rate_Pct"] = df_merged.apply(
        lambda r: (
            round((r["Expired_Stock_Loss"] / r["Total_Operating_Revenue"]) * 100.0, 2)
            if r["Total_Operating_Revenue"] > 0
            else 0.0
        ),
        axis=1,
    )

    df_merged["Net_Operating_Margin_Pct"] = df_merged.apply(
        lambda r: (
            round(
                (r["True_Net_Operational_Profit"] / r["Total_Operating_Revenue"])
                * 100.0,
                2,
            )
            if r["Total_Operating_Revenue"] > 0
            else 0.0
        ),
        axis=1,
    )

    # Inventory Health Status Flag
    def evaluate_status(row):
        ratio = row["Procurement_Ratio"]
        cash_drag = row["Inventory_Cash_Drag"]
        if ratio > 1.15:
            return f"OVER-PURCHASING ALERT (Trapped {cash_drag:,.0f} EGP in stock)"
        elif ratio < 0.85 and ratio > 0:
            return "DESTOCKING (Inventory reducing)"
        else:
            return "BALANCED (Cash flow aligned)"

    df_merged["Inventory_Health_Status"] = df_merged.apply(evaluate_status, axis=1)

    # Format Output Report
    final_report = pd.DataFrame(
        {
            "Month": df_merged["Month"],
            "POS Sales Revenue": df_merged["POS_Sales_Revenue"].round(2),
            "Service & Direct Cash Revenue": df_merged["Service_Revenue"].round(2),
            "Total Operating Revenue": df_merged["Total_Operating_Revenue"].round(2),
            "Cost of Goods Sold (COGS)": df_merged["COGS"].round(2),
            "Gross Operating Profit": df_merged["Gross_Operating_Profit"].round(2),
            "Operating Expenses": df_merged["Operating_Expenses"].round(2),
            "Expired Stock Losses": df_merged["Expired_Stock_Loss"].round(2),
            "True Net Operational Profit": df_merged[
                "True_Net_Operational_Profit"
            ].round(2),
            "New Stock Purchases": df_merged["New_Purchases"].round(2),
            "Inventory Cash Drag (Stock Delta)": df_merged["Inventory_Cash_Drag"].round(
                2
            ),
            "Net Operating Cash Flow": df_merged["Net_Cash_Flow"].round(2),
            "Procurement Ratio (Purchases/COGS)": df_merged["Procurement_Ratio"],
            "Expiry Loss Rate %": df_merged["Expiry_Loss_Rate_Pct"],
            "Net Operating Margin %": df_merged["Net_Operating_Margin_Pct"],
            "Inventory Health Status": df_merged["Inventory_Health_Status"],
        }
    )

    output_csv = "monthly_operational_audit_2025_present.csv"
    final_report.to_csv(output_csv, index=False, encoding="utf-8-sig")
    print(f"\n[SUCCESS] Generated comprehensive monthly audit report: {output_csv}")
    print(final_report.to_string(index=False))


if __name__ == "__main__":
    run_financial_audit()
