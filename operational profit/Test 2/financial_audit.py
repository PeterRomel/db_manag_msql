import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

# ==============================================================================
# 1. DATABASE CONFIGURATION
# ==============================================================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

START_DATE = "2025-01-01 00:00:00"
END_DATE = "2026-09-30 23:59:59"

# Service customer accounts to isolate into (H)
SERVICE_CUST_IDS = (4, 5, 6, 7, 18, 60, 61)

params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USERNAME};"
    f"PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

# Base monthly timeline (21 months)
MONTHS_TIMELINE = pd.period_range(start="2025-01", end="2026-09", freq="M").astype(str)

# ==============================================================================
# 2. RAW LOG SQL EXTRACTION QUERIES
# ==============================================================================

# A. Raw Completed Sales Lines (with Customer ID to separate Services)
SQL_SALES = text("""
    SELECT 
        h.sto_id,
        h.sth_id,
        h.cust_id,
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

# B. End of Shift Report (Shift_Changes_d)
SQL_SHIFTS = text("""
    SELECT 
        scd_sto_id AS sto_id,
        scd_start AS shift_date,
        CAST(ISNULL(scd_d_sales, 0) AS float) AS cash_sales,
        CAST(ISNULL(scd_o_salesVisa, 0) AS float) AS visa_sales,
        CAST(ISNULL(scd_d_cashIn, 0) AS float) AS cash_in,
        CAST(ISNULL(scd_c_cashOut, 0) AS float) AS cash_out_expenses
    FROM Shift_Changes_d
    WHERE scd_sto_id = :sto_id
      AND scd_start BETWEEN :start_date AND :end_date
""")

# C. Purchases at Wholesale Cost (pur_trans_d + pur_trans_h)
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

# D. Historical Month-End Inventory Cost & Expired Cost (ICS_Month_Close)
SQL_INVENTORY = text("""
    SELECT 
        sto_id,
        sec_insert_date,
        CAST(itm_qty AS float) AS itm_qty,
        CAST(pharm_price AS float) AS pharm_price,
        itm_expiry_date
    FROM ICS_Month_Close
    WHERE sto_id = :sto_id
      AND sec_insert_date BETWEEN :start_date AND :end_date
""")

# ==============================================================================
# 3. REPORT BUILDERS
# ==============================================================================

def generate_store_9_report():
    print("\n" + "="*70)
    print("GENERATING STORE 9 REPORT (تحت التصريف - sto_id = 48)")
    print("="*70)

    params = {"sto_id": 48, "start_date": START_DATE, "end_date": END_DATE}

    with engine.connect() as conn:
        df_sales = pd.read_sql(SQL_SALES, conn, params=params)
        df_inv = pd.read_sql(SQL_INVENTORY, conn, params=params)

    # 1. Base monthly table
    m_df = pd.DataFrame({"Month": MONTHS_TIMELINE}).set_index("Month")

    # 2. Sales & COGS
    if not df_sales.empty:
        df_sales["Net_Qty"] = df_sales["qnty"] - df_sales["itm_back"]
        df_sales = df_sales[df_sales["Net_Qty"] > 0].copy()

        df_sales["Gross_Sales"] = df_sales["Net_Qty"] * df_sales["itm_sell"]
        adj_dis_mon = (df_sales["itm_dis_mon"] * (df_sales["Net_Qty"] / df_sales["qnty"])).fillna(0)
        line_net = (df_sales["Gross_Sales"] * ((100.0 - df_sales["itm_dis_per"]) / 100.0)) - adj_dis_mon
        line_net = line_net.clip(lower=0.0)

        df_sales["Net_Sales"] = (line_net * ((100.0 - df_sales["bill_dis_per"]) / 100.0)).clip(lower=0.0)
        df_sales["Total_Discounts"] = df_sales["Gross_Sales"] - df_sales["Net_Sales"]
        df_sales["COGS"] = df_sales["Net_Qty"] * df_sales["itm_cost"]

        df_sales["sale_date"] = pd.to_datetime(df_sales["sale_date"])
        df_sales["Month"] = df_sales["sale_date"].dt.to_period("M").astype(str)

        agg_sales = df_sales.groupby("Month").agg(
            Gross_sales=("Gross_Sales", "sum"),
            total_discounts=("Total_Discounts", "sum"),
            net_sales=("Net_Sales", "sum"),
            COGS=("COGS", "sum")
        )
        m_df = m_df.join(agg_sales)
    else:
        m_df["Gross_sales"] = 0.0
        m_df["total_discounts"] = 0.0
        m_df["net_sales"] = 0.0
        m_df["COGS"] = 0.0

    m_df.fillna(0.0, inplace=True)
    m_df["pos_profit"] = m_df["net_sales"] - m_df["COGS"]
    m_df["profit_%"] = ((m_df["pos_profit"] / m_df["net_sales"]) * 100.0).replace([float('inf'), -float('inf')], 0.0).fillna(0.0).round(2)

    # 3. Inventory Snapshots from ICS_Month_Close
    if not df_inv.empty:
        df_inv["sec_insert_date"] = pd.to_datetime(df_inv["sec_insert_date"])
        df_inv["itm_expiry_date"] = pd.to_datetime(df_inv["itm_expiry_date"])
        df_inv["Month"] = df_inv["sec_insert_date"].dt.to_period("M").astype(str)

        df_inv["Line_Cost"] = df_inv["itm_qty"] * df_inv["pharm_price"]
        df_inv["Expired_Line_Cost"] = df_inv.apply(
            lambda r: r["Line_Cost"] if (pd.notnull(r["itm_expiry_date"]) and r["itm_expiry_date"] < r["sec_insert_date"]) else 0.0,
            axis=1
        )

        agg_inv = df_inv.groupby("Month").agg(
            Total_Inventory_Cost=("Line_Cost", "sum"),
            Total_Expired_Cost=("Expired_Line_Cost", "sum")
        )
        m_df = m_df.join(agg_inv)
    else:
        m_df["Total_Inventory_Cost"] = 0.0
        m_df["Total_Expired_Cost"] = 0.0

    m_df.fillna(0.0, inplace=True)
    m_df["Expired_Inventory_Pct"] = ((m_df["Total_Expired_Cost"] / m_df["Total_Inventory_Cost"]) * 100.0).replace([float('inf'), -float('inf')], 0.0).fillna(0.0).round(2)

    # Output CSV for Store 9
    out_store_9 = pd.DataFrame({
        "Month": m_df.index,
        "Gross sales": m_df["Gross_sales"].round(2),
        "total discounts": m_df["total_discounts"].round(2),
        "net sales": m_df["net_sales"].round(2),
        "COGS": m_df["COGS"].round(2),
        "pos profit": m_df["pos_profit"].round(2),
        "profit %": m_df["profit_%"],
        "[Audit] Total inventory Cost": m_df["Total_Inventory_Cost"].round(2),
        "[Audit] Total inventory Expired Stock Cost": m_df["Total_Expired_Cost"].round(2),
        "[Audit] I/H %": m_df["Expired_Inventory_Pct"]
    })

    filename = "monthly_audit_Store_9_Taht_ElTasreef.csv"
    out_store_9.to_csv(filename, index=False, encoding="utf-8-sig")
    print(f"[SUCCESS] Exported Store 9 report to: {filename}")


def generate_store_1_report():
    print("\n" + "="*70)
    print("GENERATING STORE 1 REPORT (صيدلية نبيل كمال - sto_id = 47)")
    print("="*70)

    params = {"sto_id": 47, "start_date": START_DATE, "end_date": END_DATE}

    with engine.connect() as conn:
        df_sales = pd.read_sql(SQL_SALES, conn, params=params)
        df_shifts = pd.read_sql(SQL_SHIFTS, conn, params=params)
        df_pur = pd.read_sql(SQL_PURCHASES, conn, params=params)
        df_inv = pd.read_sql(SQL_INVENTORY, conn, params=params)

    m_df = pd.DataFrame({"Month": MONTHS_TIMELINE}).set_index("Month")

    # 1. Sales & COGS & Services Separation
    if not df_sales.empty:
        df_sales["Net_Qty"] = df_sales["qnty"] - df_sales["itm_back"]
        df_sales = df_sales[df_sales["Net_Qty"] > 0].copy()

        df_sales["Gross_Sales"] = df_sales["Net_Qty"] * df_sales["itm_sell"]
        adj_dis_mon = (df_sales["itm_dis_mon"] * (df_sales["Net_Qty"] / df_sales["qnty"])).fillna(0)
        line_net = (df_sales["Gross_Sales"] * ((100.0 - df_sales["itm_dis_per"]) / 100.0)) - adj_dis_mon
        line_net = line_net.clip(lower=0.0)

        df_sales["Net_Sales"] = (line_net * ((100.0 - df_sales["bill_dis_per"]) / 100.0)).clip(lower=0.0)
        df_sales["Discount_Amount"] = df_sales["Gross_Sales"] - df_sales["Net_Sales"]
        df_sales["COGS"] = df_sales["Net_Qty"] * df_sales["itm_cost"]

        df_sales["sale_date"] = pd.to_datetime(df_sales["sale_date"])
        df_sales["Month"] = df_sales["sale_date"].dt.to_period("M").astype(str)

        # Flag service client accounts
        is_service = df_sales["cust_id"].isin(SERVICE_CUST_IDS)
        df_merch = df_sales[~is_service]
        df_srv = df_sales[is_service]

        # Aggregate merchandise sales
        agg_merch = df_merch.groupby("Month").agg(
            Gross_Sales=("Gross_Sales", "sum"),
            Discounts_Given=("Discount_Amount", "sum"),
            Net_Sales=("Net_Sales", "sum"),
            COGS=("COGS", "sum")
        )
        m_df = m_df.join(agg_merch)

        # Aggregate service revenue into (H)
        agg_srv = df_srv.groupby("Month")["Net_Sales"].sum().rename("Service_Inflows")
        m_df = m_df.join(agg_srv)
    else:
        m_df["Gross_Sales"] = 0.0
        m_df["Discounts_Given"] = 0.0
        m_df["Net_Sales"] = 0.0
        m_df["COGS"] = 0.0
        m_df["Service_Inflows"] = 0.0

    m_df.fillna(0.0, inplace=True)
    m_df["POS_Gross_Profit"] = m_df["Net_Sales"] - m_df["COGS"]
    m_df["Gross_Margin_Pct"] = ((m_df["POS_Gross_Profit"] / m_df["Net_Sales"]) * 100.0).replace([float('inf'), -float('inf')], 0.0).fillna(0.0).round(2)

    # 2. Purchases at Wholesale Cost
    if not df_pur.empty:
        base_price = df_pur["itm_pur_price"] * ((100.0 - df_pur["itm_extra_dis"]) / 100.0)
        net_unit_price = base_price * ((100.0 - df_pur["bill_dis_per"]) / 100.0)
        df_pur["Total_Purchase_Cost"] = (df_pur["qnty"] * (net_unit_price + df_pur["itm_tax_price"])) + (df_pur["bonus"] * df_pur["itm_tax_price"])

        df_pur["pur_date"] = pd.to_datetime(df_pur["pur_date"])
        df_pur["Month"] = df_pur["pur_date"].dt.to_period("M").astype(str)
        agg_pur = df_pur.groupby("Month")["Total_Purchase_Cost"].sum().rename("New_Purchases")
        m_df = m_df.join(agg_pur)
    else:
        m_df["New_Purchases"] = 0.0

    m_df.fillna(0.0, inplace=True)
    m_df["Procurement_Ratio"] = (m_df["New_Purchases"] / m_df["COGS"]).replace([float('inf'), -float('inf')], 0.0).fillna(0.0).round(2)
    m_df["Cash_Tied_In_Stock"] = m_df["New_Purchases"] - m_df["COGS"]

    def get_status(ratio):
        if ratio > 1.10:
            return "OVER-PURCHASING (Draining Liquid Cash)"
        elif ratio < 0.90:
            return "DESTOCKING (Freeing Cash / Watch Stockouts)"
        return "BALANCED (Procurement Matches Dispensing)"

    m_df["Purchasing_Status"] = m_df["Procurement_Ratio"].apply(get_status)

    # 3. End of Shift Report (Shift_Changes_d)
    if not df_shifts.empty:
        df_shifts["shift_date"] = pd.to_datetime(df_shifts["shift_date"])
        df_shifts["Month"] = df_shifts["shift_date"].dt.to_period("M").astype(str)

        agg_shifts = df_shifts.groupby("Month").agg(
            Real_Cash_Collected=("cash_sales", "sum"),
            Real_Visa_Collected=("visa_sales", "sum"),
            Operating_Expenses=("cash_out_expenses", "sum")
        )
        m_df = m_df.join(agg_shifts)
    else:
        m_df["Real_Cash_Collected"] = 0.0
        m_df["Real_Visa_Collected"] = 0.0
        m_df["Operating_Expenses"] = 0.0

    m_df.fillna(0.0, inplace=True)

    # Clean the single 8,010,871.00 EGP typo in August 2026 (documented in F_Notices Notice #89)
    if "2026-08" in m_df.index and m_df.loc["2026-08", "Real_Cash_Collected"] > 5000000.0:
        print("      [CLEANUP] Adjusted 8,010,871.00 EGP typo in August 2026 cash sales (Notice #89).")
        m_df.loc["2026-08", "Real_Cash_Collected"] -= 8010871.00

    m_df["Total_Real_Collected"] = m_df["Real_Cash_Collected"] + m_df["Real_Visa_Collected"]

    # 4. Net Operational Profit (E - G + H) & Cash Flow Impact (I - Cash Tied)
    m_df["Net_Operational_Profit"] = m_df["POS_Gross_Profit"] - m_df["Operating_Expenses"] + m_df["Service_Inflows"]
    m_df["Net_Cash_Flow_Impact"] = m_df["Net_Operational_Profit"] - m_df["Cash_Tied_In_Stock"]

    # 5. Month-End Stock Snapshots from ICS_Month_Close
    if not df_inv.empty:
        df_inv["sec_insert_date"] = pd.to_datetime(df_inv["sec_insert_date"])
        df_inv["itm_expiry_date"] = pd.to_datetime(df_inv["itm_expiry_date"])
        df_inv["Month"] = df_inv["sec_insert_date"].dt.to_period("M").astype(str)

        df_inv["Line_Cost"] = df_inv["itm_qty"] * df_inv["pharm_price"]
        df_inv["Expired_Line_Cost"] = df_inv.apply(
            lambda r: r["Line_Cost"] if (pd.notnull(r["itm_expiry_date"]) and r["itm_expiry_date"] < r["sec_insert_date"]) else 0.0,
            axis=1
        )

        agg_inv = df_inv.groupby("Month").agg(
            Total_Inventory_Cost=("Line_Cost", "sum"),
            Total_Expired_Cost=("Expired_Line_Cost", "sum")
        )
        m_df = m_df.join(agg_inv)
    else:
        m_df["Total_Inventory_Cost"] = 0.0
        m_df["Total_Expired_Cost"] = 0.0

    m_df.fillna(0.0, inplace=True)
    m_df["Expired_Inventory_Pct"] = ((m_df["Total_Expired_Cost"] / m_df["Total_Inventory_Cost"]) * 100.0).replace([float('inf'), -float('inf')], 0.0).fillna(0.0).round(2)

    # Output CSV for Store 1 (Matching column names and structure)
    out_store_1 = pd.DataFrame({
        "Month": m_df.index,
        "(A) Gross Sales": m_df["Gross_Sales"].round(2),
        "(B) Total Discounts Given": m_df["Discounts_Given"].round(2),
        "(C) Net Sales Revenue": m_df["Net_Sales"].round(2),
        "(D) Cost of Goods Sold (COGS)": m_df["COGS"].round(2),
        "(E) POS Gross Profit": m_df["POS_Gross_Profit"].round(2),
        "Gross Margin %": m_df["Gross_Margin_Pct"],
        "(F) New Purchases at Cost": m_df["New_Purchases"].round(2),
        "Procurement Ratio (F / D)": m_df["Procurement_Ratio"],
        "Purchasing Audit Status": m_df["Purchasing_Status"],
        "Cash Tied Up in Extra Stock (F - D)": m_df["Cash_Tied_In_Stock"].round(2),
        "(G) Operating Expenses": m_df["Operating_Expenses"].round(2),
        "(H) Service & Direct Cash Inflows": m_df["Service_Inflows"].round(2),
        "Real Cash Collected (from end of shift report)": m_df["Real_Cash_Collected"].round(2),
        "Real Visa Collected (from end of shift report)": m_df["Real_Visa_Collected"].round(2),
        "Total Real Collected Money(I+J) the": m_df["Total_Real_Collected"].round(2),
        "(I) Net Operational Profit (E - G + H)": m_df["Net_Operational_Profit"].round(2),
        "(J) Net Cash Flow Impact (I -G+H - Cash Tied)": m_df["Net_Cash_Flow_Impact"].round(2),
        "[Audit] Total inventory Cost": m_df["Total_Inventory_Cost"].round(2),
        "[Audit] Total inventory Expired Stock Cost": m_df["Total_Expired_Cost"].round(2),
        "[Audit] L/K %": m_df["Expired_Inventory_Pct"]
    })

    filename = "monthly_audit_Store_1_Main_Pharmacy.csv"
    out_store_1.to_csv(filename, index=False, encoding="utf-8-sig")
    print(f"[SUCCESS] Exported Store 1 report to: {filename}")


if __name__ == "__main__":
    generate_store_9_report()
    generate_store_1_report()
    print("\n[FINISHED] Both reports generated successfully.")