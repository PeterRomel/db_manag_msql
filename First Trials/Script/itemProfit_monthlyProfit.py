import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

# ==========================================
# 1. DATABASE CONFIGURATION
# ==========================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

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
# 2. NAME DECRYPTION FUNCTION
# ==========================================
def decode_legacy_bidi_bytes(raw_bytes):
    """
    Decodes obfuscated varbinary product names in Item_Catalog.
    Splits binary in half, reverses each chunk independently,
    concatenates, and decodes with windows-1256.
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
        return corrected_bytes.decode("windows-1256", errors="replace").strip()
    except Exception:
        return "Decoding Error"


# ==========================================
# 3. EXTRACTION QUERIES (RAW LOGS ONLY)
# ==========================================

# Raw completed sales transactions
SALES_QUERY = text("""
    SELECT 
        d.std_id,
        d.sth_id,
        d.itm_id,
        ic.itm_code,
        ic.itm_name_en_encrypt,
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
    INNER JOIN Item_Catalog ic ON ic.itm_id = d.itm_id
    WHERE h.sth_flag = 'C' -- Completed / Closed Invoices only
""")

# Raw operational cash expenses from the General Ledger
EXPENSES_QUERY = text("""
    SELECT 
        fh.fh_id,
        fh.fh_actual_date AS expense_date,
        CAST(fh.fh_value AS float) AS expense_amount,
        fh.fh_notes AS expense_notes,
        ISNULL(oe.oe_name_ar, ISNULL(ac.ac_name_ar, 'مصروفات تشغيلية عامة')) AS expense_category
    FROM F_Transaction_Header fh
    LEFT JOIN Other_Expenses oe ON oe.oe_id = fh.fh_to_id AND fh.fh_to_type = '5'
    LEFT JOIN F_Account_Chart ac ON ac.ac_id = fh.fh_to_id AND fh.fh_to_type = '5'
    WHERE fh.fh_trans_type = '1' -- Cash out
      AND fh.fh_from_type = '3'  -- From Cash Drawer / POS Store
      AND fh.fh_to_type = '5'    -- To Expense Accounts
""")


# ==========================================
# 4. PROCESSING LOGIC
# ==========================================
def main():
    print("[1/5] Extracting raw sales transactions from Genius ERP...")
    with engine.connect() as conn:
        df_sales = pd.read_sql(SALES_QUERY, conn)
        df_expenses = pd.read_sql(EXPENSES_QUERY, conn)

    if df_sales.empty:
        print("[ERROR] No sales records found.")
        return

    print(
        f"      Loaded {len(df_sales):,} sales lines and {len(df_expenses):,} expense entries."
    )

    # ----------------------------------------------------
    # A. VECTORIZED LINE-BY-LINE PROFIT CALCULATIONS
    # ----------------------------------------------------
    print(
        "[2/5] Calculating exact revenues, snapshot costs, and discounts in Python..."
    )

    # 1. Net active quantity (Sales minus immediate returns on invoice)
    df_sales["Net_Qty"] = df_sales["qnty"] - df_sales["itm_back"]

    # Exclude non-movement or voided line quantities
    df_sales = df_sales[df_sales["Net_Qty"] > 0].copy()

    # 2. Gross selling value
    df_sales["Gross_Sales"] = df_sales["Net_Qty"] * df_sales["itm_sell"]

    # 3. Two-Tier Discount Calculations:
    # Adjust line monetary discount proportionally if partial quantity was returned
    adj_line_mon_dis = (
        df_sales["itm_dis_mon"] * (df_sales["Net_Qty"] / df_sales["qnty"])
    ).fillna(0)
    line_net = (
        df_sales["Gross_Sales"] * ((100.0 - df_sales["itm_dis_per"]) / 100.0)
    ) - adj_line_mon_dis
    line_net = line_net.clip(lower=0.0)

    # Apply whole-invoice discount percentage
    df_sales["Final_Net_Revenue"] = (
        line_net * ((100.0 - df_sales["bill_dis_per"]) / 100.0)
    ).clip(lower=0.0)

    # Total discount given
    df_sales["Total_Discount"] = df_sales["Gross_Sales"] - df_sales["Final_Net_Revenue"]

    # 4. Purchase Cost using the ERP snapshot cost (itm_cost)
    df_sales["Total_Cost"] = df_sales["Net_Qty"] * df_sales["itm_cost"]

    # 5. Line Operational Profit
    df_sales["Operational_Profit"] = (
        df_sales["Final_Net_Revenue"] - df_sales["Total_Cost"]
    )

    # Date handling for monthly grouping
    df_sales["sale_date"] = pd.to_datetime(df_sales["sale_date"])
    df_sales["Year_Month"] = df_sales["sale_date"].dt.to_period("M").astype(str)

    # ----------------------------------------------------
    # B. REPORT 1: PRODUCT-LEVEL AGGREGATE SUMMARY
    # ----------------------------------------------------
    print("[3/5] Building Product-Level Aggregate Summary...")

    # Build unique lookup of decrypted names for items sold (optimized performance)
    unique_items = (
        df_sales[["itm_id", "itm_code", "itm_name_en_encrypt"]]
        .drop_duplicates("itm_id")
        .copy()
    )
    unique_items["Item_Name_EN"] = unique_items["itm_name_en_encrypt"].apply(
        decode_legacy_bidi_bytes
    )
    item_meta = unique_items.set_index("itm_id")[["itm_code", "Item_Name_EN"]]

    # Aggregate by Product ID
    prod_agg = (
        df_sales.groupby("itm_id")
        .agg(
            Total_Quantity_Sold=("Net_Qty", "sum"),
            Total_Gross_Sales=("Gross_Sales", "sum"),
            Total_Discount_Value=("Total_Discount", "sum"),
            Total_Net_Sales_Revenue=("Final_Net_Revenue", "sum"),
            Total_Purchase_Cost=("Total_Cost", "sum"),
            Total_Profit_Value=("Operational_Profit", "sum"),
        )
        .reset_index()
    )

    # Join metadata
    prod_agg = prod_agg.join(item_meta, on="itm_id")

    # Add margins & markups
    prod_agg["Profit_Margin_Pct"] = prod_agg.apply(
        lambda r: (
            round((r["Total_Profit_Value"] / r["Total_Net_Sales_Revenue"] * 100.0), 2)
            if r["Total_Net_Sales_Revenue"] > 0
            else 0.0
        ),
        axis=1,
    )
    prod_agg["Profit_Markup_Pct"] = prod_agg.apply(
        lambda r: (
            round((r["Total_Profit_Value"] / r["Total_Purchase_Cost"] * 100.0), 2)
            if r["Total_Purchase_Cost"] > 0
            else 0.0
        ),
        axis=1,
    )

    # Rearrange columns
    product_summary_df = pd.DataFrame(
        {
            "Database Item ID": prod_agg["itm_id"],
            "Item Code (Barcode)": prod_agg["itm_code"],
            "Item English Name": prod_agg["Item_Name_EN"],
            "Total Quantity Sold": prod_agg["Total_Quantity_Sold"].round(2),
            "Total Gross Sales": prod_agg["Total_Gross_Sales"].round(2),
            "Total Discount Given": prod_agg["Total_Discount_Value"].round(2),
            "Total Net Revenue": prod_agg["Total_Net_Sales_Revenue"].round(2),
            "Total Purchase Cost": prod_agg["Total_Purchase_Cost"].round(2),
            "Total Operational Profit": prod_agg["Total_Profit_Value"].round(2),
            "Profit Margin % (on Revenue)": prod_agg["Profit_Margin_Pct"],
            "Profit Markup % (on Cost)": prod_agg["Profit_Markup_Pct"],
        }
    ).sort_values(by="Total Operational Profit", ascending=False)

    prod_filename = "products_operational_profit_summary.csv"
    product_summary_df.to_csv(prod_filename, index=False, encoding="utf-8-sig")
    print(f"      [OK] Saved: {prod_filename} ({len(product_summary_df):,} products)")

    # ----------------------------------------------------
    # C. REPORT 2: MONTHLY OPERATIONAL PROFIT & LOSS (WITH EXPENSES)
    # ----------------------------------------------------
    print("[4/5] Building Monthly Operational Profit & Loss Report...")

    # Monthly Sales Aggregations
    monthly_sales = (
        df_sales.groupby("Year_Month")
        .agg(
            Monthly_Gross_Sales=("Gross_Sales", "sum"),
            Monthly_Discounts=("Total_Discount", "sum"),
            Monthly_Net_Revenue=("Final_Net_Revenue", "sum"),
            Monthly_COGS=("Total_Cost", "sum"),
            Monthly_Gross_Profit=("Operational_Profit", "sum"),
        )
        .reset_index()
    )

    # Monthly Expenses Aggregations
    if not df_expenses.empty:
        df_expenses["expense_date"] = pd.to_datetime(df_expenses["expense_date"])
        df_expenses["Year_Month"] = (
            df_expenses["expense_date"].dt.to_period("M").astype(str)
        )
        monthly_exp = (
            df_expenses.groupby("Year_Month")["expense_amount"].sum().reset_index()
        )
        monthly_exp.rename(
            columns={"expense_amount": "Monthly_Operating_Expenses"}, inplace=True
        )
    else:
        monthly_exp = pd.DataFrame(columns=["Year_Month", "Monthly_Operating_Expenses"])

    # Merge Monthly Sales with Monthly Expenses
    monthly_report = pd.merge(monthly_sales, monthly_exp, on="Year_Month", how="left")
    monthly_report["Monthly_Operating_Expenses"] = monthly_report[
        "Monthly_Operating_Expenses"
    ].fillna(0.0)

    # Net Operational Profit = Gross Operational Profit - Operating Expenses
    monthly_report["Net_Operational_Profit"] = (
        monthly_report["Monthly_Gross_Profit"]
        - monthly_report["Monthly_Operating_Expenses"]
    )

    # Operating Profit Margins
    monthly_report["Gross_Margin_Pct"] = (
        (
            (
                monthly_report["Monthly_Gross_Profit"]
                / monthly_report["Monthly_Net_Revenue"]
            )
            * 100.0
        )
        .round(2)
        .fillna(0.0)
    )

    monthly_report["Net_Operating_Margin_Pct"] = (
        (
            (
                monthly_report["Net_Operational_Profit"]
                / monthly_report["Monthly_Net_Revenue"]
            )
            * 100.0
        )
        .round(2)
        .fillna(0.0)
    )

    monthly_pnl_df = pd.DataFrame(
        {
            "Month": monthly_report["Year_Month"],
            "Gross Sales Revenue": monthly_report["Monthly_Gross_Sales"].round(2),
            "Total Discounts Given": monthly_report["Monthly_Discounts"].round(2),
            "Net Sales Revenue": monthly_report["Monthly_Net_Revenue"].round(2),
            "Cost of Goods Sold (COGS)": monthly_report["Monthly_COGS"].round(2),
            "Gross Operational Profit": monthly_report["Monthly_Gross_Profit"].round(2),
            "Gross Margin %": monthly_report["Gross_Margin_Pct"],
            "Operating Expenses": monthly_report["Monthly_Operating_Expenses"].round(2),
            "Net Operational Profit": monthly_report["Net_Operational_Profit"].round(2),
            "Net Operating Margin %": monthly_report["Net_Operating_Margin_Pct"],
        }
    ).sort_values(by="Month", ascending=True)

    month_filename = "monthly_operational_profit_and_loss.csv"
    monthly_pnl_df.to_csv(month_filename, index=False, encoding="utf-8-sig")
    print(f"      [OK] Saved: {month_filename} ({len(monthly_pnl_df)} months)")

    print("\n[5/5] All operations completed successfully.")


if __name__ == "__main__":
    main()
