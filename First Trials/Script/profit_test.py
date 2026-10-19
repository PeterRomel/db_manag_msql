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
TARGET_ITEM_CODE = "PA"  # Can match itm_code or itm_id
OUTPUT_CSV_FILE = f"operational_profit_item_{TARGET_ITEM_CODE}.csv"

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
# 2. DECRYPTION FUNCTION
# ==========================================
def decode_legacy_bidi_bytes(raw_bytes):
    """
    Decodes the obfuscated varbinary medicine names:
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
# 3. EXTRACTION QUERY
# ==========================================
sql_query = text("""
    SELECT 
        h.sth_id AS Invoice_ID,
        ic.itm_id,
        ic.itm_code,
        ic.itm_name_en_encrypt,
        h.sec_insert_date AS Date_Of_Sale,
        d.qnty AS Qty_Sold,
        ISNULL(d.itm_back, 0) AS Qty_Returned,
        CAST(d.itm_sell AS float) AS Unit_Original_Sell_Price,
        CAST(ISNULL(d.itm_cost, 0) AS float) AS Unit_Purchase_Price,
        CAST(ISNULL(d.itm_dis_per, 0) AS float) AS Line_Dis_Per,
        CAST(ISNULL(d.itm_dis_mon, 0) AS float) AS Line_Dis_Mon,
        CAST(ISNULL(h.total_dis_per, 0) AS float) AS Bill_Dis_Per
    FROM sales_trans_d d
    INNER JOIN sales_trans_h h ON h.sth_id = d.sth_id
    INNER JOIN Item_Catalog ic ON ic.itm_id = d.itm_id
    WHERE (ic.itm_code = :item_code OR CAST(ic.itm_id AS varchar(50)) = :item_code)
      AND h.sth_flag = 'C' -- Closed / Completed Sales only
    ORDER BY h.sec_insert_date ASC;
""")


# ==========================================
# 4. DATA PROCESSING & PROFIT CALCULATIONS
# ==========================================
def run_profit_analysis():
    print(
        f"Connecting to database and fetching records for item: {TARGET_ITEM_CODE}..."
    )
    with engine.connect() as conn:
        df = pd.read_sql(sql_query, conn, params={"item_code": TARGET_ITEM_CODE})

    if df.empty:
        print(
            f"No completed sales transactions found for item code '{TARGET_ITEM_CODE}'."
        )
        return

    print(
        f"Found {len(df)} transactions. Processing discounts and operational profit..."
    )

    # Decrypt English product name
    df["Item english name"] = df["itm_name_en_encrypt"].apply(decode_legacy_bidi_bytes)

    # 1. Effective Quantity (Net of immediate bill returns)
    df["Net_Qty"] = df["Qty_Sold"] - df["Qty_Returned"]

    # 2. Financial Pipeline
    # Gross Revenue
    df["Gross_Revenue"] = df["Net_Qty"] * df["Unit_Original_Sell_Price"]

    # Inline / Line-level Discount calculation
    # Proportionally adjust cash discount if items were returned
    df["Adj_Line_Dis_Mon"] = df.apply(
        lambda r: (
            (r["Line_Dis_Mon"] * (r["Net_Qty"] / r["Qty_Sold"]))
            if r["Qty_Sold"] > 0
            else 0.0
        ),
        axis=1,
    )
    df["Line_Net"] = (
        df["Gross_Revenue"] * ((100.0 - df["Line_Dis_Per"]) / 100.0)
    ) - df["Adj_Line_Dis_Mon"]
    df["Line_Net"] = df["Line_Net"].clip(lower=0.0)

    # Invoice / Header-level Discount calculation
    df["Final_Net_Revenue"] = df["Line_Net"] * ((100.0 - df["Bill_Dis_Per"]) / 100.0)
    df["Final_Net_Revenue"] = df["Final_Net_Revenue"].clip(lower=0.0)

    # Unit Discounted Selling Price
    df["Unit_Discounted_Sell_Price"] = df.apply(
        lambda r: (r["Final_Net_Revenue"] / r["Net_Qty"]) if r["Net_Qty"] > 0 else 0.0,
        axis=1,
    )

    # Total Discount Value & Effective Percentage
    df["Total_Discount_Value"] = df["Gross_Revenue"] - df["Final_Net_Revenue"]
    df["Effective_Discount_Percentage"] = df.apply(
        lambda r: (
            (r["Total_Discount_Value"] / r["Gross_Revenue"] * 100.0)
            if r["Gross_Revenue"] > 0
            else 0.0
        ),
        axis=1,
    )

    # Cost Calculations (Using snapshot purchase price)
    df["Total_Purchase_Price"] = df["Net_Qty"] * df["Unit_Purchase_Price"]

    # Operational Profit Calculations
    df["Profit_Value"] = df["Final_Net_Revenue"] - df["Total_Purchase_Price"]

    # Profit Margin % (Profit relative to Net Revenue)
    df["Profit_Margin_Percentage"] = df.apply(
        lambda r: (
            (r["Profit_Value"] / r["Final_Net_Revenue"] * 100.0)
            if r["Final_Net_Revenue"] > 0
            else 0.0
        ),
        axis=1,
    )

    # Profit Markup % (Profit relative to Cost)
    df["Profit_Markup_Percentage"] = df.apply(
        lambda r: (
            (r["Profit_Value"] / r["Total_Purchase_Price"] * 100.0)
            if r["Total_Purchase_Price"] > 0
            else 0.0
        ),
        axis=1,
    )

    # ==========================================
    # 5. SHAPING OUTPUT CSV
    # ==========================================
    output_df = pd.DataFrame(
        {
            "Invoice ID": df["Invoice_ID"],
            "Item code": df["itm_code"],
            "Item english name": df["Item english name"],
            "Date of sale": df["Date_Of_Sale"],
            "Quantity sold": df["Net_Qty"].round(2),
            "Unit original selling price": df["Unit_Original_Sell_Price"].round(3),
            "Line discount percentage": df["Line_Dis_Per"].round(2),
            "Invoice discount percentage": df["Bill_Dis_Per"].round(2),
            "Effective discount percentage": df["Effective_Discount_Percentage"].round(
                2
            ),
            "Total discount value": df["Total_Discount_Value"].round(3),
            "Unit discounted selling price value": df[
                "Unit_Discounted_Sell_Price"
            ].round(3),
            "Total discounted selling price value": df["Final_Net_Revenue"].round(3),
            "Unit purchase price": df["Unit_Purchase_Price"].round(3),
            "Total purchase price": df["Total_Purchase_Price"].round(3),
            "Profit value": df["Profit_Value"].round(3),
            "Profit margin percentage": df["Profit_Margin_Percentage"].round(2),
            "Profit markup percentage": df["Profit_Markup_Percentage"].round(2),
        }
    )

    # Save to CSV
    output_df.to_csv(OUTPUT_CSV_FILE, index=False, encoding="utf-8-sig")
    print(
        f"\n[SUCCESS] Exported {len(output_df)} sales transactions to: {OUTPUT_CSV_FILE}"
    )


if __name__ == "__main__":
    run_profit_analysis()
