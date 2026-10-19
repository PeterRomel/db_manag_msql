import urllib.parse
import pandas as pd
from sqlalchemy import create_engine


# --- 1. DECRYPTION FUNCTION WITH BIDI FIX ---
def decrypt_name(byte_val, lang):
    if byte_val is None:
        return ""
    try:
        # Decode from bytes to text using Arabic Windows encoding
        text = byte_val.decode("cp1256")
        # Reverse the text [::-1] and remove extra spaces
        reversed_text = text[::-1].strip()

        # Force LibreOffice/Excel to read the text in the correct direction
        if lang == "ar":
            return "\u200f" + reversed_text  # Invisible Right-to-Left mark
        else:
            return "\u200e" + reversed_text  # Invisible Left-to-Right mark

    except Exception:
        return str(byte_val)


# --- 2. DATABASE CONNECTION SETTINGS ---
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

print("Connecting to the database...")
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

# --- 3. THE SQL QUERY (NO 'TOP 100' LIMIT) ---
sql_query = """
    WITH SalesData AS (
        SELECT 
            d.itm_id,
            SUM(d.qnty) as TotalQty,
            
            -- 1. Gross Revenue
            SUM(d.qnty * d.itm_sell) as GrossRevenue,
            
            -- 2. Final Net Revenue (applying both line discounts and invoice total discounts)
            SUM( ((((d.qnty * d.itm_sell) * (100.0 - ISNULL(d.itm_dis_per, 0)) / 100.0) - ISNULL(d.itm_dis_mon, 0))) * (100.0 - ISNULL(h.total_dis_per, 0)) / 100.0 ) as FinalNetRevenue,
            
            -- 3. Total Cost
            SUM(d.qnty * ISNULL(d.itm_cost, 0)) as TotalCost,
            
            -- 4. Active Days
            COUNT(DISTINCT CONVERT(varchar(10), h.sec_insert_date, 120)) as DaysSold
            
        FROM sales_trans_d d
        JOIN sales_trans_h h ON d.sth_id = h.sth_id
        WHERE h.sth_flag = 'C' 
          AND d.itm_nexist <> 1 
        GROUP BY d.itm_id
    )
    -- Notice I removed the 'TOP 100' here to get EVERYTHING
    SELECT 
        i.itm_code AS [Barcode],
        i.itm_name_en_encrypt AS [English Name],
        i.itm_name_ar_encrypt AS [Arabic Name],
        
        sd.TotalQty AS [Total Qty Sold],
        sd.GrossRevenue AS [Gross Revenue],
        (sd.GrossRevenue - sd.FinalNetRevenue) AS [Total Discounts],
        sd.FinalNetRevenue AS [Net Revenue],
        sd.TotalCost AS [Total Cost],
        (sd.FinalNetRevenue - sd.TotalCost) AS [Net Profit],
        
        CASE WHEN sd.FinalNetRevenue > 0 
             THEN ((sd.FinalNetRevenue - sd.TotalCost) / sd.FinalNetRevenue) 
             ELSE 0 END AS [Profit Margin],
        
        sd.DaysSold AS [Active Selling Days],
        (sd.TotalQty / sd.DaysSold) AS [Velocity (Qty / Active Day)]

    FROM SalesData sd
    JOIN Item_Catalog i ON sd.itm_id = i.itm_id
    WHERE sd.DaysSold > 0
    ORDER BY sd.FinalNetRevenue DESC;
"""

print("Fetching ALL data from the database. This might take a few seconds...")
df = pd.read_sql(sql_query, engine)
print(f"Data retrieved! Found {len(df)} unique medicines sold.")

print("Decrypting Arabic and English item names...")
# Apply the decryption function and tell it which language to format for
df["English Name"] = df["English Name"].apply(lambda x: decrypt_name(x, "en"))
df["Arabic Name"] = df["Arabic Name"].apply(lambda x: decrypt_name(x, "ar"))

# --- 4. FORMATTING AND SAVING TO EXCEL ---
output_file = "Complete_Pharmacy_Sales_Report.xlsx"
print(f"Generating Excel file: {output_file}...")

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="All Sales", index=False)

    workbook = writer.book
    worksheet = writer.sheets["All Sales"]

    # Safely calculate column widths
    for i, col in enumerate(df.columns):
        max_data_len = df[col].astype(str).str.len().max()
        if pd.isna(max_data_len):
            max_data_len = 0
        column_len = int(max(max_data_len, len(str(col)))) + 4
        worksheet.set_column(i, i, column_len)

    # Number Formats
    money_format = workbook.add_format({"num_format": "#,##0.00"})
    qty_format = workbook.add_format({"num_format": "#,##0.0"})
    perc_format = workbook.add_format({"num_format": "0.00%"})

    # Set Excel reading order to help LibreOffice (1 = LTR, 2 = RTL)
    rtl_format = workbook.add_format({"reading_order": 2})
    ltr_format = workbook.add_format({"reading_order": 1})

    # Apply formats to specific columns (0-indexed)
    worksheet.set_column(1, 1, None, ltr_format)  # English Name
    worksheet.set_column(2, 2, None, rtl_format)  # Arabic Name
    worksheet.set_column(3, 3, 15, qty_format)  # Total Qty
    worksheet.set_column(4, 8, 18, money_format)  # Gross to Net Profit
    worksheet.set_column(9, 9, 15, perc_format)  # Profit Margin %
    worksheet.set_column(11, 11, 22, qty_format)  # Velocity

    # Freeze the top row so it stays visible when you scroll down!
    worksheet.freeze_panes(1, 0)

    # Add a header style
    header_format = workbook.add_format(
        {
            "bold": True,
            "valign": "top",
            "fg_color": "#4F81BD",
            "font_color": "white",
            "border": 1,
        }
    )

    # Write the headers with style
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

print("✅ Success! Your complete pharmacy report is ready.")
