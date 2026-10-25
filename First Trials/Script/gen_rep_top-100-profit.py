import urllib.parse
import pandas as pd
from sqlalchemy import create_engine


# --- 1. DECRYPTION FUNCTION ---
# e-Plus stores names backwards in Windows-1256 encoding. This fixes it.
def decrypt_name(byte_val):
    if byte_val is None:
        return ""
    try:
        # Decode from bytes to text using Arabic Windows encoding
        text = byte_val.decode("cp1256")
        # Reverse the text [::-1] and remove extra spaces
        return text[::-1].strip()
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

# --- 3. THE SQL QUERY ---
# Built to be compatible with MS SQL Server 2000
sql_query = """
    WITH SalesData AS (
        SELECT 
            d.itm_id,
            SUM(d.qnty) as TotalQty,
            
            -- 1. Gross Revenue
            SUM(d.qnty * d.itm_sell) as GrossRevenue,
            
            -- 2. Line Net before invoice discount
            SUM( ((d.qnty * d.itm_sell) * (100.0 - ISNULL(d.itm_dis_per, 0)) / 100.0) - ISNULL(d.itm_dis_mon, 0) ) as LineNetRevenue,
            
            -- 3. Final Net Revenue (applying invoice discount)
            SUM( ((((d.qnty * d.itm_sell) * (100.0 - ISNULL(d.itm_dis_per, 0)) / 100.0) - ISNULL(d.itm_dis_mon, 0))) * (100.0 - ISNULL(h.total_dis_per, 0)) / 100.0 ) as FinalNetRevenue,
            
            -- 4. Total Cost (Quantity * Cost price at the exact time of sale)
            SUM(d.qnty * ISNULL(d.itm_cost, 0)) as TotalCost,
            
            -- 5. Dates and Active Days
            MIN(h.sec_insert_date) as FirstSaleDate,
            MAX(h.sec_insert_date) as LastSaleDate,
            COUNT(DISTINCT CONVERT(varchar(10), h.sec_insert_date, 120)) as DaysSold
            
        FROM sales_trans_d d
        JOIN sales_trans_h h ON d.sth_id = h.sth_id
        WHERE h.sth_flag = 'C' 
          AND d.itm_nexist <> 1 -- Ignore temporary/non-existent items
        GROUP BY d.itm_id
    )
    SELECT TOP 100
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

print("Fetching data from the database...")
df = pd.read_sql(sql_query, engine)

print("Decrypting Arabic and English item names...")
# Apply our decryption function to the encrypted columns
df["English Name"] = df["English Name"].apply(decrypt_name)
df["Arabic Name"] = df["Arabic Name"].apply(decrypt_name)

# --- 4. FORMATTING AND SAVING TO EXCEL ---
output_file = "Top_100_Medicines_Profit_Analysis.xlsx"
print(f"Generating Excel file: {output_file}...")

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Sales & Profit Analysis", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Sales & Profit Analysis"]

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

    # Apply formats to specific columns (0-indexed)
    worksheet.set_column("D:D", 15, qty_format)  # Total Qty
    worksheet.set_column("E:I", 18, money_format)  # Financials (Gross to Profit)
    worksheet.set_column("J:J", 15, perc_format)  # Profit Margin %
    worksheet.set_column("L:L", 22, qty_format)  # Velocity

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

print("✅ Success! Your advanced profit analysis report is ready.")
