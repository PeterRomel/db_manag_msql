import urllib.parse
import pandas as pd
from sqlalchemy import create_engine

# --- 1. DATABASE CONNECTION SETTINGS ---
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

# Safely encode the password
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


# --- 2. DECRYPTION FUNCTION ---
# We modified your function slightly to take RAW BYTES directly from SQL
# instead of reading a Hex string from a CSV. It's much faster!
def decode_legacy_bidi_bytes(raw_bytes):
    # If the cell is empty/NULL, return an empty string
    if pd.isna(raw_bytes) or not raw_bytes:
        return ""

    try:
        # Calculate midpoint for the chunk split
        N = len(raw_bytes)
        L = N // 2

        # Split into Left and Right chunks
        left_chunk = raw_bytes[:L]
        right_chunk = raw_bytes[L:]

        # Reverse each chunk
        reversed_left = left_chunk[::-1]
        reversed_right = right_chunk[::-1]

        # Combine and decode using Windows-1256
        corrected_bytes = reversed_left + reversed_right
        return corrected_bytes.decode("windows-1256")
    except Exception as e:
        return f"Decoding Error"


# --- 3. THE SQL QUERY ---
# Notice we are now selecting the _encrypt columns instead of the blank ones
sql_query = """
    WITH SalesData AS (
        SELECT 
            d.itm_id,
            SUM(d.qnty) as TotalQty,
            SUM(d.qnty * d.itm_sell) as TotalRevenue,
            MIN(h.sec_insert_date) as FirstSaleDate,
            MAX(h.sec_insert_date) as LastSaleDate
        FROM sales_trans_d d
        JOIN sales_trans_h h ON d.sth_id = h.sth_id
        WHERE h.sth_flag = 'C'
        GROUP BY d.itm_id
    )
    SELECT TOP 100
        i.itm_code AS [Barcode],
        i.itm_name_en_encrypt AS [English_Binary],
        i.itm_name_ar_encrypt AS [Arabic_Binary],
        sd.TotalQty AS [Total Quantity Sold],
        sd.TotalRevenue AS [Total Revenue (EGP)],
        
        -- Calculate the number of months between first and last sale
        DATEDIFF(month, sd.FirstSaleDate, sd.LastSaleDate) + 1 AS [Months Active],
        
        -- Calculate Averages
        (sd.TotalQty / (DATEDIFF(month, sd.FirstSaleDate, sd.LastSaleDate) + 1)) AS [Avg Monthly Qty Sold],
        (sd.TotalRevenue / (DATEDIFF(month, sd.FirstSaleDate, sd.LastSaleDate) + 1)) AS [Avg Monthly Revenue]
        
    FROM SalesData sd
    JOIN Item_Catalog i ON sd.itm_id = i.itm_id
    ORDER BY sd.TotalRevenue DESC;
"""

print("Fetching data from the database...")
# Connect and pull data into pandas
df = pd.read_sql(sql_query, engine)

# --- 4. APPLY DECRYPTION TO THE DATA ---
print("Decrypting item names...")
# Create new columns with the decrypted text
df["English Name"] = df["English_Binary"].apply(decode_legacy_bidi_bytes)
df["Arabic Name"] = df["Arabic_Binary"].apply(decode_legacy_bidi_bytes)

# Drop the messy binary columns so they don't show up in Excel
df = df.drop(columns=["English_Binary", "Arabic_Binary"])

# Reorder the columns so the Names appear right after the Barcode
column_order = [
    "Barcode",
    "English Name",
    "Arabic Name",
    "Total Quantity Sold",
    "Total Revenue (EGP)",
    "Months Active",
    "Avg Monthly Qty Sold",
    "Avg Monthly Revenue",
]
df = df[column_order]

# --- 5. FORMATTING AND SAVING TO EXCEL ---
output_file = "Top_100_Medicines_Report_Decrypted.xlsx"
print(f"Generating Excel file: {output_file}...")

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Sales Analysis", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Sales Analysis"]

    # Safely calculate and set column widths
    for i, col in enumerate(df.columns):
        max_data_len = df[col].astype(str).str.len().max()
        if pd.isna(max_data_len):
            max_data_len = 0

        column_len = int(max(max_data_len, len(str(col)))) + 4
        worksheet.set_column(i, i, column_len)

    # Format numbers to have commas and 2 decimal places
    money_format = workbook.add_format({"num_format": "#,##0.00"})

    # Apply money format to Revenue columns (Columns E, H in Excel -> Index 4, 7 in Python)
    worksheet.set_column(4, 4, 20, money_format)  # Total Revenue
    worksheet.set_column(7, 7, 20, money_format)  # Avg Monthly Revenue

    # Add a header style
    header_format = workbook.add_format(
        {"bold": True, "valign": "top", "fg_color": "#D7E4BC", "border": 1}
    )

    # Write the headers with style
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

print("✅ Success! Your beautifully decoded report is ready.")
