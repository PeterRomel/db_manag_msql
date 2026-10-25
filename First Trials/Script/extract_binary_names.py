import urllib.parse
import pandas as pd
from sqlalchemy import create_engine

# --- 1. DATABASE CONNECTION SETTINGS ---
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

# --- 2. THE SQL QUERY ---
# We filter items to only include those that exist in Sales or Purchases
sql_query = """
    SELECT 
        itm_code AS [Barcode],
        itm_name_en_encrypt AS [English_Binary],
        itm_name_ar_encrypt AS [Arabic_Binary]
    FROM Item_Catalog
    WHERE itm_id IN (
        SELECT DISTINCT itm_id FROM sales_trans_d
        UNION
        SELECT DISTINCT itm_id FROM pur_trans_d
    )
"""

print("Fetching used items from the database...")
df = pd.read_sql(sql_query, engine)

# --- 3. CONVERT RAW BYTES TO HEX STRING ---
# When Python reads varbinary/image from SQL, it reads it as raw bytes.
# We need to convert it to a readable Hexadecimal string (like "0x4D4145...") so you can easily share it in a CSV.
print("Converting binary data to Hexadecimal strings...")


def bytes_to_hex(val):
    if isinstance(val, bytes):
        # Convert bytes to hex and uppercase it. Add '0x' to the front.
        return "0x" + val.hex().upper()
    return val


df["English_Binary"] = df["English_Binary"].apply(bytes_to_hex)
df["Arabic_Binary"] = df["Arabic_Binary"].apply(bytes_to_hex)

# --- 4. SAVE TO CSV ---
output_file = "Used_Items_Binary_Data.csv"
print(f"Saving to {output_file}...")

# We use utf-8 encoding and don't save the index numbers
df.to_csv(output_file, index=False, encoding="utf-8")

print("✅ Success! The CSV is ready.")
