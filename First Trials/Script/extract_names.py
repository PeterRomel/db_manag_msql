import urllib.parse
import pandas as pd
from sqlalchemy import create_engine

# --- 1. DATABASE CONNECTION ---
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

# --- 2. SQL QUERY ---
# We are grabbing the barcode and the encrypted binary columns.
# We will limit it to items that actually have an encrypted English name.
sql_query = """
    SELECT 
        itm_code, 
        itm_name_en_encrypt, 
        itm_name_ar_encrypt
    FROM Item_Catalog
    WHERE itm_name_en_encrypt IS NOT NULL
"""

print("Fetching raw binary data...")
df = pd.read_sql(sql_query, engine)

# --- 3. CONVERT BINARY TO HEXADECIMAL ---
print("Converting binary to readable Hex format...")


# This function safely takes the raw database bytes and turns them into a Hex string
def safe_hex(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, bytes):
        return val.hex().upper()  # Turns b'ABC' into '414243'
    return str(val)


# Apply the function to our columns
df["EN_HEX"] = df["itm_name_en_encrypt"].apply(safe_hex)
df["AR_HEX"] = df["itm_name_ar_encrypt"].apply(safe_hex)

# Drop the original binary columns as Excel hates raw binary
df = df.drop(columns=["itm_name_en_encrypt", "itm_name_ar_encrypt"])

# --- 4. SAVE TO EXCEL ---
output_file = "Raw_Encrypted_Names.xlsx"
print(f"Saving to {output_file}...")

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False)

    # Auto-adjust column widths for readability
    worksheet = writer.sheets["Sheet1"]
    worksheet.set_column(0, 0, 20)  # Barcode
    worksheet.set_column(1, 2, 60)  # Hex strings

print("✅ Extraction Complete! File saved.")
