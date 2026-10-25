import urllib.parse
from sqlalchemy import create_engine
import pandas as pd

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
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")


# --- 2. DECRYPTION ALGORITHM ---
def decode_legacy_bidi_bytes(raw_bytes):
    """Decrypts the obfuscated item names stored in the Genius database."""
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
        return "Decoding Error"


# --- 3. SQL QUERY ---
# We define "used" items as any item that has a history in Sales, Purchases, or the Inventory Ledger.
query = """
SELECT 
    itm_code, 
    itm_name_en_encrypt
FROM Item_Catalog
WHERE itm_id IN (
    SELECT itm_id FROM sales_trans_d
    UNION
    SELECT itm_id FROM pur_trans_d
    UNION
    SELECT itm_id FROM Item_Class
)
"""


def main():
    print("Connecting to the database and fetching used items...")

    # Fetch data into a pandas DataFrame
    df = pd.read_sql(query, engine)
    print(f"Successfully fetched {len(df)} used items. Decrypting English names...")

    # --- 4. PROCESS DATA ---
    # Apply the decryption function to the encrypted column
    df["Item_English_Name"] = df["itm_name_en_encrypt"].apply(decode_legacy_bidi_bytes)

    # Drop the raw binary column to clean up the dataframe
    df = df.drop(columns=["itm_name_en_encrypt"])

    # Rename the code column for better readability in the CSV
    df = df.rename(columns={"itm_code": "Item_Code"})

    # --- 5. EXPORT TO CSV ---
    output_filename = "used_items_extract.csv"

    # Using 'utf-8-sig' ensures that Excel reads the CSV encoding correctly
    df.to_csv(output_filename, index=False, encoding="utf-8-sig")
    print(f"Extraction complete! Data successfully saved to '{output_filename}'.")


if __name__ == "__main__":
    main()
