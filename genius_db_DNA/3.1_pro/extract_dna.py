import os
import urllib.parse
from sqlalchemy import create_engine, text
import pandas as pd

# ==========================================
# 1. DATABASE CONNECTION SETTINGS
# ==========================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

# Safely encode connection parameters
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
# 2. FOLDER STRUCTURE SETUP
# ==========================================
BASE_DIR = "Genius_Database_DNA"
DIRS = {
    "structures": os.path.join(BASE_DIR, "1_Table_Structures"),
    "samples_used": os.path.join(BASE_DIR, "2_Table_Samples", "Used_Tables"),
    "samples_unused": os.path.join(BASE_DIR, "2_Table_Samples", "Unused_Tables"),
    "code": os.path.join(BASE_DIR, "3_Database_Code"),
    "types": os.path.join(BASE_DIR, "4_Data_Types")
}

for d in DIRS.values():
    os.makedirs(d, exist_ok=True)

print(f"[*] Output directory created at: ./{BASE_DIR}")

# ==========================================
# 3. EXTRACTION FUNCTIONS
# ==========================================

def extract_table_structures():
    print("[*] Extracting Table Structures...")
    query = """
    SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, DATA_TYPE, 
           CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE, COLUMN_DEFAULT
    FROM INFORMATION_SCHEMA.COLUMNS
    ORDER BY TABLE_SCHEMA, TABLE_NAME, ORDINAL_POSITION;
    """
    df = pd.read_sql(query, engine)
    df.to_csv(os.path.join(DIRS["structures"], "All_Tables_Columns_Structure.csv"), index=False, encoding='utf-8-sig')

def extract_table_samples():
    print("[*] Analyzing Table Usage and Extracting Samples...")
    # Query to get row counts for all tables to distinguish used vs unused
    row_count_query = """
    SELECT s.name AS SchemaName, t.name AS TableName, p.rows AS RowCounts
    FROM sys.tables t
    INNER JOIN sys.indexes i ON t.object_id = i.object_id
    INNER JOIN sys.partitions p ON i.object_id = p.object_id AND i.index_id = p.index_id
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE t.is_ms_shipped = 0 AND i.type IN (0,1)
    GROUP BY s.name, t.name, p.rows;
    """
    df_counts = pd.read_sql(row_count_query, engine)
    
    unused_tables = []
    
    with engine.connect() as conn:
        for index, row in df_counts.iterrows():
            schema = row['SchemaName']
            table = row['TableName']
            rows = row['RowCounts']
            
            if rows == 0:
                unused_tables.append({'Schema': schema, 'Table': table})
            else:
                # Extract 10 rows sample for USED tables
                try:
                    sample_query = f"SELECT TOP 10 * FROM [{schema}].[{table}]"
                    df_sample = pd.read_sql(sample_query, conn)
                    safe_name = f"{schema}_{table}_sample.csv".replace("/", "_")
                    df_sample.to_csv(os.path.join(DIRS["samples_used"], safe_name), index=False, encoding='utf-8-sig')
                except Exception as e:
                    print(f"    [!] Could not sample [{schema}].[{table}]: {e}")

    # Save list of unused tables
    pd.DataFrame(unused_tables).to_csv(os.path.join(DIRS["samples_unused"], "Unused_Tables_List.csv"), index=False, encoding='utf-8-sig')

def extract_database_code():
    print("[*] Extracting Views, Procedures, and Triggers...")
    
    # 1. Views
    views_query = """
    SELECT s.name AS SchemaName, v.name AS ViewName, m.definition AS ViewDefinition
    FROM sys.views v
    JOIN sys.sql_modules m ON v.object_id = m.object_id
    JOIN sys.schemas s ON v.schema_id = s.schema_id;
    """
    df_views = pd.read_sql(views_query, engine)
    with open(os.path.join(DIRS["code"], "All_Views.sql"), "w", encoding="utf-8") as f:
        for _, row in df_views.iterrows():
            f.write(f"-- ==========================================\n")
            f.write(f"-- VIEW: [{row['SchemaName']}].[{row['ViewName']}]\n")
            f.write(f"-- ==========================================\n")
            f.write(f"{row['ViewDefinition']}\n\nGO\n\n")

    # 2. Stored Procedures
    proc_query = """
    SELECT s.name AS SchemaName, p.name AS ProcedureName, m.definition AS ProcedureDefinition
    FROM sys.procedures p
    JOIN sys.sql_modules m ON p.object_id = m.object_id
    JOIN sys.schemas s ON p.schema_id = s.schema_id;
    """
    df_procs = pd.read_sql(proc_query, engine)
    with open(os.path.join(DIRS["code"], "All_Stored_Procedures.sql"), "w", encoding="utf-8") as f:
        for _, row in df_procs.iterrows():
            f.write(f"-- ==========================================\n")
            f.write(f"-- PROCEDURE: [{row['SchemaName']}].[{row['ProcedureName']}]\n")
            f.write(f"-- ==========================================\n")
            f.write(f"{row['ProcedureDefinition']}\n\nGO\n\n")

    # 3. Triggers
    trigger_query = """
    SELECT s.name AS SchemaName, t.name AS TriggerName, tb.name AS TableName, m.definition AS TriggerDefinition
    FROM sys.triggers t
    JOIN sys.sql_modules m ON t.object_id = m.object_id
    JOIN sys.tables tb ON t.parent_id = tb.object_id
    JOIN sys.schemas s ON tb.schema_id = s.schema_id;
    """
    df_triggers = pd.read_sql(trigger_query, engine)
    with open(os.path.join(DIRS["code"], "All_Triggers.sql"), "w", encoding="utf-8") as f:
        for _, row in df_triggers.iterrows():
            f.write(f"-- ==========================================\n")
            f.write(f"-- TRIGGER: [{row['SchemaName']}].[{row['TriggerName']}] ON TABLE [{row['TableName']}]\n")
            f.write(f"-- ==========================================\n")
            f.write(f"{row['TriggerDefinition']}\n\nGO\n\n")

def extract_data_types():
    print("[*] Extracting Data Types...")
    query = """
    SELECT s.name as SchemaName, t.name as DataTypeName, 
           t.max_length, t.precision, t.scale, t.is_user_defined
    FROM sys.types t
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    ORDER BY t.is_user_defined DESC, t.name;
    """
    df = pd.read_sql(query, engine)
    df.to_csv(os.path.join(DIRS["types"], "All_Data_Types.csv"), index=False, encoding='utf-8-sig')

# ==========================================
# 4. EXECUTION
# ==========================================
if __name__ == "__main__":
    print("Starting Genius Database DNA Extraction...")
    try:
        extract_table_structures()
        extract_table_samples()
        extract_database_code()
        extract_data_types()
        print("\n[SUCCESS] DNA Extraction Complete! Check the 'Genius_Database_DNA' folder.")
    except Exception as e:
        print(f"\n[ERROR] An error occurred: {e}")