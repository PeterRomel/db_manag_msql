import urllib.parse
from sqlalchemy import create_engine, text
import pandas as pd
import os

# --- DATABASE CONNECTION SETTINGS ---
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

# Create a directory to store the output
output_dir = "Genius_DB_Blueprint"
os.makedirs(output_dir, exist_ok=True)
os.makedirs(f"{output_dir}/Table_Samples", exist_ok=True)

print("Starting Database Extraction...")

with engine.connect() as conn:
    # 1. Extract Schema (Tables, Columns, Data Types)
    print("Extracting Schema...")
    schema_query = """
    SELECT 
        t.name AS TableName, c.name AS ColumnName, ty.name AS DataType, 
        c.max_length AS MaxLength, c.is_nullable AS IsNullable,
        ISNULL(i.is_primary_key, 0) AS IsPrimaryKey
    FROM sys.tables t
    INNER JOIN sys.columns c ON t.object_id = c.object_id
    INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
    LEFT JOIN sys.index_columns ic ON ic.object_id = c.object_id AND ic.column_id = c.column_id
    LEFT JOIN sys.indexes i ON ic.object_id = i.object_id AND ic.index_id = i.index_id AND i.is_primary_key = 1
    ORDER BY t.name, c.column_id;
    """
    df_schema = pd.read_sql(text(schema_query), conn)
    df_schema.to_csv(f"{output_dir}/1_Database_Schema.csv", index=False, encoding='utf-8-sig')

    # 2. Extract Foreign Keys (Relationships)
    print("Extracting Relationships...")
    fk_query = """
    SELECT 
        fk.name AS ForeignKeyName, tp.name AS ParentTable, cp.name AS ParentColumn,
        tr.name AS ReferencedTable, cr.name AS ReferencedColumn
    FROM sys.foreign_keys fk
    INNER JOIN sys.tables tp ON fk.parent_object_id = tp.object_id
    INNER JOIN sys.tables tr ON fk.referenced_object_id = tr.object_id
    INNER JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id = fk.object_id
    INNER JOIN sys.columns cp ON fkc.parent_column_id = cp.column_id AND fkc.parent_object_id = cp.object_id
    INNER JOIN sys.columns cr ON fkc.referenced_column_id = cr.column_id AND fkc.referenced_object_id = cr.object_id;
    """
    df_fk = pd.read_sql(text(fk_query), conn)
    df_fk.to_csv(f"{output_dir}/2_Table_Relationships.csv", index=False, encoding='utf-8-sig')

    # 3. Extract Views
    print("Extracting Views...")
    views_query = "SELECT TABLE_NAME, VIEW_DEFINITION FROM INFORMATION_SCHEMA.VIEWS"
    df_views = pd.read_sql(text(views_query), conn)
    df_views.to_csv(f"{output_dir}/3_Database_Views.csv", index=False, encoding='utf-8-sig')

    # 4. Extract Stored Procedures
    print("Extracting Procedures...")
    proc_query = """
    SELECT o.name AS ProcedureName, m.definition AS ProcedureCode 
    FROM sys.objects o 
    INNER JOIN sys.sql_modules m ON o.object_id = m.object_id 
    WHERE o.type = 'P'
    """
    df_procs = pd.read_sql(text(proc_query), conn)
    df_procs.to_csv(f"{output_dir}/4_Stored_Procedures.csv", index=False, encoding='utf-8-sig')

    # 5. Extract Triggers
    print("Extracting Triggers...")
    trigger_query = """
    SELECT t.name AS TriggerName, OBJECT_NAME(t.parent_id) AS TableName, m.definition AS TriggerCode
    FROM sys.triggers t
    INNER JOIN sys.sql_modules m ON t.object_id = m.object_id
    """
    df_triggers = pd.read_sql(text(trigger_query), conn)
    df_triggers.to_csv(f"{output_dir}/5_Database_Triggers.csv", index=False, encoding='utf-8-sig')

    # 6. Extract 5-10 Rows Sample for Every Table
    print("Extracting Table Samples...")
    tables = df_schema['TableName'].unique()
    for table in tables:
        try:
            # Get top 10 rows
            sample_df = pd.read_sql(text(f"SELECT TOP 10 * FROM [{table}]"), conn)
            if not sample_df.empty:
                sample_df.to_csv(f"{output_dir}/Table_Samples/{table}_sample.csv", index=False, encoding='utf-8-sig')
        except Exception as e:
            print(f"Could not sample table {table}: {e}")

print(f"Extraction Complete! All files saved to the '{output_dir}' folder.")