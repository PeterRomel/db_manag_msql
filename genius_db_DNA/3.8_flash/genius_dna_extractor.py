import os
import sys
import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

# --- DATABASE CONNECTION SETTINGS ---
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

# --- OUTPUT DIRECTORY STRUCTURE ---
BASE_DIR = "Genius_DNA_Extraction"
DIRS = {
    "arch": os.path.join(BASE_DIR, "01_Architecture_Summary"),
    "tables": os.path.join(BASE_DIR, "02_Tables_Structure"),
    "samples_used": os.path.join(BASE_DIR, "03_Table_Samples", "used_tables"),
    "samples_overview": os.path.join(BASE_DIR, "03_Table_Samples"),
    "views": os.path.join(BASE_DIR, "04_Views"),
    "procs": os.path.join(BASE_DIR, "05_Stored_Procedures"),
    "triggers": os.path.join(BASE_DIR, "06_Triggers"),
}

for d in DIRS.values():
    os.makedirs(d, exist_ok=True)


def sanitize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Format binary/byte blobs into readable token representations to avoid breaking exports."""
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(
                lambda x: f"<HEX_BLOB:{len(x)}B>" if isinstance(x, (bytes, bytearray)) else x
            )
    return df


print("=" * 70)
print("🚀 STARTING GENIUS DATABASE COMPREHENSIVE DNA EXTRACTION")
print("=" * 70)

with engine.connect() as conn:

    # =========================================================================
    # 1. TABLE STRUCTURE & METADATA
    # =========================================================================
    print("\n[1/6] Extracting Tables Overview, Columns, Keys, and Indexes...")

    # 1.1 Tables Overview with exact live row count & disk space
    tables_overview_query = """
    SELECT 
        s.name AS [schema_name],
        t.name AS [table_name],
        p.rows AS [row_count],
        (SUM(a.total_pages) * 8) AS [total_space_kb],
        (SUM(a.used_pages) * 8) AS [used_space_kb],
        CASE WHEN p.rows > 0 THEN 'ACTIVE' ELSE 'EMPTY' END AS [status]
    FROM sys.tables t
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    INNER JOIN sys.indexes i ON t.object_id = i.object_id
    INNER JOIN sys.partitions p ON i.object_id = p.object_id AND i.index_id = p.index_id
    INNER JOIN sys.allocation_units a ON p.partition_id = a.container_id
    WHERE t.is_ms_shipped = 0 AND i.index_id <= 1
    GROUP BY s.name, t.name, p.rows
    ORDER BY p.rows DESC, t.name;
    """
    df_tables = pd.read_sql(tables_overview_query, conn)
    df_tables.to_csv(os.path.join(DIRS["tables"], "01_tables_overview.csv"), index=False, encoding="utf-8-sig")

    # 1.2 Columns Schema
    columns_query = """
    SELECT 
        s.name AS [schema_name],
        t.name AS [table_name],
        c.column_id,
        c.name AS [column_name],
        tp.name AS [data_type],
        c.max_length,
        c.precision,
        c.scale,
        c.is_nullable,
        c.is_identity,
        c.is_computed,
        OBJECT_DEFINITION(c.default_object_id) AS [default_constraint]
    FROM sys.columns c
    INNER JOIN sys.tables t ON c.object_id = t.object_id
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    INNER JOIN sys.types tp ON c.user_type_id = tp.user_type_id
    WHERE t.is_ms_shipped = 0
    ORDER BY s.name, t.name, c.column_id;
    """
    df_cols = pd.read_sql(columns_query, conn)
    df_cols.to_csv(os.path.join(DIRS["tables"], "02_columns_schema.csv"), index=False, encoding="utf-8-sig")

    # 1.3 Primary Keys
    pk_query = """
    SELECT 
        tc.TABLE_SCHEMA AS [schema_name],
        tc.TABLE_NAME AS [table_name],
        tc.CONSTRAINT_NAME AS [pk_name],
        kcu.COLUMN_NAME AS [column_name],
        kcu.ORDINAL_POSITION AS [position]
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
    JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu 
        ON tc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME 
        AND tc.TABLE_SCHEMA = kcu.TABLE_SCHEMA
    WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
    ORDER BY tc.TABLE_SCHEMA, tc.TABLE_NAME, kcu.ORDINAL_POSITION;
    """
    df_pk = pd.read_sql(pk_query, conn)
    df_pk.to_csv(os.path.join(DIRS["tables"], "03_primary_keys.csv"), index=False, encoding="utf-8-sig")

    # 1.4 Foreign Keys
    fk_query = """
    SELECT 
        fk.name AS [foreign_key_name],
        s_parent.name AS [parent_schema],
        t_parent.name AS [parent_table],
        c_parent.name AS [parent_column],
        s_ref.name AS [referenced_schema],
        t_ref.name AS [referenced_table],
        c_ref.name AS [referenced_column],
        fk.delete_referential_action_desc AS [on_delete],
        fk.update_referential_action_desc AS [on_update]
    FROM sys.foreign_keys fk
    INNER JOIN sys.foreign_key_columns fkc ON fk.object_id = fkc.constraint_object_id
    INNER JOIN sys.tables t_parent ON fkc.parent_object_id = t_parent.object_id
    INNER JOIN sys.schemas s_parent ON t_parent.schema_id = s_parent.schema_id
    INNER JOIN sys.columns c_parent ON fkc.parent_object_id = c_parent.object_id AND fkc.parent_column_id = c_parent.column_id
    INNER JOIN sys.tables t_ref ON fkc.referenced_object_id = t_ref.object_id
    INNER JOIN sys.schemas s_ref ON t_ref.schema_id = s_ref.schema_id
    INNER JOIN sys.columns c_ref ON fkc.referenced_object_id = c_ref.object_id AND fkc.referenced_column_id = c_ref.column_id
    ORDER BY t_parent.name, fk.name;
    """
    df_fk = pd.read_sql(fk_query, conn)
    df_fk.to_csv(os.path.join(DIRS["tables"], "04_foreign_keys.csv"), index=False, encoding="utf-8-sig")

    # 1.5 Indexes
    indexes_query = """
    SELECT 
        s.name AS [schema_name],
        t.name AS [table_name],
        i.name AS [index_name],
        i.type_desc AS [index_type],
        i.is_unique,
        i.is_primary_key,
        STRING_AGG(c.name, ', ') WITHIN GROUP (ORDER BY ic.key_ordinal) AS [indexed_columns]
    FROM sys.indexes i
    INNER JOIN sys.tables t ON i.object_id = t.object_id
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
    INNER JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
    WHERE t.is_ms_shipped = 0 AND i.name IS NOT NULL
    GROUP BY s.name, t.name, i.name, i.type_desc, i.is_unique, i.is_primary_key
    ORDER BY s.name, t.name, i.name;
    """
    try:
        df_indexes = pd.read_sql(indexes_query, conn)
        df_indexes.to_csv(os.path.join(DIRS["tables"], "05_indexes.csv"), index=False, encoding="utf-8-sig")
    except Exception as e:
        print(f"⚠️ Index query warning (older SQL Server syntax compatibility): {e}")

    # =========================================================================
    # 2. DATA TYPES (System & User-Defined)
    # =========================================================================
    print("\n[2/6] Extracting System and User-Defined Data Types...")
    datatypes_query = """
    SELECT 
        s.name AS [schema_name],
        t.name AS [type_name],
        st.name AS [base_system_type],
        t.max_length,
        t.precision,
        t.scale,
        t.is_nullable,
        t.is_user_defined
    FROM sys.types t
    LEFT JOIN sys.schemas s ON t.schema_id = s.schema_id
    LEFT JOIN sys.types st ON t.system_type_id = st.user_type_id
    ORDER BY t.is_user_defined DESC, t.name;
    """
    df_types = pd.read_sql(datatypes_query, conn)
    df_types.to_csv(os.path.join(DIRS["tables"], "06_all_data_types.csv"), index=False, encoding="utf-8-sig")

    # =========================================================================
    # 3. 10 ROWS SAMPLING (ACTIVE VS UNUSED TABLES)
    # =========================================================================
    print("\n[3/6] Sampling 10 rows from active tables & cataloging empty ones...")
    used_summary = []
    empty_summary = []
    consolidated_sample_path = os.path.join(DIRS["samples_overview"], "all_used_samples_consolidated.txt")

    with open(consolidated_sample_path, "w", encoding="utf-8") as cons_f:
        cons_f.write(f"GENIUS DATABASE ACTIVE TABLE SAMPLES (TOP 10 ROWS)\n{'='*70}\n\n")

        for idx, row in df_tables.iterrows():
            schema = row["schema_name"]
            table = row["table_name"]
            row_count = row["row_count"]

            if row_count == 0:
                empty_summary.append({"schema": schema, "table_name": table, "row_count": 0})
                continue

            used_summary.append({"schema": schema, "table_name": table, "row_count": row_count})

            # Fetch sample
            try:
                sample_query = f"SELECT TOP 10 * FROM [{schema}].[{table}]"
                df_sample = pd.read_sql(sample_query, conn)
                df_sample = sanitize_dataframe(df_sample)

                # Save individual table CSV
                sample_file = os.path.join(DIRS["samples_used"], f"{table}_top10.csv")
                df_sample.to_csv(sample_file, index=False, encoding="utf-8-sig")

                # Append to unified text file
                cons_f.write(f"TABLE: [{schema}].[{table}] | TOTAL ROWS: {row_count:,}\n")
                cons_f.write("-" * 70 + "\n")
                cons_f.write(df_sample.to_string(index=False) + "\n\n")

            except Exception as err:
                print(f"❌ Error sampling table [{schema}].[{table}]: {err}")

    # Output active vs unused summary CSVs
    pd.DataFrame(used_summary).to_csv(os.path.join(DIRS["samples_overview"], "active_tables_list.csv"), index=False, encoding="utf-8-sig")
    pd.DataFrame(empty_summary).to_csv(os.path.join(DIRS["samples_overview"], "unused_empty_tables_list.csv"), index=False, encoding="utf-8-sig")
    print(f"   -> Active tables sampled: {len(used_summary)}")
    print(f"   -> Empty/unused tables: {len(empty_summary)}")

    # =========================================================================
    # 4. VIEWS EXTRACTION (Metadata + DDL Code)
    # =========================================================================
    print("\n[4/6] Extracting Views and DDL SQL definitions...")
    views_query = """
    SELECT 
        s.name AS [schema_name],
        v.name AS [view_name],
        v.create_date,
        v.modify_date,
        sm.definition AS [view_definition]
    FROM sys.views v
    INNER JOIN sys.schemas s ON v.schema_id = s.schema_id
    LEFT JOIN sys.sql_modules sm ON v.object_id = sm.object_id
    WHERE v.is_ms_shipped = 0
    ORDER BY s.name, v.name;
    """
    df_views = pd.read_sql(views_query, conn)
    df_views[["schema_name", "view_name", "create_date", "modify_date"]].to_csv(
        os.path.join(DIRS["views"], "01_views_metadata.csv"), index=False, encoding="utf-8-sig"
    )

    views_sql_path = os.path.join(DIRS["views"], "all_views_definitions.sql")
    with open(views_sql_path, "w", encoding="utf-8") as f:
        f.write("-- ============================================================\n")
        f.write("-- GENIUS DATABASE: ALL VIEWS DEFINITIONS\n")
        f.write("-- ============================================================\n\n")
        for _, row in df_views.iterrows():
            f.write(f"-- VIEW: [{row['schema_name']}].[{row['view_name']}]\n")
            f.write(f"-- Modified: {row['modify_date']}\n")
            f.write((row["view_definition"] or "-- No definition available").strip())
            f.write("\nGO\n\n" + "-" * 70 + "\n\n")

    # =========================================================================
    # 5. PROCEDURES EXTRACTION (Metadata, Parameters, DDL Code)
    # =========================================================================
    print("\n[5/6] Extracting Stored Procedures, Parameters, and DDL Code...")
    procs_query = """
    SELECT 
        s.name AS [schema_name],
        p.name AS [procedure_name],
        p.create_date,
        p.modify_date,
        sm.definition AS [procedure_definition]
    FROM sys.procedures p
    INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
    LEFT JOIN sys.sql_modules sm ON p.object_id = sm.object_id
    WHERE p.is_ms_shipped = 0
    ORDER BY s.name, p.name;
    """
    df_procs = pd.read_sql(procs_query, conn)
    df_procs[["schema_name", "procedure_name", "create_date", "modify_date"]].to_csv(
        os.path.join(DIRS["procs"], "01_procedures_metadata.csv"), index=False, encoding="utf-8-sig"
    )

    procs_sql_path = os.path.join(DIRS["procs"], "all_procedures_definitions.sql")
    with open(procs_sql_path, "w", encoding="utf-8") as f:
        f.write("-- ============================================================\n")
        f.write("-- GENIUS DATABASE: ALL STORED PROCEDURES DEFINITIONS\n")
        f.write("-- ============================================================\n\n")
        for _, row in df_procs.iterrows():
            f.write(f"-- PROCEDURE: [{row['schema_name']}].[{row['procedure_name']}]\n")
            f.write(f"-- Modified: {row['modify_date']}\n")
            f.write((row["procedure_definition"] or "-- No definition available").strip())
            f.write("\nGO\n\n" + "-" * 70 + "\n\n")

    proc_params_query = """
    SELECT 
        s.name AS [schema_name],
        p.name AS [procedure_name],
        pr.parameter_id,
        pr.name AS [parameter_name],
        t.name AS [data_type],
        pr.max_length,
        pr.is_output
    FROM sys.parameters pr
    INNER JOIN sys.procedures p ON pr.object_id = p.object_id
    INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
    INNER JOIN sys.types t ON pr.user_type_id = t.user_type_id
    WHERE p.is_ms_shipped = 0
    ORDER BY s.name, p.name, pr.parameter_id;
    """
    df_proc_params = pd.read_sql(proc_params_query, conn)
    df_proc_params.to_csv(os.path.join(DIRS["procs"], "02_procedures_parameters.csv"), index=False, encoding="utf-8-sig")

    # =========================================================================
    # 6. TRIGGERS EXTRACTION (Metadata & DDL Code)
    # =========================================================================
    print("\n[6/6] Extracting Triggers and DDL SQL definitions...")
    triggers_query = """
    SELECT 
        s.name AS [parent_schema],
        OBJECT_NAME(tr.parent_id) AS [parent_table_or_view],
        tr.name AS [trigger_name],
        tr.is_disabled,
        tr.is_instead_of_trigger,
        tr.create_date,
        tr.modify_date,
        sm.definition AS [trigger_definition]
    FROM sys.triggers tr
    LEFT JOIN sys.objects o ON tr.parent_id = o.object_id
    LEFT JOIN sys.schemas s ON o.schema_id = s.schema_id
    LEFT JOIN sys.sql_modules sm ON tr.object_id = sm.object_id
    WHERE tr.is_ms_shipped = 0
    ORDER BY [parent_table_or_view], tr.name;
    """
    df_triggers = pd.read_sql(triggers_query, conn)
    df_triggers[["parent_schema", "parent_table_or_view", "trigger_name", "is_disabled", "is_instead_of_trigger", "create_date", "modify_date"]].to_csv(
        os.path.join(DIRS["triggers"], "01_triggers_metadata.csv"), index=False, encoding="utf-8-sig"
    )

    triggers_sql_path = os.path.join(DIRS["triggers"], "all_triggers_definitions.sql")
    with open(triggers_sql_path, "w", encoding="utf-8") as f:
        f.write("-- ============================================================\n")
        f.write("-- GENIUS DATABASE: ALL TRIGGERS DEFINITIONS\n")
        f.write("-- ============================================================\n\n")
        for _, row in df_triggers.iterrows():
            f.write(f"-- TRIGGER: [{row['trigger_name']}] ON [{row['parent_schema']}].[{row['parent_table_or_view']}]\n")
            f.write(f"-- Disabled: {row['is_disabled']} | Instead Of: {row['is_instead_of_trigger']}\n")
            f.write((row["trigger_definition"] or "-- No definition available").strip())
            f.write("\nGO\n\n" + "-" * 70 + "\n\n")

    # =========================================================================
    # 7. GENERATE MASTER DNA SUMMARY REPORT
    # =========================================================================
    summary_md_path = os.path.join(DIRS["arch"], "MASTER_GENIUS_DNA_SUMMARY.md")
    with open(summary_md_path, "w", encoding="utf-8") as f:
        f.write("# GENIUS (e-Plus) Database Complete Architectural Blueprint\n\n")
        f.write("## 1. High-Level Inventory\n")
        f.write(f"- **Total Tables Identified**: {len(df_tables):,}\n")
        f.write(f"  - **Active / Populated Tables**: {len(used_summary):,}\n")
        f.write(f"  - **Empty / Unused Tables**: {len(empty_summary):,}\n")
        f.write(f"- **Total Table Columns**: {len(df_cols):,}\n")
        f.write(f"- **Total Primary Key Constraints**: {len(df_pk):,}\n")
        f.write(f"- **Total Foreign Key Relationships**: {len(df_fk):,}\n")
        f.write(f"- **Total Views**: {len(df_views):,}\n")
        f.write(f"- **Total Stored Procedures**: {len(df_procs):,}\n")
        f.write(f"- **Total Triggers**: {len(df_triggers):,}\n\n")

        f.write("## 2. Top 15 Heaviest Active Tables by Row Count\n\n")
        f.write("| Schema | Table Name | Row Count | Total Size (KB) | Used Size (KB) |\n")
        f.write("| :--- | :--- | :--- | :--- | :--- |\n")
        for _, row in df_tables.head(15).iterrows():
            f.write(f"| `{row['schema_name']}` | `{row['table_name']}` | {row['row_count']:,} | {row['total_space_kb']:,} KB | {row['used_space_kb']:,} KB |\n")

        f.write("\n## 3. Directory Layout Guide\n")
        f.write("```text\n")
        f.write("Genius_DNA_Extraction/\n")
        f.write("├── 01_Architecture_Summary/       -> Master architectural executive summary\n")
        f.write("├── 02_Tables_Structure/            -> Columns, PKs, FKs, Indexes, Data Types\n")
        f.write("├── 03_Table_Samples/\n")
        f.write("│   ├── used_tables/               -> 10-row CSV samples for every active table\n")
        f.write("│   ├── active_tables_list.csv     -> Index of active tables\n")
        f.write("│   ├── unused_empty_tables_list.csv -> Index of unused tables\n")
        f.write("│   └── all_used_samples_consolidated.txt -> Unified raw text sampler\n")
        f.write("├── 04_Views/                      -> Views metadata & full SQL definitions\n")
        f.write("├── 05_Stored_Procedures/          -> Procs metadata, parameters, full SQL DDL\n")
        f.write("└── 06_Triggers/                   -> Triggers metadata & full SQL DDL\n")
        f.write("```\n")

print("\n" + "=" * 70)
print("✅ EXTRACTION COMPLETE! All files written to folder: './Genius_DNA_Extraction'")
print("=" * 70)