import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

# ==============================================================================
# 1. DATABASE CONFIGURATION
# ==============================================================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

START_DATE = "2025-01-01 00:00:00"
END_DATE = "2026-09-30 23:59:59"

params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

# ==============================================================================
# 2. SQL QUERIES WITH FULL AUDIT TRAILS & REFERENCES
# ==============================================================================

# A. OPERATING EXPENSES (Overhead Accounts in Ledger: fh_to_type = '5')
SQL_EXPENSES_DETAILED = text("""
    SELECT 
        fh.fh_id AS [Ledger_Tx_ID],
        fh.fh_actual_date AS [Transaction_Date],
        CONVERT(varchar(7), fh.fh_actual_date, 120) AS [Year_Month],
        fh.fh_trans_type AS [Trans_Type_Flag],         -- 1 = Cash Out
        'صرف نقدية' AS [Trans_Type_Description],
        fh.fh_from_type AS [From_Type_Flag],           -- 3 = Cash Store / POS Drawer
        'خزينة / نقطة بيع' AS [From_Source_Description],
        fh.fh_from_id AS [From_CashStore_DB_ID],
        ISNULL(fcs.fcs_name_ar, 'نقطة بيع') AS [CashStore_Name],
        fh.fh_to_type AS [To_Type_Flag],               -- 5 = Overhead Expense Account
        'حساب مصروفات' AS [To_Destination_Description],
        fh.fh_to_id AS [To_Expense_Account_DB_ID],
        oe.oe_code AS [Expense_Software_Code],
        ISNULL(oe.oe_name_ar, ISNULL(ac.ac_name_ar, 'مصروفات عامة')) AS [Expense_Category_Name],
        CAST(fh.fh_value AS float) AS [Amount_EGP],
        ISNULL(fh.fh_notes, '') AS [Notes],
        ISNULL(fh.fh_computer, '') AS [Workstation_Computer],
        ISNULL(emp.e_Name, fh.sec_insert_uid) AS [Created_By_User]
    FROM F_Transaction_Header fh
    LEFT JOIN Other_Expenses oe ON oe.oe_id = fh.fh_to_id AND fh.fh_to_type = '5'
    LEFT JOIN F_Account_Chart ac ON ac.ac_id = fh.fh_to_id AND fh.fh_to_type = '5'
    LEFT JOIN F_Cash_Store fcs ON fcs.fcs_id = fh.fh_from_id AND fh.fh_from_type = '3'
    LEFT JOIN Employee emp ON CAST(emp.e_id AS varchar(50)) = fh.sec_insert_uid
    WHERE fh.fh_trans_type = '1' -- Cash Transaction
      AND fh.fh_from_type = '3'  -- Paid out of POS drawer / Cash Store
      AND fh.fh_to_type = '5'    -- Destination is an Operating Expense Account
      AND fh.fh_actual_date BETWEEN :start_date AND :end_date
    ORDER BY fh.fh_actual_date ASC;
""")

# B. CASH INFLOWS (Clinical Services & Rounding Accounts: Customer IDs 4, 5, 6, 61)
SQL_INFLOWS_DETAILED = text("""
    SELECT 
        fh.fh_id AS [Ledger_Tx_ID],
        fh.fh_actual_date AS [Transaction_Date],
        CONVERT(varchar(7), fh.fh_actual_date, 120) AS [Year_Month],
        fh.fh_trans_type AS [Trans_Type_Flag],         -- 1 = Cash In
        'توريد نقدية' AS [Trans_Type_Description],
        fh.fh_from_type AS [From_Type_Flag],           -- 1 = Customer Profile
        'حساب عميل / خدمة' AS [From_Source_Description],
        fh.fh_from_id AS [Customer_DB_ID],             -- Database Primary Key
        c.cust_code AS [Customer_Software_Code],       -- The Code Visible on Screen
        c.cust_name_ar AS [Service_Account_Name],
        fh.fh_to_type AS [To_Type_Flag],               -- 3 = POS Drawer
        'خزينة / نقطة بيع' AS [To_Destination_Description],
        fh.fh_to_id AS [To_CashStore_DB_ID],
        ISNULL(fcs.fcs_name_ar, 'نقطة بيع') AS [CashStore_Name],
        CAST(fh.fh_value AS float) AS [Amount_EGP],
        ISNULL(fh.fh_notes, '') AS [Notes],
        ISNULL(fh.fh_computer, '') AS [Workstation_Computer],
        ISNULL(emp.e_Name, fh.sec_insert_uid) AS [Created_By_User]
    FROM F_Transaction_Header fh
    INNER JOIN Customer c ON c.cust_id = fh.fh_from_id AND fh.fh_from_type = '1'
    LEFT JOIN F_Cash_Store fcs ON fcs.fcs_id = fh.fh_to_id AND fh.fh_to_type = '3'
    LEFT JOIN Employee emp ON CAST(emp.e_id AS varchar(50)) = fh.sec_insert_uid
    WHERE fh.fh_trans_type = '1' -- Cash In
      AND fh.fh_to_type = '3'    -- Paid into POS Drawer
      AND fh.fh_from_type = '1'  -- From Client Account
      AND fh.fh_from_id IN (4, 5, 6, 61) -- Codes: 18 (Sugar), 35 (Pressure), 32 (Compounding), 55 (Surplus)
      AND (fh.fh_notes NOT LIKE '%فاتورة بيع%' OR fh.fh_notes IS NULL) -- Exclude routine counter sales
      AND fh.fh_actual_date BETWEEN :start_date AND :end_date
    ORDER BY fh.fh_actual_date ASC;
""")

# ==============================================================================
# 3. EXECUTION & EXPORT
# ==============================================================================
def main():
    print("="*80)
    print(f"EXTRACTING FULL AUDIT LOGS ({START_DATE[:10]} to {END_DATE[:10]})")
    print("="*80)

    query_params = {"start_date": START_DATE, "end_date": END_DATE}

    with engine.connect() as conn:
        print("\n[1/3] Extracting Operating Expenses...")
        df_expenses = pd.read_sql(SQL_EXPENSES_DETAILED, conn, params=query_params)

        print("[2/3] Extracting Service & Surplus Cash Inflows...")
        df_inflows = pd.read_sql(SQL_INFLOWS_DETAILED, conn, params=query_params)

    # 1. Export Expenses CSV
    exp_filename = "verified_operational_expenses_2025_2026.csv"
    df_expenses.to_csv(exp_filename, index=False, encoding="utf-8-sig")
    print(f"      -> [SUCCESS] Saved {len(df_expenses)} rows to: {exp_filename}")

    # 2. Export Cash Inflows CSV
    inflow_filename = "verified_cash_inflows_services_2025_2026.csv"
    df_inflows.to_csv(inflow_filename, index=False, encoding="utf-8-sig")
    print(f"      -> [SUCCESS] Saved {len(df_inflows)} rows to: {inflow_filename}")

    # ==========================================================================
    # 4. PRINT MONTHLY SUMMARY FOR INSTANT VERIFICATION
    # ==========================================================================
    print("\n" + "="*80)
    print("MONTHLY AUDIT RECONCILIATION SUMMARY (JAN 2025 - SEP 2026)")
    print("="*80)

    exp_monthly = df_expenses.groupby("Year_Month").agg(
        Expense_Tx_Count=("Ledger_Tx_ID", "count"),
        Total_Expenses_EGP=("Amount_EGP", "sum")
    )

    inflow_monthly = df_inflows.groupby("Year_Month").agg(
        Service_Tx_Count=("Ledger_Tx_ID", "count"),
        Total_Service_Inflow_EGP=("Amount_EGP", "sum")
    )

    summary = pd.concat([exp_monthly, inflow_monthly], axis=1).fillna(0.0)
    summary["Total_Expenses_EGP"] = summary["Total_Expenses_EGP"].round(2)
    summary["Total_Service_Inflow_EGP"] = summary["Total_Service_Inflow_EGP"].round(2)

    print(summary.to_string())
    print("="*80)

if __name__ == "__main__":
    main()