import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

# ==============================================================================
# DATABASE CONNECTION
# ==============================================================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

# Target Window: Last 6 Months of your data (Apr 2026 - Sep 2026)
START_DATE = "2026-04-01 00:00:00"
END_DATE = "2026-09-30 23:59:59"

params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

def run_unfiltered_audit():
    print("="*80)
    print(f"RUNNING UNFILTERED 6-MONTH AUDIT ({START_DATE[:10]} to {END_DATE[:10]})")
    print("="*80)

    with engine.connect() as conn:
        # ======================================================================
        # 1. EXTRACT ALL EXPENSES / CASH-OUTS (NO RESTRICTIVE FILTERS)
        # ======================================================================
        print("\n[1/2] Extracting ALL expense & cash-out records...")
        
        # Stream A: Shift Drawer Cash-Outs
        sql_shift_expenses = text("""
            SELECT 
                'Shift Drawer' AS Source_Module,
                d.scd_sc_id AS Reference_ID,
                d.scd_start AS Transaction_Date,
                'منصرف نقدية ومصاريف وردية' AS Category,
                CAST(d.scd_c_cashOut AS float) AS Amount,
                CAST(d.scd_c_purchases AS float) AS Purchases_Paid_From_Drawer,
                CAST(d.scd_c_salesReturn AS float) AS Cash_Returns_Refunded,
                ISNULL(sc.sc_notes, '') AS Notes
            FROM Shift_Changes_d d
            LEFT JOIN Shift_Changes sc ON sc.sc_id = d.scd_sc_id
            WHERE d.scd_sto_id = 47
              AND d.scd_start BETWEEN :start_date AND :end_date
              AND (d.scd_c_cashOut > 0 OR d.scd_c_purchases > 0 OR d.scd_c_salesReturn > 0)
        """)
        df_shift_exp = pd.read_sql(sql_shift_expenses, conn, params={"start_date": START_DATE, "end_date": END_DATE})

        # Stream B: General Ledger Cash-Outs (F_Transaction_Header)
        sql_ledger_expenses = text("""
            SELECT 
                'General Ledger' AS Source_Module,
                fh.fh_id AS Reference_ID,
                fh.fh_actual_date AS Transaction_Date,
                CASE 
                    WHEN fh.fh_to_type = '5' THEN ISNULL(oe.oe_name_ar, ISNULL(ac.ac_name_ar, 'مصروفات عامة'))
                    WHEN fh.fh_to_type = '2' THEN 'سداد لمورد نقداً'
                    WHEN fh.fh_to_type = '1' THEN 'صرف لعميل نقداً'
                    WHEN fh.fh_to_type = '3' THEN 'تحويل بين خزائن'
                    ELSE 'صرف نقدية متنوع'
                END AS Category,
                CAST(fh.fh_value AS float) AS Amount,
                0.0 AS Purchases_Paid_From_Drawer,
                0.0 AS Cash_Returns_Refunded,
                ISNULL(fh.fh_notes, '') AS Notes
            FROM F_Transaction_Header fh
            LEFT JOIN Other_Expenses oe ON oe.oe_id = fh.fh_to_id AND fh.fh_to_type = '5'
            LEFT JOIN F_Account_Chart ac ON ac.ac_id = fh.fh_to_id AND fh.fh_to_type = '5'
            WHERE fh.fh_trans_type = '1' -- Cash Out
              AND fh.fh_from_type = '3'  -- From Register / Cash Safe
              AND fh.fh_actual_date BETWEEN :start_date AND :end_date
              AND fh.fh_value > 0
        """)
        df_ledger_exp = pd.read_sql(sql_ledger_expenses, conn, params={"start_date": START_DATE, "end_date": END_DATE})

        # Combine all expenses
        df_all_expenses = pd.concat([df_shift_exp, df_ledger_exp], ignore_index=True)
        df_all_expenses.sort_values(by="Transaction_Date", ascending=True, inplace=True)
        
        exp_file = "unfiltered_expenses_last6months.csv"
        df_all_expenses.to_csv(exp_file, index=False, encoding="utf-8-sig")
        print(f"      -> Exported {len(df_all_expenses)} expense rows to: {exp_file}")

        # ======================================================================
        # 2. EXTRACT ALL CASH INFLOWS (NO RESTRICTIVE FILTERS)
        # ======================================================================
        print("\n[2/2] Extracting ALL cash inflow records...")
        
        # Stream A: Shift Drawer Cash-Ins
        sql_shift_inflows = text("""
            SELECT 
                'Shift Drawer' AS Source_Module,
                d.scd_sc_id AS Reference_ID,
                d.scd_start AS Transaction_Date,
                'وارد نقدية بالوردية' AS Category,
                CAST(d.scd_d_cashIn AS float) AS Amount,
                CAST(d.scd_d_sales AS float) AS Shift_Cash_Sales,
                CAST(d.scd_o_salesVisa AS float) AS Shift_Visa_Sales,
                ISNULL(sc.sc_notes, '') AS Notes
            FROM Shift_Changes_d d
            LEFT JOIN Shift_Changes sc ON sc.sc_id = d.scd_sc_id
            WHERE d.scd_sto_id = 47
              AND d.scd_start BETWEEN :start_date AND :end_date
              AND (d.scd_d_cashIn > 0 OR d.scd_d_sales > 0 OR d.scd_o_salesVisa > 0)
        """)
        df_shift_in = pd.read_sql(sql_shift_inflows, conn, params={"start_date": START_DATE, "end_date": END_DATE})

        # Stream B: General Ledger Cash-Ins (F_Transaction_Header)
        sql_ledger_inflows = text("""
            SELECT 
                'General Ledger' AS Source_Module,
                fh.fh_id AS Reference_ID,
                fh.fh_actual_date AS Transaction_Date,
                CASE 
                    WHEN fh.fh_from_type = '1' THEN ISNULL(c.cust_name_ar, 'توريد من عميل')
                    WHEN fh.fh_from_type = '2' THEN 'توريد من مورد'
                    WHEN fh.fh_from_type = '3' THEN 'تحويل من خزينة أخرى'
                    ELSE 'توريد نقدية عام'
                END AS Category,
                CAST(fh.fh_value AS float) AS Amount,
                0.0 AS Shift_Cash_Sales,
                0.0 AS Shift_Visa_Sales,
                ISNULL(fh.fh_notes, '') AS Notes
            FROM F_Transaction_Header fh
            LEFT JOIN Customer c ON c.cust_id = fh.fh_from_id AND fh.fh_from_type = '1'
            WHERE fh.fh_trans_type = '1' -- Cash In
              AND fh.fh_to_type = '3'    -- To Register / Safe
              AND fh.fh_actual_date BETWEEN :start_date AND :end_date
              AND fh.fh_value > 0
        """)
        df_ledger_in = pd.read_sql(sql_ledger_inflows, conn, params={"start_date": START_DATE, "end_date": END_DATE})

        # Combine all cash inflows
        df_all_inflows = pd.concat([df_shift_in, df_ledger_in], ignore_index=True)
        df_all_inflows.sort_values(by="Transaction_Date", ascending=True, inplace=True)
        
        inflow_file = "unfiltered_cash_inflows_last6months.csv"
        df_all_inflows.to_csv(inflow_file, index=False, encoding="utf-8-sig")
        print(f"      -> Exported {len(df_all_inflows)} inflow rows to: {inflow_file}")

        # ======================================================================
        # 3. TERMINAL SUMMARY BY MONTH
        # ======================================================================
        print("\n" + "="*80)
        print("MONTHLY OVERVIEW: RAW EXPENSES VS. RAW CASH INFLOWS (LAST 6 MONTHS)")
        print("="*80)

        df_all_expenses["Month"] = pd.to_datetime(df_all_expenses["Transaction_Date"]).dt.to_period("M").astype(str)
        df_all_inflows["Month"] = pd.to_datetime(df_all_inflows["Transaction_Date"]).dt.to_period("M").astype(str)

        m_exp = df_all_expenses.groupby("Month")["Amount"].sum().rename("Total_Raw_Cash_Out")
        m_in = df_all_inflows.groupby("Month")["Amount"].sum().rename("Total_Raw_Cash_In")

        overview = pd.concat([m_exp, m_in], axis=1).fillna(0.0)
        print(overview.to_string())

if __name__ == "__main__":
    run_unfiltered_audit()