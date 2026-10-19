import urllib.parse
import sys
import pandas as pd
from sqlalchemy import create_engine, text

# ==============================================================================
# DATABASE CONFIGURATION
# ==============================================================================
SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

# Configure pandas to display all columns and prevent truncation in console output
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 1000)
pd.set_option("display.max_colwidth", None)

params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")


def run_forensic_diagnostics():
    with engine.connect() as conn:
        
        # ======================================================================
        # 1. INVESTIGATING SEPTEMBER 2026 EXPENSE SPIKE
        # ======================================================================
        print("=" * 80)
        print("1. INVESTIGATING SEPTEMBER 2026 EXPENSE SPIKE (Shift_Changes_d & Shift_Changes)")
        print("=" * 80)
        q1 = text("""
        SELECT 
            d.scd_sc_id AS Shift_ID,
            d.scd_start AS Shift_Date,
            d.scd_c_cashOut AS Cash_Out_Expenses,
            d.scd_d_cashIn AS Cash_In,
            d.scd_d_sales AS Cash_Sales,
            d.scd_c_purchases AS Purchases_Paid_From_Drawer,
            sc.sc_notes AS Shift_Notes
        FROM Shift_Changes_d d
        LEFT JOIN Shift_Changes sc ON sc.sc_id = d.scd_sc_id
        WHERE d.scd_sto_id = 47
          AND d.scd_start >= '2026-09-01' AND d.scd_start <= '2026-09-30 23:59:59'
          AND (d.scd_c_cashOut > 5000 OR d.scd_d_cashIn > 5000)
        ORDER BY d.scd_c_cashOut DESC;
        """)
        df1 = pd.read_sql(q1, conn)
        
        if df1.empty:
            print("No individual shifts in Sep 2026 exceeded 5,000 EGP cash-out. Checking top 10 shifts:")
            q1_alt = text("""
            SELECT TOP 10
                d.scd_sc_id AS Shift_ID, d.scd_start, d.scd_c_cashOut, sc.sc_notes
            FROM Shift_Changes_d d
            LEFT JOIN Shift_Changes sc ON sc.sc_id = d.scd_sc_id
            WHERE d.scd_sto_id = 47 AND d.scd_start >= '2026-09-01'
            ORDER BY d.scd_c_cashOut DESC;
            """)
            df1 = pd.read_sql(q1_alt, conn)
            
        print(df1.to_string(index=False))
        df1.to_csv("september_2026_expense_spike.csv", index=False, encoding="utf-8-sig")
        print("\n[Exported to 'september_2026_expense_spike.csv']\n")


        # ======================================================================
        # 2. MAPPING CUSTOMER ACCOUNT CODES
        # ======================================================================
        print("=" * 80)
        print("2. MAPPING CUSTOMER ACCOUNT CODES: 55, 18, 32, 35 (Customer Table)")
        print("=" * 80)
        q2 = text("""
        SELECT 
            cust_id,
            cust_code,
            cust_name_ar,
            cust_name_en,
            cust_type,
            cust_current_credit,
            cust_notes
        FROM Customer
        WHERE cust_code IN ('55', '18', '32', '35', '055', '018', '032', '035')
           OR cust_id IN (55, 18, 32, 35);
        """)
        df_cust = pd.read_sql(q2, conn)
        print(df_cust.to_string(index=False))
        df_cust.to_csv("customer_mapping.csv", index=False, encoding="utf-8-sig")
        print("\n[Exported to 'customer_mapping.csv']\n")


        # ======================================================================
        # 3. SALES TRANSACTIONS FOR THESE SPECIFIC CUSTOMER CODES
        # ======================================================================
        print("=" * 80)
        print("3. SALES TRANSACTIONS FOR THESE SPECIFIC CUSTOMER CODES (2025 - 2026)")
        print("=" * 80)
        q3 = text("""
        SELECT 
            YEAR(h.sec_insert_date) AS Yr,
            MONTH(h.sec_insert_date) AS Mo,
            c.cust_code,
            c.cust_name_ar,
            COUNT(*) AS Invoices_Count,
            SUM(h.total_bill_net) AS Total_Net_Sales,
            SUM(h.sth_cash) AS Total_Cash_Paid
        FROM sales_trans_h h
        INNER JOIN Customer c ON c.cust_id = h.cust_id
        WHERE (c.cust_code IN ('55', '18', '32', '35', '055', '018', '032', '035')
           OR c.cust_id IN (55, 18, 32, 35))
          AND h.sec_insert_date BETWEEN '2025-01-01' AND '2026-09-30 23:59:59'
        GROUP BY YEAR(h.sec_insert_date), MONTH(h.sec_insert_date), c.cust_code, c.cust_name_ar
        ORDER BY Yr, Mo;
        """)
        df_sales_cust = pd.read_sql(q3, conn)
        if df_sales_cust.empty:
            print("No sales invoices found under these customer codes in sales_trans_h.")
        else:
            print(df_sales_cust.to_string(index=False))
            df_sales_cust.to_csv("customer_sales_transactions.csv", index=False, encoding="utf-8-sig")
            print("\n[Exported to 'customer_sales_transactions.csv']\n")


        # ======================================================================
        # 4. LEDGER MOVEMENTS FOR THESE CUSTOMER CODES
        # ======================================================================
        print("=" * 80)
        print("4. LEDGER MOVEMENTS FOR THESE CUSTOMER CODES (F_Transaction_Header)")
        print("=" * 80)
        q4 = text("""
        SELECT 
            fh.fh_id,
            fh.fh_actual_date,
            fh.fh_trans_type,
            fh.fh_from_type,
            fh.fh_from_id,
            fh.fh_to_type,
            fh.fh_to_id,
            fh.fh_value,
            fh.fh_notes,
            c.cust_code,
            c.cust_name_ar
        FROM F_Transaction_Header fh
        LEFT JOIN Customer c ON c.cust_id = fh.fh_from_id AND fh.fh_from_type = '1'
        WHERE (c.cust_code IN ('55', '18', '32', '35') OR fh.fh_from_id IN (55, 18, 32, 35))
          AND fh.fh_actual_date BETWEEN '2025-01-01' AND '2026-09-30 23:59:59'
        ORDER BY fh.fh_actual_date;
        """)
        df_ledger_cust = pd.read_sql(q4, conn)
        if df_ledger_cust.empty:
            print("No direct cash ledger movements found under these customer codes.")
        else:
            print(df_ledger_cust.to_string(index=False))
            df_ledger_cust.to_csv("customer_ledger_movements.csv", index=False, encoding="utf-8-sig")
            print("\n[Exported to 'customer_ledger_movements.csv']\n")


        # ======================================================================
        # 5. MONTHLY SHIFT CASH-IN VS CASH-OUT FOR STORE 1
        # ======================================================================
        print("=" * 80)
        print("5. MONTHLY SHIFT CASH-IN (وارد نقدية) VS CASH-OUT (منصرف نقدية) FOR STORE 1")
        print("=" * 80)
        q5 = text("""
        SELECT 
            YEAR(scd_start) AS Yr,
            MONTH(scd_start) AS Mo,
            COUNT(*) AS Shifts,
            SUM(scd_d_sales) AS Shift_Cash_Sales,
            SUM(scd_o_salesVisa) AS Shift_Visa_Sales,
            SUM(scd_d_cashIn) AS Shift_Total_Cash_In,
            SUM(scd_c_cashOut) AS Shift_Total_Cash_Out_Expenses
        FROM Shift_Changes_d
        WHERE scd_sto_id = 47
          AND scd_start BETWEEN '2025-01-01' AND '2026-09-30 23:59:59'
        GROUP BY YEAR(scd_start), MONTH(scd_start)
        ORDER BY Yr, Mo;
        """)
        df_monthly = pd.read_sql(q5, conn)
        print(df_monthly.to_string(index=False))
        df_monthly.to_csv("monthly_shift_cash_summary.csv", index=False, encoding="utf-8-sig")
        print("\n[Exported to 'monthly_shift_cash_summary.csv']\n")


if __name__ == "__main__":
    print("Running forensic diagnostics and exporting results to CSV files...\n")
    run_forensic_diagnostics()
    print("All tables have been successfully exported to individual CSV files.")