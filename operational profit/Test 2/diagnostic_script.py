import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text

SERVER = "localhost,1433"
DATABASE = "Genius"
USERNAME = "sa"
PASSWORD = "Emma@2015"

params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

def run_diagnostics():
    with engine.connect() as conn:
        print("="*60)
        print("1. SHIFT REPORT TOTALS (Shift_Changes_d) [2025 - 2026]")
        print("="*60)
        q1 = text("""
            SELECT 
                scd_sto_id,
                YEAR(scd_start) AS Yr,
                COUNT(*) AS Shifts_Count,
                SUM(scd_d_sales) AS Total_Cash_Sales,
                SUM(scd_o_salesVisa) AS Total_Visa_Sales,
                SUM(scd_d_cashIn) AS Total_Cash_In,
                SUM(scd_c_cashOut) AS Total_Cash_Out_Expenses
            FROM Shift_Changes_d
            WHERE scd_start >= '2025-01-01'
            GROUP BY scd_sto_id, YEAR(scd_start)
            ORDER BY scd_sto_id, Yr;
        """)
        print(pd.read_sql(q1, conn).to_string(index=False))

        print("\n" + "="*60)
        print("2. HISTORICAL MONTH-END INVENTORY (ICS_Month_Close)")
        print("="*60)
        q2 = text("""
            SELECT 
                sto_id,
                YEAR(sec_insert_date) AS Yr,
                MONTH(sec_insert_date) AS Mo,
                COUNT(*) AS Stock_Lines,
                SUM(itm_qty * pharm_price) AS Total_Stock_Cost,
                SUM(CASE WHEN itm_expiry_date < sec_insert_date THEN itm_qty * pharm_price ELSE 0 END) AS Expired_Stock_Cost
            FROM ICS_Month_Close
            WHERE sec_insert_date >= '2025-01-01'
            GROUP BY sto_id, YEAR(sec_insert_date), MONTH(sec_insert_date)
            ORDER BY sto_id, Yr, Mo;
        """)
        df_ics = pd.read_sql(q2, conn)
        if df_ics.empty:
            print("ICS_Month_Close has no rows for 2025+. Checking live Item_Class_Store...")
            q2_live = text("""
                SELECT 
                    sto_id,
                    COUNT(*) AS Active_Lines,
                    SUM(itm_qty * pharm_price) AS Current_Stock_Cost,
                    SUM(CASE WHEN itm_expiry_date < GETDATE() THEN itm_qty * pharm_price ELSE 0 END) AS Current_Expired_Cost
                FROM Item_Class_Store
                WHERE itm_qty > 0
                GROUP BY sto_id;
            """)
            print(pd.read_sql(q2_live, conn).to_string(index=False))
        else:
            print(df_ics.to_string(index=False))

        print("\n" + "="*60)
        print("3. SERVICE CUSTOMER SALES (sales_trans_h with cust_id IN (4,5,6,7,18,60,61))")
        print("="*60)
        q3 = text("""
            SELECT 
                h.sto_id,
                h.cust_id,
                c.cust_name_ar,
                COUNT(*) AS Invoices_Count,
                SUM(h.total_bill_net) AS Total_Revenue
            FROM sales_trans_h h
            LEFT JOIN Customer c ON c.cust_id = h.cust_id
            WHERE h.cust_id IN (4, 5, 6, 7, 18, 60, 61)
              AND h.sec_insert_date >= '2025-01-01'
            GROUP BY h.sto_id, h.cust_id, c.cust_name_ar;
        """)
        print(pd.read_sql(q3, conn).to_string(index=False))

        print("\n" + "="*60)
        print("4. EXPENSE CATEGORIES IN LEDGER (F_Transaction_Header)")
        print("="*60)
        q4 = text("""
            SELECT TOP 10
                fh_trans_type,
                fh_from_type,
                fh_to_type,
                COUNT(*) AS Tx_Count,
                SUM(fh_value) AS Total_Sum
            FROM F_Transaction_Header
            WHERE fh_actual_date >= '2025-01-01'
            GROUP BY fh_trans_type, fh_from_type, fh_to_type
            ORDER BY Tx_Count DESC;
        """)
        print(pd.read_sql(q4, conn).to_string(index=False))

if __name__ == "__main__":
    run_diagnostics()