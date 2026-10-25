import urllib.parse
import pandas as pd
from sqlalchemy import create_engine


# --- 1. THE 20-CHARACTER BLOCK DECRYPTION FUNCTION ---
def decrypt_eplus_name(byte_val, lang="en"):
    if byte_val is None:
        return ""
    try:
        text = byte_val.decode("cp1256")
        decrypted = ""

        # E-Plus splits names into 20-character chunks and reverses each chunk
        for i in range(0, len(text), 20):
            chunk = text[i : i + 20]
            decrypted += chunk[::-1]

        decrypted = decrypted.strip()

        # Add BiDi markers to fix LibreOffice/Excel layout issues
        if lang == "ar":
            return "\u200f" + decrypted  # Invisible Right-to-Left mark
        else:
            return "\u200e" + decrypted  # Invisible Left-to-Right mark

    except Exception:
        return str(byte_val)


# --- 2. DATABASE CONNECTION ---
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

# ==========================================
# REPORT 1: ALL ITEMS & FIXED NAMES
# ==========================================
print("Generating All Items Sales Report...")
sql_items = """
    WITH SalesData AS (
        SELECT 
            d.itm_id,
            SUM(d.qnty) as TotalQty,
            AVG(d.itm_sell) as AvgSellPrice,
            AVG(ISNULL(d.itm_cost, 0)) as AvgCostPrice,
            AVG(ISNULL(d.itm_dis_per, 0)) as AvgItemDiscountPerc,
            
            -- Gross Revenue
            SUM(d.qnty * d.itm_sell) as GrossRevenue,
            
            -- Final Net Revenue (applying both line discounts and invoice total discounts)
            SUM( ((((d.qnty * d.itm_sell) * (100.0 - ISNULL(d.itm_dis_per, 0)) / 100.0) - ISNULL(d.itm_dis_mon, 0))) * (100.0 - ISNULL(h.total_dis_per, 0)) / 100.0 ) as FinalNetRevenue,
            
            -- Total Cost
            SUM(d.qnty * ISNULL(d.itm_cost, 0)) as TotalCost,
            
            -- Active Days
            COUNT(DISTINCT CONVERT(varchar(10), h.sec_insert_date, 120)) as DaysSold
            
        FROM sales_trans_d d
        JOIN sales_trans_h h ON d.sth_id = h.sth_id
        WHERE h.sth_flag = 'C' AND d.itm_nexist <> 1 
        GROUP BY d.itm_id
    )
    SELECT 
        i.itm_code AS [Barcode],
        i.itm_name_en_encrypt AS [English Name],
        i.itm_name_ar_encrypt AS [Arabic Name],
        
        sd.TotalQty AS [Total Qty Sold],
        sd.AvgSellPrice AS [Avg Sell Price (EGP)],
        sd.AvgCostPrice AS [Avg Cost Price (EGP)],
        (sd.AvgItemDiscountPerc / 100.0) AS [Avg Discount Given (%)],
        
        sd.FinalNetRevenue AS [Net Revenue (EGP)],
        sd.TotalCost AS [Total Cost (EGP)],
        (sd.FinalNetRevenue - sd.TotalCost) AS [Net Profit (EGP)],
        
        CASE WHEN sd.FinalNetRevenue > 0 
             THEN ((sd.FinalNetRevenue - sd.TotalCost) / sd.FinalNetRevenue) 
             ELSE 0 END AS [Profit Margin (%)],
        
        sd.DaysSold AS [Active Selling Days],
        (sd.TotalQty / sd.DaysSold) AS [Velocity (Qty / Active Day)]

    FROM SalesData sd
    JOIN Item_Catalog i ON sd.itm_id = i.itm_id
    WHERE sd.DaysSold > 0
    ORDER BY sd.FinalNetRevenue DESC;
"""

df_items = pd.read_sql(sql_items, engine)
print("Decrypting item names using 20-character block logic...")
df_items["English Name"] = df_items["English Name"].apply(
    lambda x: decrypt_eplus_name(x, "en")
)
df_items["Arabic Name"] = df_items["Arabic Name"].apply(
    lambda x: decrypt_eplus_name(x, "ar")
)

with pd.ExcelWriter("1_All_Items_Sales_And_Names.xlsx", engine="xlsxwriter") as writer:
    df_items.to_excel(writer, sheet_name="All Items", index=False)
    worksheet = writer.sheets["All Items"]
    workbook = writer.book

    money_format = workbook.add_format({"num_format": "#,##0.00"})
    qty_format = workbook.add_format({"num_format": "#,##0.0"})
    perc_format = workbook.add_format({"num_format": "0.00%"})
    ltr_format = workbook.add_format({"reading_order": 1})
    rtl_format = workbook.add_format({"reading_order": 2})

    worksheet.set_column(1, 1, 40, ltr_format)
    worksheet.set_column(2, 2, 40, rtl_format)
    worksheet.set_column(3, 3, 15, qty_format)
    worksheet.set_column(4, 5, 18, money_format)
    worksheet.set_column(6, 6, 20, perc_format)
    worksheet.set_column(7, 9, 20, money_format)
    worksheet.set_column(10, 10, 18, perc_format)
    worksheet.set_column(12, 12, 22, qty_format)
    worksheet.freeze_panes(1, 0)

    header_format = workbook.add_format(
        {"bold": True, "fg_color": "#4F81BD", "font_color": "white", "border": 1}
    )
    for col_num, value in enumerate(df_items.columns.values):
        worksheet.write(0, col_num, value, header_format)

print("-> Saved 1_All_Items_Sales_And_Names.xlsx")

# ==========================================
# REPORT 2: WEEKLY PROFIT & LOSS
# ==========================================
print("Generating Weekly Profit Report...")
sql_profit = """
    WITH Sales AS (
        SELECT 
            CAST(YEAR(h.sec_insert_date) AS VARCHAR(4)) + '-W' + RIGHT('0' + CAST(DATEPART(wk, h.sec_insert_date) AS VARCHAR(2)), 2) AS WeekGroup,
            SUM( ((((d.qnty * d.itm_sell) * (100.0 - ISNULL(d.itm_dis_per,0))/100.0) - ISNULL(d.itm_dis_mon,0))) * (100.0 - ISNULL(h.total_dis_per,0))/100.0 ) as Revenue,
            SUM( d.qnty * ISNULL(d.itm_cost, 0) ) as COGS,
            SUM( ((((d.itm_back * ISNULL(d.itm_back_price, d.itm_sell)) * (100.0 - ISNULL(d.itm_dis_per,0))/100.0) - ISNULL(d.itm_dis_mon,0))) * (100.0 - ISNULL(h.total_dis_per,0))/100.0 ) as ReturnRevenue,
            SUM( d.itm_back * ISNULL(d.itm_aver_cost, d.itm_cost) ) as ReturnCOGS
        FROM sales_trans_d d
        JOIN sales_trans_h h ON d.sth_id = h.sth_id
        WHERE h.sth_flag = 'C' AND d.itm_nexist <> 1
        GROUP BY CAST(YEAR(h.sec_insert_date) AS VARCHAR(4)) + '-W' + RIGHT('0' + CAST(DATEPART(wk, h.sec_insert_date) AS VARCHAR(2)), 2)
    ),
    Expenses AS (
        SELECT 
            CAST(YEAR(sec_insert_date) AS VARCHAR(4)) + '-W' + RIGHT('0' + CAST(DATEPART(wk, sec_insert_date) AS VARCHAR(2)), 2) AS WeekGroup,
            SUM(fh_value) as TotalExpenses
        FROM F_Transaction_Header
        WHERE fh_to_type = '1' AND fh_trans_type = '1' 
        GROUP BY CAST(YEAR(sec_insert_date) AS VARCHAR(4)) + '-W' + RIGHT('0' + CAST(DATEPART(wk, sec_insert_date) AS VARCHAR(2)), 2)
    )
    SELECT 
        COALESCE(s.WeekGroup, e.WeekGroup) AS [Year-Week],
        ISNULL(s.Revenue, 0) - ISNULL(s.ReturnRevenue, 0) AS [Net Sales Revenue (EGP)],
        ISNULL(s.COGS, 0) - ISNULL(s.ReturnCOGS, 0) AS [Cost of Goods Sold (EGP)],
        (ISNULL(s.Revenue, 0) - ISNULL(s.ReturnRevenue, 0)) - (ISNULL(s.COGS, 0) - ISNULL(s.ReturnCOGS, 0)) AS [Gross Profit (EGP)],
        ISNULL(e.TotalExpenses, 0) AS [Pharmacy Expenses (EGP)],
        ((ISNULL(s.Revenue, 0) - ISNULL(s.ReturnRevenue, 0)) - (ISNULL(s.COGS, 0) - ISNULL(s.ReturnCOGS, 0))) - ISNULL(e.TotalExpenses, 0) AS [Net Profit (EGP)]
    FROM Sales s
    FULL OUTER JOIN Expenses e ON s.WeekGroup = e.WeekGroup
    ORDER BY [Year-Week] DESC;
"""

df_profit = pd.read_sql(sql_profit, engine)
with pd.ExcelWriter("2_Weekly_Profit_Report.xlsx", engine="xlsxwriter") as writer:
    df_profit.to_excel(writer, sheet_name="Weekly Profit", index=False)
    worksheet = writer.sheets["Weekly Profit"]
    workbook = writer.book
    money_format = workbook.add_format({"num_format": "#,##0.00"})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:F", 25, money_format)

    header_format = workbook.add_format(
        {"bold": True, "fg_color": "#4F81BD", "font_color": "white", "border": 1}
    )
    for col_num, value in enumerate(df_profit.columns.values):
        worksheet.write(0, col_num, value, header_format)

print("-> Saved 2_Weekly_Profit_Report.xlsx")
print("✅ Done! Go check out the files!")
