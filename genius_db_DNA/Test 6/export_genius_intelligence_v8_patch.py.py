import re

path = "export_genius_intelligence_v8.py"
with open(path, "r", encoding="utf-8") as f:
    code = f.read()

# 1. Fix line 152 syntax error
code = code.replace("except ValueError, TypeError:", "except (ValueError, TypeError):")

# 2. Integer formatting fix for counts, days, ranks
code = re.sub(
    r'(\["qty",\s*"quantity",\s*"count",\s*"items",\s*"no_",\s*"units",\s*"lines",\s*"days",\s*"hours",\s*"rank",?\s*\]\s*\):\s*return\s*"#,##0\.00")',
    '["rank", "days", "dayes", "lines", "count", "hour", "hours"]): return "#,##0"\n    if any(term in h for term in ["qty", "quantity", "items", "no_", "units"]): return "#,##0.00"',
    code
)

# 3. Pro-rata cash discount in sales_line_items
code = code.replace(
    "- ISNULL(ISNULL(h.total_des_mon, 0) / NULLIF(h.no_of_items, 0), 0)",
    "- (ISNULL(h.total_des_mon, 0) * (((d.qnty - ISNULL(d.itm_back, 0)) * d.itm_sell) / NULLIF(h.total_bill, 0)))"
)

# 4. Retain 2024 ONLY on 8 heavy transaction tables
heavy = {
    'q_sales_d', 'q_sales_h', 'q_payments', 'q_treasury_trans',
    'q_gl_engine', 'q_forensic_stock', 'q_vendor_pricing', 'q_pur_d'
}

def clean_block(m):
    qname, body = m.group(1), m.group(2)
    if qname in heavy:
        return f'{qname} = """{body}"""'
    b = body.replace(' (Excl. 2024)', '').replace(', Excl. 2024', '')
    b = re.sub(r'\s*WHERE\s*\([^)]*2024[^)]*\)', '', b)
    b = re.sub(r'\s*AND\s*\([^)]*2024[^)]*\)', '', b)
    b = re.sub(r'\s*WHERE\s+sal\.sal_year\s*<>\s*2024', '', b)
    b = re.sub(r'\s*AND\s+YEAR\([^)]+\)\s*<>\s*2024', '', b)
    b = re.sub(r'\s*WHERE\s+YEAR\([^)]+\)\s*<>\s*2024\s*;', ';', b)
    b = re.sub(r'\s*WHERE\s+YEAR\([^)]+\)\s*<>\s*2024', '', b)
    return f'{qname} = """{b}"""'

code = re.sub(r'(q_[a-zA-Z0-9_]+)\s*=\s*"""([\s\S]*?)"""', clean_block, code)

with open(path, "w", encoding="utf-8") as f:
    f.write(code)

print(">>> export_genius_intelligence_v8.py successfully patched!")