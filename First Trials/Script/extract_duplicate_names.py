import re
import pandas as pd

# Load your extract file
df = pd.read_csv("used_items_extract.csv")


def clean_and_extract_signature(text):
    if not isinstance(text, str):
        return "", "", "", "", ""

    orig = text.strip()
    s = text.upper()

    # 1. Remove obvious ERP noise & pricing artifacts
    s = re.sub(r"[\@\#\*\!\?\[\]\(\)]", " ", s)
    s = re.sub(r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP|جنيه)\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+\s*LE\b", " ", s, flags=re.I)

    # Remove metadata annotations
    noise_patterns = [
        r"\bNEW\s*PRICE\b",
        r"\bOLD\s*PRICE\b",
        r"\bN\.?P\.?\d*\b",
        r"\bN\/PRICE\b",
        r"\bOLD\/PRICE\b",
        r"\bNEU\s*\.?PRICE\b",
        r"سعر\s*جديد",
        r"سعر\s*قديم",
        r"بدون\s*خصم",
        r"موقوف",
        r"\bOFFER\b",
        r"\bFREE\b",
        r"\bPRES\b",
        r"\bEGY\b",
    ]
    for np in noise_patterns:
        s = re.sub(np, " ", s, flags=re.I)

    # 2. Standardize Units and Forms
    s = re.sub(r"\bTABLETS?\b", "TAB", s)
    s = re.sub(r"\bCAPSULES?\b", "CAP", s)
    s = re.sub(r"\bSYRUP\b", "SYP", s)
    s = re.sub(r"\bSUSPENSION\b", "SUSP", s)
    s = re.sub(r"\bAMPOULES?\b", "AMP", s)
    s = re.sub(r"\bOINTMENT\b", "OINT", s)
    s = re.sub(r"\bGM\b|\bGRAMS?\b", "G", s)
    s = re.sub(r"\bML\b|\bMILLILITERS?\b", "ML", s)

    # 3. Extract Strength (e.g. 500MG, 10/25MG, 0.5%, 1G)
    strength_match = re.search(
        r"\b(\d+(\.\d+)?(\s*\/\s*\d+(\.\d+)?)?\s*(MG|MCG|G|IU|%|M|UI))\b", s
    )
    strength = strength_match.group(0).replace(" ", "") if strength_match else ""
    if strength:
        s = s.replace(strength_match.group(0), " ")

    # 4. Extract Form
    forms = [
        "TAB",
        "CAP",
        "SYP",
        "SUSP",
        "DROPS?",
        "DROP",
        "CREAM",
        "OINT",
        "GEL",
        "AMP",
        "VIAL",
        "SPRAY",
        "SACHET",
        "LOTION",
        "WIPES?",
    ]
    form_found = ""
    for f in forms:
        if re.search(r"\b" + f + r"\b", s):
            form_found = f.rstrip("S")  # singularize
            s = re.sub(r"\b" + f + r"\b", " ", s)
            break

    # 5. Extract Modifiers (Plus, Forte, Extra, Compound, XR, SR, Retard)
    modifiers = [
        "PLUS",
        "FORTE",
        "EXTRA",
        "XR",
        "SR",
        "RETARD",
        "COMP",
        "ADVANCE",
        "D",
        "COR",
    ]
    mod_found = []
    for m in modifiers:
        if re.search(r"\b" + m + r"\b", s):
            mod_found.append(m)
            s = re.sub(r"\b" + m + r"\b", " ", s)
    modifier_str = "_".join(sorted(mod_found))

    # 6. Remaining words form the Root Brand Name
    s = re.sub(r"\b\d+\s*(TAB|CAP|AMP|SACHET|DOSE|ML|G)\b", " ", s)  # pack sizes
    s = re.sub(r"\b\d+\b", " ", s)  # standalone digits
    brand_tokens = [w for w in re.split(r"\s+", s) if len(w) > 1]
    brand_root = "_".join(brand_tokens[:2]) if brand_tokens else "UNKNOWN"

    # Construct Canonical Group Key
    signature = f"{brand_root}|{modifier_str}|{strength}|{form_found}"
    return signature, brand_root, modifier_str, strength, form_found


# Apply transformation
df[["Signature", "Brand", "Modifier", "Strength", "Form"]] = df[
    "Item_English_Name"
].apply(lambda x: pd.Series(clean_and_extract_signature(x)))

# Filter for duplicates (groups with count > 1)
group_counts = df.groupby("Signature")["Item_Code"].transform("count")
df["Duplicate_Count"] = group_counts

duplicates_df = df[df["Duplicate_Count"] > 1].sort_values(by=["Signature", "Item_Code"])

# Save results
duplicates_df.to_csv(
    "grouped_duplicates_for_review.csv", index=False, encoding="utf-8-sig"
)

print(f"Total rows: {len(df)}")
print(f"Identified clustered duplicate candidates: {len(duplicates_df)}")
print(f"Unique canonical master products: {duplicates_df['Signature'].nunique()}")
