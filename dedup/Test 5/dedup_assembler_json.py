"""ASSEMBLER FOR AI STUDIO JSON CLUSTERS

Reads the original raw 6,772-row extract and maps the duplicate clusters
generated from AI Studio into the finalized 6-column catalog schema.
"""

import csv
import json
import os
import re
import pandas as pd

RAW_EXTRACT_CSV = "used_items_extract.csv"
AI_CLUSTERS_JSON = "ai_studio_clusters.json"
OUTPUT_CSV = "catalog_deduplicated_from_aistudio.csv"


def clean_master_name(raw_name: str) -> str:
    s = raw_name.strip()
    # Normalize typos and strip noise/prices
    s = re.sub(r"[@#*!?\[\]{}()]", " ", s)
    s = re.sub(r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP|L\.E\.|جنيه)\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+LE\b", " ", s, flags=re.I)
    s = re.sub(
        r"(\b(?:TAB|CAP|TABLETS|CAPSULES|ML|GM|G|SACH|SACHETS|AMP|VIAL)\b\s*)\d{2,4}$",
        r"\1",
        s,
        flags=re.I,
    )
    s = re.sub(
        r"\s*(NEW(\.|\s*)?PRICE.*|OLD(\.|\s*)?PRICE.*|N\.?P.*|سعر جديد.*|سعر قديم.*|موقوف.*)$",
        "",
        s,
        flags=re.I,
    )
    s = re.sub(
        r"\s+\b(m|k|pres|aug|t3a|sedico|pharco|amryia|amrya|mup|eipico)\b\s*$",
        "",
        s,
        flags=re.I,
    )
    return re.sub(r"\s+", " ", s).strip().upper()


def build_catalog_from_json():
    print(f"[*] Reading {RAW_EXTRACT_CSV}...")
    df = pd.read_csv(RAW_EXTRACT_CSV)

    code_col = (
        "Item_Code"
        if "Item_Code" in df.columns
        else ("code" if "code" in df.columns else df.columns[0])
    )
    name_col = (
        "Item_English_Name"
        if "Item_English_Name" in df.columns
        else ("name" if "name" in df.columns else df.columns[1])
    )

    items = {}
    for _, row in df.iterrows():
        c = str(row[code_col]).strip()
        n = str(row[name_col]).strip()
        items[c] = n

    print(f"[*] Loaded {len(items)} items from raw extract.")

    print(f"[*] Reading clusters from {AI_CLUSTERS_JSON}...")
    with open(AI_CLUSTERS_JSON, "r", encoding="utf-8") as f:
        clusters = json.load(f)

    # Map each code to its cluster
    code_to_cluster = {}
    for cluster in clusters:
        clean_cluster = sorted(
            list(set(str(c).strip() for c in cluster if str(c).strip())),
            key=lambda x: (0, int(x)) if x.isdigit() else (1, x),
        )
        if len(clean_cluster) > 1:
            for c in clean_cluster:
                code_to_cluster[c] = clean_cluster

    # Assemble rows
    final_rows = []
    for code, raw_name in items.items():
        if code in code_to_cluster:
            sibling_codes = code_to_cluster[code]
            master_code = sibling_codes[0]
            master_name = clean_master_name(items.get(master_code, raw_name))
            dup_count = len(sibling_codes) - 1
            dup_str = "; ".join(sibling_codes)
            comment = f"Duplicate record clustered into '{master_name}'"
        else:
            master_name = clean_master_name(raw_name)
            dup_count = 0
            dup_str = ""
            comment = "Unique SKU"

        final_rows.append(
            {
                "code": code,
                "name": raw_name,
                "number of duplicates": dup_count,
                "duplicate_codes": dup_str,
                "proposed master name": master_name,
                "extra comments": comment,
            }
        )

    # Sort alphabetically by master name, then by code
    final_rows.sort(
        key=lambda x: (
            x["proposed master name"],
            (0, int(x["code"])) if x["code"].isdigit() else (1, x["code"]),
        )
    )

    fieldnames = [
        "code",
        "name",
        "number of duplicates",
        "duplicate_codes",
        "proposed master name",
        "extra comments",
    ]
    with open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(final_rows)

    print("\n" + "=" * 60)
    print(f"[✔] Successfully exported {len(final_rows)} records to: {OUTPUT_CSV}")
    dups = sum(1 for r in final_rows if int(r.get("number of duplicates", 0)) > 0)
    print(f"[✔] Total Duplicate Rows Mapped: {dups}")
    print(f"[✔] Total Unique SKUs:           {len(final_rows) - dups}")
    print("=" * 60)


if __name__ == "__main__":
    build_catalog_from_json()
