import json
import os
import re
import pandas as pd

CHECKPOINT_FILE = "checkpoint_parsed_items.json"
OUTPUT_USER_TARGET_FORMAT = "final_deduplicated_catalog_user_format.csv"
OUTPUT_DUPLICATES_ONLY = "duplicate_clusters_for_review.csv"
OUTPUT_ALL = "deduplicated_pharmacy_master.csv"


def safe_sort_key(code):
    """
    Sorts numeric codes by their integer value and string codes alphabetically
    without ever comparing int to str directly.
    """
    s = str(code).strip()
    if s.isdigit():
        return (0, int(s), "")
    return (1, 0, s)


def finalize_catalog():
    if not os.path.exists(CHECKPOINT_FILE):
        raise FileNotFoundError(
            f"Could not find '{CHECKPOINT_FILE}'. Make sure it is in the same folder."
        )

    print(f"[*] Reading parsed items from '{CHECKPOINT_FILE}'...")
    with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    parsed_df = pd.DataFrame(data)
    print(f"[*] Successfully loaded {len(parsed_df)} parsed items.")

    # 1. Clean and normalize string fields
    parsed_df["canonical_brand"] = (
        parsed_df["canonical_brand"]
        .fillna("UNKNOWN")
        .astype(str)
        .str.upper()
        .str.strip()
    )
    parsed_df["strength"] = (
        parsed_df["strength"].fillna("").astype(str).str.upper().str.strip()
    )
    parsed_df["dosage_form"] = (
        parsed_df["dosage_form"].fillna("OTHER").astype(str).str.upper().str.strip()
    )
    parsed_df["pack_size"] = (
        parsed_df["pack_size"].fillna("").astype(str).str.upper().str.strip()
    )
    parsed_df["item_code"] = parsed_df["item_code"].astype(str).str.strip()
    parsed_df["original_input"] = parsed_df["original_input"].astype(str).str.strip()

    # 2. Build the unique SKU matching signature
    parsed_df["Canonical_SKU_Signature"] = (
        parsed_df["canonical_brand"]
        + " | "
        + parsed_df["strength"]
        + " | "
        + parsed_df["dosage_form"]
        + " | "
        + parsed_df["pack_size"]
    )

    # 3. Calculate Duplicate Count (Cluster size minus 1)
    parsed_df["Duplicate_Count"] = (
        parsed_df.groupby("Canonical_SKU_Signature")["item_code"].transform("count") - 1
    )

    # 4. Aggregate all sibling codes in each cluster (safely sorted)
    def collect_codes(series):
        unique_codes = list(series.unique())
        sorted_codes = sorted(unique_codes, key=safe_sort_key)
        return "; ".join(sorted_codes)

    sibling_codes_map = (
        parsed_df.groupby("Canonical_SKU_Signature")["item_code"]
        .apply(collect_codes)
        .to_dict()
    )
    parsed_df["duplicate_codes"] = parsed_df["Canonical_SKU_Signature"].map(
        sibling_codes_map
    )

    # 5. Build clean Canonical Master Display Name
    parsed_df["proposed_master_name"] = (
        (
            parsed_df["canonical_brand"]
            + " "
            + parsed_df["strength"]
            + " "
            + parsed_df["dosage_form"]
            + " "
            + parsed_df["pack_size"]
        )
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # 6. Build the 6-column target format requested
    target_df = parsed_df.rename(
        columns={
            "item_code": "code",
            "original_input": "name",
            "Duplicate_Count": "number of duplicates",
            "proposed_master_name": "proposed master name",
        }
    )

    # Sort alphabetically by proposed master name, then by code
    # This guarantees that duplicate items appear right next to each other
    target_df["sort_code_helper"] = target_df["code"].apply(safe_sort_key)
    target_df = target_df.sort_values(
        by=["proposed master name", "sort_code_helper"]
    ).drop(columns=["sort_code_helper"])

    target_df["extra comments"] = target_df.apply(
        lambda r: (
            f"Duplicate record clustered into '{r['proposed master name']}'"
            if r["number of duplicates"] > 0
            else "Unique SKU"
        ),
        axis=1,
    )

    # Select final columns
    final_export = target_df[
        [
            "code",
            "name",
            "number of duplicates",
            "duplicate_codes",
            "proposed master name",
            "extra comments",
        ]
    ]

    # Export Files
    final_export.to_csv(OUTPUT_USER_TARGET_FORMAT, index=False, encoding="utf-8-sig")
    parsed_df.to_csv(OUTPUT_ALL, index=False, encoding="utf-8-sig")

    duplicates_df = final_export[final_export["number of duplicates"] > 0]
    duplicates_df.to_csv(OUTPUT_DUPLICATES_ONLY, index=False, encoding="utf-8-sig")

    print("\n" + "=" * 70)
    print("DEDUPLICATION AUDIT COMPLETE")
    print("=" * 70)
    print(f"Total Rows Processed:                    {len(final_export)}")
    print(
        f"Unique Canonical Master SKUs:            {parsed_df['Canonical_SKU_Signature'].nunique()}"
    )
    print(f"Total Duplicate Rows Identified:         {len(duplicates_df)}")
    print(
        f"Total Unique/Single Rows:                {len(final_export) - len(duplicates_df)}"
    )
    print(f"\n[+] Master CSV (All Items, User Format): '{OUTPUT_USER_TARGET_FORMAT}'")
    print(f"[+] Duplicates Review File:              '{OUTPUT_DUPLICATES_ONLY}'")
    print(f"[+] Full Audit Ledger:                   '{OUTPUT_ALL}'")
    print("=" * 70)


if __name__ == "__main__":
    finalize_catalog()
