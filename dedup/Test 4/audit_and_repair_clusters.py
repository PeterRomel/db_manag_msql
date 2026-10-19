"""POST-PROCESSING PHARMACEUTICAL INTEGRITY VALIDATOR

Audits and repairs final_deduplicated_catalog_PERFECT.csv locally (0 API Calls).
Prevents illegal merges: Strength clashes, Gel-to-Tablet merges, Mono vs Combo clashes.
"""

import csv
import os
import re
from collections import defaultdict

INPUT_CSV = "final_deduplicated_catalog_PERFECT.csv"
OUTPUT_CSV = "final_deduplicated_catalog_VETTED.csv"


def extract_dosage_form_category(text: str) -> str:
    s = text.upper()
    if re.search(r"\b(GEL|CREAM|OINT|OINTMENT|LOTION|PAINT)\b", s):
        return "TOPICAL"
    if re.search(r"\b(TAB|CAP|TABLET|CAPSULE|PILL|LACTAB|DRAGEE)S?\b", s):
        return "ORAL_SOLID"
    if re.search(r"\b(SYRUP|SUSP|SUSPENSION|ORAL SOL|DROPS|SOLUTION)\b", s):
        return "ORAL_LIQUID"
    if re.search(r"\b(AMP|VIAL|INJECTION|SERINGUE|SYRINGE)S?\b", s):
        return "INJECTABLE"
    if re.search(r"\b(EYE|OPHTH|EAR|OTIC|NASAL|SPRAY)\b", s):
        return "DROPS_SPRAY"
    return "OTHER"


def extract_strength(text: str) -> str:
    # Matches patterns like '50 MG', '1000 IU', '0.5%', '25/250 MG'
    match = re.search(
        r"\b(\d+(\.\d+)?(\/\d+(\.\d+)?)?)\s*(MG|MCG|G|GM|IU|U|%)\b",
        text,
        flags=re.I,
    )
    return match.group(0).upper().replace(" ", "") if match else ""


def extract_pack_count(text: str) -> str:
    # Matches patterns like '20 TAB', '30 CAP', '5 AMP', '100 ML'
    match = re.search(
        r"\b(\d+)\s*(TAB|CAP|TABLET|CAPSULE|AMP|SACH|VIAL|ML|GM)S?\b",
        text,
        flags=re.I,
    )
    return match.group(0).upper().replace(" ", "") if match else ""


def has_combo_marker(text: str) -> bool:
    return bool(
        re.search(
            r"\b(PLUS|CO|COMP|COMPOUND|HCT|FORTE|MAX|EXTRA)\b",
            text,
            flags=re.I,
        )
    )


def audit_and_repair_catalog(input_path: str, output_path: str):
    print(f"[*] Auditing pharmaceutical integrity of: {input_path}...")
    with open(input_path, mode="r", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))

    # Group records by their assigned duplicate_codes cluster
    clusters = defaultdict(list)
    singletons = []

    for idx, r in enumerate(rows):
        dup_str = r.get("duplicate_codes", "").strip()
        if dup_str and int(r.get("number of duplicates", 0)) > 0:
            clusters[dup_str].append(idx)
        else:
            singletons.append(idx)

    repaired_clusters = 0
    split_items_count = 0

    # Validate each multi-item cluster
    for dup_key, indices in list(clusters.items()):
        first_idx = indices[0]
        base_name = rows[first_idx]["name"]
        base_cat = extract_dosage_form_category(base_name)
        base_str = extract_strength(base_name)
        base_pack = extract_pack_count(base_name)
        base_combo = has_combo_marker(base_name)

        invalid_indices = []

        for i in indices[1:]:
            curr_name = rows[i]["name"]
            curr_cat = extract_dosage_form_category(curr_name)
            curr_str = extract_strength(curr_name)
            curr_pack = extract_pack_count(curr_name)
            curr_combo = has_combo_marker(curr_name)

            # RULE 1: Form category clash (e.g. Gel vs Tablet)
            if base_cat != "OTHER" and curr_cat != "OTHER" and base_cat != curr_cat:
                invalid_indices.append(i)
                continue

            # RULE 2: Explicit Strength mismatch (e.g. 5000 IU vs 50000 IU)
            if base_str and curr_str and base_str != curr_str:
                invalid_indices.append(i)
                continue

            # RULE 3: Combo marker clash (e.g. Plain vs PLUS)
            if base_combo != curr_combo:
                invalid_indices.append(i)
                continue

            # RULE 4: Pack size mismatch (e.g. 20 TAB vs 30 TAB, 5 AMP vs 6 AMP)
            if base_pack and curr_pack and base_pack != curr_pack:
                invalid_indices.append(i)
                continue

        # If illegal merges were detected, split them out
        if invalid_indices:
            repaired_clusters += 1
            split_items_count += len(invalid_indices)

            valid_indices = [i for i in indices if i not in invalid_indices]

            # Re-cluster the remaining valid items
            if len(valid_indices) > 1:
                valid_codes = sorted(
                    [rows[i]["code"].strip() for i in valid_indices],
                    key=lambda x: (0, int(x)) if x.isdigit() else (1, x),
                )
                v_dup_str = "; ".join(valid_codes)
                v_master_name = rows[valid_indices[0]]["proposed master name"]
                for vi in valid_indices:
                    rows[vi]["number of duplicates"] = len(valid_codes) - 1
                    rows[vi]["duplicate_codes"] = v_dup_str
                    rows[vi][
                        "extra comments"
                    ] = f"Duplicate record clustered into '{v_master_name}'"
            elif len(valid_indices) == 1:
                idx_only = valid_indices[0]
                rows[idx_only]["number of duplicates"] = "0"
                rows[idx_only]["duplicate_codes"] = ""
                rows[idx_only]["extra comments"] = "Unique SKU"

            # Convert separated items back to clean individual SKUs
            for inv_i in invalid_indices:
                rows[inv_i]["number of duplicates"] = "0"
                rows[inv_i]["duplicate_codes"] = ""
                # Keep their individual trade name cleaned up
                rows[inv_i]["proposed master name"] = (
                    rows[inv_i]["name"].strip().upper()
                )
                rows[inv_i][
                    "extra comments"
                ] = "Unique SKU (Unmerged by Pharmaceutical Safety Validator)"

    # Sort and Write the final file
    fieldnames = [
        "code",
        "name",
        "number of duplicates",
        "duplicate_codes",
        "proposed master name",
        "extra comments",
    ]

    def safe_sort(row):
        c = str(row["code"]).strip()
        return (row["proposed master name"], (0, int(c)) if c.isdigit() else (1, c))

    rows.sort(key=safe_sort)

    with open(output_path, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print("\n" + "=" * 60)
    print("PHARMACEUTICAL INTEGRITY AUDIT COMPLETE")
    print("=" * 60)
    print(f"Total Clusters Audited:       {len(clusters)}")
    print(f"Illegal Clusters Corrected:   {repaired_clusters}")
    print(f"Mismatched Items Split Out:   {split_items_count}")
    print(f"Final Vetted File Written:    '{os.path.abspath(output_path)}'")
    print("=" * 60)


if __name__ == "__main__":
    audit_and_repair_catalog(INPUT_CSV, OUTPUT_CSV)
