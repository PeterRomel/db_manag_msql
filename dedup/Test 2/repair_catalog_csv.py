import re
import pandas as pd

INPUT_FILE = "final_deduplicated_catalog.csv"
CLEAN_FILE = "final_deduplicated_catalog_PERFECT.csv"


def repair_and_standardize_csv():
    print(f"[*] Reading and repairing '{INPUT_FILE}'...")

    clean_rows = []

    with open(INPUT_FILE, "r", encoding="utf-8-sig", errors="ignore") as f:
        lines = f.readlines()

    header_skipped = False
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue

        # Skip repeated header lines
        if "proposed master name" in line_str.lower() and not header_skipped:
            header_skipped = True
            continue
        elif "proposed master name" in line_str.lower() and header_skipped:
            continue

        # Fix specific model glitch: 2;"84870; 84830" -> 2,"84870; 84830"
        line_str = re.sub(r'(\d+);"', r'\1,"', line_str)

        # Parse CSV line using regex to respect quotes
        parts = re.split(r',(?=(?:[^"]*"[^"]*")*[^"]*$)', line_str)

        if len(parts) >= 2:
            code = parts[0].strip(" \"'")
            name = parts[1].strip(" \"'")

            # Extract number of duplicates
            dup_count = "0"
            if len(parts) > 2 and parts[2].strip(" \"'").isdigit():
                dup_count = parts[2].strip(" \"'")

            # Extract duplicate codes
            dup_codes = ""
            if len(parts) > 3:
                dup_codes = parts[3].strip(" \"'")

            # Extract proposed master name
            master_name = ""
            if len(parts) > 4:
                master_name = parts[4].strip(" \"'")

            # Extract comments
            comments = ""
            if len(parts) > 5:
                comments = parts[5].strip(" \"'")

            # Fallback for blank master names (the 0,,,, glitch)
            if not master_name:
                # Clean the raw name as fallback
                clean_n = re.sub(
                    r"\b\d+(\.\d+)?\s*(LE|L\.E|EGP)\b", "", name, flags=re.I
                )
                clean_n = re.sub(r"[@#*!?]", "", clean_n).strip()
                master_name = clean_n.upper()

            clean_rows.append(
                {
                    "code": code,
                    "name": name,
                    "raw_dup_count": dup_count,
                    "duplicate_codes": dup_codes,
                    "proposed master name": master_name,
                    "extra comments": comments,
                }
            )

    df = pd.DataFrame(clean_rows)
    print(f"[*] Parsed {len(df)} lines. Harmonizing duplicate clusters...")

    # --- Symmetrical Duplicate Healing & Math Correction ---
    # Group by identical proposed master name to ensure mathematical perfection
    df["norm_master"] = df["proposed master name"].str.upper().str.strip()

    # Calculate real duplicate count based on master name grouping
    df["real_count"] = df.groupby("norm_master")["code"].transform("count") - 1

    # Collect all sibling codes for each cluster
    def get_cluster_codes(series):
        codes = sorted(list(set(str(c).strip() for c in series if str(c).strip())))
        return "; ".join(codes)

    cluster_map = df.groupby("norm_master")["code"].apply(get_cluster_codes).to_dict()

    df["number of duplicates"] = df["real_count"]
    df["duplicate_codes"] = df.apply(
        lambda r: (
            cluster_map[r["norm_master"]] if r["number of duplicates"] > 0 else ""
        ),
        axis=1,
    )

    df["extra comments"] = df.apply(
        lambda r: (
            f"Duplicate record clustered into '{r['proposed master name']}'"
            if r["number of duplicates"] > 0
            else "Unique SKU"
        ),
        axis=1,
    )

    final_df = df[
        [
            "code",
            "name",
            "number of duplicates",
            "duplicate_codes",
            "proposed master name",
            "extra comments",
        ]
    ]

    # Sort strictly alphabetically
    final_df = final_df.sort_values(by=["proposed master name", "code"]).reset_index(
        drop=True
    )

    # Save with UTF-8 BOM so Excel on Android displays all columns & Arabic perfectly
    final_df.to_csv(CLEAN_FILE, index=False, encoding="utf-8-sig")

    print("\n" + "=" * 60)
    print(f"[✔] REPAIR COMPLETE: Saved to '{CLEAN_FILE}'")
    print(f"Total Rows: {len(final_df)}")
    print(
        f"Total Duplicate Rows: {len(final_df[final_df['number of duplicates'] > 0])}"
    )
    print(
        f"Total Unique Rows:    {len(final_df[final_df['number of duplicates'] == 0])}"
    )
    print("=" * 60)


if __name__ == "__main__":
    repair_and_standardize_csv()
