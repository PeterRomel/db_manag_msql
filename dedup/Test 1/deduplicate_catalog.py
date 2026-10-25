import os
import re
import time
import pandas as pd
from google import genai
from google.genai import types

# ---------------------------------------------------------
# 1. SETUP & CLIENT INITIALIZATION
# ---------------------------------------------------------
API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("Please set your GEMINI_API_KEY environment variable.")

client = genai.Client(api_key=API_KEY)

# Use the modern Gemini 3.8 Flash model
MODEL_ID = "gemini-3.8-flash"

SYSTEM_INSTRUCTION = """
You are an expert Egyptian Pharmacy ERP catalog entity-resolution and deduplication engine.
Your task is to deduplicate Egyptian pharmacy catalog items adhering strictly to these clinical & ERP rules:

1. STRICT COMMERCIAL BRAND ISOLATION:
   - Trade brand names must NEVER be merged across different commercial products, even if they share active molecules and strengths (e.g., Paramol, Panadol, Cetal, Abimol must remain distinct master SKUs).

2. RULE 1-A (STRICT PACK SIZE MATCHING):
   - Items with differing pack sizes, blister counts, or volumes (e.g. 20 TAB vs 30 TAB, or 60 ML vs 100 ML) MUST REMAIN SEPARATE SKUs.
   - Merging is strictly reserved for items with identical: Brand + Active Strength + Dosage Form + Pack Size.

3. METADATA SANITIZATION:
   - Strip retail prices (e.g., '18LE', '22.5le', '357', '78le').
   - Strip ERP status flags and noise ('@', '@@', '#', '*', '!', '?').
   - Strip cashier lifecycle notes ('NEW PRICE', 'OLD PRICE', 'N.P', 'سعر جديد', 'موقوف', 'بدون خصم').
   - Correct dosage typos ('1000 mMG' -> '1000 MG', 'tables' -> 'TAB', 'cabsules' -> 'CAP').
   - Standardize Arabic names into clean English canonical master names where applicable.

4. MASTER RECORD SELECTION:
   - Choose the oldest/lowest numerical Item_Code as the primary master code.

5. OUTPUT FORMAT:
   - Return ONLY raw CSV lines (no markdown, no backticks, no conversational text).
   - Columns:
     code,name,number of duplicates,duplicate_codes,proposed master name,extra comments
   - 'duplicate_codes' must list all clustered codes separated by semicolons (; ) inside double quotes.
   - Duplicates must appear adjacent to each other.
"""


# ---------------------------------------------------------
# 2. INTELLIGENT ROOT-SAFE CHUNKING
# ---------------------------------------------------------
def get_clean_root(name: str) -> str:
    """Extracts the primary alphabetical root word to keep brand families together."""
    if not isinstance(name, str):
        return ""
    # Strip leading non-alphabetics, special characters, and digits
    clean = re.sub(r"^[^a-zA-Z\u0621-\u064A]+", "", name.strip())
    tokens = re.split(r"[\s\-_/]+", clean)
    return tokens[0].upper() if tokens else ""


def create_root_safe_batches(df: pd.DataFrame, target_batch_size: int = 100):
    """
    Sorts catalog and groups by Brand Root. Never splits a shared brand root across batch boundaries.
    """
    df["clean_root"] = df["Item_English_Name"].apply(get_clean_root)
    # Sort alphabetically by root, then by raw name
    df = df.sort_values(by=["clean_root", "Item_English_Name"]).reset_index(drop=True)

    batches = []
    current_batch = []

    # Group rows by their common brand root
    grouped = df.groupby("clean_root", sort=False)

    for root, group in grouped:
        group_rows = group.to_dict("records")

        # If adding this brand family exceeds target size and current_batch is not empty, flush current batch
        if (
            len(current_batch) + len(group_rows) > target_batch_size
            and len(current_batch) > 0
        ):
            batches.append(pd.DataFrame(current_batch))
            current_batch = []

        current_batch.extend(group_rows)

    if current_batch:
        batches.append(pd.DataFrame(current_batch))

    return batches


# ---------------------------------------------------------
# 3. PROCESSING & EXECUTION PIPELINE
# ---------------------------------------------------------
def run_deduplication(input_csv_path: str, output_csv_path: str):
    print(f"Reading catalog from: {input_csv_path}")
    raw_df = pd.read_csv(input_csv_path)

    # Ensure expected columns exist
    if "Item_Code" not in raw_df.columns or "Item_English_Name" not in raw_df.columns:
        raise ValueError(
            "Input CSV must contain 'Item_Code' and 'Item_English_Name' columns."
        )

    batches = create_root_safe_batches(raw_df, target_batch_size=100)
    total_batches = len(batches)
    print(f"Created {total_batches} root-safe batches for processing.")

    # Initialize output file with CSV header
    with open(output_csv_path, "w", encoding="utf-8-sig") as out_f:
        out_f.write(
            "code,name,number of duplicates,duplicate_codes,proposed master name,extra comments\n"
        )

    for idx, batch_df in enumerate(batches, start=1):
        batch_first_item = batch_df.iloc[0]["Item_English_Name"]
        batch_last_item = batch_df.iloc[-1]["Item_English_Name"]
        print(
            f"[{idx}/{total_batches}] Processing {len(batch_df)} items ({batch_first_item[:15]}... -> {batch_last_item[:15]}...)"
        )

        csv_payload = batch_df[["Item_Code", "Item_English_Name"]].to_csv(index=False)
        user_prompt = (
            f"Process and deduplicate this batch of catalog items:\n\n{csv_payload}"
        )

        success = False
        retries = 3
        while not success and retries > 0:
            try:
                response = client.models.generate_content(
                    model=MODEL_ID,
                    contents=user_prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        temperature=0.0,
                        # Low-to-moderate thinking budget for clinical edge cases
                        thinking_config=types.ThinkingConfig(thinking_budget=1024),
                    ),
                )

                # Sanitize response
                text_result = (
                    response.text.replace("```csv", "").replace("```", "").strip()
                )

                # Filter out header if model re-emitted it
                lines = [
                    line
                    for line in text_result.splitlines()
                    if not line.lower().startswith("code,name,number of duplicates")
                ]

                if lines:
                    with open(output_csv_path, "a", encoding="utf-8-sig") as out_f:
                        out_f.write("\n".join(lines) + "\n")

                success = True
                # Pause 4.5 seconds to comply with 15 RPM Free Tier quota
                time.sleep(4.5)

            except Exception as e:
                retries -= 1
                print(
                    f"  Warning: Batch {idx} failed with error: {e}. Retries left: {retries}"
                )
                time.sleep(8)

        if not success:
            print(
                f"  [ERROR]: Batch {idx} failed permanently after retries. Logged for manual inspection."
            )

    print(f"\nExecution finished successfully!")
    print(f"Complete master catalog saved to: {os.path.abspath(output_csv_path)}")


if __name__ == "__main__":
    INPUT_FILE = "used_items_extract.csv"
    OUTPUT_FILE = "deduplicated_pharmacy_master.csv"
    run_deduplication(INPUT_FILE, OUTPUT_FILE)
