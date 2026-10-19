import os
from dotenv import load_dotenv

load_dotenv()

import os
import re
import time
import json
import pandas as pd
from google import genai
from google.genai import types
from google.genai.errors import ClientError

# ==============================================================================
# 1. EXACT GEMINI CONFIGURATION & ENGINE SETTINGS (From your file)
# ==============================================================================
API_KEY = os.getenv(
    "GEMINI_API_KEY", os.getenv("GCP_API_KEY")
)
MODEL_ID = "gemini-3.5-flash-lite"

BATCH_SIZE = 100
REQUEST_DELAY_SECONDS = 10  # Enforces safe throughput under 15 RPM
MAX_RETRIES = 5

INPUT_FILE = "used_items_extract.csv"
OUTPUT_FILE = "final_deduplicated_catalog.csv"
CHECKPOINT_FILE = "checkpoint_processed_batches.json"

# ==============================================================================
# 2. PHARMACEUTICAL MASTER DOSSIER INSTRUCTIONS
# ==============================================================================
SYSTEM_INSTRUCTION = """You are an expert Egyptian Pharmacy ERP catalog entity-resolution and deduplication engine.
Your task is to deduplicate Egyptian pharmacy catalog items adhering strictly to these clinical & ERP rules:

1. STRICT COMMERCIAL BRAND ISOLATION:
   - Trade brand names must NEVER be merged across different commercial products, even if they share active molecules and strengths (e.g., Paramol, Panadol, Cetal, Abimol must remain distinct master SKUs).

2. RULE 1-A (STRICT PACK SIZE MATCHING):
   - Items with differing pack sizes, blister counts, or volumes (e.g., 20 TAB vs 30 TAB, or 60 ML vs 100 ML) MUST REMAIN SEPARATE SKUs.
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
   - Return ONLY raw CSV lines (no markdown, no backticks, no conversational preamble).
   - Columns:
     code,name,number of duplicates,duplicate_codes,proposed master name,extra comments
   - 'duplicate_codes' must list all clustered codes separated by semicolons (; ) inside double quotes.
   - Sibling duplicate rows must appear adjacent to each other.
"""


# ==============================================================================
# 3. ROOT-SAFE SEMANTIC GROUPING (Prevents Boundary Splitting)
# ==============================================================================
def get_clean_root(name: str) -> str:
    """Extracts the primary alphabetical root word to keep brand families together."""
    if not isinstance(name, str):
        return ""
    clean = re.sub(r"^[^a-zA-Z\u0621-\u064A]+", "", name.strip())
    tokens = re.split(r"[\s\-_/]+", clean)
    return tokens[0].upper() if tokens else ""


def create_root_safe_batches(
    df: pd.DataFrame, target_batch_size: int = 100
) -> list[pd.DataFrame]:
    """
    Sorts catalog and groups by Brand Root.
    Guarantees no brand family (e.g., all CEFOTAX or PANADOL records) is cut across batch borders.
    """
    df["clean_root"] = df["Item_English_Name"].apply(get_clean_root)
    df = df.sort_values(by=["clean_root", "Item_English_Name"]).reset_index(drop=True)

    batches = []
    current_batch = []
    grouped = df.groupby("clean_root", sort=False)

    for _, group in grouped:
        group_rows = group.to_dict("records")
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


# ==============================================================================
# 4. GEMINI SERVER INFERENCE WITH EXACT RETRY & ERROR ENGINE
# ==============================================================================
def execute_gemini_batch_with_retry(client: genai.Client, csv_payload: str) -> str:
    user_prompt = (
        f"Process and deduplicate this batch of catalog items:\n\n{csv_payload}"
    )

    for attempt in range(MAX_RETRIES):
        try:
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_INSTRUCTION,
                    temperature=0.0,
                ),
            )
            return response.text

        except ClientError as e:
            error_str = str(e)
            if "limit: 0" in error_str or "NOT_FOUND" in error_str:
                raise RuntimeError(f"Permanent configuration error: {error_str}")

            # Parse dynamic wait time provided by Google API
            retry_match = re.search(r"retryDelay':\s*'(\d+)s", error_str)
            if not retry_match:
                retry_match = re.search(r"retry in (\d+(\.\d+)?)s", error_str)

            wait_time = int(float(retry_match.group(1))) + 2 if retry_match else 20
            print(
                f"\n[!] Rate limit reached. Pausing {wait_time}s as requested by API "
                f"(Attempt {attempt + 1}/{MAX_RETRIES})..."
            )
            time.sleep(wait_time)

        except Exception as e:
            print(
                f"\n[!] Unexpected error: {e}. Pausing 10s (Attempt {attempt + 1}/{MAX_RETRIES})..."
            )
            time.sleep(10)

    raise RuntimeError(f"Batch failed permanently after {MAX_RETRIES} attempts.")


# ==============================================================================
# 5. PIPELINE EXECUTION WITH RESUME CHECKPOINTS
# ==============================================================================
def run_pipeline():
    print(f"[*] Initializing Gemini Client with model: '{MODEL_ID}'...")
    client = genai.Client(api_key=API_KEY)

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(
            f"Cannot find input file '{INPUT_FILE}'. Place it in the script directory."
        )

    raw_df = pd.read_csv(INPUT_FILE)
    if "Item_Code" not in raw_df.columns or "Item_English_Name" not in raw_df.columns:
        raise ValueError(
            "Input CSV must contain 'Item_Code' and 'Item_English_Name' columns."
        )

    print(f"[*] Loaded {len(raw_df)} total records from '{INPUT_FILE}'.")
    batches = create_root_safe_batches(raw_df, target_batch_size=BATCH_SIZE)
    total_batches = len(batches)
    print(
        f"[*] Clustered catalog into {total_batches} root-safe batches (zero boundary splits)."
    )

    # Load checkpoint if resuming an interrupted session
    completed_batch_indices = set()
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                completed_batch_indices = set(json.load(f))
            print(
                f"[*] Resuming run: {len(completed_batch_indices)}/{total_batches} batches already processed."
            )
        except Exception:
            completed_batch_indices = set()

    # Initialize output file header if starting fresh
    if not os.path.exists(OUTPUT_FILE) or len(completed_batch_indices) == 0:
        with open(OUTPUT_FILE, "w", encoding="utf-8-sig") as out_f:
            out_f.write(
                "code,name,number of duplicates,duplicate_codes,proposed master name,extra comments\n"
            )

    # Process batches
    for idx, batch_df in enumerate(batches, start=1):
        if idx in completed_batch_indices:
            continue

        first_item = batch_df.iloc[0]["Item_English_Name"]
        last_item = batch_df.iloc[-1]["Item_English_Name"]
        print(
            f"[{idx}/{total_batches}] Processing {len(batch_df)} items ({first_item[:16]}... -> {last_item[:16]}...)"
        )

        csv_payload = batch_df[["Item_Code", "Item_English_Name"]].to_csv(index=False)
        raw_response = execute_gemini_batch_with_retry(client, csv_payload)

        # Clean markdown code blocks if emitted
        clean_text = raw_response.replace("```csv", "").replace("```", "").strip()

        # Filter out accidental repeated header rows
        valid_lines = [
            line
            for line in clean_text.splitlines()
            if line.strip()
            and not line.lower().startswith("code,name,number of duplicates")
        ]

        # Write directly to disk
        if valid_lines:
            with open(OUTPUT_FILE, "a", encoding="utf-8-sig") as out_f:
                out_f.write("\n".join(valid_lines) + "\n")

        # Update checkpoint
        completed_batch_indices.add(idx)
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            json.dump(list(completed_batch_indices), f)

        time.sleep(REQUEST_DELAY_SECONDS)

    print("\n" + "=" * 70)
    print("DEDUPLICATION RUN FINISHED SUCCESSFULLY")
    print("=" * 70)
    print(f"[+] Output File: '{os.path.abspath(OUTPUT_FILE)}'")
    print(f"[+] Checkpoint:  '{os.path.abspath(CHECKPOINT_FILE)}'")


if __name__ == "__main__":
    run_pipeline()
