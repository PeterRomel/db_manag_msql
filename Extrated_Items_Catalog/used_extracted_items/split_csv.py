import csv
import math
import os


def split_csv_into_nine(input_filename):
    # Ensure utf-8 encoding is used safely
    with open(input_filename, mode="r", encoding="utf-8", newline="") as infile:
        reader = csv.reader(infile)
        header = next(reader)  # Read and store the header row
        rows = list(reader)  # Load all data rows into memory

    total_rows = len(rows)
    num_files = 9
    rows_per_file = math.ceil(total_rows / num_files)

    base_name, ext = os.path.splitext(input_filename)

    for i in range(num_files):
        start_idx = i * rows_per_file
        end_idx = start_idx + rows_per_file
        chunk_rows = rows[start_idx:end_idx]

        # If a chunk is empty (e.g. if rows < 9), break early
        if not chunk_rows:
            break

        output_filename = f"{base_name}_part_{i+1}{ext}"

        with open(output_filename, mode="w", encoding="utf-8", newline="") as outfile:
            writer = csv.writer(outfile)
            writer.writerow(header)
            writer.writerows(chunk_rows)

        print(f"Created {output_filename} with {len(chunk_rows)} rows (plus header).")


if __name__ == "__main__":
    # Replace 'input.data.csv' with the path to your actual CSV file
    input_file = "used_items_extract.csv"
    split_csv_into_nine(input_file)
