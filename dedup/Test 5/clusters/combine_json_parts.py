"""COMBINE AI STUDIO JSON PARTS INTO A SINGLE FILE

Usage:
  1. Save your AI Studio outputs as part1.json, part2.json, part3.json, etc.
     (or part1.txt, part2.txt) in the same folder.
  2. Run: python combine_json_parts.py
"""

import glob
import json
import os
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "ai_studio_clusters.json")


def clean_and_parse_json_text(raw_text: str):
    """Strips markdown code fences and parses JSON safely."""
    # Remove markdown code fences like ```json ... ```
    cleaned = re.sub(r"```(json)?", "", raw_text).strip("` \n\r\t")

    # If the text has extra text around the JSON, find the first '[' and last ']'
    start_idx = cleaned.find("[")
    end_idx = cleaned.rfind("]")

    if start_idx != -1 and end_idx != -1:
        cleaned = cleaned[start_idx : end_idx + 1]

    return json.loads(cleaned)


def merge_parts():
    # Finds all files named part*.json or part*.txt
    part_files = sorted(glob.glob(os.path.join(SCRIPT_DIR, "part*.*")))

    if not part_files:
        print(
            "[!] No files matching 'part*.json' or 'part*.txt' found in the current folder."
        )
        print("    Please save your turns as part1.json, part2.json, etc.")
        return

    print(f"[*] Found {len(part_files)} part files to merge: {', '.join(part_files)}")

    combined_clusters = []
    seen_clusters = set()

    for filename in part_files:
        print(f"[*] Reading '{filename}'...")
        try:
            with open(filename, "r", encoding="utf-8") as f:
                content = f.read()

            parsed = clean_and_parse_json_text(content)

            # Handle if the AI wrapped it inside {"clusters": [...]}
            if isinstance(parsed, dict) and "clusters" in parsed:
                parsed = parsed["clusters"]

            if isinstance(parsed, list):
                for item in parsed:
                    # If item is a cluster of codes (e.g. ["84662", "85169"])
                    if isinstance(item, list):
                        codes = sorted(
                            list(set(str(c).strip() for c in item if str(c).strip()))
                        )
                        if len(codes) > 1:
                            cluster_key = tuple(codes)
                            if cluster_key not in seen_clusters:
                                seen_clusters.add(cluster_key)
                                combined_clusters.append(codes)
                    # If item was a dict (e.g. {"codes": [...]})
                    elif isinstance(item, dict) and "codes" in item:
                        codes = sorted(
                            list(
                                set(
                                    str(c).strip()
                                    for c in item["codes"]
                                    if str(c).strip()
                                )
                            )
                        )
                        if len(codes) > 1:
                            cluster_key = tuple(codes)
                            if cluster_key not in seen_clusters:
                                seen_clusters.add(cluster_key)
                                combined_clusters.append(codes)
        except Exception as e:
            print(f"[ERROR] Failed to parse '{filename}': {e}")

    # Write the unified JSON file
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(combined_clusters, f, indent=2, ensure_ascii=False)

    print("\n" + "=" * 50)
    print(f"[✔] Successfully created: {OUTPUT_FILE}")
    print(f"[✔] Total Unique Duplicate Clusters: {len(combined_clusters)}")
    print("=" * 50)


if __name__ == "__main__":
    merge_parts()
