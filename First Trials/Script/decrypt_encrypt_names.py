import pandas as pd


def decrypt_hex(hex_str: str) -> str:
    """
    Decrypts database hex representation back to plain text string.
    """
    if pd.isna(hex_str) or not isinstance(hex_str, str):
        return ""

    # Clean 0x prefix and any whitespace
    hex_str = hex_str.strip()
    if hex_str.startswith("0x") or hex_str.startswith("0X"):
        hex_str = hex_str[2:]

    try:
        # Convert hex string to byte array
        raw_bytes = bytes.fromhex(hex_str)
        # Reverse the bytes and decode with Windows-1256
        return raw_bytes[::-1].decode("cp1256", errors="replace")
    except Exception as e:
        return f"[Decryption Error: {e}]"


def encrypt_text(text_str: str) -> str:
    """
    Encrypts plain text string to database hex representation.
    """
    if not isinstance(text_str, str):
        return "0x"

    # Encode string to Windows-1256 bytes
    raw_bytes = text_str.encode("cp1256", errors="replace")
    # Reverse byte sequence
    reversed_bytes = raw_bytes[::-1]
    # Return formatted hex string with 0x prefix
    return "0x" + reversed_bytes.hex().upper()


# ==========================================
# Example Usage & File Batch Processing
# ==========================================
if __name__ == "__main__":
    # Test cases from database PDF comparison
    sample_decrypted = "AMIKACIN LOTION SPRAY 100 ML NEW PRICE"
    sample_hex = "0x4F4C46204545524620544E455641C9EDE4C7CCE320C9E4EDCAEDCA2057"

    print("--- Test Decryption ---")
    print("Decrypted Sample:", decrypt_hex(sample_hex))

    print("\n--- Test Encryption ---")
    print("Encrypted Sample:", encrypt_text(sample_decrypted))

    # Batch process CSV file
    try:
        csv_path = "raw_encrypted_names.csv"
        df = pd.read_csv(csv_path)

        # Decrypt English and Arabic encrypted columns
        df["itm_name_en_decrypted"] = df["itm_name_en_encrypt"].apply(decrypt_hex)
        df["itm_name_ar_decrypted"] = df["itm_name_ar_encrypt"].apply(decrypt_hex)

        # Save output to clean CSV
        output_path = "decrypted_items.csv"
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
        print(f"\nSuccessfully decrypted {len(df)} records to '{output_path}'.")
    except Exception as err:
        print(f"\nCould not process CSV file: {err}")
