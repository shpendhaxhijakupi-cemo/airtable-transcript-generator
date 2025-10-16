import os
import sys

def main():
    record_id = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
    if not record_id:
        print("Missing RECORD_ID")
        sys.exit(1)

    # For now just echo; later we’ll pull Airtable + build PDF
    print(f"[OK] Received RECORD_ID: {record_id}")

if __name__ == "__main__":
    main()
