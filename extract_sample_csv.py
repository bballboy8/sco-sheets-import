"""
Extract the first 25 rows from one CSV file in the ZIP and save to a local CSV file.
"""
import csv
import io
import os
import tempfile
import urllib.request
import ssl
from zipfile import ZipFile

def download_zip(url: str) -> str:
    """Download ZIP file to temporary location."""
    print(f"Downloading {url}...")
    ssl_context = ssl.create_default_context()
    ssl_context.check_hostname = False
    ssl_context.verify_mode = ssl.CERT_NONE
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://claimit.ca.gov/",
    }
    
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, context=ssl_context, timeout=600) as response:
        fd, path = tempfile.mkstemp(suffix=".zip")
        with os.fdopen(fd, "wb") as f:
            f.write(response.read())
        print(f"Downloaded to: {path}")
        return path

def extract_first_rows(zip_path: str, output_file: str, num_rows: int = 25):
    """Extract first N rows from first CSV in ZIP and save to output file."""
    print(f"\nExtracting first {num_rows} rows from ZIP...")
    
    with ZipFile(zip_path, mode="r", allowZip64=True) as zf:
        csv_files = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not csv_files:
            print("No CSV files found in ZIP!")
            return
        
        csv_name = csv_files[0]
        print(f"Processing: {csv_name}")
        
        with zf.open(csv_name, "r") as fbin:
            try:
                text = io.TextIOWrapper(fbin, encoding="utf-8-sig", newline="")
                reader = csv.reader(text)
            except UnicodeDecodeError:
                fbin.seek(0)
                text = io.TextIOWrapper(fbin, encoding="latin-1", newline="")
                reader = csv.reader(text)
            
            # Read header
            header = next(reader, None)
            if header is None:
                print("No header found!")
                return
            
            # Read first N rows
            rows = []
            for i, row in enumerate(reader, 1):
                if i > num_rows:
                    break
                rows.append(row)
            
            # Write to output file
            print(f"Writing {len(rows)} rows to {output_file}...")
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(header)
                writer.writerows(rows)
            
            print(f"✓ Saved {len(rows)} rows (plus header) to {output_file}")

if __name__ == "__main__":
    test_url = "https://claimit.ca.gov/upd-property-records/04_From_500_To_Beyond.zip"
    output_file = "raw_sample_data.csv"
    
    try:
        zip_path = download_zip(test_url)
        extract_first_rows(zip_path, output_file, num_rows=25)
        os.remove(zip_path)
        print(f"\n✓ Done! Check {output_file}")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
