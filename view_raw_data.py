"""
Download and display raw CSV data from the ZIP file before any processing.
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

def view_raw_csv_data(zip_path: str, max_rows: int = 20):
    """Extract and display raw CSV data."""
    print(f"\nExtracting and viewing raw CSV data from: {zip_path}\n")
    print("="*80)
    
    with ZipFile(zip_path, mode="r", allowZip64=True) as zf:
        csv_files = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        print(f"Found {len(csv_files)} CSV file(s): {', '.join(csv_files)}\n")
        
        for csv_name in csv_files[:1]:  # Just show first CSV file
            print(f"\n{'='*80}")
            print(f"FILE: {csv_name}")
            print(f"{'='*80}\n")
            
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
                    continue
                
                print(f"HEADER ({len(header)} columns):")
                print("-" * 80)
                for i, col in enumerate(header, 1):
                    print(f"  {i:2d}. {col}")
                
                print(f"\n\nRAW DATA (first {max_rows} rows):")
                print("-" * 80)
                
                # Show first few rows
                for row_num, row in enumerate(reader, 1):
                    if row_num > max_rows:
                        break
                    
                    print(f"\nRow {row_num}:")
                    print("-" * 80)
                    for i, (col_name, value) in enumerate(zip(header, row)):
                        if value and str(value).strip():  # Only show non-empty values
                            print(f"  {col_name}: {value}")
                    
                    # Also show as CSV row for reference
                    print(f"\n  CSV format: {','.join(row[:5])}...")  # First 5 columns
                
                # Count total rows
                remaining_rows = sum(1 for _ in reader)
                total_rows = row_num + remaining_rows
                print(f"\n\nTotal rows in file: {total_rows:,}")
                print(f"Shown: {min(row_num, max_rows)} rows")

if __name__ == "__main__":
    # Use the smaller test file
    test_url = "https://claimit.ca.gov/upd-property-records/04_From_500_To_Beyond.zip"
    
    try:
        zip_path = download_zip(test_url)
        view_raw_csv_data(zip_path, max_rows=10)
        
        print(f"\n\n{'='*80}")
        print("Raw ZIP file location (will be cleaned up):")
        print(zip_path)
        print(f"{'='*80}")
        
        # Ask if user wants to keep the file
        keep = input("\nKeep the downloaded ZIP file? (y/n): ").strip().lower()
        if keep != 'y':
            os.remove(zip_path)
            print("ZIP file removed.")
        else:
            print(f"ZIP file kept at: {zip_path}")
            
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()



