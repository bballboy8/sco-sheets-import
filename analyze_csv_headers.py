"""Quick script to download and analyze CSV headers from SCO source files."""
import csv
import io
import tempfile
import os
from zipfile import ZipFile
import requests
import urllib3

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Use the first ZIP file for analysis
url = "https://claimit.ca.gov/upd-property-records/01_From_0_To_Below_10.zip"

print(f"Downloading {url}...")
try:
    session = requests.Session()
    session.verify = False
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://claimit.ca.gov/",
    }
    
    with session.get(url, headers=headers, stream=True, timeout=60) as r:
        r.raise_for_status()
        fd, path = tempfile.mkstemp(suffix=".zip")
        with os.fdopen(fd, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)
    
    print(f"Downloaded to {path}")
    print("\nAnalyzing CSV files in ZIP...")
    
    with ZipFile(path, mode="r", allowZip64=True) as zf:
        csv_files = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        print(f"Found {len(csv_files)} CSV file(s)\n")
        
        for csv_name in csv_files[:3]:  # Analyze first 3 CSV files
            print(f"\n{'='*80}")
            print(f"CSV File: {csv_name}")
            print(f"{'='*80}")
            
            with zf.open(csv_name, "r") as fbin:
                try:
                    text = io.TextIOWrapper(fbin, encoding="utf-8-sig", newline="")
                except UnicodeDecodeError:
                    fbin.seek(0)
                    text = io.TextIOWrapper(fbin, encoding="latin-1", newline="")
                
                reader = csv.reader(text)
                header = next(reader, None)
                
                if header:
                    print(f"\nTotal columns: {len(header)}\n")
                    print("Column names:")
                    for i, col in enumerate(header, 1):
                        print(f"  {i:2d}. {col}")
                    
                    # Look for phone/email related fields
                    print("\n" + "-"*80)
                    print("Fields containing 'PHONE', 'EMAIL', 'MOBILE', 'CELL', 'CONTACT':")
                    phone_email_fields = []
                    for col in header:
                        col_upper = col.upper()
                        if any(term in col_upper for term in ['PHONE', 'EMAIL', 'MOBILE', 'CELL', 'CONTACT', 'TEL']):
                            phone_email_fields.append(col)
                            print(f"  - {col}")
                    
                    if not phone_email_fields:
                        print("  (None found)")
                    
                    # Show first data row for reference
                    first_row = next(reader, None)
                    if first_row:
                        print("\n" + "-"*80)
                        print("Sample data row (first few columns):")
                        for i, (col_name, value) in enumerate(zip(header[:10], first_row[:10])):
                            print(f"  {col_name}: {value[:50] if value else '(empty)'}")
                else:
                    print("  (No header found)")
    
    # Cleanup
    try:
        os.remove(path)
    except:
        pass
        
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()




