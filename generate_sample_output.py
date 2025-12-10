"""
Generate a sample output CSV file showing what data would be included from the downloaded ZIP files.
This script processes a sample ZIP file and outputs a CSV with the same format that would be written to Google Sheets.
"""
import csv
import io
import os
import re
import tempfile
import urllib.request
import ssl
from typing import List
from zipfile import ZipFile
from datetime import datetime

# Copy business terms from sco_import.py
BUSINESS_TERMS = [
    "LLC","INC","INC.","CORP","CORPORATION","LLP","LIMITED","LTD","COMPANY","CO ",
    "APC","PC","PLC","LP","L.P.","DBA","FIRM","ASSOCIATES","HOSPITAL","ASSOCIATION",
    "FOUNDATION","CHURCH","UNIVERSITY","UNIV","COLLEGE","SCHOOL","CHARTER","MINISTRY",
    "UNION","CLUB","CENTER","CENTRE","TEAM","WORLD","INTERNATIONAL","SERVICES",
    "SOLUTIONS","MANAGEMENT","CONSULTING","HOLDINGS","HOLDING","PROPERTIES","PROPERTY",
    "VENTURES","SYSTEMS","TECHNOLOGIES","TECH","REAL ESTATE","LABS","LABORATORIES",
    "METAL","SUPPLY","SUPPLIES","MANUFACTUR","MANUFACTURI","LOGISTICS","TRANSPORT",
    "TRANSPORTAT","NETWORK","NETWORKS","STUDIOS","PRODUCTIONS","ENTERTAINME","BANK",
    "CREDIT UNION","INSURANCE","MORTGAGE","FUND","INVESTMENTS","INVESTMENT","CAPITAL",
    "ADVISORS","SECURITIES","MEDICAL GRO","HEALTHCARE","HEALTH CARE","AUTO","BODY",
    "APPAREL","PEDIATRICS","RECOVERY","FOOD","FOODS","SALES","CONSTRUCTIO","CARPET",
    "TILE","GLASS","PERFORMANC","CAR","CARS","BOAT","BOATS","ESCROW","CATERING",
    "TRUCKING","TRUCK","TRUCKS","MARKET","PACKING","PACKAGING","OF ","THE ","A ",
    "COMMUNICAT","COUNTY","CITY","STATE","TREASURER","DEPARTMENT","ELEMENATRY",
    "DISTRICT","GROUP","SVCS","&","AND ","VOLUNTEER"
]

def _matches_with_boundaries(text: str, term: str) -> bool:
    """Check if a term appears as a whole word in the text."""
    if not term or not text:
        return False
    escaped_term = re.escape(term)
    if term.endswith(' '):
        pattern = r'(?<![A-Za-z0-9])' + escaped_term
    else:
        pattern = r'(?<![A-Za-z0-9])' + escaped_term + r'(?![A-Za-z0-9])'
    return bool(re.search(pattern, text))

def is_business(owner_name: str) -> bool:
    if not owner_name or not owner_name.strip():
        return False
    name = owner_name.upper()
    for term in BUSINESS_TERMS:
        if _matches_with_boundaries(name, term):
            return True
    return False

def should_include_record(row: List[str], header: List[str]) -> bool:
    try:
        cash_idx  = header.index("CURRENT_CASH_BALANCE") if "CURRENT_CASH_BALANCE" in header else -1
        owner_cnt = header.index("NO_OF_OWNERS")         if "NO_OF_OWNERS"          in header else -1
        name_idx  = header.index("OWNER_NAME")           if "OWNER_NAME"            in header else -1
        addr_idx  = header.index("OWNER_STREET_1")       if "OWNER_STREET_1"        in header else -1

        cash   = float(row[cash_idx]) if cash_idx >= 0 and row[cash_idx] else 0.0
        owners = int(row[owner_cnt])  if owner_cnt >= 0 and row[owner_cnt] else 1
        name   = (row[name_idx]  if name_idx  >= 0 else "") or ""
        addr   = (row[addr_idx]  if addr_idx  >= 0 else "") or ""

        if cash <= 6000: return False
        if owners >= 3 and cash < 17999: return False
        if owners == 2 and cash < 11999: return False
        if not name or name.strip().upper() in {"BLANK","UNKNOWN",""}: return False
        if not addr or addr.strip().upper() in {"BLANK","UNKNOWN",""}: return False
        return True
    except (ValueError, IndexError):
        return False

def add_metadata_columns(row: List[str], header: List[str]) -> List[str]:
    current_date = datetime.now().strftime("%Y-%m-%d")
    name_idx = header.index("OWNER_NAME") if "OWNER_NAME" in header else -1
    owner = row[name_idx] if name_idx >= 0 else ""
    record_type = "Business" if is_business(owner) else "Individual"
    
    # Look for phone and email fields in the source CSV
    mobile_phone = ""
    other_phone = ""
    email = ""
    
    # Try to find mobile phone field
    mobile_field_names = ["MOBILE_PHONE", "MOBILE", "OWNER_MOBILE_PHONE", "OWNER_MOBILE", "CELL_PHONE", "CELL"]
    for field_name in mobile_field_names:
        if field_name in header:
            idx = header.index(field_name)
            if idx < len(row) and row[idx]:
                mobile_phone = str(row[idx]).strip()
                break
    
    # Try to find other phone field
    other_phone_field_names = ["OTHER_PHONE", "PHONE", "OWNER_PHONE", "HOME_PHONE", "WORK_PHONE", "TELEPHONE"]
    for field_name in other_phone_field_names:
        if field_name in header:
            idx = header.index(field_name)
            if idx < len(row) and row[idx]:
                phone_value = str(row[idx]).strip()
                if field_name == "OTHER_PHONE":
                    other_phone = phone_value
                    break
                elif not mobile_phone and not other_phone:
                    if field_name not in mobile_field_names:
                        other_phone = phone_value
                        break
    
    # Try to find email field
    email_field_names = ["EMAIL", "OWNER_EMAIL", "EMAIL_ADDRESS", "E_MAIL", "E-MAIL"]
    for field_name in email_field_names:
        if field_name in header:
            idx = header.index(field_name)
            if idx < len(row) and row[idx]:
                email = str(row[idx]).strip()
                break
    
    # Build metadata columns
    metadata_cols = [current_date, record_type, "100%", "Leads", mobile_phone, other_phone, email]
    
    return row + metadata_cols

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
        return path

def process_zip_and_generate_output(zip_path: str, output_csv: str, max_rows: int = 100):
    """Process ZIP file and generate sample output CSV."""
    print(f"Processing ZIP file: {zip_path}")
    
    included_rows = []
    filtered_count = 0
    total_count = 0
    
    with ZipFile(zip_path, mode="r", allowZip64=True) as zf:
        csv_files = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        print(f"Found {len(csv_files)} CSV file(s)")
        
        for csv_name in csv_files:
            print(f"\nProcessing {csv_name}...")
            with zf.open(csv_name, "r") as fbin:
                try:
                    text = io.TextIOWrapper(fbin, encoding="utf-8-sig", newline="")
                    reader = csv.reader(text)
                except UnicodeDecodeError:
                    fbin.seek(0)
                    text = io.TextIOWrapper(fbin, encoding="latin-1", newline="")
                    reader = csv.reader(text)
                
                header = next(reader, None)
                if header is None:
                    continue
                
                print(f"  Header: {len(header)} columns")
                
                # Add metadata column headers
                metadata_headers = ["CREATED_BY_DATE", "TYPE", "CONFIDENCE_LEVEL", "STAGE", "MOBILE PHONE", "OTHER PHONE", "EMAIL"]
                full_header = header + metadata_headers
                
                for row in reader:
                    total_count += 1
                    if not row:
                        continue
                    
                    if not should_include_record(row, header):
                        filtered_count += 1
                        continue
                    
                    extended_row = add_metadata_columns(row, header)
                    included_rows.append(extended_row)
                    
                    if len(included_rows) >= max_rows:
                        print(f"  Reached max_rows limit ({max_rows}), stopping...")
                        break
                
                if len(included_rows) >= max_rows:
                    break
    
    # Write output CSV
    print(f"\nWriting {len(included_rows)} rows to {output_csv}...")
    with open(output_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(full_header)
        writer.writerows(included_rows)
    
    print(f"\nSummary:")
    print(f"  Total records processed: {total_count:,}")
    print(f"  Records filtered out: {filtered_count:,}")
    print(f"  Records included: {len(included_rows):,}")
    print(f"  Output file: {output_csv}")

if __name__ == "__main__":
    # Use the smaller test file
    test_url = "https://claimit.ca.gov/upd-property-records/04_From_500_To_Beyond.zip"
    output_file = "sample_output.csv"
    
    try:
        zip_path = download_zip(test_url)
        process_zip_and_generate_output(zip_path, output_file, max_rows=100)
        os.remove(zip_path)
        print(f"\n✓ Sample output generated: {output_file}")
        print(f"  You can open this file in Excel or any CSV viewer to see the data structure.")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()

