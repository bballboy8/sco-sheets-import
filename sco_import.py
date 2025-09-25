import csv, io, os, json, time, tempfile
from typing import List, Optional
from datetime import datetime
import re
import random

import requests
import gspread
from google.oauth2.service_account import Credentials
from zipfile import ZipFile
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# =================== CONFIG ===================
SHEET_ID = os.environ["SHEET_ID"]
GOOGLE_CREDENTIALS = os.environ["GOOGLE_CREDENTIALS"]

# QUICK TEST: start with the smallest/first ZIP to verify end-to-end.
# After it works, set USE_ALL_ZIPS = True
USE_ALL_ZIPS = True

ALL_ZIPS = [
    "https://dpupd.sco.ca.gov/01_From_0_To_Below_10.zip",
    "https://dpupd.sco.ca.gov/02_From_10_To_Below_100.zip",
    "https://dpupd.sco.ca.gov/03_From_100_To_Below_500.zip",
    "https://dpupd.sco.ca.gov/04_From_500_To_Beyond.zip",
]
SCO_ZIPS = ALL_ZIPS if USE_ALL_ZIPS else ALL_ZIPS[:1]

APPEND_BATCH_SIZE = 1000
API_PAUSE_SEC = 1.5
MAX_RECORD_ROWS = 500_000

RECORDS_TAB = "Records"

# Business classification terms
BUSINESS_TERMS = [
    "LLC", "INC", "INC.", "CORP", "CORPORATION", "LLP", "LIMITED", "LTD", "COMPANY", "CO ", 
    "APC", "PC", "PLC", "LP", "L.P.", "DBA", "FIRM", "ASSOCIATES", "HOSPITAL", "ASSOCIATION", 
    "FOUNDATION", "CHURCH", "UNIVERSITY", "UNIV", "COLLEGE", "SCHOOL", "CHARTER", "MINISTRY", 
    "UNION", "CLUB", "CENTER", "CENTRE", "TEAM", "WORLD", "INTERNATIONAL", "SERVICES", 
    "SOLUTIONS", "MANAGEMENT", "CONSULTING", "HOLDINGS", "HOLDING", "PROPERTIES", "PROPERTY", 
    "VENTURES", "SYSTEMS", "TECHNOLOGIES", "TECH", "REAL ESTATE", "LABS", "LABORATORIES", 
    "METAL", "SUPPLY", "SUPPLIES", "MANUFACTUR", "MANUFACTURI", "LOGISTICS", "TRANSPORT", 
    "TRANSPORTAT", "NETWORK", "NETWORKS", "STUDIOS", "PRODUCTIONS", "ENTERTAINME", "BANK", 
    "CREDIT UNION", "INSURANCE", "MORTGAGE", "FUND", "INVESTMENTS", "INVESTMENT", "CAPITAL", 
    "ADVISORS", "SECURITIES", "MEDICAL GRO", "HEALTHCARE", "HEALTH CARE", "AUTO", "BODY", 
    "APPAREL", "PEDIATRICS", "RECOVERY", "FOOD", "FOODS", "SALES", "CONSTRUCTIO", "CARPET", 
    "TILE", "GLASS", "PERFORMANC", "CAR", "CARS", "BOAT", "BOATS", "ESCROW", "CATERING", 
    "TRUCKING", "TRUCK", "TRUCKS", "MARKET", "PACKING", "PACKAGING", 
    "COMMUNICAT", "COUNTY", "CITY", "STATE", "TREASURER", "DEPARTMENT", "ELEMENATRY", 
    "DISTRICT", "GROUP", "SVCS", "VOLUNTEER"
]

# =================== RETRY LOGIC ===================
def retry_with_backoff(func, max_retries=3, base_delay=1, max_delay=60):
    """Retry a function with exponential backoff."""
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"[error] Final attempt failed: {e}", flush=True)
                raise
            
            # Calculate delay with exponential backoff and jitter
            delay = min(base_delay * (2 ** attempt) + random.uniform(0, 1), max_delay)
            print(f"[retry] Attempt {attempt + 1} failed: {e}", flush=True)
            print(f"[retry] Waiting {delay:.1f}s before retry...", flush=True)
            time.sleep(delay)

def safe_append_rows(ws, rows, value_input_option="RAW"):
    """Safely append rows with retry logic."""
    def _append():
        return ws.append_rows(rows, value_input_option=value_input_option)
    return retry_with_backoff(_append)

def safe_update(ws, values, range_name):
    """Safely update cells with retry logic."""
    def _update():
        return ws.update(values=values, range_name=range_name)
    return retry_with_backoff(_update)

# =================== GOOGLE AUTH ===================
def gs_client():
    info = json.loads(GOOGLE_CREDENTIALS)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

# =================== SHEET HELPERS ===================
def get_or_create_records_ws(sh) -> gspread.Worksheet:
    try:
        return sh.worksheet(RECORDS_TAB)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=RECORDS_TAB, rows=1, cols=26)

def find_key_col(header: List[str]) -> int:
    for i, h in enumerate(header):
        if str(h).strip().lower() == "property id" or str(h).strip().upper() == "PROPERTY_ID":
            return i
    for i, h in enumerate(header):
        hs = str(h).lower()
        if "property" in hs and "id" in hs:
            return i
    return 0

def load_existing_keys_from_records(ws: gspread.Worksheet, key_col_idx: int) -> set:
    col = key_col_idx + 1
    vals = ws.col_values(col)[1:]  # Skip header
    # Only include non-empty values
    return set((v or "").strip() for v in vals if v and str(v).strip())

def is_business(owner_name: str) -> bool:
    """Determine if owner is a business based on business terms in the name."""
    if not owner_name or not owner_name.strip():
        return False
    
    owner_upper = owner_name.upper()
    for term in BUSINESS_TERMS:
        if term in owner_upper:
            return True
    return False

def should_include_record(row: List[str], header: List[str]) -> bool:
    """Apply filtering rules to determine if record should be included."""
    try:
        # Get column indices
        cash_balance_idx = header.index("CURRENT_CASH_BALANCE") if "CURRENT_CASH_BALANCE" in header else -1
        owners_idx = header.index("NO_OF_OWNERS") if "NO_OF_OWNERS" in header else -1
        owner_name_idx = header.index("OWNER_NAME") if "OWNER_NAME" in header else -1
        owner_street_idx = header.index("OWNER_STREET_1") if "OWNER_STREET_1" in header else -1
        
        # Get values
        cash_balance = float(row[cash_balance_idx]) if cash_balance_idx >= 0 and row[cash_balance_idx] else 0.0
        num_owners = int(row[owners_idx]) if owners_idx >= 0 and row[owners_idx] else 1
        owner_name = row[owner_name_idx] if owner_name_idx >= 0 else ""
        owner_street = row[owner_street_idx] if owner_street_idx >= 0 else ""
        
        # Filter 1: Dollar value must be over $6,000
        if cash_balance <= 6000:
            return False
        
        # Filter 2: Remove claims with 3+ owners and value under $17,999
        if num_owners >= 3 and cash_balance < 17999:
            return False
        
        # Filter 3: Remove claims with 2 owners and value under $11,999
        if num_owners == 2 and cash_balance < 11999:
            return False
        
        # Filter 4: Remove records with missing/blank owner information
        if not owner_name or owner_name.strip().upper() in ["BLANK", "UNKNOWN", ""]:
            return False
        
        # Filter 5: Remove records with missing/blank address information
        if not owner_street or owner_street.strip().upper() in ["BLANK", "UNKNOWN", ""]:
            return False
        
        return True
        
    except (ValueError, IndexError):
        return False

def add_metadata_columns(row: List[str], header: List[str]) -> List[str]:
    """Add the required metadata columns to a row."""
    # Get current date
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    # Determine if business or individual
    owner_name_idx = header.index("OWNER_NAME") if "OWNER_NAME" in header else -1
    owner_name = row[owner_name_idx] if owner_name_idx >= 0 else ""
    record_type = "Business" if is_business(owner_name) else "Individual"
    
    # Add new columns: Created By Date, Type, Confidence Level, Stage
    new_columns = [current_date, record_type, "100%", "Leads"]
    
    return row + new_columns

def append_in_batches(ws: gspread.Worksheet, rows: List[List[str]], width: int):
    if not rows:
        return
    out = []
    for r in rows:
        rr = list(r)
        if len(rr) < width:
            rr.extend([""] * (width - len(rr)))
        elif len(rr) > width:
            rr = rr[:width]
        out.append(rr)
        if len(out) >= APPEND_BATCH_SIZE:
            safe_append_rows(ws, out, value_input_option="RAW")
            print(f"[append] {len(out)} rows", flush=True)
            out.clear()
            time.sleep(API_PAUSE_SEC)
    if out:
        safe_append_rows(ws, out, value_input_option="RAW")
        print(f"[append] {len(out)} rows", flush=True)

# =================== SCO DOWNLOAD / CSV STREAM ===================
def download_to_temp(url: str) -> str:
    def _download():
        # Try HEAD for size if server provides it
        size_hint = None
        try:
            hr = requests.head(url, timeout=30, allow_redirects=True)
            if hr.ok:
                size_hint = int(hr.headers.get("Content-Length") or 0) or None
        except Exception:
            pass

        headers = {
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://sco.ca.gov/upd_download_property_records.html",
            "Accept": "*/*",
        }

        print(f"[download] starting → {url}", flush=True)
        with requests.get(url, headers=headers, stream=True, timeout=600) as r:
            r.raise_for_status()
            fd, path = tempfile.mkstemp(suffix=".zip")
            total = 0
            next_marker = 50 * 1024 * 1024  # 50 MB
            with os.fdopen(fd, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 1024):
                    if not chunk:
                        continue
                    f.write(chunk)
                    total += len(chunk)
                    if total >= next_marker:
                        mb = total / (1024 * 1024)
                        if size_hint:
                            pct = 100.0 * total / size_hint
                            print(f"[download] {mb:.0f} MB ({pct:.1f}%)", flush=True)
                        else:
                            print(f"[download] {mb:.0f} MB", flush=True)
                        next_marker += 50 * 1024 * 1024

        size_mb = os.path.getsize(path) / (1024 * 1024)
        if size_hint:
            hint_mb = size_hint / (1024 * 1024)
            print(f"[downloaded] {size_mb:.1f} MB (server hint {hint_mb:.1f} MB)", flush=True)
        else:
            print(f"[downloaded] {size_mb:.1f} MB", flush=True)
        return path
    
    return retry_with_backoff(_download, max_retries=3, base_delay=2, max_delay=30)

def csv_reader_from_zip(zip_path: str):
    with ZipFile(zip_path, mode="r", allowZip64=True) as zf:
        names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        print(f"[zip] {len(names)} CSV file(s): {', '.join(names[:3])}{' …' if len(names) > 3 else ''}", flush=True)
        for name in names:
            with zf.open(name, "r") as fbin:
                try:
                    text = io.TextIOWrapper(fbin, encoding="utf-8-sig", newline="")
                    reader = csv.reader(text)
                    yield name, reader
                except UnicodeDecodeError:
                    fbin.seek(0)
                    text = io.TextIOWrapper(fbin, encoding="latin-1", newline="")
                    reader = csv.reader(text)
                    yield name, reader

# =================== MAIN ===================
def main():
    print("[run] started", flush=True)
    gc = gs_client()
    sh = gc.open_by_key(SHEET_ID)
    ws = get_or_create_records_ws(sh)

    existing_header = ws.row_values(1)
    header_in_sheet = bool(existing_header)
    sheet_width = len(existing_header) if existing_header else None
    key_col_idx: Optional[int] = None

    # Load dedupe set lazily (after we know the key column)
    existing_keys: set = set()

    # Current rows (header not counted) - only count non-empty rows
    try:
        all_values = ws.col_values(1)
        # Count only non-empty rows (skip header and empty cells)
        rows_so_far = len([v for v in all_values[1:] if v and str(v).strip()])
    except Exception:
        rows_so_far = 0
    print(f"[sheet] Records has ~{rows_so_far:,} non-empty rows", flush=True)

    total_new = 0

    for url in SCO_ZIPS:
        print(f"Processing: {url}", flush=True)
        zpath = download_to_temp(url)
        try:
            for entry_name, reader in csv_reader_from_zip(zpath):
                header = next(reader, None)
                if header is None:
                    continue

                if not header_in_sheet:
                    # Write original header first
                    safe_update(ws, [header], "A1")
                    print(f"[header] written from {entry_name}", flush=True)
                    
                    # Add new columns one by one to avoid cell limit issues
                    new_columns = ["CREATED_BY_DATE", "TYPE", "CONFIDENCE_LEVEL", "STAGE"]
                    for i, col_name in enumerate(new_columns):
                        col_num = len(header) + i + 1
                        col_letter = gspread.utils.rowcol_to_a1(1, col_num).replace('1', '')
                        safe_update(ws, [[col_name]], f"{col_letter}1")
                    
                    existing_header = header + new_columns
                    sheet_width = len(existing_header)
                    key_col_idx = find_key_col(header)  # Use original header for key column
                    header_in_sheet = True
                    # Load keys now that we know which column to read
                    try:
                        existing_keys |= load_existing_keys_from_records(ws, key_col_idx)
                        print(f"[dedupe] loaded {len(existing_keys):,} keys from current sheet", flush=True)
                    except Exception as e:
                        print(f"[warn] could not load existing keys: {e}", flush=True)
                else:
                    if sheet_width is None:
                        sheet_width = len(existing_header)
                    if key_col_idx is None:
                        key_col_idx = find_key_col(existing_header)
                    if not existing_keys:
                        try:
                            existing_keys |= load_existing_keys_from_records(ws, key_col_idx)
                            print(f"[dedupe] loaded {len(existing_keys):,} keys from current sheet", flush=True)
                        except Exception as e:
                            print(f"[warn] could not load existing keys: {e}", flush=True)

                batch = []
                filtered_count = 0
                for row in reader:
                    if not row:
                        continue
                    
                    # Apply filtering rules
                    if not should_include_record(row, header):
                        filtered_count += 1
                        continue
                    
                    key = (row[key_col_idx] if key_col_idx < len(row) else "").strip()
                    key_norm = key or json.dumps(row, ensure_ascii=False)
                    if key_norm in existing_keys:
                        continue

                    # Add metadata columns
                    extended_row = add_metadata_columns(row, header)

                    if rows_so_far + len(batch) + 1 > MAX_RECORD_ROWS:
                        if batch:
                            append_in_batches(ws, batch, sheet_width)
                            rows_so_far += len(batch)
                            total_new += len(batch)
                            batch.clear()
                        print(f"[stop] capacity at ~{rows_so_far:,} rows", flush=True)
                        print(f"[done] new rows this run: {total_new:,}", flush=True)
                        return

                    batch.append(extended_row)
                    existing_keys.add(key_norm)

                    if len(batch) >= APPEND_BATCH_SIZE:
                        append_in_batches(ws, batch, sheet_width)
                        rows_so_far += len(batch)
                        total_new += len(batch)
                        batch.clear()
                
                if filtered_count > 0:
                    print(f"[filter] excluded {filtered_count:,} records that didn't meet criteria", flush=True)

                if batch:
                    append_in_batches(ws, batch, sheet_width)
                    rows_so_far += len(batch)
                    total_new += len(batch)
                    batch.clear()
        finally:
            try:
                os.remove(zpath)
            except OSError:
                pass

    print(f"[done] new rows this run: {total_new:,}", flush=True)

if __name__ == "__main__":
    main()
