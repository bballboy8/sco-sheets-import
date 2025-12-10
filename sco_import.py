import csv, io, os, json, time, tempfile, re
from typing import List, Optional
from datetime import datetime
import random
import ssl
import urllib.request
import urllib.error

import requests
import gspread
from google.oauth2.service_account import Credentials
from zipfile import ZipFile
from dotenv import load_dotenv

# ---------- env ----------
load_dotenv()
SHEET_ID              = os.environ["SHEET_ID"]                          # main spreadsheet (Records tab)
SEEN_IDS_SHEET_ID     = os.environ["SEEN_IDS_SHEET_ID"]         # Seen Property IDs spreadsheet (IDs tab)
GOOGLE_CREDENTIALS    = os.environ["GOOGLE_CREDENTIALS"]

# ---------- config ----------
USE_ALL_ZIPS = True
ALL_ZIPS = [
    "https://claimit.ca.gov/upd-property-records/04_From_500_To_Beyond.zip",
    "https://claimit.ca.gov/upd-property-records/03_From_100_To_Below_500.zip", 
    "https://claimit.ca.gov/upd-property-records/02_From_10_To_Below_100.zip",
    "https://claimit.ca.gov/upd-property-records/01_From_0_To_Below_10.zip",
]
# Note: There's also "https://claimit.ca.gov/upd-property-records/00_All_Records.zip" for all records
SCO_ZIPS = ALL_ZIPS if USE_ALL_ZIPS else ALL_ZIPS[:1]

APPEND_BATCH_SIZE = 1000
API_PAUSE_SEC     = 1.5
MAX_RECORD_ROWS   = 500_000

RECORDS_TAB = "Records"
SEEN_TAB    = "IDs"           # inside the Seen IDs spreadsheet
SEEN_HDR    = "PROPERTY_ID"   # header text in A1

# ---------- business terms & filters ----------
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

# ---------- retry ----------
def retry_with_backoff(func, max_retries=3, base_delay=1, max_delay=60):
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"[error] Final attempt failed: {e}", flush=True)
                raise
            delay = min(base_delay * (2 ** attempt) + random.uniform(0, 1), max_delay)
            print(f"[retry] Attempt {attempt + 1} failed: {e}", flush=True)
            print(f"[retry] Waiting {delay:.1f}s before retry...", flush=True)
            time.sleep(delay)

def safe_append_rows(ws, rows, value_input_option="RAW"):
    def _append(): return ws.append_rows(rows, value_input_option=value_input_option)
    return retry_with_backoff(_append)

def safe_update(ws, values, range_name):
    def _update(): return ws.update(values=values, range_name=range_name)
    return retry_with_backoff(_update)

def safe_batch_update(ws, requests):
    def _batch_update(): return ws.batch_update(requests)
    return retry_with_backoff(_batch_update)

# ---------- Data validation helpers ----------
def get_data_validation_from_cell(ws, cell_address):
    """Get data validation rule from a specific cell (e.g., AH2)."""
    try:
        # For now, we'll assume no existing validation and return None
        # The Google Sheets API doesn't have a direct way to get data validation rules
        # We'll create a default rule instead
        return None
    except Exception as e:
        print(f"[warn] Could not get data validation from {cell_address}: {e}", flush=True)
        return None

def apply_data_validation_to_range(ws, start_row, end_row, validation_rule, spreadsheet=None):
    """Apply data validation rule to a range of cells in column AH."""
    if not validation_rule:
        return
    
    try:
        # Create the data validation request using the correct format
        request = {
            "requests": [
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,  # Convert to 0-based
                            "endRowIndex": end_row,
                            "startColumnIndex": 33,  # Column AH (0-based)
                            "endColumnIndex": 34
                        },
                        "rule": validation_rule
                    }
                }
            ]
        }
        
        # Apply the validation rule using the spreadsheet's batch_update method
        if spreadsheet:
            spreadsheet.batch_update(request)
        else:
            # Try to get the spreadsheet from the worksheet
            try:
                spreadsheet = ws.spreadsheet
                spreadsheet.batch_update(request)
            except AttributeError:
                print(f"[warn] Could not access spreadsheet object for data validation", flush=True)
                return
        
        print(f"[validation] Applied data validation to rows {start_row}-{end_row} in column AH", flush=True)
        
    except Exception as e:
        print(f"[warn] Could not apply data validation to rows {start_row}-{end_row}: {e}", flush=True)

def create_default_data_validation_rule():
    """Create a default dropdown validation rule with YES/NO options."""
    return {
        "condition": {
            "type": "ONE_OF_LIST",
            "values": [
                {"userEnteredValue": "YES"},
                {"userEnteredValue": "NO"}
            ]
        },
        "showCustomUi": True,
        "strict": True
    }

def ensure_ah2_has_validation(ws, spreadsheet=None):
    """Ensure AH2 has data validation rule, create one if it doesn't exist."""
    try:
        validation_rule = get_data_validation_from_cell(ws, "AH2")
        
        if not validation_rule:
            print("[validation] Creating default YES/NO dropdown validation in AH2", flush=True)
            default_rule = create_default_data_validation_rule()
            apply_data_validation_to_range(ws, 2, 2, default_rule, spreadsheet)  # Apply to AH2 only
            return default_rule
        else:
            print("[validation] Found existing data validation rule in AH2", flush=True)
            return validation_rule
    except Exception as e:
        print(f"[warn] Could not ensure AH2 has validation: {e}", flush=True)
        return None

def copy_data_validation_to_new_rows(ws, start_row, num_rows, spreadsheet=None):
    """Copy data validation from AH2 to new rows starting at start_row."""
    try:
        # Ensure AH2 has validation rule
        validation_rule = ensure_ah2_has_validation(ws, spreadsheet)
        
        if validation_rule:
            # Apply the validation rule to the new rows
            end_row = start_row + num_rows - 1
            apply_data_validation_to_range(ws, start_row, end_row, validation_rule, spreadsheet)
        else:
            print(f"[warn] Could not create or find data validation rule for AH2", flush=True)
    except Exception as e:
        print(f"[warn] Could not copy data validation to new rows {start_row}-{start_row + num_rows - 1}: {e}", flush=True)
        # Don't raise - allow the main process to continue even if validation fails

def apply_validation_to_all_rows_in_ah(ws, spreadsheet=None):
    """Apply data validation to all rows in column AH from row 2 to the bottom of the sheet."""
    try:
        # Ensure AH2 has validation rule
        validation_rule = ensure_ah2_has_validation(ws, spreadsheet)
        
        if not validation_rule:
            print(f"[warn] Could not get validation rule for AH2", flush=True)
            return
        
        # Find the last row - try multiple methods
        last_row = None
        
        # Method 1: Try to get row count from worksheet properties
        try:
            # Get worksheet metadata which includes gridProperties
            sheet_metadata = ws.spreadsheet.get_worksheet_by_id(ws.id)
            if hasattr(sheet_metadata, 'row_count'):
                last_row = sheet_metadata.row_count
        except Exception:
            pass
        
        # Method 2: Try ws.row_count property (if available)
        if last_row is None:
            try:
                if hasattr(ws, 'row_count'):
                    last_row = ws.row_count
            except Exception:
                pass
        
        # Method 3: Get all values from column A and find last non-empty row
        if last_row is None:
            try:
                col_a_values = ws.col_values(1)  # Get all values from column A
                if col_a_values:
                    # Find the last non-empty row
                    for i in range(len(col_a_values) - 1, -1, -1):
                        if col_a_values[i] and str(col_a_values[i]).strip():
                            last_row = i + 1  # Convert to 1-based
                            break
                    # If no non-empty rows found, use length
                    if last_row is None:
                        last_row = len(col_a_values)
            except Exception:
                pass
        
        # Method 4: Use get_all_values() to get all rows (more reliable than col_values)
        if last_row is None:
            try:
                all_values = ws.get_all_values()
                if all_values:
                    # Find last non-empty row (checking column A)
                    for i in range(len(all_values) - 1, -1, -1):
                        if i < len(all_values) and all_values[i] and len(all_values[i]) > 0:
                            if all_values[i][0] and str(all_values[i][0]).strip():
                                last_row = i + 1  # Convert to 1-based (row 1 is header)
                                break
                    if last_row is None:
                        # If no non-empty rows, use the total number of rows
                        last_row = len(all_values)
            except Exception as e:
                print(f"[warn] get_all_values() failed: {e}", flush=True)
                pass
        
        # Method 5: Use MAX_RECORD_ROWS as fallback (we know max capacity)
        if last_row is None:
            print(f"[warn] Could not determine last row, using MAX_RECORD_ROWS + 1", flush=True)
            last_row = MAX_RECORD_ROWS + 1
        else:
            # Ensure we don't exceed MAX_RECORD_ROWS
            last_row = min(last_row, MAX_RECORD_ROWS + 1)
        
        # Apply validation from row 2 to last_row
        if last_row >= 2:
            print(f"[validation] Applying validation to column AH from row 2 to row {last_row}", flush=True)
            apply_data_validation_to_range(ws, 2, last_row, validation_rule, spreadsheet)
            print(f"[validation] Applied validation to column AH from row 2 to row {last_row}", flush=True)
        else:
            print(f"[validation] No data rows found, validation only applied to AH2", flush=True)
    except Exception as e:
        print(f"[warn] Could not apply validation to all rows in column AH: {e}", flush=True)
        import traceback
        print(f"[warn] Traceback: {traceback.format_exc()}", flush=True)
        # Don't raise - allow the main process to continue even if validation fails

# ---------- Google auth ----------
def gs_client():
    info = json.loads(GOOGLE_CREDENTIALS)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

# ---------- Sheets helpers ----------
def get_or_create_records_ws(sh) -> gspread.Worksheet:
    try:
        return sh.worksheet(RECORDS_TAB)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=RECORDS_TAB, rows=1, cols=26)

def open_seen_ids_ws(gc) -> Optional[gspread.Worksheet]:
    """Open Seen IDs worksheet (SEEN_TAB), ensure header. Returns None if SEEN_IDS_SHEET_ID unset."""
    if not SEEN_IDS_SHEET_ID:
        return None
    sh = gc.open_by_key(SEEN_IDS_SHEET_ID)
    try:
        ws = sh.worksheet(SEEN_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SEEN_TAB, rows=1, cols=1)
    # ensure header
    try:
        first_cell = ws.cell(1,1).value
        if not first_cell or first_cell.strip().upper() != SEEN_HDR:
            safe_update(ws, [[SEEN_HDR]], "A1")
    except Exception:
        safe_update(ws, [[SEEN_HDR]], "A1")
    return ws

def load_seen_property_ids(gc) -> set:
    """Load processed PROPERTY_IDs from the Seen IDs sheet."""
    ws = open_seen_ids_ws(gc)
    if not ws:
        print("[seen] SEEN_IDS_SHEET_ID not set; skipping external seen list.", flush=True)
        return set()
    col = ws.col_values(1)[1:]  # skip header
    seen = set()
    for v in col:
        s = (v or "").strip()
        if s:
            seen.add(s)
    print(f"[seen] loaded {len(seen):,} IDs from Seen IDs sheet", flush=True)
    return seen

def find_key_col(header: List[str]) -> int:
    for i, h in enumerate(header):
        if str(h).strip().upper() == "PROPERTY_ID": return i
    for i, h in enumerate(header):
        hs = str(h).lower()
        if "property" in hs and "id" in hs: return i
    return 0

def load_existing_keys_from_records(ws: gspread.Worksheet, key_col_idx: int) -> set:
    """Return a set of PROPERTY_IDs already on the Records sheet (non-empty only)."""
    col = key_col_idx + 1
    vals = ws.col_values(col)[1:]
    return set((v or "").strip() for v in vals if v and str(v).strip())

# ---------- rules ----------
def _matches_with_boundaries(text: str, term: str) -> bool:
    """
    Check if a term appears as a whole word in the text.
    A whole word means the term is surrounded by word boundaries (non-alphanumeric characters
    or start/end of string).
    
    For terms ending with a space (like "A ", "THE "), the trailing space serves as the
    word boundary, so we don't require an additional non-alphanumeric character after it.
    """
    if not term or not text:
        return False
    
    # Escape special regex characters in the term
    escaped_term = re.escape(term)
    
    # Check if term ends with a space - if so, the space itself is the boundary
    if term.endswith(' '):
        # For terms ending with space, only check that it's not preceded by alphanumeric
        # The trailing space ensures it's followed by a word boundary
        pattern = r'(?<![A-Za-z0-9])' + escaped_term
    else:
        # For other terms, require word boundaries on both sides
        pattern = r'(?<![A-Za-z0-9])' + escaped_term + r'(?![A-Za-z0-9])'
    
    return bool(re.search(pattern, text))

def is_business(owner_name: str) -> bool:
    if not owner_name or not owner_name.strip():
        return False
    name = owner_name.upper()
    
    for term in BUSINESS_TERMS:
        # Use whole word matching for all terms
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
    # Try common field name variations
    mobile_phone = ""
    other_phone = ""
    email = ""
    
    # Try to find mobile phone field (prioritize mobile-specific fields)
    mobile_field_names = ["MOBILE_PHONE", "MOBILE", "OWNER_MOBILE_PHONE", "OWNER_MOBILE", "CELL_PHONE", "CELL"]
    for field_name in mobile_field_names:
        if field_name in header:
            idx = header.index(field_name)
            if idx < len(row) and row[idx]:
                mobile_phone = str(row[idx]).strip()
                break
    
    # Try to find other phone field (prioritize "OTHER_PHONE", then generic phone fields)
    other_phone_field_names = ["OTHER_PHONE", "PHONE", "OWNER_PHONE", "HOME_PHONE", "WORK_PHONE", "TELEPHONE"]
    for field_name in other_phone_field_names:
        if field_name in header:
            idx = header.index(field_name)
            if idx < len(row) and row[idx]:
                phone_value = str(row[idx]).strip()
                # If this is "OTHER_PHONE", always use it
                if field_name == "OTHER_PHONE":
                    other_phone = phone_value
                    break
                # For generic "PHONE" fields, only use if mobile_phone wasn't found from a mobile-specific field
                # and we haven't set other_phone yet
                elif not mobile_phone and not other_phone:
                    # Check if this field was already used for mobile_phone
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
    
    # Build metadata columns - always add all 7 columns since EXPORT_TO_CRM is now in AH
    metadata_cols = [current_date, record_type, "100%", "Leads", mobile_phone, other_phone, email]
    
    return row + metadata_cols

def find_first_empty_row(ws: gspread.Worksheet) -> int:
    """Find the first empty row after the header row (row 1). Returns 1-based row number."""
    try:
        col_a_values = ws.col_values(1)  # Get all values from column A
        # Start from row 2 (index 1) since row 1 is always headers
        for i in range(1, len(col_a_values)):
            value = col_a_values[i]
            if not value or not str(value).strip():
                return i + 1  # Return 1-based row number
        # If all rows after header have values, return next row
        return len(col_a_values) + 1
    except Exception:
        return 2  # If error, start from row 2 (after header)

def write_in_batches(ws: gspread.Worksheet, rows: List[List[str]], data_width: int, spreadsheet=None):
    if not rows: return
    out = []
    current_start_row = find_first_empty_row(ws)
    
    # Force writing only to columns A-AG (33 columns) to include new metadata columns
    # Note: EXPORT_TO_CRM validation is in column AH (34), which is outside this range
    max_cols = 33
    
    for r in rows:
        rr = list(r)
        # Only pad to max_cols, don't extend beyond AG
        if len(rr) < max_cols: 
            rr.extend([""] * (max_cols - len(rr)))
        elif len(rr) > max_cols: 
            rr = rr[:max_cols]
        out.append(rr)
        
        if len(out) >= APPEND_BATCH_SIZE:
            # Write batch starting at current_start_row, up to column AG
            end_row = current_start_row + len(out) - 1
            range_name = f"A{current_start_row}:AG{end_row}"
            safe_update(ws, out, range_name)
            print(f"[write] {len(out)} rows at row {current_start_row} (columns A-AG)", flush=True)
            
            # Copy data validation from AH2 to the new rows
            copy_data_validation_to_new_rows(ws, current_start_row, len(out), spreadsheet)
            
            current_start_row = end_row + 1
            out.clear()
            time.sleep(API_PAUSE_SEC)
    
    if out:
        # Write remaining rows
        end_row = current_start_row + len(out) - 1
        range_name = f"A{current_start_row}:AG{end_row}"
        safe_update(ws, out, range_name)
        print(f"[write] {len(out)} rows at row {current_start_row} (columns A-AG)", flush=True)
        
        # Copy data validation from AH2 to the new rows
        copy_data_validation_to_new_rows(ws, current_start_row, len(out), spreadsheet)

# ---------- download / unzip ----------
def download_to_temp(url: str) -> str:
    def _download():
        import socket
        import urllib3
        
        # Test connection first
        try:
            hostname = url.split("//")[1].split("/")[0]
            print(f"[network] Testing connection to {hostname}...", flush=True)
            
            # Try to resolve DNS
            try:
                ip = socket.gethostbyname(hostname)
                print(f"[network] DNS resolved {hostname} to {ip}", flush=True)
            except socket.gaierror as e:
                print(f"[network] DNS resolution failed: {e}", flush=True)
                raise
            
            # Try to connect
            try:
                sock = socket.create_connection((hostname, 443), timeout=10)
                sock.close()
                print(f"[network] TCP connection test successful", flush=True)
            except (socket.timeout, ConnectionRefusedError) as e:
                print(f"[network] TCP connection failed: {e}", flush=True)
                print(f"[network] Note: The server may be blocking connections or down", flush=True)
        except Exception as e:
            print(f"[network] Connection test error: {e}", flush=True)
        
        size_hint = None
        try:
            # Try HEAD request with SSL verification disabled
            session = requests.Session()
            session.verify = False
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            
            hr = session.head(url, timeout=30, allow_redirects=True)
            if hr.ok: 
                size_hint = int(hr.headers.get("Content-Length") or 0) or None
                print(f"[network] HEAD request successful, size hint: {size_hint/1024/1024:.1f} MB" if size_hint else "[network] HEAD request successful", flush=True)
        except Exception as e:
            print(f"[network] HEAD request failed: {e}", flush=True)

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Referer": "https://claimit.ca.gov/",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate, br",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        print(f"[download] starting -> {url}", flush=True)
        print(f"[download] Using SSL verification: False (bypassing certificate check)", flush=True)
        
        # Create session with disabled SSL verification
        session = requests.Session()
        session.verify = False
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        with session.get(url, headers=headers, stream=True, timeout=600) as r:
            r.raise_for_status()
            fd, path = tempfile.mkstemp(suffix=".zip")
            total = 0
            next_marker = 50 * 1024 * 1024
            with os.fdopen(fd, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 1024):
                    if not chunk: continue
                    f.write(chunk); total += len(chunk)
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

def download_with_urllib(url: str) -> str:
    """Fallback download method using urllib if requests fails."""
    print(f"[fallback] Trying urllib download method for {url}", flush=True)
    
    # Create SSL context that doesn't verify certificates
    ssl_context = ssl.create_default_context()
    ssl_context.check_hostname = False
    ssl_context.verify_mode = ssl.CERT_NONE
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://claimit.ca.gov/",
    }
    
    req = urllib.request.Request(url, headers=headers)
    
    try:
        with urllib.request.urlopen(req, context=ssl_context, timeout=600) as response:
            fd, path = tempfile.mkstemp(suffix=".zip")
            total = 0
            next_marker = 50 * 1024 * 1024
            
            with os.fdopen(fd, "wb") as f:
                while True:
                    chunk = response.read(1024 * 1024)  # 1MB chunks
                    if not chunk:
                        break
                    f.write(chunk)
                    total += len(chunk)
                    
                    if total >= next_marker:
                        mb = total / (1024 * 1024)
                        print(f"[fallback download] {mb:.0f} MB", flush=True)
                        next_marker += 50 * 1024 * 1024
            
            size_mb = os.path.getsize(path) / (1024 * 1024)
            print(f"[fallback downloaded] {size_mb:.1f} MB", flush=True)
            return path
            
    except urllib.error.URLError as e:
        print(f"[fallback] urllib download failed: {e}", flush=True)
        raise

def csv_reader_from_zip(zip_path: str):
    with ZipFile(zip_path, mode="r", allowZip64=True) as zf:
        names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        print(f"[zip] {len(names)} CSV file(s): {', '.join(names[:3])}{' …' if len(names) > 3 else ''}", flush=True)
        for name in names:
            with zf.open(name, "r") as fbin:
                try:
                    text = io.TextIOWrapper(fbin, encoding="utf-8-sig", newline="")
                    reader = csv.reader(text); yield name, reader
                except UnicodeDecodeError:
                    fbin.seek(0)
                    text = io.TextIOWrapper(fbin, encoding="latin-1", newline="")
                    reader = csv.reader(text); yield name, reader

# ---------- main ----------
def main():
    print("[run] started", flush=True)
    gc = gs_client()

    # open main sheet
    sh = gc.open_by_key(SHEET_ID)
    ws = get_or_create_records_ws(sh)
    
    # Ensure AH2 has data validation rule and apply it to all rows
    ensure_ah2_has_validation(ws, sh)
    apply_validation_to_all_rows_in_ah(ws, sh)

    # dedupe sets
    existing_header = ws.row_values(1)
    header_in_sheet = bool(existing_header)
    sheet_width = len(existing_header) if existing_header else None
    key_col_idx: Optional[int] = None

    # count non-empty rows (skip header)
    try:
        col1 = ws.col_values(1)[1:]
        rows_so_far = sum(1 for v in col1 if (v or "").strip())
    except Exception:
        rows_so_far = 0
    print(f"[sheet] Records has ~{rows_so_far:,} non-empty rows", flush=True)

    # IDs we've already processed in the separate Seen IDs sheet
    seen_external_ids = load_seen_property_ids(gc)

    # property IDs already on the Records sheet (avoid importing duplicates into the same sheet)
    existing_sheet_pids: set = set()

    total_new = 0

    for url in SCO_ZIPS:
        print(f"Processing: {url}", flush=True)
        
        # Try regular download first, then fallback to urllib if it fails
        try:
            zpath = download_to_temp(url)
        except Exception as e:
            print(f"[error] Regular download failed: {e}", flush=True)
            print(f"[fallback] Attempting alternative download method...", flush=True)
            try:
                zpath = download_with_urllib(url)
            except Exception as e2:
                print(f"[error] All download methods failed for {url}", flush=True)
                print(f"[error] Last error: {e2}", flush=True)
                print(f"[skip] Skipping this ZIP file and continuing with next one...", flush=True)
                continue
        try:
            for entry_name, reader in csv_reader_from_zip(zpath):
                header = next(reader, None)
                if header is None: continue
                
                # Print header fields for analysis (only once per CSV file)
                print(f"[header] CSV file '{entry_name}' has {len(header)} columns:", flush=True)
                print(f"[header] Columns: {', '.join(header)}", flush=True)

                if not header_in_sheet:
                    # Always write headers since row 1 is always reserved for headers
                    print(f"[header] Writing headers", flush=True)
                    safe_update(ws, [header], "A1")
                    new_cols = ["CREATED_BY_DATE", "TYPE", "CONFIDENCE_LEVEL", "STAGE", "MOBILE PHONE", "OTHER PHONE", "EMAIL"]
                    
                    for i, name in enumerate(new_cols, start=1):
                        col_letter = gspread.utils.rowcol_to_a1(1, len(header)+i).rstrip("1")
                        safe_update(ws, [[name]], f"{col_letter}1")
                    existing_header = header + new_cols
                    
                    # Limit sheet_width to AG (33 columns) to include new metadata columns
                    sheet_width = min(len(existing_header), 33)  # Write to columns A-AG
                    key_col_idx = find_key_col(header)
                    header_in_sheet = True
                    try:
                        existing_sheet_pids |= load_existing_keys_from_records(ws, key_col_idx)
                        print(f"[dedupe] loaded {len(existing_sheet_pids):,} PROPERTY_IDs from Records sheet", flush=True)
                    except Exception as e:
                        print(f"[warn] could not load Records IDs: {e}", flush=True)
                else:
                    if sheet_width is None: sheet_width = min(len(existing_header), 33)  # Limit to A-AG
                    if key_col_idx is None: key_col_idx = find_key_col(existing_header)
                    if not existing_sheet_pids:
                        try:
                            existing_sheet_pids |= load_existing_keys_from_records(ws, key_col_idx)
                            print(f"[dedupe] loaded {len(existing_sheet_pids):,} PROPERTY_IDs from Records sheet", flush=True)
                        except Exception as e:
                            print(f"[warn] could not load Records IDs: {e}", flush=True)

                batch = []
                filtered_count = 0
                for row in reader:
                    if not row: continue
                    if not should_include_record(row, header):
                        filtered_count += 1; continue

                    pid = (row[key_col_idx] if key_col_idx < len(row) else "").strip()
                    if not pid: continue

                    # *** NEW: skip if we've already processed this ID in Seen IDs sheet ***
                    if pid in seen_external_ids:
                        continue

                    # also skip if already in Records sheet
                    if pid in existing_sheet_pids:
                        continue

                    extended_row = add_metadata_columns(row, header)

                    # capacity guard
                    if rows_so_far + len(batch) + 1 > MAX_RECORD_ROWS:
                        if batch:
                            write_in_batches(ws, batch, sheet_width, sh)
                            cnt = len(batch)
                            rows_so_far += cnt; total_new += cnt; batch.clear()
                        print(f"[stop] capacity at ~{rows_so_far:,} rows", flush=True)
                        print(f"[done] new rows this run: {total_new:,}", flush=True)
                        return

                    batch.append(extended_row)
                    existing_sheet_pids.add(pid)  # mark as present to avoid intra-run dupes

                    if len(batch) >= APPEND_BATCH_SIZE:
                        write_in_batches(ws, batch, sheet_width, sh)
                        cnt = len(batch)
                        rows_so_far += cnt; total_new += cnt; batch.clear()

                if filtered_count > 0:
                    print(f"[filter] excluded {filtered_count:,} records that didn't meet criteria", flush=True)

                if batch:
                    write_in_batches(ws, batch, sheet_width, sh)
                    cnt = len(batch)
                    rows_so_far += cnt; total_new += cnt; batch.clear()
        finally:
            try: os.remove(zpath)
            except OSError: pass

    print(f"[done] new rows this run: {total_new:,}", flush=True)

if __name__ == "__main__":
    main()
