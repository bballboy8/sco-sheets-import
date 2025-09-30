import csv, io, os, json, time, tempfile
from typing import List, Optional
from datetime import datetime
import random

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
    "https://dpupd.sco.ca.gov/01_From_0_To_Below_10.zip",
    "https://dpupd.sco.ca.gov/02_From_10_To_Below_100.zip",
    "https://dpupd.sco.ca.gov/03_From_100_To_Below_500.zip",
    "https://dpupd.sco.ca.gov/04_From_500_To_Beyond.zip",
]
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
    if ws.get_last_row() == 0:
        safe_update(ws, [[SEEN_HDR]], "A1")
    elif (ws.cell(1,1).value or "").strip().upper() != SEEN_HDR:
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
def is_business(owner_name: str) -> bool:
    if not owner_name or not owner_name.strip():
        return False
    name = owner_name.upper()
    space_required_terms = {"OF ", "AND "}
    beginning_terms = {"A ", "THE "}
    for term in BUSINESS_TERMS:
        if term in beginning_terms:
            if name.startswith(term): return True
        elif term in space_required_terms:
            start = 0
            while True:
                pos = name.find(term, start)
                if pos == -1: break
                before_ok = (pos == 0) or (name[pos-1] == " ")
                after_ok  = (pos + len(term) >= len(name)) or (name[pos+len(term)] == " ")
                if before_ok and after_ok: return True
                start = pos + 1
        else:
            if term in name: return True
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
    return row + [current_date, record_type, "100%", "Leads"]

def append_in_batches(ws: gspread.Worksheet, rows: List[List[str]], width: int):
    if not rows: return
    out = []
    for r in rows:
        rr = list(r)
        if len(rr) < width: rr.extend([""] * (width - len(rr)))
        elif len(rr) > width: rr = rr[:width]
        out.append(rr)
        if len(out) >= APPEND_BATCH_SIZE:
            safe_append_rows(ws, out, value_input_option="RAW")
            print(f"[append] {len(out)} rows", flush=True)
            out.clear()
            time.sleep(API_PAUSE_SEC)
    if out:
        safe_append_rows(ws, out, value_input_option="RAW")
        print(f"[append] {len(out)} rows", flush=True)

# ---------- download / unzip ----------
def download_to_temp(url: str) -> str:
    def _download():
        size_hint = None
        try:
            hr = requests.head(url, timeout=30, allow_redirects=True)
            if hr.ok: size_hint = int(hr.headers.get("Content-Length") or 0) or None
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
        zpath = download_to_temp(url)
        try:
            for entry_name, reader in csv_reader_from_zip(zpath):
                header = next(reader, None)
                if header is None: continue

                if not header_in_sheet:
                    # Write original header and our extra columns
                    safe_update(ws, [header], "A1")
                    new_cols = ["CREATED_BY_DATE", "TYPE", "CONFIDENCE_LEVEL", "STAGE"]
                    for i, name in enumerate(new_cols, start=1):
                        col_letter = gspread.utils.rowcol_to_a1(1, len(header)+i).rstrip("1")
                        safe_update(ws, [[name]], f"{col_letter}1")
                    existing_header = header + new_cols
                    sheet_width = len(existing_header)
                    key_col_idx = find_key_col(header)
                    header_in_sheet = True
                    try:
                        existing_sheet_pids |= load_existing_keys_from_records(ws, key_col_idx)
                        print(f"[dedupe] loaded {len(existing_sheet_pids):,} PROPERTY_IDs from Records sheet", flush=True)
                    except Exception as e:
                        print(f"[warn] could not load Records IDs: {e}", flush=True)
                else:
                    if sheet_width is None: sheet_width = len(existing_header)
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
                            append_in_batches(ws, batch, sheet_width)
                            cnt = len(batch)
                            rows_so_far += cnt; total_new += cnt; batch.clear()
                        print(f"[stop] capacity at ~{rows_so_far:,} rows", flush=True)
                        print(f"[done] new rows this run: {total_new:,}", flush=True)
                        return

                    batch.append(extended_row)
                    existing_sheet_pids.add(pid)  # mark as present to avoid intra-run dupes

                    if len(batch) >= APPEND_BATCH_SIZE:
                        append_in_batches(ws, batch, sheet_width)
                        cnt = len(batch)
                        rows_so_far += cnt; total_new += cnt; batch.clear()

                if filtered_count > 0:
                    print(f"[filter] excluded {filtered_count:,} records that didn't meet criteria", flush=True)

                if batch:
                    append_in_batches(ws, batch, sheet_width)
                    cnt = len(batch)
                    rows_so_far += cnt; total_new += cnt; batch.clear()
        finally:
            try: os.remove(zpath)
            except OSError: pass

    print(f"[done] new rows this run: {total_new:,}", flush=True)

if __name__ == "__main__":
    main()
