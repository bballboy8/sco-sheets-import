import csv, io, os, json, time, tempfile
from typing import List, Optional

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
API_PAUSE_SEC = 0.8
MAX_RECORD_ROWS = 500_000

RECORDS_TAB = "Records"

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
            ws.append_rows(out, value_input_option="RAW")
            print(f"[append] {len(out)} rows", flush=True)
            out.clear()
            time.sleep(API_PAUSE_SEC)
    if out:
        ws.append_rows(out, value_input_option="RAW")
        print(f"[append] {len(out)} rows", flush=True)

# =================== SCO DOWNLOAD / CSV STREAM ===================
def download_to_temp(url: str) -> str:
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
                    ws.update(values=[header], range_name="A1")
                    print(f"[header] written from {entry_name}", flush=True)
                    existing_header = header
                    sheet_width = len(header)
                    key_col_idx = find_key_col(header)
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
                for row in reader:
                    if not row:
                        continue
                    key = (row[key_col_idx] if key_col_idx < len(row) else "").strip()
                    key_norm = key or json.dumps(row, ensure_ascii=False)
                    if key_norm in existing_keys:
                        continue

                    if rows_so_far + len(batch) + 1 > MAX_RECORD_ROWS:
                        if batch:
                            append_in_batches(ws, batch, sheet_width)
                            rows_so_far += len(batch)
                            total_new += len(batch)
                            batch.clear()
                        print(f"[stop] capacity at ~{rows_so_far:,} rows", flush=True)
                        print(f"[done] new rows this run: {total_new:,}", flush=True)
                        return

                    batch.append(row)
                    existing_keys.add(key_norm)

                    if len(batch) >= APPEND_BATCH_SIZE:
                        append_in_batches(ws, batch, sheet_width)
                        rows_so_far += len(batch)
                        total_new += len(batch)
                        batch.clear()

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
