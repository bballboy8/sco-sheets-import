import csv, io, os, json, time, tempfile
from typing import List, Optional, Tuple

import requests
import gspread
from google.oauth2.service_account import Credentials
from zipfile import ZipFile  # ZIP64 supported

# =================== CONFIG ===================
SHEET_ID = os.environ["SHEET_ID"]                       # main spreadsheet (has Records + Meta)
GOOGLE_CREDENTIALS = os.environ["GOOGLE_CREDENTIALS"]   # full JSON as a single string (GitHub secret)

# Optional: sheet that stores ALL seen PROPERTY_IDs (one column, tab name "IDs")
# Create an empty spreadsheet, add a sheet named "IDs", share with the service account, and set its ID here.
SEEN_IDS_SHEET_ID = os.environ.get("SEEN_IDS_SHEET_ID", "").strip()

# The 4 segmented archives (avoid All_Records.zip)
SCO_ZIPS = [
    "https://dpupd.sco.ca.gov/01_From_0_To_Below_10.zip",
    "https://dpupd.sco.ca.gov/02_From_10_To_Below_100.zip",
    "https://dpupd.sco.ca.gov/03_From_100_To_Below_500.zip",
    "https://dpupd.sco.ca.gov/04_From_500_To_Beyond.zip",
]

# Batching / pacing
APPEND_BATCH_SIZE = 5000
IDS_APPEND_BATCH  = 20000  # write IDs to archive in bigger chunks (1 col, cheap)
API_PAUSE_SEC     = 1.0

# Capacity backstop:
# Google Sheets hard cap = 10,000,000 cells. We stop by rows to stay safe.
# Roughly: if your header width is ~60 columns, 150k rows ≈ 9M cells.
MAX_RECORD_ROWS = 180_000   # header not counted; adjust to your sheet's width tolerance

# Names
RECORDS_TAB = "Records"
META_TAB    = "Meta"    # hidden sheet used for checkpoints/logs
IDS_TAB     = "IDs"     # in the SEEN_IDS_SHEET_ID spreadsheet


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
def open_or_create_meta(sh) -> gspread.Worksheet:
    try:
        ws = sh.worksheet(META_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=META_TAB, rows=100, cols=6)
        ws.update(values=[["zip_url","entry_index","row_offset","notes","ts"]], range_name="A1")
    return ws

def get_or_create_records_ws(sh) -> gspread.Worksheet:
    try:
        ws = sh.worksheet(RECORDS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=RECORDS_TAB, rows=1, cols=26)
    return ws

def open_ids_archive(gc) -> Optional[gspread.Worksheet]:
    if not SEEN_IDS_SHEET_ID:
        return None
    sh = gc.open_by_key(SEEN_IDS_SHEET_ID)
    try:
        ws = sh.worksheet(IDS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=IDS_TAB, rows=1, cols=1)
        ws.update(values=[["PROPERTY_ID"]], range_name="A1")
    return ws

def log_meta(meta_ws: gspread.Worksheet, msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    meta_ws.append_row([ "", "", "", msg, ts ], value_input_option="RAW")

def read_checkpoint(meta_ws: gspread.Worksheet, zip_url: str) -> Tuple[int,int]:
    # Returns (entry_index, row_offset); defaults to (0, 0)
    try:
        vals = meta_ws.get_all_values()
    except Exception:
        return (0, 0)
    for r in vals[1:]:
        if r and r[0] == zip_url:
            ei = int(r[1] or "0")
            ro = int(r[2] or "0")
            return (ei, ro)
    return (0, 0)

def write_checkpoint(meta_ws: gspread.Worksheet, zip_url: str, entry_index: int, row_offset: int, note: str=""):
    vals = meta_ws.get_all_values()
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    if len(vals) <= 1:
        meta_ws.update(values=[["zip_url","entry_index","row_offset","notes","ts"]], range_name="A1")
        vals = [["zip_url","entry_index","row_offset","notes","ts"]]
    # find row
    row_idx = None
    for i, r in enumerate(vals[1:], start=2):
        if r and r[0] == zip_url:
            row_idx = i
            break
    data = [[zip_url, str(entry_index), str(row_offset), note, ts]]
    if row_idx:
        meta_ws.update(values=data, range_name=f"A{row_idx}:E{row_idx}")
    else:
        meta_ws.append_row(data[0], value_input_option="RAW")

def clear_checkpoint(meta_ws: gspread.Worksheet, zip_url: str):
    vals = meta_ws.get_all_values()
    if len(vals) <= 1:
        return
    for i, r in enumerate(vals[1:], start=2):
        if r and r[0] == zip_url:
            meta_ws.delete_rows(i)
            break

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
    vals = ws.col_values(col)[1:]
    return set((v or "").strip() for v in vals)

def load_existing_keys_from_archive(ids_ws: Optional[gspread.Worksheet]) -> set:
    if not ids_ws:
        return set()
    vals = ids_ws.col_values(1)[1:]
    return set((v or "").strip() for v in vals)

def append_ids_to_archive(ids_ws: Optional[gspread.Worksheet], ids_batch: List[str]):
    if not ids_ws or not ids_batch:
        return
    # unique and non-empty
    uniq = [i for i in dict.fromkeys(i for i in ids_batch if i)]
    # chunk large inserts
    out = []
    for pid in uniq:
        out.append([pid])
        if len(out) >= IDS_APPEND_BATCH:
            ids_ws.append_rows(out, value_input_option="RAW")
            out.clear()
            time.sleep(API_PAUSE_SEC)
    if out:
        ids_ws.append_rows(out, value_input_option="RAW")

def append_in_batches(ws: gspread.Worksheet, rows: List[List[str]], width: int):
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
            out.clear()
            time.sleep(API_PAUSE_SEC)
    if out:
        ws.append_rows(out, value_input_option="RAW")


# =================== SCO DOWNLOAD / CSV STREAM ===================
def download_to_temp(url: str) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://sco.ca.gov/upd_download_property_records.html",
        "Accept": "*/*",
    }
    with requests.get(url, headers=headers, stream=True, timeout=180) as r:
        r.raise_for_status()
        fd, path = tempfile.mkstemp(suffix=".zip")
        with os.fdopen(fd, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)
        return path

def csv_reader_from_zip(zip_path: str):
    with ZipFile(zip_path, mode="r", allowZip64=True) as zf:
        for name in zf.namelist():
            if not name.lower().endswith(".csv"):
                continue
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
    gc = gs_client()
    main_sh = gc.open_by_key(SHEET_ID)
    meta_ws = open_or_create_meta(main_sh)
    records_ws = get_or_create_records_ws(main_sh)
    ids_ws = open_ids_archive(gc)  # may be None

    # Header state
    existing_header = records_ws.row_values(1)
    header_in_sheet = bool(existing_header)
    sheet_width = len(existing_header) if existing_header else None
    key_col_idx: Optional[int] = None

    # De-dupe set: load from Records + Archive IDs
    # Note: This set can be large; for very large archives consider sharding IDs archive.
    existing_keys = set()
    if header_in_sheet:
        # We'll set key_col_idx once we know the header mapping
        pass

    if ids_ws:
        existing_keys |= load_existing_keys_from_archive(ids_ws)

    # Current rows (header not counted)
    try:
        rows_so_far = len(records_ws.col_values(1)) - 1
    except Exception:
        rows_so_far = 0

    total_new = 0
    ids_to_archive_batch: List[str] = []

    for url in SCO_ZIPS:
        print(f"Processing: {url}", flush=True)
        entry_start, row_offset = read_checkpoint(meta_ws, url)

        zpath = download_to_temp(url)
        try:
            entry_idx = -1
            for entry_name, reader in csv_reader_from_zip(zpath):
                entry_idx += 1

                # Skip entries before checkpoint
                if entry_idx < entry_start:
                    # consume header to keep reader aligned
                    _ = next(reader, None)
                    continue

                header = next(reader, None)
                if header is None:
                    continue

                if not header_in_sheet:
                    # First time: write header, detect key column
                    records_ws.update(values=[header], range_name="A1")
                    existing_header = header
                    sheet_width = len(header)
                    key_col_idx = find_key_col(header)
                    # Now we can load keys from Records column (existing content is just header)
                    # existing_keys unchanged here
                    header_in_sheet = True
                else:
                    if sheet_width is None:
                        sheet_width = len(existing_header)
                    if key_col_idx is None:
                        key_col_idx = find_key_col(existing_header)

                # If not already loaded, load keys from Records (now that we know key column)
                if not existing_keys:
                    try:
                        existing_keys |= load_existing_keys_from_records(records_ws, key_col_idx)
                    except Exception:
                        pass

                # Fast-forward rows within current entry to checkpoint offset
                skipped = 0
                if entry_idx == entry_start and row_offset > 0:
                    for _ in range(row_offset):
                        try:
                            next(reader)
                            skipped += 1
                        except StopIteration:
                            break

                batch = []
                local_offset = row_offset if entry_idx == entry_start else 0

                for row in reader:
                    local_offset += 1
                    if not row:
                        continue

                    key = (row[key_col_idx] if key_col_idx < len(row) else "").strip()
                    key_norm = key or json.dumps(row, ensure_ascii=False)

                    if key_norm in existing_keys:
                        continue

                    # Capacity guard by rows
                    if rows_so_far + len(batch) + 1 > MAX_RECORD_ROWS:
                        # Flush pending work
                        if batch:
                            append_in_batches(records_ws, batch, sheet_width)
                            rows_so_far += len(batch)
                            total_new += len(batch)
                            # collect IDs for archive
                            for r in batch:
                                pid = (r[key_col_idx] if key_col_idx < len(r) else "").strip()
                                if pid:
                                    ids_to_archive_batch.append(pid)
                            batch.clear()
                            append_ids_to_archive(ids_ws, ids_to_archive_batch)
                            ids_to_archive_batch.clear()

                        write_checkpoint(meta_ws, url, entry_idx, local_offset, note="capacity stop")
                        print(f"Capacity stop at ~{rows_so_far} rows. Checkpoint saved ({url} @ entry {entry_idx}, offset {local_offset}).", flush=True)
                        print(f"Done. New rows added: {total_new}", flush=True)
                        return

                    # Accept row
                    batch.append(row)
                    existing_keys.add(key_norm)
                    if key:
                        ids_to_archive_batch.append(key)

                    if len(batch) >= APPEND_BATCH_SIZE:
                        append_in_batches(records_ws, batch, sheet_width)
                        rows_so_far += len(batch)
                        total_new += len(batch)
                        batch.clear()
                        # archive IDs periodically
                        if len(ids_to_archive_batch) >= IDS_APPEND_BATCH:
                            append_ids_to_archive(ids_ws, ids_to_archive_batch)
                            ids_to_archive_batch.clear()

                # End of entry: persist checkpoint progress within this ZIP
                write_checkpoint(meta_ws, url, entry_idx, 0, note=f"entry {entry_idx} complete")

            # Finished this ZIP
            clear_checkpoint(meta_ws, url)

        finally:
            try:
                os.remove(zpath)
            except OSError:
                pass

    # Final flush of any pending IDs
    append_ids_to_archive(ids_ws, ids_to_archive_batch)
    print(f"Done. New rows added: {total_new}", flush=True)


if __name__ == "__main__":
    main()
