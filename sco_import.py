import csv, io, os, sys, json, time, tempfile
from typing import List
import requests
import gspread
from google.oauth2.service_account import Credentials
from zipfile import ZipFile, ZIP_DEFLATED  # ZIP64 automatically supported

SHEET_ID = os.environ["SHEET_ID"]  # set in GitHub Secrets
GOOGLE_CREDENTIALS = os.environ["GOOGLE_CREDENTIALS"]  # JSON string in Secrets

# The 4 segmented archives (avoid the huge All_Records.zip)
SCO_ZIPS = [
    "https://dpupd.sco.ca.gov/01_From_0_To_Below_10.zip",
    "https://dpupd.sco.ca.gov/02_From_10_To_Below_100.zip",
    "https://dpupd.sco.ca.gov/03_From_100_To_Below_500.zip",
    "https://dpupd.sco.ca.gov/04_From_500_To_Beyond.zip",
]

# ------------- Google auth -------------
def gs_client():
    info = json.loads(GOOGLE_CREDENTIALS)
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

# ------------- Sheet helpers -------------
def get_or_create_records_ws(gc):
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("Records")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Records", rows=1, cols=26)
    return ws

def find_key_col(header: List[str]) -> int:
    # Prefer exact "Property ID", else anything containing both "property" and "id"
    for i, h in enumerate(header):
        if str(h).strip().lower() == "property id":
            return i
    for i, h in enumerate(header):
        hs = str(h).lower()
        if "property" in hs and "id" in hs:
            return i
    return 0  # fallback

def load_existing_keys(ws, key_col_idx: int) -> set:
    # Read header (row 1) to know sheet width; then fetch only the key column below header
    # gspread is 1-indexed for columns
    col = key_col_idx + 1
    # .col_values returns entire column; slice off header
    vals = ws.col_values(col)[1:]
    # normalize to strings
    return set((v or "").strip() for v in vals)

def append_in_batches(ws, rows: List[List[str]], width: int, batch_size: int = 3000):
    # normalize row width and append in chunks
    out = []
    for r in rows:
        rr = list(r)
        if len(rr) < width:
            rr.extend([""] * (width - len(rr)))
        elif len(rr) > width:
            rr = rr[:width]
        out.append(rr)
        if len(out) >= batch_size:
            ws.append_rows(out, value_input_option="RAW")
            out.clear()
            # polite pacing to avoid API quota bursts
            time.sleep(1)
    if out:
        ws.append_rows(out, value_input_option="RAW")

# ------------- SCO download/unzip/stream -------------
def download_to_temp(url: str) -> str:
    # stream to disk (keeps RAM low)
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://sco.ca.gov/upd_download_property_records.html",
        "Accept": "*/*",
    }
    with requests.get(url, headers=headers, stream=True, timeout=120) as r:
        r.raise_for_status()
        fd, path = tempfile.mkstemp(suffix=".zip")
        with os.fdopen(fd, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)
        return path

def csv_reader_from_zip(zip_path: str):
    # Iterate CSV entries one by one, streaming rows
    with ZipFile(zip_path, mode="r", allowZip64=True) as zf:
        for name in zf.namelist():
            if not name.lower().endswith(".csv"):
                continue
            with zf.open(name, "r") as fbin:
                # try utf-8-sig; fallback to latin-1
                try:
                    text = io.TextIOWrapper(fbin, encoding="utf-8-sig", newline="")
                    reader = csv.reader(text)
                    yield name, reader
                except UnicodeDecodeError:
                    fbin.seek(0)
                    text = io.TextIOWrapper(fbin, encoding="latin-1", newline="")
                    reader = csv.reader(text)
                    yield name, reader

def main():
    gc = gs_client()
    ws = get_or_create_records_ws(gc)

    # If sheet empty, we’ll write header when we see it the first time
    existing_header = ws.row_values(1)
    header_in_sheet = bool(existing_header)
    if not existing_header:
        existing_header = []

    # We’ll collect existing keys after we know key column
    existing_keys = None
    key_col_idx = None
    sheet_width = len(existing_header) if existing_header else None

    total_new = 0

    for url in SCO_ZIPS:
        print(f"Processing: {url}", flush=True)
        zpath = download_to_temp(url)
        try:
            for entry_name, reader in csv_reader_from_zip(zpath):
                header = next(reader, None)
                if header is None:
                    continue

                # write header if sheet is blank
                if not header_in_sheet:
                    ws.update("A1", [header])
                    existing_header = header
                    sheet_width = len(header)
                    key_col_idx = find_key_col(header)
                    existing_keys = load_existing_keys(ws, key_col_idx)
                    header_in_sheet = True
                else:
                    # if header already present, align width + find key col if not set
                    if sheet_width is None:
                        sheet_width = len(existing_header)
                    if key_col_idx is None:
                        key_col_idx = find_key_col(existing_header)
                    if existing_keys is None:
                        existing_keys = load_existing_keys(ws, key_col_idx)

                batch = []
                for row in reader:
                    if not row:
                        continue
                    key = (row[key_col_idx] if key_col_idx < len(row) else "").strip()
                    # fallback: if there is somehow no Property ID, dedupe on full row
                    key_norm = key or json.dumps(row, ensure_ascii=False)
                    if key_norm not in existing_keys:
                        batch.append(row)
                        existing_keys.add(key_norm)

                    if len(batch) >= 5000:
                        append_in_batches(ws, batch, sheet_width)
                        total_new += len(batch)
                        batch.clear()

                if batch:
                    append_in_batches(ws, batch, sheet_width)
                    total_new += len(batch)
                    batch.clear()

        finally:
            try:
                os.remove(zpath)
            except OSError:
                pass

    print(f"Done. New rows added: {total_new}", flush=True)

if __name__ == "__main__":
    main()
