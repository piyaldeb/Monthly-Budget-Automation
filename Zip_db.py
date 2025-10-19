import os
import re
import json
import time
import glob
import traceback
from datetime import datetime
import requests
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
# Load local .env file
load_dotenv()
# ===============================
# Config (env first, fallback)
# ===============================
BASE_DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", "downloads")
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

ODOO_URL   = os.getenv("ODOO_URL")
DB         = os.getenv("ODOO_DB")
USERNAME   = os.getenv("ODOO_USERNAME")
PASSWORD   = os.getenv("ODOO_PASSWORD")


# Wizard / report config
MODEL                 = os.environ.get("ODOO_WIZARD_MODEL", "mrp.report.custom")
REPORT_BUTTON_METHOD  = os.environ.get("ODOO_REPORT_BUTTON", "action_generate_xlsx_report")
REPORT_TYPE           = os.environ.get("ODOO_REPORT_TYPE",   "invs")      # 'pi' | 'pir' | 'r_invs'
COMPANY_ID            = int(os.environ.get("ODOO_COMPANY_ID", "1"))         # 3 = Metal (per your context)
TZ                    = os.environ.get("ODOO_TZ", "Asia/Dhaka")

# Dates (YYYY-MM-DD for Odoo)
from datetime import datetime, timedelta

# Calculate yesterday's date
yesterday = datetime.now() - timedelta(days=1)

# If yesterday is in a different month than today, use yesterday's month
# Otherwise, use current month
if yesterday.month != datetime.now().month:
    # Yesterday was last day of previous month
    DATE_FROM = yesterday.replace(day=1).strftime("%Y-%m-%d")
    DATE_TO = yesterday.strftime("%Y-%m-%d")
else:
    # Yesterday is in current month
    DATE_FROM = datetime.now().replace(day=1).strftime("%Y-%m-%d")
    DATE_TO = yesterday.strftime("%Y-%m-%d")

# You can still override with environment variables if needed
DATE_FROM = os.environ.get("DATE_FROM", DATE_FROM)
DATE_TO = os.environ.get("DATE_TO", DATE_TO)
# Google Sheets
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID       = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME           = os.environ.get("SHEET_NAME", "Zip")
PASTE_COLUMNS        = int(os.environ.get("PASTE_COLUMNS", "25"))  # A:I for 9 columns

# ===============================
# Utilities
# ===============================
def log(msg: str):
    print(f"{datetime.now()} {msg}", flush=True)

def newest_file(directory, patterns=("*.xlsx","*.xls")):
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(directory, p)))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

# ===============================
# Google Sheets helpers
# ===============================
def get_google_sheets_service_values():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account JSON not found: {SERVICE_ACCOUNT_FILE}")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds).spreadsheets().values()

def update_google_sheet_with_file(file_path: str, sheet_name: str, paste_cols: int):
    df = pd.read_excel(file_path, engine="openpyxl")
    # Numeric rounding on columns 2..paste_cols (index 1..paste_cols-1)
    for col in df.columns[1:paste_cols]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    # Replace NaNs with blank for Sheets
    df = df.where(pd.notnull(df), "")
    df_to_paste = df.iloc[:, 0:paste_cols]

    svc = get_google_sheets_service_values()
    
    # Get existing data from sheet
    try:
        result = svc.get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A:Z"
        ).execute()
        existing_values = result.get('values', [])
    except:
        existing_values = []
    
    if len(existing_values) <= 1:
        # Sheet is empty or only has headers - write with headers
        values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()
        svc.update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()
        log(f"✅ Google Sheet '{sheet_name}' initialized with {len(values)-1} rows.")
    else:
        # Sheet has data - merge with duplicate handling
        # Convert existing data to DataFrame (skip header row)
        headers = existing_values[0]
        existing_df = pd.DataFrame(existing_values[1:], columns=headers)
        
        # Assume first column is the unique identifier (product/item name)
        # Get new data without headers
        new_df = df_to_paste.copy()
        new_df.columns = df_to_paste.columns.tolist()
        
        # Identify the key column (usually first column - product/item name)
        key_col = df_to_paste.columns[0]
        
        # Remove rows from existing_df that match keys in new_df
        if key_col in existing_df.columns and len(existing_df) > 0:
            new_keys = set(new_df[key_col].astype(str))
            existing_df = existing_df[~existing_df[key_col].astype(str).isin(new_keys)]
        
        # Combine: existing (without duplicates) + new data
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        
        # Prepare final values with headers
        values = [headers] + combined_df.values.tolist()
        
        # Clear and write all data
        last_col_letter = chr(ord('A') + paste_cols - 1)
        svc.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:{last_col_letter}").execute()
        svc.update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()
        
        log(f"✅ Google Sheet '{sheet_name}' updated: {len(new_df)} new/updated rows, {len(values)-1} total rows.")

# ===============================
# Odoo (requests) flow
# ===============================
def odoo_login(session: requests.Session) -> int:
    url = f"{ODOO_URL}/web/session/authenticate"
    payload = {
        "jsonrpc": "2.0",
        "params": {"db": DB, "login": USERNAME, "password": PASSWORD},
    }
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    uid = r.json().get("result", {}).get("uid")
    if not uid:
        raise RuntimeError(f"Login failed: {r.text[:500]}")
    log(f"✅ Logged in, uid={uid}")
    return uid

def get_csrf_token(session: requests.Session) -> str:
    # Not always needed for GET download, but we fetch it for the POST fallback route
    r = session.get(f"{ODOO_URL}/web", timeout=60)
    r.raise_for_status()
    m = re.search(r'csrf_token:\s*"([A-Za-z0-9]+)"', r.text)
    token = m.group(1) if m else ""
    log(f"ℹ️ CSRF token: {token or '(none found)'}")
    return token

def wizard_create(session: requests.Session, uid: int) -> int:
    url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/create"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "create",
            "args": [{}],
            "kwargs": {"context": {"uid": uid}},
        },
    }
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    wiz_id = r.json().get("result")
    if not wiz_id:
        raise RuntimeError(f"Wizard create failed: {r.text[:500]}")
    log(f"✅ Wizard created: id={wiz_id}")
    return wiz_id

def wizard_save(session: requests.Session, uid: int, report_type: str, date_from: str, date_to: str) -> int:
    url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "web_save",
            "args": [[], {
                "report_type": report_type,
                "date_from": date_from,
                "date_to": date_to
            }],
            "kwargs": {
                "context": {
                    "lang": "en_US",
                    "tz": TZ,
                    "uid": uid,
                    "allowed_company_ids": [COMPANY_ID],
                },
                "specification": {"report_type": {}, "date_from": {}, "date_to": {}},
            },
        },
    }
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    result = r.json().get("result") or []
    wiz_id = result[0].get("id") if result and isinstance(result, list) else None
    if not wiz_id:
        raise RuntimeError(f"web_save failed: {r.text[:500]}")
    log(f"✅ Wizard saved: id={wiz_id} ({report_type} {date_from}→{date_to})")
    return wiz_id

def press_report_button(session: requests.Session, uid: int, wiz_id: int) -> dict:
    url = f"{ODOO_URL}/web/dataset/call_button"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": REPORT_BUTTON_METHOD,
            "args": [[wiz_id]],
            "kwargs": {
                "context": {
                    "lang": "en_US",
                    "tz": TZ,
                    "uid": uid,
                    "allowed_company_ids": [COMPANY_ID],
                }
            },
        },
    }
    r = session.post(url, json=payload, timeout=60)
    r.raise_for_status()
    result = r.json().get("result") or {}
    # Expect an ir.actions.report dict
    if not result or result.get("type") not in ("ir.actions.report", "ir_actions_report"):
        log(f"⚠️ Unexpected button result (printing first 600 chars): {json.dumps(result, indent=2)[:600]}")
    else:
        log(f"✅ Button returned report action: {result.get('report_name')} ({result.get('report_type')})")
    return result

def build_report_paths(report_info: dict, date_from: str, date_to: str, uid: int, wiz_id: int):
    """
    Returns (relative_report_path, options, context) for xlsx download.
    We prefer direct GET to /report/xlsx/<report_name>?options=...&context=...
    """
    report_name = report_info.get("report_name")
    if not report_name:
        # If server didn't include it, you must set a sensible default/mapping per report_type
        # Fallback (update to the correct template on your server if needed):
        fallback = {
            # "pi":   "taps_manufacturing.pi_xls_template",
            # "pir":  "taps_manufacturing.packing_invoice_summery_separated",
            "invs":"taps_manufacturing.packing_invoice_summary",
        }
        report_name = fallback.get(REPORT_TYPE)
        if not report_name:
            raise RuntimeError("No report_name in action and no fallback mapping available. Please set one.")

    options = {"date_from": date_from, "date_to": date_to, "company_id": COMPANY_ID}
    context = {
        "lang": "en_US",
        "tz": TZ,
        "uid": uid,
        "allowed_company_ids": [COMPANY_ID],
        "active_model": MODEL,
        "active_id": wiz_id,
        "active_ids": [wiz_id],
    }
    rel = f"/report/xlsx/{report_name}?options={json.dumps(options)}&context={json.dumps(context)}"
    return rel, options, context

def try_direct_get_xlsx(session: requests.Session, relative_path: str, out_path: str, timeout: int = 300) -> bool:
    """
    Try to download via direct GET with increased timeout and progressive waiting.
    """
    url = f"{ODOO_URL}{relative_path}"
    
    # Try with longer timeout and stream=True for large files
    try:
        r = session.get(url, stream=True, timeout=timeout)
        
        if r.status_code == 200 and r.headers.get("content-type", "").startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            with open(out_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            return True
        
        log(f"Direct GET failed: {r.status_code} {r.headers.get('content-type','')}")
        if r.status_code == 502:
            log(f"⚠️ 502 Bad Gateway - Server timeout generating report")
        return False
        
    except requests.exceptions.Timeout:
        log(f"⏱️ Request timed out after {timeout} seconds")
        return False
    except Exception as e:
        log(f"❌ GET request error: {e}")
        return False

def try_report_download_post(session: requests.Session, csrf_token: str, relative_path: str, out_path: str, timeout: int = 300) -> bool:
    """
    Fallback to /report/download with form-encoded fields, with increased timeout.
    """
    url = f"{ODOO_URL}/report/download"
    data = {
        "data": json.dumps([relative_path, "xlsx"]),
        "context": "{}",
        "token": "filetoken-123",
        "csrf_token": csrf_token or "",
    }
    headers = {"Referer": f"{ODOO_URL}/web"}
    
    try:
        r = session.post(url, data=data, headers=headers, timeout=timeout)
        
        if r.status_code == 200 and r.headers.get("content-type", "").startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            with open(out_path, "wb") as f:
                f.write(r.content)
            return True
        
        log(f"/report/download failed: {r.status_code}")
        if r.status_code == 502:
            log(f"⚠️ 502 Bad Gateway - Server timeout generating report")
        return False
        
    except requests.exceptions.Timeout:
        log(f"⏱️ POST request timed out after {timeout} seconds")
        return False
    except Exception as e:
        log(f"❌ POST request error: {e}")
        return False

def download_from_odoo(date_from: str, date_to: str) -> str:
    """
    Full flow → returns the path of the downloaded XLSX file.
    """
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

    uid = odoo_login(session)
    csrf = get_csrf_token(session)

    wiz_id = wizard_create(session, uid)
    wiz_id = wizard_save(session, uid, REPORT_TYPE, date_from, date_to)
    report_info = press_report_button(session, uid, wiz_id)

    # Give the server time to start generating the report
    log("⏳ Waiting 10 seconds for server to prepare report...")
    time.sleep(10)

    rel_path, _options, _context = build_report_paths(report_info, date_from, date_to, uid, wiz_id)
    out_name = f"{REPORT_TYPE}_{date_from}_to_{date_to}.xlsx".replace(":", "-")
    out_path = os.path.join(BASE_DOWNLOAD_DIR, out_name)

    # Try with progressively longer timeouts
    timeouts = [180, 300, 420]  # 3 min, 5 min, 7 min
    
    for attempt, timeout in enumerate(timeouts, 1):
        log(f"⬇️  Attempt {attempt}/{len(timeouts)}: Direct GET download (timeout: {timeout}s)...")
        if try_direct_get_xlsx(session, rel_path, out_path, timeout):
            log(f"✅ Downloaded via direct GET → {out_path}")
            return out_path
        
        if attempt < len(timeouts):
            log(f"⏳ Waiting 15 seconds before next attempt...")
            time.sleep(15)

    # Final fallback to POST
    log("↩️  Final attempt: /report/download POST (timeout: 420s)...")
    if try_report_download_post(session, csrf, rel_path, out_path, 420):
        log(f"✅ Downloaded via /report/download → {out_path}")
        return out_path

    raise RuntimeError("Download failed: Server consistently returns 502 Bad Gateway. The report may be too large or complex for the server to generate within timeout limits.")

# ===============================
# Main with retry
# ===============================
from datetime import datetime, timedelta

def get_date_chunks(start_date_str: str, end_date_str: str, chunk_days: int = 7):
    """Split date range into smaller chunks to avoid server timeout"""
    start = datetime.strptime(start_date_str, "%Y-%m-%d")
    end = datetime.strptime(end_date_str, "%Y-%m-%d")
    
    chunks = []
    current = start
    
    while current <= end:
        chunk_end = min(current + timedelta(days=chunk_days - 1), end)
        chunks.append((
            current.strftime("%Y-%m-%d"),
            chunk_end.strftime("%Y-%m-%d")
        ))
        current = chunk_end + timedelta(days=1)
    
    return chunks

def main():
    max_retries = 3
    
    # Split the date range into weekly chunks
    date_chunks = get_date_chunks(DATE_FROM, DATE_TO, chunk_days=7)
    log(f"📅 Split date range into {len(date_chunks)} chunks: {DATE_FROM} → {DATE_TO}")
    
    all_dataframes = []
    
    for chunk_idx, (chunk_from, chunk_to) in enumerate(date_chunks, 1):
        log(f"\n{'='*60}")
        log(f"📦 Processing chunk {chunk_idx}/{len(date_chunks)}: {chunk_from} → {chunk_to}")
        log(f"{'='*60}")
        
        for attempt in range(1, max_retries + 1):
            try:
                log(f"▶️ Attempt {attempt}/{max_retries} for chunk {chunk_idx}...")
                
                xlsx_path = download_from_odoo(chunk_from, chunk_to)
                
                # Read the Excel file into DataFrame
                df = pd.read_excel(xlsx_path, engine="openpyxl")
                all_dataframes.append(df)
                
                # Clean up the chunk file
                if os.path.exists(xlsx_path):
                    os.remove(xlsx_path)
                
                log(f"✅ Chunk {chunk_idx} completed ({len(df)} rows)")
                break  # Success, move to next chunk
                
            except Exception as e:
                log(f"❌ Error on attempt {attempt} for chunk {chunk_idx}: {e}")
                if attempt < max_retries:
                    wait_time = 30 * (2 ** (attempt - 1))
                    log(f"⏳ Retrying chunk {chunk_idx} in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    log(f"🚨 Chunk {chunk_idx} failed after {max_retries} attempts")
                    raise SystemExit(1)
    
    # Combine all chunks into one DataFrame
    if not all_dataframes:
        log("❌ No data collected from any chunks")
        raise SystemExit(1)
    
    log(f"\n{'='*60}")
    log(f"📊 Combining {len(all_dataframes)} chunks...")
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    
    # Remove duplicates if any (based on first column)
    first_col = combined_df.columns[0]
    before_dedup = len(combined_df)
    combined_df = combined_df.drop_duplicates(subset=[first_col], keep='last')
    after_dedup = len(combined_df)
    
    if before_dedup > after_dedup:
        log(f"🔄 Removed {before_dedup - after_dedup} duplicate rows")
    
    # Save combined data temporarily
    temp_file = os.path.join(BASE_DOWNLOAD_DIR, f"combined_{DATE_FROM}_to_{DATE_TO}.xlsx")
    combined_df.to_excel(temp_file, index=False, engine="openpyxl")
    log(f"💾 Combined file saved: {temp_file} ({len(combined_df)} rows)")
    
    # Upload to Google Sheets
    log(f"☁️ Uploading to Google Sheets...")
    update_google_sheet_with_file(temp_file, SHEET_NAME, PASTE_COLUMNS)
    
    # Clean up combined file
    if os.path.exists(temp_file):
        os.remove(temp_file)
    
    log("🎉 Done. Zipper")