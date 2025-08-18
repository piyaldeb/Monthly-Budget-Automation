import os
import time
import glob
import json
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import chromedriver_autoinstaller
from google.oauth2 import service_account
from googleapiclient.discovery import build

# =========================
# Configuration
# =========================
DOWNLOAD_DIR = "downloads"
ARTIFACTS_DIR = "artifacts"  # debug artifacts (screens, html, logs)
ODOO_URL = "https://taps.odoo.com"
ODOO_USERNAME = "ranak@texzipperbd.com"
ODOO_PASSWORD = "2326"

SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = '1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I'
SHEET_NAME = "Mt"
PASTE_COLUMNS = 9  # A-I

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(ARTIFACTS_DIR, exist_ok=True)

# =========================
# Locators
# =========================
DATE_FROM_INPUT_XPATH = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[2]/div/div/input"
DATE_TO_INPUT_XPATH   = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[3]/div[2]/div/div/input"
EXPORT_BTN_XPATH      = "/html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[1]"

# =========================
# Debug helpers
# =========================
def hlog(run_dir, level, msg):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{stamp} [{level}] {msg}"
    print(line, flush=True)
    with open(os.path.join(run_dir, "debug.log"), "a", encoding="utf-8") as f:
        f.write(line + "\n")

def dump_console_logs(driver, run_dir, label):
    try:
        for log_type in ("browser", "performance"):
            try:
                entries = driver.get_log(log_type)
            except Exception:
                entries = []
            if entries:
                p = os.path.join(run_dir, f"{label}_console_{log_type}.json")
                with open(p, "w", encoding="utf-8") as fp:
                    json.dump(entries, fp, indent=2)
    except Exception:
        pass

def dump_page(driver, run_dir, label):
    try:
        sshot = os.path.join(run_dir, f"{label}.png")
        htmlp = os.path.join(run_dir, f"{label}.html")
        driver.save_screenshot(sshot)
        with open(htmlp, "w", encoding="utf-8") as fp:
            fp.write(driver.page_source)
        dump_console_logs(driver, run_dir, label)
    except Exception:
        pass

def ready_state(driver):
    try:
        return driver.execute_script("return document.readyState")
    except Exception:
        return "unknown"

def safe_click(driver, elem, run_dir, label):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(0.2)
        elem.click()
        return True
    except Exception as e1:
        hlog(run_dir, "WARN", f"{label}: normal click failed: {e1}; trying JS click")
        try:
            driver.execute_script("arguments[0].click();", elem)
            return True
        except Exception as e2:
            hlog(run_dir, "ERR", f"{label}: JS click failed: {e2}")
            return False

# =========================
# Google Sheets
# =========================
def get_google_sheets_service(run_dir):
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account JSON not found: {SERVICE_ACCOUNT_FILE}")
    hlog(run_dir, "STEP", "Authenticating Google Sheets")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    service = build("sheets", "v4", credentials=creds).spreadsheets().values()
    hlog(run_dir, "OK", "Google Sheets authentication done")
    return service

# =========================
# Download wait (with heartbeat + timeout)
# =========================
def wait_for_download_complete(download_dir, run_dir, timeout=240, heartbeat=10):
    start = time.time()
    hlog(run_dir, "STEP", f"Waiting for XLSX (timeout={timeout}s)")
    initial = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
    last_beat = 0
    while True:
        time.sleep(1)
        current = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
        newfiles = current - initial
        if newfiles:
            latest = max(newfiles, key=os.path.getctime)
            time.sleep(5)  # let it finalize write
            hlog(run_dir, "OK", f"Download detected: {latest}")
            return latest
        # heartbeat
        if time.time() - last_beat >= heartbeat:
            cr = glob.glob(os.path.join(download_dir, "*.crdownload"))
            xl = glob.glob(os.path.join(download_dir, "*.xlsx"))
            hlog(run_dir, "INFO", f"Heartbeat: {len(cr)} .crdownload, {len(xl)} .xlsx in dir")
            last_beat = time.time()
        if time.time() - start > timeout:
            # final listing
            all_files = glob.glob(os.path.join(download_dir, "*"))
            listing = [os.path.basename(f) for f in all_files]
            hlog(run_dir, "ERR", f"Download timeout. Dir listing: {listing}")
            return None

# =========================
# Helpers
# =========================
def _norm_ddmmyy(s: str) -> str:
    """Accept 'DD/MM/YYYY' or 'DD/MM/YY' and return 'DD/MM/YY'."""
    s = s.strip()
    parts = s.split("/")
    if len(parts) == 3 and len(parts[2]) == 4:
        return f"{parts[0]}/{parts[1]}/{parts[2][-2:]}"
    return s

def _allow_headless_downloads(driver, download_dir, run_dir):
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": os.path.abspath(download_dir)
        })
        hlog(run_dir, "OK", "Headless download path set via CDP")
    except Exception as e:
        hlog(run_dir, "WARN", f"Unable to set download behavior: {e}")

# =========================
# Selenium: Download Report
# =========================
def download_from_odoo(company="Metal", date_from="01/08/2025", date_to="31/08/2025"):
    run_dir = os.path.join(ARTIFACTS_DIR, datetime.now().strftime("run_%Y%m%d_%H%M%S"))
    os.makedirs(run_dir, exist_ok=True)
    hlog(run_dir, "STEP", f"Run dir: {run_dir}")

    # Chrome setup with logs
    chromedriver_autoinstaller.install()
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-features=InfiniteSessionRestore,TranslateUI")
    options.add_argument("--log-level=0")
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    # capture console + perf logs
    options.set_capability("goog:loggingPrefs", {"performance": "ALL", "browser": "ALL"})

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(120)
    driver.set_script_timeout(120)
    wait = WebDriverWait(driver, 45)

    _allow_headless_downloads(driver, DOWNLOAD_DIR, run_dir)

    try:
        # 1) Open login
        hlog(run_dir, "STEP", "1. Opening Odoo")
        driver.get(ODOO_URL)
        hlog(run_dir, "INFO", f"URL={driver.current_url}")
        hlog(run_dir, "INFO", f"readyState={ready_state(driver)} title={driver.title}")
        dump_page(driver, run_dir, "01_open")

        # 2) Login
        hlog(run_dir, "STEP", "2. Logging in")
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys(ODOO_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(ODOO_PASSWORD)
        login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Log in')]")))
        safe_click(driver, login_btn, run_dir, "login_button")
        time.sleep(3)
        hlog(run_dir, "INFO", f"Post-login URL={driver.current_url} readyState={ready_state(driver)}")
        dump_page(driver, run_dir, "02_after_login")

        # 3) Switch company
        hlog(run_dir, "STEP", f"3. Switching company -> {company}")
        switcher = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div.o_menu_systray div.o_switch_company_menu > button > span")
        ))
        safe_click(driver, switcher, run_dir, "company_switcher")
        time.sleep(1)
        target_div = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//div[contains(@class, 'log_into')][span[contains(text(), '{company}')]]")
        ))
        safe_click(driver, target_div, run_dir, "company_target")
        time.sleep(3)
        hlog(run_dir, "INFO", f"After company switch URL={driver.current_url} readyState={ready_state(driver)}")
        dump_page(driver, run_dir, "03_after_company_switch")

        # 4) Navigate to MRP REPORTS
        hlog(run_dir, "STEP", "4. Navigating to MRP REPORTS")
        body = wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        body.send_keys("MRP REPORTS")
        body.send_keys(Keys.ENTER)
        time.sleep(4)
        hlog(run_dir, "INFO", f"After nav URL={driver.current_url} readyState={ready_state(driver)}")
        dump_page(driver, run_dir, "04_after_nav")

        # 5) Select report
        hlog(run_dir, "STEP", "5. Selecting 'Invoice Summary'")
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//select")))
        safe_click(driver, dropdown, run_dir, "report_dropdown")
        dropdown.send_keys("Invoice Summary")
        dropdown.send_keys(Keys.ENTER)
        time.sleep(2)
        dump_page(driver, run_dir, "05_after_select_report")

        # 6) Set dates
        df = _norm_ddmmyy(date_from)
        dt = _norm_ddmmyy(date_to)
        hlog(run_dir, "STEP", f"6. Setting dates: FROM={df} TO={dt}")

        date_from_input = wait.until(EC.presence_of_element_located((By.XPATH, DATE_FROM_INPUT_XPATH)))
        date_from_input.clear(); time.sleep(0.1)
        date_from_input.send_keys(df); date_from_input.send_keys(Keys.ENTER)
        time.sleep(1)

        date_to_input = wait.until(EC.presence_of_element_located((By.XPATH, DATE_TO_INPUT_XPATH)))
        date_to_input.clear(); time.sleep(0.1)
        date_to_input.send_keys(dt); date_to_input.send_keys(Keys.ENTER)
        time.sleep(1)
        dump_page(driver, run_dir, "06_after_dates")

        # 7) Export
        hlog(run_dir, "STEP", "7. Clicking export")
        export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, EXPORT_BTN_XPATH)))
        if not safe_click(driver, export_btn, run_dir, "export_button"):
            hlog(run_dir, "ERR", "Export button click failed")
            dump_page(driver, run_dir, "07_export_click_failed")
            return None

        # 8) Wait for file
        hlog(run_dir, "STEP", "8. Waiting for download")
        fpath = wait_for_download_complete(DOWNLOAD_DIR, run_dir, timeout=300, heartbeat=10)
        dump_page(driver, run_dir, "08_after_download_attempt")
        return fpath

    except Exception:
        hlog(run_dir, "ERR", f"Exception:\n{traceback.format_exc()}")
        dump_page(driver, run_dir, "zz_exception")
        return None
    finally:
        try:
            dump_console_logs(driver, run_dir, "final")
        except Exception:
            pass
        driver.quit()
        hlog(run_dir, "OK", "Browser closed")

# =========================
# Update Google Sheet
# =========================
def update_google_sheet_with_file(file_path, sheet_name, run_dir):
    try:
        hlog(run_dir, "STEP", f"Reading Excel: {file_path}")
        df = pd.read_excel(file_path, engine="openpyxl")

        hlog(run_dir, "STEP", "Processing data")
        for col in df.columns[1:PASTE_COLUMNS]:
            df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
        df = df.where(pd.notnull(df), "")
        df_to_paste = df.iloc[:, 0:PASTE_COLUMNS]
        values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()

        service = get_google_sheets_service(run_dir)
        hlog(run_dir, "STEP", "Clearing existing sheet range A:I")
        service.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:I").execute()
        hlog(run_dir, "STEP", "Updating Google Sheet with new data")
        service.update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()
        hlog(run_dir, "OK", f"Google Sheet '{sheet_name}' updated")

        if os.path.exists(file_path):
            os.remove(file_path)
            hlog(run_dir, "OK", f"Deleted local file: {file_path}")

    except Exception:
        hlog(run_dir, "ERR", f"Error updating Google Sheets:\n{traceback.format_exc()}")

# =========================
# Main
# =========================
def main():
    date_from = "01/08/2025"
    date_to   = "31/08/2025"  # fixed end date
    run_dir = os.path.join(ARTIFACTS_DIR, datetime.now().strftime("run_%Y%m%d_%H%M%S_main"))
    os.makedirs(run_dir, exist_ok=True)

    hlog(run_dir, "STEP", "Starting Odoo -> Sheets process")
    downloaded_file = download_from_odoo(company="Metal", date_from=date_from, date_to=date_to)
    if downloaded_file:
        update_google_sheet_with_file(downloaded_file, sheet_name=SHEET_NAME, run_dir=run_dir)
    else:
        hlog(run_dir, "ERR", "Download failed, nothing to update")

if __name__ == "__main__":
    main()
