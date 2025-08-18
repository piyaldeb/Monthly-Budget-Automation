import os
import time
import glob
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

# -------------------------
# Config (logic unchanged)
# -------------------------
BASE_DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", "downloads")   # base folder
ODOO_URL = "https://taps.odoo.com"
ODOO_USERNAME = os.environ.get("ODOO_USERNAME", "ranak@texzipperbd.com")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "2326")

SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME = os.environ.get("SHEET_NAME", "Zip")
PASTE_COLUMNS = int(os.environ.get("PASTE_COLUMNS", "25"))

os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

# Keep your exact locators/flow
DATE_INPUT_XPATH = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[2]/div/div/input"
EXPORT_BTN_XPATH = "/html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[1]"

def log(msg: str):
    print(f"{datetime.now()} {msg}", flush=True)

# -------------------------
# Google Sheets (unchanged)
# -------------------------
def get_google_sheets_service():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account JSON not found: {SERVICE_ACCOUNT_FILE}")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds).spreadsheets().values()

# -------------------------
# Download helpers (CI only)
# -------------------------
def _list_files(d, patterns=("*.xlsx","*.xls","*.crdownload")):
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(d, p)))
    return sorted(files, key=os.path.getmtime)

def _human_size(path):
    try:
        return f"{os.path.getsize(path)}B"
    except Exception:
        return "?"

def allow_headless_downloads(driver, download_dir):
    # point headless Chrome to our unique run folder
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": os.path.abspath(download_dir)
    })

def wait_for_download_complete(download_dir, start_time, timeout=180, quiet_gap=2):
    """
    Wait for a new or updated .xlsx/.xls in a *fresh* run folder.
    Accept file if ctime/mtime >= export click time.
    Show size growth logs for visibility in CI.
    """
    log(f"Waiting for download in {download_dir} (timeout {timeout}s)...")
    end = time.time() + timeout
    last_log = 0
    seen_sizes = {}

    while time.time() < end:
        time.sleep(1)
        finished = _list_files(download_dir, ("*.xlsx","*.xls"))
        # Accept any finished file with mtime >= click time
        for f in finished:
            try:
                mt = os.path.getmtime(f)
                ct = os.path.getctime(f)
            except FileNotFoundError:
                continue
            if mt >= start_time or ct >= start_time:
                time.sleep(quiet_gap)  # flush
                log(f"Download completed: {os.path.basename(f)} ({_human_size(f)})")
                return f

        # heartbeat + size growth for .crdownload
        temps = _list_files(download_dir, ("*.crdownload",))
        if temps:
            # log every ~10s
            if time.time() - last_log > 10:
                parts = []
                for t in temps:
                    sz = _human_size(t)
                    parts.append(f"{os.path.basename(t)} [{sz}]")
                log("Still downloading... temp: " + ", ".join(parts))
                last_log = time.time()
        else:
            if time.time() - last_log > 10:
                contents = [os.path.basename(x)+"["+_human_size(x)+"]" for x in _list_files(download_dir)]
                if contents:
                    log("Downloads dir status: " + ", ".join(contents))
                else:
                    log("Downloads dir is empty; waiting for export to start.")
                last_log = time.time()

    log("Download timed out.")
    return None

# -------------------------
# Selenium: Download Report (UI logic preserved)
# -------------------------
def download_from_odoo(company="Zipper", date_from="01/01/2025", date_to=None):
    if not date_to:
        date_to = (datetime.today() - timedelta(days=1)).strftime("%m/%d/%Y")

    # Create a unique folder per run to avoid same-name collisions
    run_dir = os.path.join(BASE_DOWNLOAD_DIR, datetime.now().strftime("run_%Y%m%d_%H%M%S"))
    os.makedirs(run_dir, exist_ok=True)

    chromedriver_autoinstaller.install()
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath(run_dir),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_argument("--log-level=3")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(90)
    driver.set_script_timeout(60)
    wait = WebDriverWait(driver, 40)

    # Make sure headless actually writes to our run_dir
    try:
        allow_headless_downloads(driver, run_dir)
        log("Headless downloads enabled via CDP.")
    except Exception as e:
        log(f"CDP download permission failed (continuing): {e}")

    try:
        log("Opening Odoo login page...")
        driver.get(ODOO_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys(ODOO_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(ODOO_PASSWORD)
        driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]").click()
        time.sleep(2)

        # Switch company (same selectors)
        log(f"Switching company to '{company}'...")
        switcher = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div.o_menu_systray div.o_switch_company_menu > button > span")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", switcher)
        switcher.click()
        time.sleep(1)
        target_div = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//div[contains(@class, 'log_into')][span[contains(text(), '{company}')]]")
        ))
        target_div.click()
        time.sleep(2)

        # Navigate to MRP Reports (unchanged)
        log("Navigating to MRP REPORTS...")
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys("MRP REPORTS")
        body.send_keys(Keys.ENTER)
        time.sleep(2)

        # Select "Invoice Summary" (unchanged)
        log("Selecting 'Invoice Summary'...")
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//select")))
        dropdown.click()
        dropdown.send_keys("Invoice Summary")
        dropdown.send_keys(Keys.ENTER)
        time.sleep(1)

        # Set date (keep your original)
        log("Setting date (as in original logic)...")
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, DATE_INPUT_XPATH)))
        date_input.clear()
        date_input.send_keys("01/08/25")
        date_input.send_keys(Keys.ENTER)
        time.sleep(1)

        # Export (unchanged) — record click time and wait
        log("Clicking export...")
        export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, EXPORT_BTN_XPATH)))

        for i in range(1, 4):  # up to 3 tries
            click_time = time.time()
            driver.execute_script("arguments[0].click();", export_btn)
            log(f"Export click attempt {i}/3...")
            f = wait_for_download_complete(run_dir, start_time=click_time, timeout=60 if i < 3 else 180)
            if f:
                return f
            time.sleep(2)

        return None

    except Exception:
        log(f"Error during Odoo interaction:\n{traceback.format_exc()}")
        # Optional: dump artifacts for debugging
        try:
            driver.save_screenshot(os.path.join(run_dir, "zip_error.png"))
            with open(os.path.join(run_dir, "zip_error.html"), "w", encoding="utf-8") as fp:
                fp.write(driver.page_source)
        except Exception:
            pass
        return None
    finally:
        driver.quit()
        log("Browser closed")

# -------------------------
# Update Google Sheet (unchanged logic)
# -------------------------
def update_google_sheet_with_file(file_path, sheet_name):
    df = pd.read_excel(file_path, engine="openpyxl")
    for col in df.columns[1:PASTE_COLUMNS]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    df = df.where(pd.notnull(df), "")
    df_to_paste = df.iloc[:, 0:PASTE_COLUMNS]
    values = [df_to_paste.columns.tolist()] + df_to_paste.values.tolist()
    service = get_google_sheets_service()
    service.clear(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A:Y").execute()
    service.update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()
    if os.path.exists(file_path):
        os.remove(file_path)

# -------------------------
# Main (unchanged dates/company)
# -------------------------
def main():
    date_from = "01/08/2025"
    date_to = "31/08/2025"
    log("Starting Zip run...")
    downloaded_file = download_from_odoo(company="Zipper", date_from=date_from, date_to=date_to)
    if downloaded_file:
        update_google_sheet_with_file(downloaded_file, sheet_name=SHEET_NAME)
        log("Zip sheet updated successfully.")
    else:
        log("Download failed; exiting with non-zero for CI.")
        raise SystemExit(1)

if __name__ == "__main__":
    main()
