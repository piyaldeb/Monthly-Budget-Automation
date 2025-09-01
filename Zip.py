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
ODOO_URL   = os.getenv("ODOO_URL")
DB         = os.getenv("ODOO_DB")
USERNAME   = os.getenv("ODOO_USERNAME")
PASSWORD   = os.getenv("ODOO_PASSWORD")
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1f5pdh23Lxrxkdtm7vOeufxWXBvMR8HYIlRcucBZ994I")
SHEET_NAME = os.environ.get("SHEET_NAME", "Zip")
PASTE_COLUMNS = int(os.environ.get("PASTE_COLUMNS", "25"))

os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

# Separate XPaths for Date From and Export button (Date To is NOT used)
DATE_FROM_INPUT_XPATH = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[2]/div/div/input"

# Footer scope where we will search for the "Download Excel" button
FOOTER_BASE_XPATH = "/html/body/div[2]/div[2]/div/div/div/div/footer"

EXPORT_BTN_XPATH = (
    FOOTER_BASE_XPATH
    + "//button[normalize-space(text())='Download Excel' "
      "or normalize-space(span/text())='Download Excel' "
      "or contains(normalize-space(.), 'Download Excel')]"
)

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
# Download helpers
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
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": os.path.abspath(download_dir)
    })

def wait_for_download_complete(download_dir, start_time, timeout=180, quiet_gap=2):
    log(f"Waiting for download in {download_dir} (timeout {timeout}s)...")
    end = time.time() + timeout
    last_log = 0

    while time.time() < end:
        time.sleep(1)
        finished = _list_files(download_dir, ("*.xlsx","*.xls"))
        for f in finished:
            try:
                mt = os.path.getmtime(f)
                ct = os.path.getctime(f)
            except FileNotFoundError:
                continue
            if mt >= start_time or ct >= start_time:
                time.sleep(quiet_gap)
                log(f"Download completed: {os.path.basename(f)} ({_human_size(f)})")
                return f
        # heartbeat logging
        if time.time() - last_log > 10:
            temps = _list_files(download_dir)
            contents = [os.path.basename(x)+"["+_human_size(x)+"]" for x in temps]
            if contents:
                log("Downloads dir status: " + ", ".join(contents))
            else:
                log("Downloads dir empty; waiting for export.")
            last_log = time.time()
    log("Download timed out.")
    return None

def _norm_ddmmyy(s: str) -> str:
    s = s.strip()
    parts = s.split("/")
    if len(parts) == 3 and len(parts[2]) == 4:
        return f"{parts[0]}/{parts[1]}/{parts[2][-2:]}"
    return s

# -------------------------
# Resilient interaction helpers
# -------------------------
def safe_click(driver, wait, xpath, tries=3, sleep=0.8):
    for i in range(1, tries+1):
        el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            el.click()
            return True
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", el)
                return True
            except Exception:
                time.sleep(sleep)
    raise RuntimeError(f"Failed to click {xpath}")

def type_xpath(driver, wait, xpath, text, press_enter=True):
    el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    el.click()
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.BACKSPACE)
    el.send_keys(text)
    if press_enter:
        el.send_keys(Keys.ENTER)
    return el

# -------------------------
# Selenium: Download Report
# -------------------------
def download_from_odoo(company="Zipper", date_from="01/08/2025"):
    run_dir = os.path.join(BASE_DOWNLOAD_DIR, datetime.now().strftime("run_%Y%m%d_%H%M%S"))
    os.makedirs(run_dir, exist_ok=True)

    chromedriver_autoinstaller.install()
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath(run_dir),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    })
    options.add_argument("--log-level=3")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(90)
    driver.set_script_timeout(60)
    wait = WebDriverWait(driver, 40)

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

        log(f"Switching company to '{company}'...")
        switcher = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div.o_menu_systray div.o_switch_company_menu > button > span")
        ))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", switcher)
        switcher.click()
        time.sleep(1)
        target_div = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//div[contains(@class, 'log_into')][span[contains(text(), '{company}')]]")
        ))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_div)
        target_div.click()
        time.sleep(2)

        log("Navigating to MRP REPORTS...")
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys("MRP REPORTS")
        body.send_keys(Keys.ENTER)
        time.sleep(2)

        log("Selecting 'Invoice Summary'...")
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//select")))
        dropdown.click()
        dropdown.send_keys("Invoice Summary")
        dropdown.send_keys(Keys.ENTER)
        time.sleep(1)

        log("Setting From date...")
        df = _norm_ddmmyy(date_from)
        date_from_input = wait.until(EC.presence_of_element_located((By.XPATH, DATE_FROM_INPUT_XPATH)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", date_from_input)
        date_from_input.click()
        date_from_input.send_keys(Keys.CONTROL, "a")
        date_from_input.send_keys(Keys.BACKSPACE)
        date_from_input.send_keys(df)
        date_from_input.send_keys(Keys.ENTER)
        time.sleep(0.5)

        log("Clicking Download Excel...")
        footer_el = wait.until(EC.presence_of_element_located((By.XPATH, FOOTER_BASE_XPATH)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", footer_el)

        export_btn = wait.until(EC.presence_of_element_located((By.XPATH, EXPORT_BTN_XPATH)))

        for _ in range(10):
            if not export_btn.get_attribute("disabled"):
                break
            time.sleep(0.2)

        for i in range(1, 4):
            click_time = time.time()
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, EXPORT_BTN_XPATH)))
                export_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", export_btn)

            log(f"Download Excel click attempt {i}/3...")
            f = wait_for_download_complete(run_dir, start_time=click_time, timeout=60 if i < 3 else 180)
            if f:
                return f
            time.sleep(2)

        return None

    except Exception:
        log(f"Error during Odoo interaction:\n{traceback.format_exc()}")
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
# Main (fixed date/company)
# -------------------------
def main():
    date_from = "01/08/2025"
    log("Starting Zip run...")
    downloaded_file = download_from_odoo(company="Zipper", date_from=date_from)
    if downloaded_file:
        update_google_sheet_with_file(downloaded_file, sheet_name=SHEET_NAME)
        log("Zip sheet updated successfully.")
    else:
        log("Download failed; exiting with non-zero for CI.")
        raise SystemExit(1)

if __name__ == "__main__":
    main()
