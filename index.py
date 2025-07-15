import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import chromedriver_autoinstaller
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configuration
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
FILE_NAME = "Packing and Invoice Summary.xlsx"
DOWNLOADED_FILE = os.path.join(DOWNLOAD_DIR, FILE_NAME)
ODOO_URL = "https://taps.odoo.com"
ODOO_USERNAME = "ranak@texzipperbd.com"
ODOO_PASSWORD = "2326"

# Google Sheets API setup
SERVICE_ACCOUNT_FILE = 'path_to_your_service_account_credentials.json'  # Path to your service account file
SPREADSHEET_ID = '13NDGH7zTB0gAp9lWnHYeegWase-OoN2Cg5CG63U_iyM'  # Your Google Sheets ID
RANGE_NAME = 'Sheet1!A1:Z1000'  # Update with the desired range in your sheet

# Ensure download directory exists
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Function to authenticate and get the Google Sheets service
def get_google_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    service = build("sheets", "v4", credentials=creds)
    return service.spreadsheets()

# Function to wait for download to complete
def wait_for_download_complete(download_dir, file_name):
    while not os.path.exists(file_name):
        time.sleep(1)

# Function to download file from Odoo
def download_from_odoo(company="Zipper"):
    if os.path.exists(DOWNLOADED_FILE):
        try:
            os.remove(DOWNLOADED_FILE)
            print(f"🧹 Deleted existing local file: {DOWNLOADED_FILE}")
        except Exception as e:
            print(f"⚠️ Could not delete existing file: {e}")

    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "directory_upgrade": True
    })

    chromedriver_autoinstaller.install()

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    try:
        driver.get(ODOO_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys(ODOO_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(ODOO_PASSWORD)
        driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]").click()
        time.sleep(4)

        # Switch company
        switcher = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.o_menu_systray div.o_switch_company_menu > button > span")))
        driver.execute_script("arguments[0].scrollIntoView(true);", switcher)
        switcher.click()
        time.sleep(2)

        # Select company
        target_div = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class, 'log_into')][span[contains(text(), '{company}')]]")))
        target_div.click()
        time.sleep(2)

        # Navigate to MRP Reports
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys("MRP REPORTS")
        body.send_keys(Keys.ENTER)
        time.sleep(5)

        # Select invoice report type
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[1]/div[2]/div/select")))
        dropdown.click()
        dropdown.send_keys("Invoice")
        dropdown.send_keys(Keys.ENTER)
        time.sleep(2)

        # Set the date filter (assuming you need to use today's date for simplicity)
        date_input = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[2]/div/div/input")))
        date_input.clear()
        date_input.send_keys(datetime.today().strftime("%d/%m/%Y"))
        date_input.send_keys(Keys.ENTER)
        time.sleep(1)

        # Click export button to download the file
        export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[1]")))
        export_btn.click()

        # Wait for download to complete
        wait_for_download_complete(DOWNLOAD_DIR, FILE_NAME)

    except Exception as e:
        print("❌ Error during Odoo interaction:", e)
    finally:
        driver.quit()

# Function to update Google Sheets with the extracted values
def update_google_sheet(data):
    try:
        service = get_google_sheets_service()

        # Prepare the data to be updated
        values = [
            ["Date", "Value"],  # Add headers if necessary
            [data["date"], data["value"]]  # Add the extracted data (from C27, C10, etc.)
        ]

        body = {
            "values": values
        }

        # Update the sheet with the values
        service.values().update(
            spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME,
            valueInputOption="RAW", body=body
        ).execute()
        print("✅ Google Sheet updated successfully.")

    except HttpError as err:
        print(f"❌ Error updating Google Sheets: {err}")

def main():
    """Main function to download and process data."""
    # Download for Zipper company
    download_from_odoo(company="Zipper")

    # Example data that might be extracted from your Odoo file
    extracted_data = {
        "date": "01/07/2025",  # Example extracted date
        "value": "1000"  # Example extracted value (like C27)
    }

    # Update Google Sheet with extracted data
    update_google_sheet(extracted_data)

    # Download for Metal Trims company
    download_from_odoo(company="Metal Trims")

    # Example data for Metal Trims
    extracted_data = {
        "date": "02/07/2025",  # Example extracted date for Metal Trims
        "value": "1500"  # Example value for Metal Trims
    }

    # Update Google Sheet with Metal Trims data
    update_google_sheet(extracted_data)

if __name__ == "__main__":
    main()
