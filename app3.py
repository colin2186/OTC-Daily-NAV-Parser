import os
import time
import glob
import tempfile
from datetime import datetime, timedelta
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait as wdw
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# Constants
CLIENT = "ASGARD"
DOWNLOAD_TIMEOUT = 30  # seconds


def read_credentials():
    """Read credentials from Streamlit secrets or fallback to local files"""
    try:
        return st.secrets["gopricing"]["username"], st.secrets["gopricing"]["password"]
    except:
        # Fallback for local testing (remove in production)
        try:
            with open('unicorn.txt') as user_file, open('pineapple.txt') as pw_file:
                return user_file.read().strip(), pw_file.read().strip()
        except FileNotFoundError:
            st.error("Credentials not found. Please configure Streamlit secrets.")
            st.stop()


def setup_driver():
    """Configure ChromeDriver for cloud or local use"""
    chrome_options = Options()

    # Cloud-optimized configuration
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    # Set download directory to temp folder
    prefs = {
        "download.default_directory": tempfile.gettempdir(),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    try:
        # Try using webdriver_manager first
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        st.error(f"Failed to initialize ChromeDriver: {str(e)}")
        st.stop()


def is_business_day(date_to_check):
    """Check if the date is a business day (weekday)"""
    return date_to_check.weekday() < 5  # Monday=0, Sunday=6


def download_report(driver, target_date_str):
    """Handle the report download process"""
    username, password = read_credentials()
    url = f"https://{username}:{password}@gopricing.ssnc-corp.cloud/clients/reports/otcReportSummary.go?action=handleSummaryFilter"

    try:
        driver.get(url)
        driver.maximize_window()
        time.sleep(2)

        # Wait for and interact with date selector
        wdw(driver, 15).until(ec.presence_of_element_located((By.XPATH, '//*[@id="valuationDate"]')))
        date_selector = driver.find_element(By.XPATH, '//*[@id="valuationDate"]')
        date_selector.clear()
        date_selector.send_keys(target_date_str)

        # Client search interaction
        client_search = driver.find_element(By.XPATH, "//*[@id='clientSearch']")
        client_search.clear()
        client_search.send_keys(CLIENT)
        time.sleep(2)
        ActionChains(driver).key_down(Keys.CONTROL).click(client_search).perform()
        client_search.send_keys(Keys.DOWN)
        time.sleep(2)
        client_search.send_keys(Keys.ENTER)
        time.sleep(2)

        # Apply filter
        filter_button = driver.find_element(By.XPATH,
                                            '//*[@id="runSummaryFlag"]/table/tbody/tr/td/table/tbody/tr[1]/td/input')
        filter_button.click()

        # Initiate download
        driver.find_element(By.XPATH, "//input[@name='handleExcelDownLoad']").click()
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="notificationMsg"]/a').click()

        # Wait for report generation
        time.sleep(DOWNLOAD_TIMEOUT)

        # Final download click
        driver.find_element(By.XPATH, '//*[@id="mainContent"]/div[1]/div/table/tbody/tr[9]/td[2]/a/img').click()

        # Wait for download to complete
        time.sleep(10)

        # Get downloaded file content
        downloaded_files = glob.glob(os.path.join(tempfile.gettempdir(), '*.xlsx'))
        if not downloaded_files:
            raise FileNotFoundError("No files were downloaded")

        newest_file = max(downloaded_files, key=os.path.getctime)

        with open(newest_file, 'rb') as f:
            file_content = f.read()

        # Clean up
        os.remove(newest_file)

        return file_content

    except Exception as e:
        st.error(f"Error during report download: {str(e)}")
        if driver:
            driver.quit()
        st.stop()


def main():
    st.title(f"{CLIENT} GoPricing Report Downloader")

    # Date range selection
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start Date", datetime.now() - timedelta(days=7))
    with col2:
        end_date = st.date_input("End Date", datetime.now())

    if st.button("Download Reports"):
        if start_date > end_date:
            st.error("End date must be after start date")
            return

        current_date = start_date
        while current_date <= end_date:
            if is_business_day(current_date):
                target_date_str = current_date.strftime("%d-%b-%Y")
                with st.spinner(f"Processing {target_date_str}..."):
                    try:
                        driver = setup_driver()
                        file_content = download_report(driver, target_date_str)

                        st.download_button(
                            label=f"Download {target_date_str} Report",
                            data=file_content,
                            file_name=f"{CLIENT}_ALL_OTC_{target_date_str}.xlsx",
                            mime="application/vnd.ms-excel",
                            key=f"download_{target_date_str}"
                        )
                        st.success(f"Successfully processed {target_date_str}")

                    except Exception as e:
                        st.error(f"Failed to process {target_date_str}: {str(e)}")
                    finally:
                        if driver:
                            driver.quit()

                time.sleep(5)  # Be polite to the server

            current_date += timedelta(days=1)


if __name__ == "__main__":
    main()