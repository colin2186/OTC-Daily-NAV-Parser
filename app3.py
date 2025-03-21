import streamlit as st
from datetime import datetime, timedelta
import time
import glob
import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait as wdw
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# Custom CSS for styling
st.markdown(
    """
    <style>
    .stButton button {
        background-color: #4CAF50;
        color: white;
        font-size: 16px;
        padding: 10px 24px;
        border-radius: 5px;
        border: none;
    }
    .stButton button:hover {
        background-color: #45a049;
    }
    .stTextInput input {
        font-size: 16px;
    }
    .stDateInput input {
        font-size: 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# Title and subheader
st.title("📊 Automated Report Downloader")
st.subheader("Download daily pricing reports for your clients")

# Input fields (blank by default)
col1, col2 = st.columns(2)
with col1:
    client = st.text_input("Enter Client Name")  # No default value
with col2:
    st.write("")  # Spacer for alignment

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Select Start Date")  # No default value
with col2:
    end_date = st.date_input("Select End Date")  # No default value

# Divider
st.divider()

# Button to run the script
if st.button("Run Report Downloader", key="run_button"):
    if not client or not start_date or not end_date:
        st.error("Please fill in all fields (Client Name, Start Date, and End Date).")
    else:
        st.write("Starting the report download process...")

        # Function to read credentials (replace with your actual implementation)
        def read_credentials():
            with open('unicorn.txt') as user_file, open('pineapple.txt') as pw_file:
                return user_file.read().strip(), pw_file.read().strip()

        # Function to set up the Chrome driver
        def setup_driver():
            chrome_driver_path = 'C:\\Drivers\\chromedriver.exe'
            options = Options()
            options.add_experimental_option('detach', True)  # Keeps Chrome page open once code is done
            return webdriver.Chrome(service=Service(chrome_driver_path), options=options)

        # Function to check if a date is a business day
        def is_business_day(date_to_check):
            if date_to_check.weekday() >= 5:  # Saturday=5, Sunday=6
                return False
            return True

        # Function to download reports for a given date range
        def gopx(target_date_str):
            username, password = read_credentials()
            url = f"https://{username}:{password}@gopricing.ssnc-corp.cloud/clients/reports/otcReportSummary.go?action=handleSummaryFilter"
            driver = setup_driver()
            driver.get(url)
            driver.maximize_window()
            driver.get(url)
            time.sleep(2)

            # Wait for the date selector to be available
            wdw(driver, 15).until(ec.presence_of_element_located((By.XPATH, '//*[@id="valuationDate"]')))
            date_selector = driver.find_element(By.XPATH, '//*[@id="valuationDate"]')
            date_selector.clear()
            date_selector.send_keys(target_date_str)

            # Find and interact with client search
            client_search = driver.find_element(By.XPATH, "//*[@id='clientSearch']")
            client_search.clear()
            client_search.send_keys(client)
            time.sleep(2)
            ActionChains(driver).key_down(Keys.CONTROL).click(client_search).perform()
            client_search.send_keys(Keys.DOWN)
            time.sleep(2)
            client_search.send_keys(Keys.ENTER)
            time.sleep(2)

            # Apply filter
            filter_button = driver.find_element(By.XPATH, '//*[@id="runSummaryFlag"]/table/tbody/tr/td/table/tbody/tr[1]/td/input')
            filter_button.click()

            # Download the report
            driver.find_element(By.XPATH, "//input[@name='handleExcelDownLoad']").click()
            time.sleep(5)
            driver.find_element(By.XPATH, '//*[@id="notificationMsg"]/a').click()

            # Wait for the report to be ready
            time.sleep(30)  # Initial wait

            # Refresh the page after 30 seconds to check if the report is ready
            driver.refresh()
            time.sleep(30)  # Wait again after refresh

            # Click the Excel image to download the report
            driver.find_element(By.XPATH, '//*[@id="mainContent"]/div[1]/div/table/tbody/tr[9]/td[2]/a/img').click()
            time.sleep(10)

            # Save the downloaded file
            dl_folder = os.path.join('C:\\Users', username, 'Downloads')
            file_name = 'C:\\' + client + '\\Daily Pricing\\'
            newest_xlsx = max(glob.glob(os.path.join(dl_folder, '*.xlsx')), key=os.path.getctime)
            destination_file = os.path.join(file_name, os.path.basename(newest_xlsx))
            shutil.copy(newest_xlsx, destination_file)
            new_name = file_name + '\\' + client + '_ALL_OTC_' + target_date_str + '.xlsx'

            if os.path.exists(new_name):
                os.remove(new_name)
            os.rename(destination_file, new_name)

            os.remove(newest_xlsx)
            driver.quit()


        # Progress bar
        progress_bar = st.progress(0)
        total_days = (end_date - start_date).days + 1
        processed_days = 0

        # Loop through each date in the range
        current_date = start_date
        while current_date <= end_date:
            if is_business_day(current_date):  # Check if the date is a business day
                target_date_str = current_date.strftime("%d-%b-%Y")  # Format the date
                st.write(f"Processing data for {target_date_str}")
                try:
                    gopx(target_date_str)
                    st.success(f"Successfully downloaded report for {target_date_str}")
                except Exception as e:
                    st.error(f"Failed to process data for {target_date_str}: {e}")
                processed_days += 1
                progress_bar.progress(processed_days / total_days)
            current_date += timedelta(days=1)

        st.write("Report download process completed!")