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
from datetime import date, datetime, timedelta

client = "ASGARD"


def read_credentials():
    with open('unicorn.txt') as user_file, open('pineapple.txt') as pw_file:
        return user_file.read().strip(), pw_file.read().strip()


def setup_driver():
    chrome_driver_path = 'C:\\Drivers\\chromedriver.exe'
    options = Options()
    options.add_experimental_option('detach', True)  # Keeps Chrome page open once code is done
    return webdriver.Chrome(service=Service(chrome_driver_path), options=options)


def is_business_day(date_to_check):
    # Check if the date is a weekend (Saturday=5, Sunday=6)
    if date_to_check.weekday() >= 5:
        return False
    # Add holiday check here if needed
    return True


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
    date_selector.send_keys(target_date_str)  # Use target_date_str instead of today

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
    time.sleep(2)

    #Select Report

    select_report =  driver.find_element(By.XPATH,'//*[@id="excelDownloadTag"]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/select')
    select_report.click()
    select_report.send_keys(Keys.DOWN)
    select_report.click()



    # Download the report
    driver.find_element(By.XPATH, "//input[@name='handleExcelDownLoad']").click()
    time.sleep(5)  # Wait for the download to initiate
    driver.find_element(By.XPATH, '//*[@id="notificationMsg"]/a').click()

    # Wait for the report to be ready (assume 30 seconds is enough)
    time.sleep(30)

    # Click the Excel image to download the report
    driver.find_element(By.XPATH, '//*[@id="mainContent"]/div[1]/div/table/tbody/tr[9]/td[2]/a/img').click()
    time.sleep(10)  # Wait for the download to complete

    # Save the downloaded file
    dl_folder = os.path.join('C:\\Users', username, 'Downloads')
    file_name = 'C:\\' + client + '\\Daily Pricing\\'
    newest_xlsx = max(glob.glob(os.path.join(dl_folder, '*.xlsx')), key=os.path.getctime)
    destination_file = os.path.join(file_name, os.path.basename(newest_xlsx))
    shutil.copy(newest_xlsx, destination_file)
    new_name = file_name + '\\' + client + '_ALL_OTC_' + target_date_str + '.xlsx'  # Use target_date_str instead of today
    if os.path.exists(new_name):
        os.remove(new_name)
    os.rename(destination_file, new_name)

    os.remove(newest_xlsx)  # Remove the original downloaded file from the Downloads folder

    driver.quit()


def main():
    # Define date range
    start_date = datetime(2025, 3, 7)  # Start date
    end_date = datetime(2025, 3, 21)  # End date

    # Loop through each date in the range
    current_date = start_date
    while current_date <= end_date:
        if is_business_day(current_date):  # Check if the date is a business day
            target_date_str = current_date.strftime("%d-%b-%Y")  # Format the date
            print(f"Processing data for {target_date_str}")
            try:
                gopx(target_date_str)  # Call the gopx function with the target date
                time.sleep(30)  # Add a delay between requests to avoid overloading the server
            except Exception as e:
                print(f"Failed to process data for {target_date_str}: {e}")
        current_date += timedelta(days=1)  # Move to the next day


if __name__ == "__main__":
    main()