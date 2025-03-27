import time
import glob
import os
import shutil
from datetime import datetime, timedelta
import PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait as wdw
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import UnexpectedAlertPresentException

# Constants
client = "ASGARD"
MAX_RETRIES = 3
RETRY_DELAY = 10  # seconds


# ---------------------------- Core Functions (UNCHANGED) ----------------------------

def read_credentials():
    cred_folder = r'C:\Credentials'
    os.makedirs(cred_folder, exist_ok=True)  # Creates folder if missing

    user_file = os.path.join(cred_folder, 'unicorn.txt')
    pw_file = os.path.join(cred_folder, 'pineapple.txt')

    with open(user_file) as uf, open(pw_file) as pf:
        return uf.read().strip(), pf.read().strip()


def setup_driver():
    chrome_driver_path = 'C:\\Drivers\\chromedriver.exe'
    options = Options()
    options.add_experimental_option('detach', True)  # Keeps Chrome page open once code is done
    return webdriver.Chrome(service=Service(chrome_driver_path), options=options)


def is_business_day(date_to_check):
    """Check if date is a weekday (Mon-Fri)"""
    return date_to_check.weekday() < 5


def gopx(target_date_str, client):
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
    filter_button = driver.find_element(By.XPATH,
                                        '//*[@id="runSummaryFlag"]/table/tbody/tr/td/table/tbody/tr[1]/td/input')
    filter_button.click()
    time.sleep(2)

    # Select Report
    select_report = driver.find_element(By.XPATH,
                                        '//*[@id="excelDownloadTag"]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/select')
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
    new_name = file_name + '\\' + client + '_ALL_OTC_' + target_date_str + '.xlsx'
    if os.path.exists(new_name):
        os.remove(new_name)
    os.rename(destination_file, new_name)

    os.remove(newest_xlsx)  # Remove the original downloaded file from the Downloads folder

    driver.quit()


# ---------------------------- Enhanced GUI ----------------------------
def create_window():
    sg.theme('DarkBlue3')  # Professional dark theme

    # Calculate default dates (previous Monday to current Friday if today is weekday)
    today = datetime.now()
    start_date = today - timedelta(days=today.weekday())
    if not is_business_day(today):
        start_date = today - timedelta(days=(today.weekday() + 7))
    end_date = start_date + timedelta(days=4)

    layout = [
        [sg.Text('Daily NAV Report Downloader', font=('Helvetica', 16), justification='center', expand_x=True)],
        [sg.HorizontalSeparator()],
        [sg.Text('Client:', size=10), sg.Input(key='-CLIENT-', default_text='ASGARD', size=20)],  # Added this line
        [sg.Text('Date Range Selection', font=('Helvetica', 12))],
        [
            sg.Text('Start Date:', size=10),
            sg.Input(key='-START-', size=12, default_text=start_date.strftime('%d-%b-%Y')),
            sg.CalendarButton('📅', target='-START-', format='%d-%b-%Y', no_titlebar=False),
            sg.Text('End Date:', size=10),
            sg.Input(key='-END-', size=12, default_text=end_date.strftime('%d-%b-%Y')),
            sg.CalendarButton('📅', target='-END-', format='%d-%b-%Y', no_titlebar=False)
        ],
        [sg.Text('Download Folder:', size=15),
         sg.Input(key='-FOLDER-', default_text=f'C:\\{client}\\Daily Pricing\\', expand_x=True),
         sg.FolderBrowse()],
        [sg.Checkbox('Include weekends (normally excluded)', key='-INCL_WEEKENDS-')],
        [sg.Checkbox('Pause between downloads (recommended)', default=True, key='-PAUSE-')],
        [sg.HorizontalSeparator()],
        [sg.ProgressBar(100, orientation='h', size=(50, 20), key='-PROGRESS-', expand_x=True)],
        [sg.Text('Ready', key='-STATUS-', size=50, relief=sg.RELIEF_SUNKEN)],
        [sg.Button('Run Reports', size=12), sg.Button('Test Connection', size=12), sg.Push(),
         sg.Button('Exit', size=12)],
        [sg.Multiline(size=(80, 15), key='-LOG-', autoscroll=True, disabled=True, expand_x=True, expand_y=True)]
    ]
    return sg.Window('Daily Report Downloader', layout, resizable=True, finalize=True)


def log_message(window, message):
    """Add a timestamped message to the log"""
    timestamp = datetime.now().strftime('%H:%M:%S')
    window['-LOG-'].print(f'[{timestamp}] {message}')
    window.refresh()


def validate_dates(start_str, end_str):
    """Validate date inputs and return datetime objects"""
    try:
        start_date = datetime.strptime(start_str, '%d-%b-%Y')
        end_date = datetime.strptime(end_str, '%d-%b-%Y')
        if start_date > end_date:
            return None, None, "Start date must be before end date"
        return start_date, end_date, None
    except ValueError:
        return None, None, "Invalid date format (use DD-MMM-YYYY, e.g. 07-Mar-2025)"


def run_reports(window, values):
    """Main function to run the reports with GUI updates"""
    start_date, end_date, error = validate_dates(values['-START-'], values['-END-'])
    if error:
        sg.popup_error(error)
        return

    include_weekends = values['-INCL_WEEKENDS-']
    pause_between = values['-PAUSE-']
    download_folder = values['-FOLDER-']

    # Get client from GUI input (ADD THIS LINE)
    client = values['-CLIENT-'].strip()

    # Verify download folder exists
    if not os.path.exists(download_folder):

        try:
            os.makedirs(download_folder)
            log_message(window, f"Created download folder: {download_folder}")
        except Exception as e:
            log_message(window, f"Error creating folder: {e}")
            return


    # Calculate total business days for progress tracking
    total_days = (end_date - start_date).days + 1
    business_days = sum(1 for day in (start_date + timedelta(n) for n in range(total_days))
                        if is_business_day(day) or include_weekends)

    if business_days == 0:
        sg.popup_error("No valid dates selected for processing")
        return

    log_message(window, f"Starting report download for {business_days} days")
    log_message(window, f"From {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")

    processed = 0
    current_date = start_date
    while current_date <= end_date:
        if is_business_day(current_date) or include_weekends:
            target_date_str = current_date.strftime("%d-%b-%Y")
            progress = int((processed / business_days) * 100)
            window['-PROGRESS-'].update(progress)
            window['-STATUS-'].update(f"Processing {target_date_str}...")
            log_message(window, f"Processing data for {target_date_str}")

            try:
                gopx(target_date_str, client)
                processed += 1
                log_message(window, f"Successfully downloaded report for {target_date_str}")

                if pause_between and processed < business_days:
                    window['-STATUS-'].update(f"Pausing for 30 seconds...")
                    log_message(window, "Pausing between downloads...")
                    for i in range(30, 0, -1):
                        window['-STATUS-'].update(f"Resuming in {i} seconds...")
                        time.sleep(1)
                        if window.TKroot.winfo_exists() and sg.Window.active(window):
                            window.refresh()
                        else:
                            return  # Window closed during pause

            except Exception as e:
                log_message(window, f"Failed to process data for {target_date_str}: {str(e)}")
                if MAX_RETRIES > 0:
                    for attempt in range(MAX_RETRIES):
                        log_message(window, f"Retry attempt {attempt + 1} of {MAX_RETRIES}")
                        time.sleep(RETRY_DELAY)
                        try:
                            gopx(target_date_str)
                            processed += 1
                            log_message(window, f"Successfully downloaded report on retry")
                            break
                        except Exception as retry_e:
                            log_message(window, f"Retry failed: {str(retry_e)}")
                    else:
                        log_message(window, f"Max retries reached for {target_date_str}")

        current_date += timedelta(days=1)

    window['-PROGRESS-'].update(100)
    window['-STATUS-'].update(f"Completed! Processed {processed} of {business_days} days")
    log_message(window, "Report download process completed")
    sg.popup(f"Process completed!\nDownloaded {processed} reports.", title='Complete')


def test_connection(window):
    """Test the connection to the reporting website"""
    log_message(window, "Testing connection to reporting website...")
    try:
        username, password = read_credentials()
        url = f"https://{username}:{password}@gopricing.ssnc-corp.cloud/clients/reports/otcReportSummary.go?action=handleSummaryFilter"
        driver = setup_driver()
        driver.get(url)
        driver.maximize_window()

        # Check if we can access the main page elements
        try:
            wdw(driver, 15).until(ec.presence_of_element_located((By.XPATH, '//*[@id="valuationDate"]')))
            log_message(window, "Connection test successful - page elements loaded")
            driver.quit()
            window['-STATUS-'].update("Connection test successful")
            sg.popup("Connection test successful!", title='Success')
        except Exception as e:
            log_message(window, f"Connection test failed - could not load page elements: {str(e)}")
            window['-STATUS-'].update("Connection test failed")
            sg.popup_error("Connection test failed - could not load page elements")
            driver.quit()
    except Exception as e:
        log_message(window, f"Connection test failed: {str(e)}")
        window['-STATUS-'].update("Connection test failed")
        sg.popup_error(f"Connection test failed: {str(e)}")


def main():
    window = create_window()

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == 'Run Reports':
            run_reports(window, values)
        elif event == 'Test Connection':
            test_connection(window)

    window.close()


if __name__ == "__main__":
    main()