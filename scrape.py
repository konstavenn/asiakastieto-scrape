import platform
import time
import _thread
from xml.dom.minidom import Element

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import beepy
import PySimpleGUI as sg

# Constants
TIMEOUT = 10
WINDOW_SIZE = (500, 500)

# Set ChromeDriver path based on operating system
if "mac" in platform.platform():
    CHROME_DRIVER_PATH = "driver/mac_chromedriver"
elif "windows" in platform.platform():
    CHROME_DRIVER_PATH = "driver/win_chromedriver"
else:
    CHROME_DRIVER_PATH = "driver/linux_chromedriver"

def play_notification_sound():
    """Play a notification sound."""
    beepy.beep(0)

def create_gui_layout():
    """Create the GUI layout for the application."""
    return [
        [sg.Text('Person of interest:'), sg.InputText("", key='Input1')],
        [sg.Button("Search", button_color='green')],
        [sg.Multiline(autoscroll=True, size=(80, 60), reroute_stdout=True, write_only=True, disabled=True, key='-OUT-')],
    ]

def initialize_driver(headless=True):
    """Initialize and return a Chrome WebDriver."""
    options = webdriver.ChromeOptions()
    options.headless = headless
    service = Service(ChromeDriverManager().install())
    print("Opening Chrome...")
    try:
        return webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"Error initializing Chrome driver: {e}")
        return None

def open_webpage(driver, url):
    """Open a webpage and handle cookie consent."""
    driver.get(url)
    print("Searching for cookie consent...")
    reject_cookie_xpath = '//*[@id="coi-tcf-modal-main"]/div[2]/div/div[3]/button[1]'
    WebDriverWait(driver, TIMEOUT).until(EC.element_to_be_clickable((By.XPATH, reject_cookie_xpath))).click()
    return driver

def create_report_file(person_name):
    """Create and return the path for a report file."""
    clean_name = person_name.strip()
    file_path = f"dd/{clean_name}.docx"
    with open(file_path, 'w') as f:
        f.write(clean_name)
    return file_path

def search_for_person(driver, person_name):
    """Perform a search for a person on the website."""
    driver.switch_to.window(driver.window_handles[0])
    select = Select(driver.find_element(By.CLASS_NAME, "search-select"))
    select.select_by_value("PERSON")
    
    search_input_xpath = '/html/body/div[2]/div[2]/div/div/form/input[2]'
    search_input = WebDriverWait(driver, TIMEOUT).until(EC.presence_of_element_located((By.XPATH, search_input_xpath)))
    search_input.clear()
    search_input.send_keys(person_name)
    search_input.send_keys(Keys.ENTER)
    
    driver.switch_to.window(driver.window_handles[1])
    return driver

def count_unique_boards(driver, total_count):
    """Count unique boards a person is on."""
    unique_companies = set()
    for i in range(1, total_count + 1):
        company_name_xpath = f'/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{i}]/span[3]'
        company_name = WebDriverWait(driver, TIMEOUT).until(EC.presence_of_element_located((By.XPATH, company_name_xpath))).text
        unique_companies.add(company_name)
    return len(unique_companies)

def log_to_file(file_path, message, extra=None):
    """Log a message to both console and file."""
    if extra:
        print(extra, message)
    else:
        print(message)
    with open(file_path, 'a') as f:
        f.write(message + '\n')

def scrape_board_info(driver, file_path, is_foundation=False):
    """Scrape and log board information for companies or foundations."""
    total_count_xpath = '/html/body/div[3]/main/section/div[3]/div/div/div[1]/div/div/div[1]/span'
    total_count = int(WebDriverWait(driver, TIMEOUT).until(EC.presence_of_element_located((By.XPATH, total_count_xpath))).text.split()[0])
    
    total_count = min(total_count, 100)  # Limit to 100 results
    actual_board_count = count_unique_boards(driver, total_count)
    
    entity_type = "foundations" if is_foundation else "boards"
    log_to_file(file_path, f" is on {actual_board_count} {entity_type}.")

    processed_companies = set()
    problematic_companies = []

    for i in range(1, total_count + 1):
        company_xpath = f'/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{i}]'
        company_name = WebDriverWait(driver, TIMEOUT).until(EC.presence_of_element_located((By.XPATH, company_xpath))).text
        
        if is_foundation and " sr " not in company_name:
            continue
        
        if company_name not in processed_companies:
            processed_companies.add(company_name)
            log_to_file(file_path, f"\t{len(processed_companies)}/{actual_board_count}: {company_name}")
            
            WebDriverWait(driver, TIMEOUT).until(EC.element_to_be_clickable((By.XPATH, company_xpath))).click()
            
            try:
                board_members_xpath = '//*[@id="row--persons-in-charge"]/div[1]/div/table/tbody/tr'
                board_members = driver.find_elements(By.XPATH, board_members_xpath)
                
                log_to_file(file_path, f"\t\tBoard has {len(board_members)} members.")
                
                for member in board_members:
                    role = member.find_element(By.XPATH, './td[1]').text
                    name = member.find_element(By.XPATH, './td[2]').text
                    log_to_file(file_path, f"\t\t\t{role}: {name}")
                
                log_to_file(file_path, "\n")
            except:
                print(f"Error processing {company_name}")
                problematic_companies.append(company_name)
            
            driver.execute_script("window.history.go(-1)")

    driver.close()
    return driver, problematic_companies

def perform_search(people):
    """Perform the search for given people."""
    try:
        driver = initialize_driver(False)
        if not driver:
            print("Failed to initialize the Chrome driver. Exiting search.")
            return

        driver = open_webpage(driver, "https://www.asiakastieto.fi/web/fi/")

        for person in people:
            report_file = create_report_file(person)
            
            driver, problematic_companies = scrape_board_info(driver, report_file)
            
            person_short_name = " ".join(person.strip().split(" ")[::len(person.split()) - 1])
            driver, problematic_foundations = scrape_board_info(driver, report_file, is_foundation=True)

        driver.quit()
        _thread.start_new_thread(play_notification_sound, ())
        print('Search complete. The report files can be found in the "dd" folder.\n')

        if problematic_companies:
            print("Issues occurred with the following companies:")
            for company in problematic_companies:
                print(f"\t{company}")

        if problematic_foundations:
            print("Issues occurred with the following foundations:")
            for foundation in problematic_foundations:
                print(f"\t{foundation}")
    except Exception as e:
        print(f"An error occurred during the search: {e}")
    finally:
        if 'driver' in locals() and driver:
            driver.quit()


def main():
    """Main function to run the GUI and handle user input."""
    layout = create_gui_layout()
    window = sg.Window("Asiakastieto Scraper", layout, finalize=True, element_justification='center', size=WINDOW_SIZE)
    window['Input1'].bind("<Return>", "_Enter")

    print("""Disclaimer: The author takes no responsibility for the functionality or legality of use cases.
If this software proves useful, consider buying the author a beer if your paths cross.
For issues, please refer to the README.txt file.

Author: Konsta Venn
    """)
    print('\nFor multiple searches, separate names by comma: "John Doe, Jane Smith".\n')

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        if event in ("Search", "Input1_Enter"):
            people = values['Input1'].split(",")
            _thread.start_new_thread(perform_search, (people,))

    window.close()

if __name__ == "__main__":
    main()