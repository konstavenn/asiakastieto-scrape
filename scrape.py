from xml.dom.minidom import Element
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import beepy
import PySimpleGUI as sg
import _thread
import time
import datetime
from selenium.webdriver.chrome.options import Options
import platform


timeout=10
if "mac" in platform.platform():
    chrome_driver_path = "driver/mac_chromedriver"

elif "windows" in platform.platform():
    chrome_driver_path = "driver/win_chromedriver"

else:
    chrome_driver_path = "driver/linux_chromedriver"


def mybeep():
    beepy.beep(0)
    return



layout = [
    [sg.Text('Person of interest:'), sg.InputText("", key='Input1')],
    [sg.Button("Search", button_color='green')],
    #[sg.Output(key='-OUT-', size=(80, 20), )],
    [sg.Multiline(autoscroll=True, size=(80, 60), reroute_stdout=True, write_only=True, disabled=True, key='-OUT-')],
]

window = sg.Window("https://www.asiakastieto.fi/web/fi/ scrape", layout, finalize=True, element_justification='center', size=(500, 500))
window['Input1'].bind("<Return>", "_Enter")
print("""The author takes no responsibility on if the code works or the legality of your use cases. 

If this software makes your life easier, feel free to offer me a beer if our paths ever cross.

If you have problems check the README.txt file.

author: 
Konsta Venn
Taika Tech

""")
print('\nFor multiple searches, seperate names by comma: "konsta venn, pekka aho".\n')

def get_element(driver, path):
    element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, path)))
    return driver, element

def init_driver(headless = True):
    options = Options()
    options.headless = headless
    print("Opening Chrome...")
    driver = webdriver.Chrome(chrome_driver_path, options=options)
    return driver

def open_page(driver, page):
    #open the page
    driver.get(page)
    print("Searching for cookie...")

    #accept the cookie
    reject_cookie = '//*[@id="coi-tcf-modal-main"]/div[2]/div/div[3]/button[1]'
    accept_cookie = '//*[@id="acceptButton"]'
    WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH,reject_cookie))).click()
    return driver

def create_dd_file(person):
    person_name = person.lstrip().rstrip()
    dd_file = "dd/" + person_name + ".docx"
    with open(dd_file, 'w') as f:
        f.write(person_name)
    return dd_file

def search_for_person(driver, person_name):
    #open search tab
    driver.switch_to.window(driver.window_handles[0])

    #search for people
    select = Select(driver.find_element(By.CLASS_NAME, "search-select"))
    select.select_by_value("PERSON")

    #search for person
    fullxpath = '/html/body/div[2]/div[2]/div/div/form/input[2]'
    driver, text_area = get_element(driver, fullxpath)
    text_area.clear()

    #text_area.send_keys(Keys.CONTROL, "a")
    text_area.send_keys(person_name)
    text_area.send_keys(Keys.ENTER)

    #go to new tab
    driver.switch_to.window(driver.window_handles[1])
    return driver

def drop_duplicates(driver, count):
    firman_nimet = []
    boards_actually = 0
    for company in range(count):
        company+=1
        #get first company name
        driver, firma_nimi = get_element(driver, '/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{}]/span[3]'.format(company))
        firma_nimi = firma_nimi.text
        if firma_nimi not in firman_nimet:
            firman_nimet.append(firma_nimi)
            boards_actually+=1
    return driver, boards_actually


def log(dd_file, string, extra=False):
    if extra != False:
        print(extra, string)
    else:
        print(string)
    with open(dd_file, 'a') as f:
        f.write(string + '\n')
    return

def person_search(driver, person, dd_file):
    search_for_person(driver, person)

    #get count of boards
    driver, full_company_count = get_element(driver, '/html/body/div[3]/main/section/div[3]/div/div/div[1]/div/div/div[1]/span')
    full_company_count = int(full_company_count.text.split()[0])

    #make sure all are on same page
    if full_company_count > 100:
        full_company_count = 100

    #get correct count without duplicates
    driver, boards_actually = drop_duplicates(driver, full_company_count)

    log(dd_file, " on mukana {} hallituksessa.".format(boards_actually), person)

    running_company_count=0
    #To not log duplicates
    company_names = []
    list_of_fucked_up_companies = []
    for company in range(full_company_count):
        company+=1
        running_company_count+=1

        #get first company name
        company_path = '/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{}]/span[3]'.format(company)
        driver, company_name = get_element(driver, company_path)
        company_name = company_name.text
        
        #check for duplicate
        if company_name in company_names:
            running_company_count-=1

        else:
            company_names.append(company_name)
            log(dd_file, "\t{}/{}: {}".format(running_company_count, boards_actually, company_name))

            #open company
            path = "/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{}]".format(company)

            #path = "/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{}]".format(company)
            link = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, path)))
            link.click()
            #driver, link = get_element(driver, path)
            #link.click()

            try:
                #get table dimensions
                row_xpath='//*[@id="row--persons-in-charge"]/div[1]/div/table/tbody/tr'
                col_xpath='//*[@id="row--persons-in-charge"]/div[1]/div/table/tbody/tr/td'
                driver, _ = get_element(driver, row_xpath)
                
                rows = len(driver.find_elements(By.XPATH, row_xpath))
                cols = int(len(driver.find_elements(By.XPATH, col_xpath)) / rows)
                print("\t\tHallituksessa on {} jäsentä.".format(rows))
                
                #loop through whole board
                for i in range(rows):
                    row = []
                    for j in range(cols-1):
                        path = '//*[@id="row--persons-in-charge"]/div[1]/div/table/tbody/tr[{}]/td[{}]'.format(i+1, j+1)
                        #path = "/html/body/div[{}]/main/section/div[3]/div[1]/div/div[1]/div/table/tbody/tr[{}]/td[{}]".format(fuckyouval, i+1, j+1)
                        val = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, path))).text
                        row.append(val)
                    log(dd_file, "\t\t\t" + row[1] +  ": " + row[0])
                log(dd_file, "\n")
            except:
                print(company_name)
                list_of_fucked_up_companies.append(company_name)

            #return to all boards of the person
            driver.execute_script("window.history.go(-1)")

    #close tab after searching for person        
    driver.close()
    return driver, list_of_fucked_up_companies









def count_foundation(driver, full_foundation_count):
    boards_actually = 0
    for company in range(full_foundation_count):
        company+=1
        #get first company name
        driver, firma_nimi = get_element(driver, '/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{}]/span[3]'.format(company))
        firma_nimi = firma_nimi.text
        if " sr " in firma_nimi:
            boards_actually+=1
    return driver, boards_actually


def foundation_search(driver, person, dd_file):
    search_for_person(driver, person)

    #get count of boards
    driver, full_foundation_count = get_element(driver, '/html/body/div[3]/main/section/div[3]/div/div/div[1]/div/div/div[1]/span')
    full_foundation_count = int(full_foundation_count.text.split()[0])

    #make sure all are on same page
    if full_foundation_count > 100:
        full_foundation_count = 100
    #get correct count without duplicates
    driver, boards_actually = count_foundation(driver, full_foundation_count)
    list_of_fucked_up_companies = []
    if boards_actually > 0:
        log(dd_file, "{} on mukana {} foundationssä.".format(person, boards_actually))

        running_company_count=0
        #To not log duplicates
        company_names = []
    
        for company in range(full_foundation_count):
            company+=1
            running_company_count+=1

            #get first company name
            company_path = '/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{}]/span[3]'.format(company)
            driver, company_name = get_element(driver, company_path)
            company_name = company_name.text
            
            #check for duplicate
            if " sr " not in company_name:
                running_company_count-=1

            else:
                company_names.append(company_name)
                log(dd_file, "\t{}/{}: {}".format(running_company_count, boards_actually, company_name))

                #open company
                path = "/html/body/div[3]/main/section/div[3]/div/div/div[2]/div/div/div[2]/a[{}]".format(company)
                driver, link = get_element(driver, path)
                link.click()
                
                try:
                    #get table dimensions
                    row_xpath='//*[@id="row--persons-in-charge"]/div[1]/div/table/tbody/tr'
                    col_xpath='//*[@id="row--persons-in-charge"]/div[1]/div/table/tbody/tr/td'
                    driver, _ = get_element(driver, row_xpath)
                    
                    rows = len(driver.find_elements(By.XPATH, row_xpath))
                    cols = int(len(driver.find_elements(By.XPATH, col_xpath)) / rows)
                    print("\t\tHallituksessa on {} jäsentä.".format(rows))
                    
                    #loop through whole board
                    for i in range(rows):
                        row = []
                        for j in range(cols-1):
                            path = '//*[@id="row--persons-in-charge"]/div[1]/div/table/tbody/tr[{}]/td[{}]'.format(i+1, j+1)
                            #path = "/html/body/div[{}]/main/section/div[3]/div[1]/div/div[1]/div/table/tbody/tr[{}]/td[{}]".format(fuckyouval, i+1, j+1)
                            val = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, path))).text
                            row.append(val)
                        log(dd_file, "\t\t\t" + row[1] +  ": " + row[0])
                    log(dd_file, "\n")
                except:
                    print(company_name)
                    list_of_fucked_up_companies.append(company_name)

                #return to all boards of the person
                driver.execute_script("window.history.go(-1)")

    #close tab after searching for person 
    driver.close()
    return driver, list_of_fucked_up_companies



def do_search(people):
    driver = init_driver(False)
    driver = open_page(driver, "https://www.asiakastieto.fi/web/fi/")

    #loop through people
    for person in people:
        dd_file = create_dd_file(person)

        #look for ompanies
        driver, list_of_fucked_up_companies = person_search(driver, person, dd_file)

        #look for foundations
        person_full_name = person.lstrip().rstrip().split(" ")
        person_short_name = person_full_name[0] + " " + person_full_name[-1]

        driver, list_of_fucked_up_foundation = foundation_search(driver, person_short_name, dd_file)

    #close chrome after full search
    driver.quit()
    _thread.start_new_thread(mybeep,())
    print('Search complete. The dd file can be found from the "dd" folder.\n\n')

    if len(list_of_fucked_up_companies):
        print("Something might have gone wrong with following companies:")
    for firma in list_of_fucked_up_companies:
        print("\t", firma)

    if len(list_of_fucked_up_foundation):
        print("Something might have gone wrong with following companies:")
    for foundation in list_of_fucked_up_foundation:
        print("\t", foundation)

    return

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == "Hae" or event == "Input1" + "_Enter":
        people = values['Input1'].split(",")
        _thread.start_new_thread(do_search, (people, ))


        
