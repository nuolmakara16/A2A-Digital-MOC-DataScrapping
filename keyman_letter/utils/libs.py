import time

import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

MAXIMUM_DELAY_TIME = 0
TOTAL_PAGE = 0
LIST_CLASSNAME = ''
ROWS_CLASSNAME = ''
DRIVER_PATH = ''
ROWS = []
MINIMIZE = True
DOCKER = False

def closeBanner(DRIVER):
    global MAXIMUM_DELAY_TIME
    banner_xpath = "/html/body/div/div[1]/div/div[6]/div[2]/div/button"
    element = EC.element_to_be_clickable((By.XPATH, banner_xpath))
    WebDriverWait(DRIVER, MAXIMUM_DELAY_TIME).until(element).click()


def appendData(DRIVER):
    global LIST_CLASSNAME
    global ROWS_CLASSNAME
    list_container = DRIVER.find_element_by_class_name(LIST_CLASSNAME)
    rows = list_container.find_elements_by_class_name(ROWS_CLASSNAME)
    for row in rows:
        data = row.text.split('\n')
        if len(data) == 3:
            title, tag, desc = data
            ROWS.append({
                'title': title,
                'tag': tag,
                'desc': desc
            })
        else:
            title, desc = data
            ROWS.append({
                'title': title,
                'tag': '',
                'desc': desc
            })
    print(ROWS)


def findDataRows(DRIVER, PAGE_NO):
    global TOTAL_PAGE
    global MINIMIZE
    url = f'https://keyman-db.smart-letter.com/departments/human_resources?page={PAGE_NO}'
    DRIVER.get(url)

    if MINIMIZE:
        DRIVER.minimize_window()
    else:
        DRIVER.maximize_window()

    closeBanner(DRIVER)

    if PAGE_NO > 1:
        for page in range(TOTAL_PAGE - 1):
            appendData(DRIVER=DRIVER)
            PAGE_NO += 1
            if PAGE_NO != TOTAL_PAGE:
                time.sleep(1)
                xpath = '/html/body/div/div[1]/div/div[1]/div/div/main/div[5]/button[2]'
                element = EC.element_to_be_clickable((By.XPATH, xpath))
                WebDriverWait(DRIVER, MAXIMUM_DELAY_TIME).until(element).click()
    else:
        appendData(DRIVER=DRIVER)
    print()


def scrapFirstPage():
    global DRIVER_PATH
    global DOCKER

    if DOCKER:
        driver = webdriver.Remote(command_executor='http://127.0.0.1:4444/wd/hub', desired_capabilities=DesiredCapabilities.CHROME)
    else:
        driver = webdriver.Chrome(executable_path=DRIVER_PATH)

    findDataRows(DRIVER=driver, PAGE_NO=1)
    driver.quit()


def scrapRemainingPage():
    global DRIVER_PATH
    if DOCKER:
        driver = webdriver.Remote(command_executor='http://127.0.0.1:4444/wd/hub', desired_capabilities=DesiredCapabilities.CHROME)
    else:
        driver = webdriver.Chrome(executable_path=DRIVER_PATH)
    findDataRows(DRIVER=driver, PAGE_NO=2)
    driver.quit()


def setVariables(max_delay_time, total_page, list_classname, rows_classname, driver_path, minimize_windows, docker):
    global MAXIMUM_DELAY_TIME
    global TOTAL_PAGE
    global LIST_CLASSNAME
    global ROWS_CLASSNAME
    global DRIVER_PATH
    global MINIMIZE
    global DOCKER

    MAXIMUM_DELAY_TIME = max_delay_time
    TOTAL_PAGE = total_page
    LIST_CLASSNAME = list_classname
    ROWS_CLASSNAME = rows_classname
    DRIVER_PATH = driver_path
    MINIMIZE = minimize_windows
    DOCKER = docker

def exportExcel(data_set):
    row_num = 0
    with xlsxwriter.Workbook('./outputs/output.xlsx') as workbook:
        worksheet = workbook.add_worksheet()
        worksheet.write_row(row_num, 0, ['No', 'Company Name', 'Tags', 'Headline'])
        row_num += 1
        for data in data_set:
            data['no'] = row_num
            worksheet.write_row(row_num, 0, [data['no'], data['title'], data['tag'], data['desc']])
            row_num += 1

def getResults():
    return ROWS
