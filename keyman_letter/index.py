import time

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter

def check_element(value):
    return value.text if value else " "

url = 'https://keyman-db.smart-letter.com/departments/human_resources?page=334'

driver = webdriver.Chrome('../drivers/chromedriver.exe')
driver.get(url)
driver.maximize_window()
delay = 20

xpath = "/html/body/div/div[1]/div/div[6]/div[2]/div/button"
element = EC.element_to_be_clickable((By.XPATH, xpath))
WebDriverWait(driver, delay).until(element).click()

main = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="gatsby-focus-wrapper"]/div/div[1]/div/div/main/div[3]')))
rows = main.find_elements_by_class_name("jss159")

result = []
row_num = 1
for page_num in range(1):
    for row in rows:
        try:
            title = check_element(row.find_element_by_class_name("jss164"))
            tag = check_element(row.find_element_by_class_name("jss173"))
            desc = check_element(row.find_element_by_class_name("jss167"))
            temp = {
                "title": title,
                "tag": tag,
                "desc": desc,
            }
            result.append(temp)
            print(row_num)
            print(f"Title: {title}")
            print(f"Tag: {tag}")
            print(f"Desc: {desc}")
            print('--------------------------------------')
            row_num += 1
        except NoSuchElementException as e:
            title = check_element(row.find_element_by_class_name("jss164"))
            tag = ''
            desc = check_element(row.find_element_by_class_name("jss167"))
            temp = {
                "title": title,
                "tag": tag,
                "desc": desc,
            }
            result.append(temp)
            print(row_num)
            print(f"Title: {title}")
            print(f"Tag: {tag}")
            print(f"Desc: {desc}")
            print('--------------------------------------')
            row_num += 1
    # try:
    #     xpath = '/html/body/div[1]/div[1]/div/div[1]/div/div/main/div[4]/button[2]'
    #     element = EC.element_to_be_clickable((By.XPATH, xpath))
    #     WebDriverWait(driver, delay).until(element).click()
    # except NoSuchElementException as e:
    #     print(e)

row_num = 0

with xlsxwriter.Workbook('../output.xlsx') as workbook:
    worksheet = workbook.add_worksheet()
    worksheet.write_row(row_num, 0, ['No', 'Company Name', 'Name', 'Headline'])

    row_num += 1
    for data in result:
        data['no'] = row_num
        worksheet.write_row(row_num, 0, [data['no'], data['title'], data['tag'], data['desc']])
        row_num += 1
