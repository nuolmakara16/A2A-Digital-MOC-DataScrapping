from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from
TARGET_URL = f'https://www.green-japan.com/search_key/01'
DRIVER_PATH = '../drivers/chromedriver.exe'
DRIVER = webdriver.Chrome(executable_path=DRIVER_PATH)
DRIVER.get(TARGET_URL)
DRIVER.maximize_window()

MAXIMUM_DELAY_TIME = 20

# job_high_div_id = 's2id_job_high'
# element = EC.element_to_be_clickable((By.ID, job_high_div_id))
# WebDriverWait(DRIVER, MAXIMUM_DELAY_TIME).until(element).click()
#
# job_high_ul = DRIVER.find_element_by_xpath('//*[@id="select2-drop"]/ul')
# job_high_li = job_high_ul.find_elements_by_tag_name("li")
# for index, job_high in enumerate(job_high_li):
#     print(f'{index + 1}.{job_high.text}')

def click_on_element(element_name):
    print('Hello world!')


company_feature_search = 's2id_company_feature'
element = EC.element_to_be_clickable((By.ID, company_feature_search))
WebDriverWait(DRIVER, MAXIMUM_DELAY_TIME).until(element).click()


company_feat_ul = element.find_elements_by_xpath('//*[@id="select2-drop"]/ul')
company_feat_li = company_feat_ul.find_elements_by_tag_name("li")
for index, company_feat in enumerate(company_feat_li):
    print(f'{index + 1}.{company_feat.text}'

