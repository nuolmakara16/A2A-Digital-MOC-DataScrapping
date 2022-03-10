from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

TARGET_URL = f'https://www.green-japan.com/search_key/01'
DRIVER_PATH = '../drivers/chromedriver.exe'
DRIVER = webdriver.Chrome(executable_path=DRIVER_PATH)
DRIVER.get(TARGET_URL)

MAXIMUM_DELAY_TIME = 20
element_id = 's2id_job_high'
element = EC.element_to_be_clickable((By.ID, element_id))
WebDriverWait(DRIVER, MAXIMUM_DELAY_TIME).until(element).click()
