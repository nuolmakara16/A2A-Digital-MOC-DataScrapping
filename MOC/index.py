from selenium import webdriver
from time import sleep
import xlsxwriter

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime

from utils._style import *

import threading

# Get ID from Json file by using panda
start_time = time.time()

# Declare variable

# General_details
company_name_khs = []
company_name_ens = []
original_entity_identifiers = []
company_statuses = []
incorporation_dates = []
re_registration_dates = []
tax_identification_numbers = []
tax_registration_dates = []
annual_return_last_filed_ons = []
# Business Activities
list_objective = []
list_main_business_activities = []
business_objectives = []
main_business_activities = []
# ៊Number of Employees
males = []
females = []
number_of_cambodian_employees = []
number_of_foreign_employees = []

# Physical Registered Office Address
physical_registered_office_addresses = []
physical_start_dates = []
# Postal Registered Office Address
postal_registered_office_addresses = []
postal_start_dates = []
postal_contact_emails = []
postal_contact_telephone_numbers = []

# list directors of company
director_name_khs = []
director_name_ens = []
director_postal_registered_office_addresses = []
director_telephone_numbers = []
chairman_of_the_board_of_directors = []

# declare list of directors of company
list_director_name_khs = []
list_director_name_ens = []
list_director_postal_registered_office_addresses = []
list_director_telephone_numbers = []
list_chairman_of_the_board_of_directors = []

# list of company id from json file
company_ids = []

# Start and Ending number for the program to run
start_at = 0
stop_at = 3

hasData = True
delay = 20

# %.8d % i will convert number to 8 digit number
data = ["%.8d" % i for i in range(start_at, stop_at)]

outWorkbook = xlsxwriter.Workbook(f"outputs/{start_at}-{stop_at}.xlsx")
outSheet = outWorkbook.add_worksheet()

driver = webdriver.Chrome('../drivers/chromedriver.exe')


def check_element(value):
    return value if value else " "


def Go_To_Search_Page(start_loop):
    try:
        driver.get("https://www.businessregistration.moc.gov.kh/")
        driver.maximize_window()

        # Click on ONLINE SERVICE dropdown
        xpath = "//*[@id='the_menu_triggers']/a[3]"
        element = EC.element_to_be_clickable((By.XPATH, xpath))
        WebDriverWait(driver, delay).until(element).click()

        # Click on SEARCH ENTITY link
        xpath = "//*[@id='appMainNavigation']/div/ul/li[1]/ul/li/ul/li[1]/a/span[2]"
        element = EC.element_to_be_clickable((By.XPATH, xpath))
        WebDriverWait(driver, delay).until(element).click()
    except:
        end_loop = time.time()
        print(
            f"{style.RED + datetime.now().strftime('%d/%m/%Y %H:%M:%S')}: Failed to search ({round(end_loop - start_loop, 2)} sec)")


def Fill_Identification_Number(identification_number):
    # Fill Entity Name or Identifier with data_value ( eg: '00000000')
    xpath = "//*[@id='QueryString']"
    Identification_TextBox = driver.find_element_by_xpath(xpath)
    Identification_TextBox.clear()
    Identification_TextBox.send_keys(identification_number)

    # Wait until search button is clickable then click search
    xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div/a[3]/span[2]"
    element = EC.element_to_be_clickable((By.XPATH, xpath))
    WebDriverWait(driver, delay).until(element).click()


# Main loop
for index in range(len(data)):
    start_loop = time.time()
    identification_number = data[index]
    if hasData:
        Go_To_Search_Page(start_loop)

    Fill_Identification_Number(identification_number)

    # After click search let it sleep for 1 second to wait for the result
    sleep(1)

    src = driver.page_source
    if "No results found" in src:
        hasData = False
        end_loop = time.time()
        print(
            f"{style.YELLOW + datetime.now().strftime('%d/%m/%Y %H:%M:%S')}: {identification_number} ({round(end_loop - start_loop, 2)} sec)")
        sleep(1)
        continue
    else:
        hasData = True
        company_ids.append(identification_number)
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[4]/div[4]/div/div[2]/div/div/div/div[2]/div[1]/div[1]/a"
            element = EC.element_to_be_clickable((By.XPATH, xpath))
            WebDriverWait(driver, delay).until(element).click()
        except:
            end_loop = time.time()
            print(
                f"{style.RED + datetime.now().strftime('%d/%m/%Y %H:%M:%S')}: No Url Found ({round(end_loop - start_loop, 2)} sec)")

        # Let the code sleep for 1 second to wait for page to load
        sleep(1)

        # General_details page
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            company_name_khs.append(result)
        except:
            company_name_khs.append(" ")
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            company_name_ens.append(result)
        except:
            company_name_ens.append(" ")
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[2]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            original_entity_identifiers.append(result)
        except:
            original_entity_identifiers.append(" ")
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[3]/div/div[1]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            company_statuses.append(result)
        except:
            company_statuses.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[3]/div/div[2]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            incorporation_dates.append(result)
        except:
            incorporation_dates.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[4]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            re_registration_dates.append(result)
        except:
            re_registration_dates.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div[1]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            tax_identification_numbers.append(result)
        except:
            tax_identification_numbers.append("")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div[2]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            tax_registration_dates.append(result)
        except:
            tax_registration_dates.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[1]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            annual_return_last_filed_ons.append(result)
        except:
            annual_return_last_filed_ons.append("")

        # Business Activities
        business_objectives = []
        main_business_activities = []

        for activity in range(1, 20):
            try:
                xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div[1]/div[" + str(
                    activity) + "]/div/div/div[1]/div[2]"
                element = driver.find_element_by_xpath(xpath).text
                result = check_element(element)
                business_objectives.append(result)
            except:
                business_objectives.append(" ")

            try:
                xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div[1]/div[" + str(
                    activity) + "]/div/div/div[2]/div[2]"
                element = driver.find_element_by_xpath(xpath).text
                result = check_element(element)
                main_business_activities.append(result)
            except:
                main_business_activities.append(" ")

        list_main_business_activities.append(main_business_activities)
        list_objective.append(business_objectives)

        # Number of Employees
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[3]/div[2]/div/div/div[1]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            males.append(result)
        except:
            males.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[3]/div[2]/div/div/div[2]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            females.append(result)
        except:
            females.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[3]/div[2]/div/div/div[3]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            number_of_cambodian_employees.append(result)
        except:
            number_of_cambodian_employees.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[3]/div[2]/div/div/div[4]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            number_of_foreign_employees.append(result)
        except:
            number_of_foreign_employees.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/ul/li[2]/a/span"
            element = EC.element_to_be_clickable((By.XPATH, xpath))
            WebDriverWait(driver, 10).until(element).click()
        except:
            end_loop = time.time()
            print(
                f"{style.RED + datetime.now().strftime('%d/%m/%Y %H:%M:%S')}: Employee Url Not Found ({round(end_loop - start_loop, 2)} sec)")

        sleep(1)

        # Physical Registered Office Address
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[1]/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            physical_registered_office_addresses.append(result)
        except:
            physical_registered_office_addresses.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[1]/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            physical_start_dates.append(result)
        except:
            physical_start_dates.append(" ")

        # Postal Registered Office Address
        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[2]/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            postal_registered_office_addresses.append(result)
        except:
            postal_registered_office_addresses.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[2]/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            postal_start_dates.append(result)
        except:
            postal_start_dates.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[2]/div/div[3]/div/div/div/div/div/div/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            postal_contact_emails.append(result)
        except:
            postal_contact_emails.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[2]/div/div[4]/div/div/div/div/div/div/div[2]"
            element = driver.find_element_by_xpath(xpath).text
            result = check_element(element)
            postal_contact_telephone_numbers.append(result)
        except:
            postal_contact_telephone_numbers.append(" ")

        try:
            xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/ul/li[3]/a/span"
            element = EC.element_to_be_clickable((By.XPATH, xpath))
            WebDriverWait(driver, 10).until(element).click()
        except:
            end_loop = time.time()
            print(
                f"{style.RED + datetime.now().strftime('%d/%m/%Y %H:%M:%S')}: Address Url Not Found ({round(end_loop - start_loop, 2)} sec)")

        # Sleep for 1 second to wait for the page to load
        sleep(1)

        # list directors of company
        director_name_khs = []
        director_name_ens = []
        director_postal_registered_office_addresses = []
        director_telephone_numbers = []
        chairman_of_the_board_of_directors = []
        for director in range(1, 6):
            try:
                xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div/div/div[2]/div[" + str(
                    director) + "]/div[2]/div/div/div/div[1]/div/div/div/div[1]/div[2]"
                element = driver.find_element_by_xpath(xpath).text
                result = check_element(element)
                director_name_khs.append(result)
            except:
                director_name_khs.append(" ")

            try:
                xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div/div/div[2]/div[" + str(
                    director) + "]/div[2]/div/div/div/div[1]/div/div/div/div[2]/div[2]"
                element = driver.find_element_by_xpath(xpath).text
                result = check_element(element)
                director_name_ens.append(result)
            except:
                director_name_ens.append(" ")

            try:
                xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[" + str(
                    director) + "]/div/div/div/div[2]/div/div/div/div/div[2]"
                element = driver.find_element_by_xpath(xpath).text
                result = check_element(element)
                director_postal_registered_office_addresses.append(result)
            except:
                director_postal_registered_office_addresses.append(" ")

            try:
                xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div/div/div[2]/div[" + str(
                    director) + "]/div[2]/div/div/div/div[3]/div/div/div[2]"
                element = driver.find_element_by_xpath(xpath).text
                result = check_element(element)
                director_telephone_numbers.append(result)
            except:
                director_telephone_numbers.append(" ")

            try:
                xpath = "/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div/div/div[2]/div[" + str(
                    director) + "]/div[2]/div/div/div/div[4]/div[2]"
                element = driver.find_element_by_xpath(xpath).text
                result = check_element(element)
                chairman_of_the_board_of_directors.append(result)
            except:
                chairman_of_the_board_of_directors.append(" ")
        list_director_name_khs.append(director_name_khs)
        list_director_name_ens.append(director_name_ens)
        list_director_postal_registered_office_addresses.append(director_postal_registered_office_addresses)
        list_director_telephone_numbers.append(director_telephone_numbers)
        list_chairman_of_the_board_of_directors.append(chairman_of_the_board_of_directors)

        end_loop = time.time()
        print(
            f"{style.GREEN + datetime.now().strftime('%d/%m/%Y %H:%M:%S')}: {identification_number} ({round(end_loop - start_loop, 2)} sec)")

driver.close()


def write_company_ids():
    for item in range(len(company_ids)):
        outSheet.write(item + 1, 0, company_ids[item])


def write_company_name_khs():
    for item in range(len(company_name_khs)):
        outSheet.write(item + 1, 1, company_name_khs[item])


def write_company_ens():
    for item in range(len(company_name_ens)):
        outSheet.write(item + 1, 2, company_name_ens[item])


def write_original_entity_identifiers():
    for item in range(len(original_entity_identifiers)):
        outSheet.write(item + 1, 3, original_entity_identifiers[item])


def write_company_statuses():
    for item in range(len(company_statuses)):
        outSheet.write(item + 1, 4, company_statuses[item])


def write_incorporation_dates():
    for item in range(len(incorporation_dates)):
        outSheet.write(item + 1, 5, incorporation_dates[item])


def write_re_registration_dates():
    for item in range(len(re_registration_dates)):
        outSheet.write(item + 1, 6, re_registration_dates[item])


def write_tax_identification_numbers():
    for item in range(len(tax_identification_numbers)):
        outSheet.write(item + 1, 7, tax_identification_numbers[item])


def write_tax_registration_dates():
    for item in range(len(tax_registration_dates)):
        outSheet.write(item + 1, 8, tax_registration_dates[item])


def write_annual_return_last_filed_ons():
    for item in range(len(annual_return_last_filed_ons)):
        outSheet.write(item + 1, 9, annual_return_last_filed_ons[item])


# Number of Employees
def write_males():
    for item in range(len(males)):
        outSheet.write(item + 1, 10, males[item])


def write_females():
    for item in range(len(females)):
        outSheet.write(item + 1, 11, females[item])


def write_number_of_cambodian_employees():
    for item in range(len(number_of_cambodian_employees)):
        outSheet.write(item + 1, 12, number_of_cambodian_employees[item])


def write_number_of_foreign_employees():
    for item in range(len(number_of_foreign_employees)):
        outSheet.write(item + 1, 13, number_of_foreign_employees[item])


# Postal Registered Office Address
def write_postal_contact_emails():
    for item in range(len(postal_contact_emails)):
        outSheet.write(item + 1, 14, postal_contact_emails[item])


def write_postal_contact_telephone_numbers():
    for item in range(len(postal_contact_telephone_numbers)):
        outSheet.write(item + 1, 15, postal_contact_telephone_numbers[item])


def write_physical_registered_office_addresses():
    for item in range(len(physical_registered_office_addresses)):
        outSheet.write(item + 1, 16, physical_registered_office_addresses[item])


def write_physical_start_dates():
    for item in range(len(physical_start_dates)):
        outSheet.write(item + 1, 17, physical_start_dates[item])


# Physical Registered Office Address
def write_postal_registered_office_addresses():
    for item in range(len(postal_registered_office_addresses)):
        outSheet.write(item + 1, 18, postal_registered_office_addresses[item])


def write_postal_start_dates():
    for item in range(len(postal_start_dates)):
        outSheet.write(item + 1, 19, postal_start_dates[item])


def write_list_director_name_khs():
    for item in range(len(list_director_name_khs)):
        initial_value = 0
        for value_in_item in range(len(list_director_name_khs[item])):
            initial_value = initial_value + 5
            outSheet.write(item + 1, 15 + initial_value, list_director_name_khs[item][value_in_item])


def write_list_director_name_ens():
    for item in range(len(list_director_name_ens)):
        initial_value = 0
        for value_in_item in range(len(list_director_name_ens[item])):
            initial_value = initial_value + 5
            outSheet.write(item + 1, 16 + initial_value, list_director_name_ens[item][value_in_item])


def write_list_director_postal_registered_office_addresses():
    for item in range(len(list_director_postal_registered_office_addresses)):
        initial_value = 0
        for value_in_item in range(len(list_director_postal_registered_office_addresses[item])):
            initial_value = initial_value + 5
            outSheet.write(item + 1, 17 + initial_value,
                           list_director_postal_registered_office_addresses[item][value_in_item])


def write_list_director_telephone_numbers():
    for item in range(len(list_director_telephone_numbers)):
        initial_value = 0
        for value_in_item in range(len(list_director_telephone_numbers[item])):
            initial_value = initial_value + 5
            outSheet.write(item + 1, 17 + initial_value, list_director_telephone_numbers[item][value_in_item])


def write_list_chairman_of_the_board_of_directors():
    for item in range(len(list_chairman_of_the_board_of_directors)):
        initial_value = 0
        for value_in_item in range(len(list_chairman_of_the_board_of_directors[item])):
            initial_value = initial_value + 5
            outSheet.write(item + 1, 19 + initial_value, list_chairman_of_the_board_of_directors[item][value_in_item])


# Business Activities
def write_list_objective():
    for item in range(len(list_objective)):
        initial_value = 0
        for value_in_item in range(len(list_objective[item])):
            initial_value = initial_value + 2
            outSheet.write(item + 1, 43 + initial_value, list_objective[item][value_in_item])


def write_list_main_business_activities():
    for item in range(len(list_main_business_activities)):
        initial_value = 0
        for value_in_item in range(len(list_main_business_activities[item])):
            initial_value = initial_value + 2
            outSheet.write(item + 1, 44 + initial_value, list_main_business_activities[item][value_in_item])


# Thread

start_thread = time.time()
t1 = threading.Thread(target=write_company_ids)
t2 = threading.Thread(target=write_company_name_khs)
t3 = threading.Thread(target=write_company_ens)
t4 = threading.Thread(target=write_original_entity_identifiers)
t5 = threading.Thread(target=write_company_statuses)
t6 = threading.Thread(target=write_incorporation_dates)
t7 = threading.Thread(target=write_re_registration_dates)
t8 = threading.Thread(target=write_tax_identification_numbers)
t9 = threading.Thread(target=write_tax_registration_dates)
t10 = threading.Thread(target=write_annual_return_last_filed_ons)
t11 = threading.Thread(target=write_males)
t12 = threading.Thread(target=write_females)
t13 = threading.Thread(target=write_number_of_cambodian_employees)
t14 = threading.Thread(target=write_number_of_foreign_employees)
t15 = threading.Thread(target=write_postal_contact_emails)
t16 = threading.Thread(target=write_postal_contact_telephone_numbers)
t17 = threading.Thread(target=write_physical_registered_office_addresses)
t18 = threading.Thread(target=write_physical_start_dates)
t19 = threading.Thread(target=write_postal_registered_office_addresses)
t20 = threading.Thread(target=write_postal_start_dates)
t21 = threading.Thread(target=write_list_director_name_khs)
t22 = threading.Thread(target=write_list_director_name_ens)
t23 = threading.Thread(target=write_list_director_postal_registered_office_addresses)
t24 = threading.Thread(target=write_list_director_telephone_numbers)
t25 = threading.Thread(target=write_list_chairman_of_the_board_of_directors)
t26 = threading.Thread(target=write_list_objective)
t27 = threading.Thread(target=write_list_main_business_activities)

t1.start()
t2.start()
t3.start()
t4.start()
t5.start()
t6.start()
t7.start()
t8.start()
t9.start()
t10.start()
t11.start()
t12.start()
t13.start()
t14.start()
t15.start()
t16.start()
t17.start()
t18.start()
t19.start()
t20.start()
t21.start()
t22.start()
t23.start()
t24.start()
t25.start()
t26.start()
t27.start()

t1.join()
t2.join()
t3.join()
t4.join()
t5.join()
t6.join()
t7.join()
t8.join()
t9.join()
t10.join()
t11.join()
t12.join()
t13.join()
t14.join()
t15.join()
t16.join()
t17.join()
t18.join()
t19.join()
t20.join()
t21.join()
t22.join()
t23.join()
t24.join()
t25.join()
t27.join()
t27.join()
end_thread = time.time()
print(style.BLUE + "Finished thread if %s seconds" % round((end_thread - start_thread), 2))
outWorkbook.close()
print(style.GREEN + "Finished in %s seconds" % round((time.time() - start_time), 2))