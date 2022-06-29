# coding=utf-8


import configparser
import csv
import os
import re
import socket
import time
from datetime import date

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait

# create a browser_instance instance with infobar disabled and password manager disabled
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
prefs = {"credentials_enable_service": False, "profile.password_manager_enabled": False}
chrome_options.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
os.system('cls' if os.name == 'nt' else 'clear')
# maximizing the browser_instance window make all the links visible.
browser.maximize_window()



# e-mutaion login credentials
login_url = 'https://office.land.gov.bd/dashboard/login'

# Prompt for username and password
# username =input("Username: ")
# password = input("Password: ")

username = 'palashmondal'
password = 'ACL@Mohanagar008'

# login
browser.get(login_url)
browser.find_element_by_id('username').send_keys(username)
browser.find_element_by_id('password').send_keys(password)
browser.find_element_by_id('submit').click()

# Click the নামজারি application_details_page_url from the left side menu
mutation = browser.find_element_by_link_text("নামজারি")
mutation.click()

# goto all all cases listing page
all_case_url = 'https://office.land.gov.bd/mutation/running_applications'

# define wait: 15 seconds timeout to load any page before it shows the error message
wait = WebDriverWait(browser, 15)

# exit warning
print("If you want to exit the script, enter 'e'")

while True:

    # prompt for case number before going to the all case page
    case_number = input("Case Number: ")

    # if anyone want to exit the script
    if case_number == 'e':
        browser.stop_client()
        browser.close()
        break

    # goto all case page
    browser.get(all_case_url)

    # now search for a case number
    search_case = browser.find_element_by_id('case-no')
    search_case.send_keys(case_number)

    # now click the 'খুঁজুন' button
    action_button = browser.find_element_by_xpath("//*[@id='search_case_no_section']/form/div[2]/div[4]/button")
    action_button.click()

    # get all search result page
    rows = len(browser.find_elements_by_xpath("//*[@id='sample_6']/tbody/tr"))
    # don't need the info of case table, so comment it out for now.
    # cols = len (browser_instance.find_elements_by_xpath("//*[@id='sample_6']/tbody/tr[2]/td"))

    # if there is no rows in the acceptable table, then we are done!
    if rows == 0:
        print("No such case found! Try again.")
        continue

    # check if the element exists, then if it is on ac land's account, it will show 'পদক্ষেপ নিন'.
    try:
        case_found = browser.find_element_by_link_text("পদক্ষেপ নিন")
    except NoSuchElementException:
        print("No such case found! Try again.")
        continue

    # get the basic information from the table
    # get application number
    # application_number = browser_instance.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[1]").text
    # get the single case details page url
    application_details_page_url = case_found.get_attribute('href')

    print('Great! Case Found! ')

    # go to the desired application details page
    browser.get(application_details_page_url)

