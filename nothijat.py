# coding=utf-8


import configparser
import csv
import os
import sys
import re
import socket
import time
from datetime import date
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


# parse configuration file
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')


print(chr(27) + "[2J")



def is_connected(hostname):
    try:
        # see if we can resolve the host name -- tells us if there is
        # a DNS listening
        host = socket.gethostbyname(hostname)
        # connect to the host -- tells us if the host is actually
        # reachable
        s = socket.create_connection((host, 80), 2)
        s.close()
        return True
    except:
        pass
    return False


def credentials():
    # get username and password from config file
    username = config.get('LOGIN', 'username')
    password = config.get('LOGIN', 'password')
    return {'username': username, 'password': password}


def prepare_browser():
    # create a browser_instance instance with infobar disabled and password manager disabled
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
    prefs = {"credentials_enable_service": False, "profile.password_manager_enabled": False}
    chrome_options.add_experimental_option("prefs", prefs)
    browser = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    os.system('cls' if os.name == 'nt' else 'clear')
    # maximizing the browser_instance window make all the links visible.
    browser.maximize_window()
    return browser


def close_session():
    browser.stop_client()
    browser.quit()


def login(browser, credentials):
    # land on login page
    login_url = config.get('URLS', 'login_page')
    welcome_page = config.get('URLS', 'welcome_page')

    browser.get(login_url)

    username = credentials['username']
    password = credentials['password']

    # try login
    browser.find_element_by_id('username').send_keys(username)
    browser.find_element_by_id('password').send_keys(password)
    browser.find_element_by_id('submit').click()

    # if login unsuccessful
    if browser.current_url != welcome_page:
        print('login error! ')
        close_session()
        return False
    else:
        print("Login successful! ")
        return True


def click_by_linktext(linktext):
    element = browser.find_element_by_link_text(linktext)
    browser.execute_script("arguments[0].click();", element)


def application_details():
    details = {}
    # get the basic information from the table
    details['application_number'] = browser.find_element_by_xpath(
        "//*[@id='sample_6']/tbody/tr[1]/td[1]").text
    details['application_filling_date'] = browser.find_element_by_xpath(
        "//*[@id='sample_6']/tbody/tr[1]/td[2]").text
    details['applicant_name'] = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[5]").text
    details['mouza'] = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[6]").text

    return details


def show_info_in_terminal(details):
    # Terminal output
    print("----------------------------")
    print("Application Number: " + details['application_number'])
    print("Application Filling Date: " + details['application_filling_date'])
    print("Applicant Name: " + details['applicant_name'])
    print("Mouza: " + details['mouza'])
    print("Case Number: " + details['new_case_number'])
    print("----------------------------")


def get_cases_from_txt():
    file = open('nothijat.txt', 'r')
    cases = file.readlines()
    file.close()
    return cases



def search_case():
    # now search for a case number
    search_case = browser.find_element_by_id('case-no')
    search_case.send_keys(case_number)

    # now enter the '??????????????????' button
    action_button = browser.find_element_by_xpath(
        "//*[@id='search_case_no_section']/form/div[2]/div[4]/button")
    action_button.click()

    # get search result
    rows = len(browser.find_elements_by_xpath("//*[@id='sample_6']/tbody/tr"))

    if rows ==0:
        return False
    else:
        return True


def check_case_position():

    # get the basic information from the table
    #application_number = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[1]").text

    # get the url of the case
    try:
        application_details_page_url = browser.find_element_by_link_text("????????????????????? ?????????").get_attribute(
            'href')
        return application_details_page_url
    except NoSuchElementException:
        # case is not in acland's account.
        return False


def add_order():

    nothijat_text = config.get('ORDERS','nothijat_text')

    # click '???????????? ??????????????????????????????' link to add 1st order
    browser.find_element_by_id('add_order').click()

    # Sleep for 1s
    time.sleep(2)

    # select the order with "?????????????????????/?????????????????????????????? ??????????????????????????? ??? ????????????????????? ???????????? ????????????", option id =23
    select = Select(browser.find_element_by_id('status-id'))
    select.select_by_value('9')

    # change the order text
    frame = browser.find_element(By.CLASS_NAME, 'cke_wysiwyg_frame')
    browser.switch_to.frame(frame)
    time.sleep(1.5)
    body = browser.find_element(By.TAG_NAME, 'body')
    body.clear()
    time.sleep(2)
    #body.click()
    body.send_keys(nothijat_text)
    body.send_keys(Keys.RETURN)
    time.sleep(2)
    browser.switch_to.parent_frame()

    # now click '?????????????????????' button to save the positive 3rd order
    order_positive_submit_button = browser.find_element_by_xpath("//*[@id='submit_order_btn']/button")
    browser.execute_script("arguments[0].click();", order_positive_submit_button)

    # sleep for 2s
    time.sleep(1.5)

    # refresh the browser_instance to load all dynamically loaded elements as static.
    browser.refresh()
   
   #send
   
       # click '??????????????????' button to open the sending modal
    send_button = browser.find_element_by_css_selector(
        "body > div.page-container > div.page-content-wrapper > div > div > div > div > div.portlet-title > div.actions > button")
    browser.execute_script("arguments[0].click();", send_button)

    time.sleep(1)
    
    # select the destination from radio buttons
    browser.find_element_by_xpath("//input[@name='users_tagged_receive[3]']").click()
    
    # Send button xpath
    send_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='submit-btn']")))
    # click the send send button finally

    send_button.click()

    
    


def remove_case_from_txt_file(case_number):

    with open("nothijat.txt", "r") as f:
        lines = f.readlines()

    with open("nothijat.txt", "w") as f:
        for line in lines:

            if line.strip("\n") != str(case_number):
                f.write(line)
    f.close()

if is_connected("one.one.one.one"):

    browser = prepare_browser()
    # 10 seconds timeout to load any page before it shows the error message
    wait = WebDriverWait(browser, 10)
    credentials = credentials()

    if login(browser, credentials) is True:

        # Click ????????????????????? link
        click_by_linktext("?????????????????????")
        
        time.sleep(1)
        browser.refresh()

        cases = get_cases_from_txt()

        all_case_url = config.get('URLS', 'all_case_page')

        problem_cases = []

        if len(get_cases_from_txt()) !=0:
            # got some cases
            # do everything


            # for each case from the nothijat.txt file on the current directory.
            for case_number in cases:

                # goto all case page
                browser.get(all_case_url)

                if search_case():
                    # case found.

                    if check_case_position() is not False:
                        # case is in ACL's account.
                        application_details_page_url = check_case_position()

                        # go to the first application details page
                        browser.get(application_details_page_url)

                        if add_order() is True:
                            # send
                            #send()
                            # Terminal output of the basic info
                            print("Case Number: " + case_number + ' is archived!')
                            # now repeat the job for the rest of the acceptable application list
                        else:
                            # could not add order,
                            print("Case Number: " + case_number + " - 3rd order already there! Work manually.")

                    else:
                        # not at acland's account.
                        print("Case Number: " + case_number + ' - case is not at AC(Land)\'s account!')
                else:
                    # case not found.
                    print("Case Number: " + case_number + ' - case not found!')
        else:
            print("No case found in text file.")

            close_session()

    # end it
    close_session()

else:
    print("No Internet! Fix it first.")