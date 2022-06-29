

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

# parse configuration file
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')


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
    details['application_details_page_url'] = browser.find_element_by_link_text(
        "পদক্ষেপ নিন").get_attribute('href')

    return details


def show_info_in_terminal(details):
    # Terminal output
    print("----------------------------")
    print("Application Number: " + details['application_number'])
    print("Application Filling Date: " + details['application_filling_date'])
    print("Applicant Name: " + details['applicant_name'])
    print("Mouza: " + details['mouza'])
    print("----------------------------")


def rejectable_application_found(url, processed=False):

    # open the rejectable application list page
    browser.get(url)

    # get all the cases loaded on the first page
    cases = len(browser.find_elements_by_xpath("//*[@id='sample_6']/tbody/tr"))

    if cases == 0:
        if processed:
            print("All rejectable applications processed. Job done! ")
        else:
            print("No rejectable application found.")

        browser.stop_client()
        browser.close()
        return False

    return True


if is_connected("one.one.one.one"):

    browser = prepare_browser()
    # 10 seconds timeout to load any page before it shows the error message
    wait = WebDriverWait(browser, 10)
    credentials = credentials()

    if login(browser, credentials) is True:

        # Click নামজারি link
        click_by_linktext("নামজারি")
                
        time.sleep(1)
        browser.refresh()

        # rejectable application list as sorted by application id as decending
        rejectable_applications_url = config.get('URLS', 'rejectable_applications_list')
        reject_text = config.get('ORDERS','reject_text')

        if rejectable_application_found(url = rejectable_applications_url):
            while True:

                # go to the first application details page
                browser.get(rejectable_applications_url)

                if rejectable_application_found(url=rejectable_applications_url, processed=True):
                    # rejectable application found

                    # get the basic information from the table
                    details = application_details()

                    # go to the first application details page
                    browser.get(details['application_details_page_url'])

                    print('i am here')
                    time.sleep(1)

                    # on the notice signature popup modal, first get the 'বাতিল' button
                    reject_button = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='portlet_tab1']/div[2]/button[5]")))
                    # Click the 'বাতিল' button
                    browser.execute_script("arguments[0].click();", reject_button)

                    time.sleep(1)

                    # enter the text 'প্রয়োজনীয় কাগজাদি না থাকায় নথিজাত করা হল।' on opinion text area.
                    opinion_text = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//*[@id='cancelable-final-reasons']")))
                    opinion_text.clear()
                    opinion_text.send_keys(reject_text)

                    # on the notice signature popup modal, first get the 'বাতিল' button
                    yes_button = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='canceled_form']/div[2]/div/div[3]/button[1]")))
                    # Click the 'বাতিল' button
                    browser.execute_script("arguments[0].click();", yes_button)
                    # done !

                    show_info_in_terminal(details)

        else:
            close_session()
    # end it
    close_session()

else:
    print("No Internet! Fix it first.")
