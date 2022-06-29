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



def show_info_in_terminal(details):
    # Terminal output
    print("Case: " +details['desk']+ ' is at '+ details['desk'] +' desk.')



def get_cases_from_txt():
    file = open('nothijat.txt', 'r')
    cases = file.readlines()
    file.close()
    return cases



def search_case():
    # now search for a case number
    search_case = browser.find_element_by_id('case-no')
    search_case.send_keys(case_number)

    # now enter the 'খুঁজুন' button
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
    details['case'] = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[4]").text
    details['desk'] = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[7]").text
    
    return details

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

        cases = get_cases_from_txt()

        all_case_url = config.get('URLS', 'all_case_page')

        problem_cases = []

        if len(get_cases_from_txt()) !=0:
            # got some cases

            # for each case from the nothijat.txt file on the current directory.
            for case_number in cases:

                # goto all case page
                browser.get(all_case_url)
                
                details = {}
                
                

                if search_case():
                    # case found.

                        check_case_position() 
                        
                        show_info_in_terminal(details)
                        
                        data = details['case']+ ';'+ details['desk']
                        
                        
                        with open('desk.txt', 'a', encoding="utf-8") as fd:
                            fd.write(f'\n{data}')
                            
                        fd.close()

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