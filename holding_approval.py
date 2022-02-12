# coding=utf-8

import os
import time
import socket
import datetime
import configparser
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import UnexpectedAlertPresentException, NoSuchElementException

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
    username = config.get('ULAO', 'username')
    password = config.get('ULAO', 'password')
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


def click_by_xpath(xpath):
    element = browser.find_element_by_xpath(xpath)
    browser.execute_script("arguments[0].click();", element)


def click_by_linktext(linktext):
    element = browser.find_element_by_link_text(linktext)
    element.click()

    # browser.execute_script("arguments[0].click();", element)


def new_case_found(processed=False):
    # time.sleep(1.5)  # temp
    try:
        if browser.find_element_by_link_text("ওপেন"):
            # wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='list']//tbody/tr[1]/td[5]/a")))
            cases = len(browser.find_elements_by_xpath("//*[@id='list']//tbody/tr"))
            # if there is no case, end here.
            if cases == 0:
                if processed:
                    print("All entry processed. Job done! ")
                else:
                    print("No new entry found.")
                browser.stop_client()
                browser.close()
                return False
    except NoSuchElementException:
        browser.get("https://ekhatian.land.gov.bd/dashboard")
        click_by_linktext("অনুমোদন এর জন্য অপেক্ষমাণ")
        return True
    except UnexpectedAlertPresentException:
        alert = browser.switch_to.alert()
        alert.send_keys('8080')
        alert.dismiss()
        browser.get("https://ekhatian.land.gov.bd/dashboard")
        click_by_linktext("অনুমোদন এর জন্য অপেক্ষমাণ")
        return True

    return True


def application_details(details):
    # get the basic information from the table
    details['application_details'] = browser.find_element_by_xpath(
        "//*[@id='list']/tbody/tr[1]/td[2]").text
    details['application_details_page_url'] = browser.find_element_by_link_text(
        "ওপেন").get_attribute('href')
    return details


def send():
    # option click

    click_element = browser.find_element_by_xpath("//*[@id='uniform-checked_khatian']")
    browser.execute_script("arguments[0].click();", click_element)

    # Sleep for 1s
    time.sleep(1)

    # select the 1st order select option, option id =3
    select = Select(browser.find_element_by_id('application_status'))
    select.select_by_value('4')

    time.sleep(0.5)

    # Send button xpath
    send_button = browser.find_element_by_xpath("//*[@id='batch-process-form']/div/div/div/div[7]/div[2]/div/button")
    browser.execute_script("arguments[0].click();", send_button)


def show_info_in_terminal(details):
    # Terminal output
    details['new_time'] = datetime.datetime.now()
    time_lapsed = details['new_time'] - details['old_time']
    print(details['application_details'] + ' ===> Approved! (' + str(time_lapsed.seconds) + ' seconds)')




# first check the internet

if is_connected("one.one.one.one"):

    browser = prepare_browser()

    # 10 seconds timeout to load any page before it shows the error message
    wait = WebDriverWait(browser, 5)

    credentials = credentials()

    if login(browser, credentials) is True:

        browser.implicitly_wait(5)

        # Click খারিজ খতিয়ান link
        click_by_linktext("ভূমি উন্নয়ন কর")

        # Click অনুমোদন এর জন্য অপেক্ষমাণ link
        click_by_linktext("অনুমোদন এর জন্য অপেক্ষমাণ")

        # if new case is found-
        if new_case_found() is True:

            while True:
                # time.sleep(1)
                details = {'old_time': datetime.datetime.now()}
                if new_case_found(processed=True):
                    # new case found

                    application_details(details)

                    # go to the first application details page
                    browser.get(details['application_details_page_url'])
                    # add_first_order()

                    # send
                    send()

                    show_info_in_terminal(details)
                else:
                    # no new case.
                    break

    # end it
    close_session()

else:
    print("No Internet! Fix it first.")
