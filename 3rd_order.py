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
    return True


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


def click_by_linktext(linktext):
    element = browser.find_element_by_link_text(linktext)
    browser.execute_script("arguments[0].click();", element)

def wait_and_click_by_xpath(wait, xpath):
    button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    # click the send send button finally
    browser.execute_script("arguments[0].click();", button)


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
    wait_and_click_by_xpath(wait, "//*[@id='loginBtn']")
    current_url = browser.current_url
    WebDriverWait(browser, 15).until(EC.url_changes(current_url))
    
    # if login unsuccessful
    if browser.current_url != welcome_page:
        print('login error! ')
        close_session()
        #browser.get(welcome_page)
        return True
    else:
        print("Login successful! ")
        return True


# goto all all cases listing page
all_case_page = config.get('URLS','all_case_page')

# exit warning
print("If you want to exit the script, enter 'e'")



# first check the internet

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

        # acceptable application list as sorted by application id as decending
        acceptable_applications = config.get('URLS', 'acceptable_applications_page_asc')

 
        while True:

            # prompt for case number before going to the all case page
            case_number = input("Case Number: ")

            # if anyone want to exit the script
            if case_number == 'e':
                browser.stop_client()
                browser.close()
                break

            # goto all case page
            browser.get(all_case_page)

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

            # click the 'প্রস্তাবপত্র' tab
            browser.find_element_by_link_text("প্রস্তাবপত্র").click()

            # proposal_tab = browser_instance.find_element_by_link_text("প্রস্তাবপত্র")
            # browser_instance.execute_script("arguments[0].click();", proposal_tab)

            # zoom proposal page popup modal
            browser.execute_script("document.body.style.zoom='130%'")

            # click 'প্রশস্ত করে দেখুন' button

            proposal_big_view_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='portlet_tab3']/div[2]/button[2]")))
            browser.execute_script("arguments[0].click();", proposal_big_view_button)

            # get the decision from terminal
            print("Specify your proposal decision: y=yes, n=no, p = pending.")
            proposal_decision = input("Decision: ")

            # close the 'প্রশস্ত করে দেখুন' popup modal
            proposal_big_view_close = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='view_proposal_large']/div[2]/div/div[3]/button[3]")))
            

            while proposal_decision and proposal_decision[0].lower() not in ['y', 'n', 'p']:
                
                print("Specify your proposal decision: y=yes, n=no, p = pending.")
                proposal_decision = input("Decision: ")
                
               
                # now ask for the decision of the proposal
                if proposal_decision == 'y':
                
                    # close the popup modal
                    browser.execute_script("arguments[0].click();", proposal_big_view_close)

                    # click 'নামজারি খতিয়ান' tab
                    khatian_tab = browser.find_element_by_link_text("নামজারি খতিয়ান")
                    browser.execute_script("arguments[0].click();", khatian_tab)

                    # 'প্রশস্ত করে দেখুন' on 'নামজারি খতিয়ান' tab
                    khatian_big_view_button = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='portlet_tab6']/div[2]/button")))
                    browser.execute_script("arguments[0].click();", khatian_big_view_button)

                    # get the khatian decision from terminal
                    print("Specify your khatian decision: y=yes, n=no")
                    khatian_decision = input("Decision: ")

                    khatian_popup_close = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='addKhotianForm']/div[4]/button")))
                    browser.execute_script("arguments[0].click();", khatian_popup_close)

                    if proposal_decision == 'y' and khatian_decision == 'y':
                        #click মতামত দিন on নামজারি খতিয়ান
                        k_sign_button = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//*[@id='portlet_tab6']/div[2]/button[2]")))
                        browser.execute_script("arguments[0].click();", k_sign_button)

                        time.sleep(0.5)

                        # enter text মতামত for khatian
                        khatian_text = wait.until(EC.presence_of_element_located(
                            (By.XPATH, "//*[@id='khotian_comment']/div[2]/div/form/div[2]/textarea")))
                        khatian_text.clear()
                        khatian_text.send_keys('অনুমোদিত।')

                        # save the নামজারি মতামত
                        khatian_comment_save_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='khotian_comment']/div[2]/div/form/div[3]/button[2]")))
                        browser.execute_script("arguments[0].click();", khatian_comment_save_button)

                        time.sleep(0.5)
                        # click the browser_instance alert box 'yes'
                        WebDriverWait(browser, 10).until(EC.alert_is_present())
                        browser.switch_to.alert.accept()

                        time.sleep(0.5)
                        # this tab takes a little bit longer to load
                        longer_wait = WebDriverWait(browser, 25)

                        # now click 'স্বাক্ষর প্রদান করুন' button
                        #khatian_opinion_sign_button = longer_wait.until(
                        #   EC.element_to_be_clickable((By.XPATH, "//*[@id='portlet_tab6']/div[1]/div[1]/div[3]/div[2]/a")))
                        # Click the 'স্বাক্ষর প্রদান করুন' button
                        #browser.execute_script("arguments[0].click();", khatian_opinion_sign_button)

                        # now click 'সংরক্ষণ' button to save the khatian opinion
                        #khatian_opinion_submit_button = browser.find_element_by_xpath("//*[@id='signature']/form/div[2]/button")
                        #browser.execute_script("arguments[0].click();", khatian_opinion_submit_button)

                        # before doing anything, just reset the page
                        browser.get(application_details_page_url)

                        # on প্রস্তাবপত্র tab

                        # click the 'প্রস্তাবপত্র' tab again to make sure it is on the right context
                        proposal_tab = browser.find_element_by_link_text("প্রস্তাবপত্র")
                        browser.execute_script("arguments[0].click();", proposal_tab)

                        # click 'মতামত দিন' button on 'প্রস্তাবপত্র' tab
                        opinion_button = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//*[@id='portlet_tab3']/div[2]/button[3]")))
                        browser.execute_script("arguments[0].click();", opinion_button)

                        time.sleep(0.5)

                        # enter the text 'প্রস্তাব অনুমোদিত।' on opinion text area.
                        opinion_text = wait.until(
                            EC.presence_of_element_located((By.XPATH, "//*[@id='add_comment']/div[2]/textarea")))
                        opinion_text.clear()
                        opinion_text.send_keys('প্রস্তাব অনুমোদিত।')

                        # save the opinion
                        opinion_save_button = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//*[@id='add_comment']/div[3]/button[2]")))
                        browser.execute_script("arguments[0].click();", opinion_save_button)

                        time.sleep(0.5)

                        # click the 'প্রস্তাবপত্র' tab again to make sure it is on the right context
                        proposal_tab = browser.find_element_by_link_text("প্রস্তাবপত্র")
                        browser.execute_script("arguments[0].click();", proposal_tab)

                        # click 'স্বাক্ষর প্রদান করুন' button
                        #proposal_sign_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='comment']/div/div[2]/a")))
                        #browser.execute_script("arguments[0].click();", proposal_sign_button)

                        # now click 'সংরক্ষণ' button to save opinion
                        #proposal_sign_submit_button = browser.find_element_by_xpath("//*[@id='signature']/form/div[2]/button")
                        #browser.execute_script("arguments[0].click();", proposal_sign_submit_button)

                        # now go the the single case page again
                        browser.get(application_details_page_url)

                        # click 'আদেশ সংযুক্তকরণ' link to add positive 3rd order
                        browser.find_element_by_id('add_order').click()

                        # Sleep for 1s
                        time.sleep(1)

                        # select the order with "শুনানি গ্রহণ ও অনুমোদনের আদেশ", option id =19
                        order_positive_select = Select(browser.find_element_by_id('status-id'))
                        order_positive_select.select_by_value('19')

                        # to load the full order on ckeditor, wait two seconds
                        time.sleep(1)

                        # now click 'সংরক্ষণ' button to save the positive 3rd order
                        order_positive_submit_button = browser.find_element_by_xpath("//*[@id='submit_order_btn']/button")
                        browser.execute_script("arguments[0].click();", order_positive_submit_button)

                        # sleep for 2s
                        time.sleep(1)

                        # refresh the browser_instance to load all dynamically loaded elements as static.
                        browser.refresh()

                        # click signature button to open signature modal
                        #order_signature_link = browser.find_element_by_xpath("//*[@id='order_data']/tr[3]/td[2]/div[2]/a")
                        #browser.execute_script("arguments[0].click();", order_signature_link)

                        # on the signature popup modal, first get the 'স্বাক্ষর প্রদান করুন' button
                        #order_signature_button = wait.until(
                        #   EC.element_to_be_clickable((By.XPATH, "//*[@id='signature']/form/div[2]/button")))
                        #order_signature_button.click()

                        # click 'প্রেরণ' button to open the sending modal
                        send_button = browser.find_element_by_css_selector(
                            "body > div.page-container > div.page-content-wrapper > div > div > div > div > div.portlet-title > div.actions > button")
                        browser.execute_script("arguments[0].click();", send_button)

                        time.sleep(1)

                        # select the recipients প্রেরণ = অফিস সহকারী, অনুলিপি = সহকারী কমিশনার (ভূমি)
                        browser.find_element_by_xpath("//input[@name='users_tagged_receive[3]']").click()
                        # browser_instance.find_element_by_xpath("//input[@name='users_tagged_view[4]']").click()

                        # Finally SEND
                        send_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='submit-btn']")))
                        send_button.click()

                    if khatian_decision == 'n':
                        continue

                elif proposal_decision == 'n':

                    # before doing anything, just reset the page
                    browser.get(application_details_page_url)

                    # click 'আদেশ সংযুক্তকরণ' link to add negative 3rd order
                    add_negative_order_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='add_order']")))
                    browser.execute_script("arguments[0].click();", add_negative_order_button)

                    time.sleep(1)

                    # select the order with "নামঞ্জুর", option id =9
                    select_negative_order = Select(browser.find_element_by_id('status-id'))
                    select_negative_order.select_by_value('9')

                    # change the order text
                    frame = browser.find_element(By.CLASS_NAME, 'cke_wysiwyg_frame')
                    browser.switch_to.frame(frame)
                    time.sleep(1)
                    body = browser.find_element(By.TAG_NAME, 'body')
                    body.clear()
                    body.click()
                    body.send_keys('রেকর্ডের সাথে মিল না থাকায় ও প্রয়োজনীয় তথ্য না থাকায় আবেদনটি বাতিল করা হলো ।')
                    body.send_keys(Keys.RETURN)
                    time.sleep(1)
                    browser.switch_to.parent_frame()

                    # now click 'সংরক্ষণ' button to save the negative 3rd order
                    order_negative_submit_button = browser.find_element_by_xpath("//*[@id='submit_order_btn']/button")
                    browser.execute_script("arguments[0].click();", order_negative_submit_button)

                    # sleep for 2s
                    time.sleep(1)

                    # refresh the browser_instance to load all dynamically loaded elements as static.
                    browser.refresh()

                    # signature modal link
                    order_negative_signature_link = browser.find_element_by_xpath("//*[@id='order_data']/tr[3]/td[2]/div[1]/a")
                    browser.execute_script("arguments[0].click();", order_negative_signature_link)

                    # on the signature popup modal, click 'স্বাক্ষর প্রদান করুন' button
                    #sign_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='signature']/form/div[2]/button")))
                    #sign_button.click()

                    # open send popup modal
                    popup_button = browser.find_element_by_css_selector(
                        "body > div.page-container > div.page-content-wrapper > div > div > div > div > div.portlet-title > div.actions > button")
                    browser.execute_script("arguments[0].click();", popup_button)
                    time.sleep(2)
                    # select the destination from radio buttons
                    destination_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='users-tagged-receive-3']")))
                    destination_button.click()

                    # Finally SEND
                    send_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='submit-btn']")))
                    send_button.click()

                elif proposal_decision == 'p':

                    continue

                else:
                    print('Invalid Decision. Please try again! ')
                    continue

                print('---------------------')
