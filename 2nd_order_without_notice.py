# import essential libraries
import sys
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

# create a browser_instance instance with infobar disabled and password manager disabled
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
prefs = {"credentials_enable_service": False, "profile.password_manager_enabled": False}
chrome_options.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome('venv/chromedriver', options=chrome_options)

# maximizing the browser_instance window make all the links visible.
browser.maximize_window()

# e-mutation login credentials
login_url = 'https://office.land.gov.bd/dashboard/login'

# Prompt for username and password
# username =input("Username: ")
# password = input("Password: ")

username = 'palashmondal'
password = 'ACL@Balaganj008'

# login
browser.get(login_url)
browser.find_element_by_id('username').send_keys(username)
browser.find_element_by_id('password').send_keys(password)
browser.find_element_by_id('submit').click()

# Click the নামজারি application_details_page_url from the left side menu
mutation = browser.find_element_by_link_text("নামজারি")
mutation.click()

# acceptable application list as sorted by application id as decending
proposal_list_url = [
    'https://office.land.gov.bd/mutation/applications/index/24/receive?sort=id&direction=asc',
    'https://office.land.gov.bd/mutation/applications/index/6/receive?sort=created&direction=asc'
]

for proposal_url in proposal_list_url:
    # for each url in url list of 2nd order
    while True:
        # open the acceptable application list page
        browser.get(proposal_url)

        # get all the cases loaded on the first page
        rows = len(browser.find_elements_by_xpath("//*[@id='sample_6']/tbody/tr"))

        # print("Number of cases: " + repr(rows))
        # cols = len (browser_instance.find_elements_by_xpath("//*[@id='sample_6']/tbody/tr[2]/td"))

        # if there is no rows in the acceptable table, then we are done!
        if rows == 0:
            print('All 2nd orders are sent. Job done!\n\n')
            break

        # Let's start working on the first application

        # get the basic information from the table
        application_number = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[1]").text
        application_filling_date = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[2]").text
        applicant_name = browser.find_element_by_xpath("//*[@id='sample_6']/tbody/tr[1]/td[5]").text
        application_details_page_url = browser.find_element_by_link_text("পদক্ষেপ নিন").get_attribute('href')

        # Terminal output of the basic info
        print("----------------------------")
        print("Application Number: " + application_number)
        print("Application Filling Date: " + application_filling_date)
        print("Applicant Name: " + applicant_name)

        # go to the first application details page
        browser.get(application_details_page_url)

        # 10 seconds timeout to load any page before it shows the error message
        wait = WebDriverWait(browser, 15)

        try:
            browser.find_element_by_xpath("//*[@id='order_data']/tr[2]/td[2]")
            print("2nd Order already exists. First fix it manually.")
            sys.exit()
        except NoSuchElementException:
            pass

        # # click 'আদেশ সংযুক্তকরণ' link to add 1st order
        browser.find_element_by_id('add_order').click()

        # # Sleep for 1s
        time.sleep(1)

        # # select the order with "কানুনগো/সার্ভেয়রের প্রতিবেদন ও শুনানির জন্য আদেশ", option id =23
        select = Select(browser.find_element_by_id('status-id'))
        select.select_by_value('23')

        # # to load the full order on ckeditor, wait three seconds
        time.sleep(3)

        # now click 'সংরক্ষণ' button to save the 1st order
        order_submit_button = browser.find_element_by_xpath("//*[@id='submit_order_btn']/button")
        browser.execute_script("arguments[0].click();", order_submit_button)

        # sleep for 2s
        time.sleep(2)

        # refresh the browser_instance to load all dynamically loaded elements as static.
        browser.refresh()

        try:
            # signature modal link xpath
            signature_link = browser.find_element_by_xpath(
                "//*[@id='order_data']/tr[2]/td[2]/div[1]/a")
            # click the signature modal link
            browser.execute_script("arguments[0].click();", signature_link)
        except NoSuchElementException:
            print("Could not put signature on 2nd order.")
            continue

        # on the signature popup modal, first get the 'স্বাক্ষর প্রদান করুন' button
        sign_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='signature']/form/div[2]/button")))
        # Click the 'স্বাক্ষর প্রদান করুন' button
        sign_button.click()

        # get the new mutation case number
        case_number = browser.find_element_by_xpath("//*[@id='case_change_no']").text

        # notice creation and sending

        # click the notice tab on right side
        notice_tab = browser.find_element_by_link_text("নোটিশ/শুনানি")
        notice_tab.click()

        # click the notice full page link within the tab
        # notice_detail_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='view_notice_link_id']")))
        # browser_instance.execute_script("arguments[0].click();", notice_detail_button)

        time.sleep(2)

        # notice submit for the first time
        dakhil_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='notice_submit_id']")))
        # click the send send button finally
        browser.execute_script("arguments[0].click();", dakhil_button)

        # sleep for 1 sec
        time.sleep(1)

        browser.refresh()

        # click the notice tab on right side
        notice_tab = browser.find_element_by_link_text("নোটিশ/শুনানি")
        notice_tab.click()

        # click the notice full page link within the tab
        notice_detail_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='view_notice_link_id']")))
        browser.execute_script("arguments[0].click();", notice_detail_button)

        # signature modal for notice link xpath
        notice_signature_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='portlet_tab5']/div[2]/div/div/table/tbody/tr[10]/td[3]/div[1]/a")))
        # click the signature modal link
        browser.execute_script("arguments[0].click();", notice_signature_link)

        # on the notice signature popup modal, first get the 'স্বাক্ষর প্রদান করুন' button
        sign_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='signature']/form/div[2]/button")))
        # Click the 'স্বাক্ষর প্রদান করুন' button
        sign_button.click()

        # click the notice full page link within the tab
        notice_detail_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='view_notice_link_id']")))
        browser.execute_script("arguments[0].click();", notice_detail_button)

        # click the send notification sms button
        notification_sms_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='portlet_tab5']/div[4]/button[4]")))
        browser.execute_script("arguments[0].click();", notification_sms_button)

        # sleep for 1 sec
        time.sleep(1)

        # close the modal
        notification_sms_button_close = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='send_email_sms']/div[2]/div/div[3]/button")))
        browser.execute_script("arguments[0].click();", notification_sms_button_close)

        # terminal output of assigned information
        print("Case Number: " + case_number)
        print("----------------------------\n")

        # open send popup modal
        popup_button = browser.find_element_by_css_selector(
            "body > div.page-container > div.page-content-wrapper > div > div > div > div > div.portlet-title > "
            "div.actions > button")
        browser.execute_script("arguments[0].click();", popup_button)

        # default receipant is ok

        # Send button xpath
        send_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='submit-btn']")))
        # click the send send button finally
        send_button.click()

        # now repeat the job for the rest of the acceptable application list

# close browser_instance
browser.stop_client()
browser.close()

input("Press Any key to exit")
