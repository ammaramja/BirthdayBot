'''
Birthday Bot by
- Ammar Amjawala
'''
import json
import time
from datetime import date

import win32com.client as wc
from selenium import webdriver
import selenium.webdriver.support.expected_conditions as EC
import selenium.webdriver.support.ui as ui
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("disable-infobars")
chrome = webdriver.Chrome(r"C:\Users\amjawala_am\PycharmProjects\chromedriver.exe", chrome_options=options)


# return True if element is visible within 60 seconds, otherwise False
def is_visible(locator, timeout=600):
    try:
        ui.WebDriverWait(chrome, timeout).until(EC.visibility_of_element_located((By.XPATH, locator)))
        return True
    except TimeoutException:
        return False


def login_xornet():
    chrome.get('https://xornet.xoriant.com')

    data = json.load(open(r'timesheet_data\db.json'))

    username_xpath = '//*[@id="userNameInput"]'
    if not is_visible(username_xpath): return "Load error"

    username_field = chrome.find_element_by_xpath(username_xpath)
    username_field.send_keys(data["username"])

    password_xpath = '//*[@id="passwordInput"]'
    if not is_visible(password_xpath): return "Load error"

    password_field = chrome.find_element_by_xpath(password_xpath)
    password_field.send_keys(data["password"])

    submit_btn = chrome.find_element_by_id('submitButton')
    submit_btn.click()


def find_birthdays():
    birthday_count_xpath = '//*[@id="birthdays-block-form--2"]/div/div[1]/span[2]'
    if not is_visible(birthday_count_xpath): return "Load error"
    birthday_count = chrome.find_element_by_xpath(birthday_count_xpath).text
    birthday_count = int(birthday_count.split("(")[-1][:-1])
    chrome.get("https://xornet.xoriant.com/birthdays")

    mailing_list = []

    for count in range(birthday_count):
        a_tag_xpath = '//*[@id="mail_content_' + str(count) + '"]'
        # if not is_visible(a_tag_xpath): return "Could not load a tag xpath"
        check_a_selection = chrome.find_element_by_xpath(a_tag_xpath)
        email_id = check_a_selection.get_attribute('href').split('?')[0][7:]
        if "xoriant.com" in email_id or "Xoriant.Com" in email_id:
            mailing_list.append(email_id)

    return mailing_list


def fire_wishes(mailing_list):
    outlook = wc.Dispatch('Outlook.Application')
    flag = 0
    for i in mailing_list:
        msg = outlook.CreateItem(0)
        msg.To = i
        msg.Subject = 'Happy Birthday!'
        msg.Body = "Hey " + i.split('.')[
            0].capitalize() + "!\n\nBirthdays are a new start, a fresh beginning and a time to pursue new endeavors with new goals." \
                 " Move forward with confidence and courage." \
                 " May today and all of your days be amazing!\nWish you and your family a very safe and health future. Once again, Happy Birthday :) \n\n" \
                 "Best Regards,\nAmmar Amjawala "
        msg.Send()
        print("Mail to " + i + " has been sent successfully!!!")
    print('\nBot has done its work...\n\nNow you do yours!')


if __name__ == '__main__':
    login_xornet()
    mailing_list = find_birthdays()
    fire_wishes(mailing_list)
    chrome.close()
