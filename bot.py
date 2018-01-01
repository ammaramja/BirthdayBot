'''
Birthday Bot by
- Ammar Amjawala
'''
from selenium import webdriver
from datetime import date
import win32com.client as wc
import json

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

chrome = webdriver.Chrome(r"C:\Users\amjawala_am\PycharmProjects\First\chromedriver.exe", chrome_options=options)

def login_xornet():

    chrome.get('https://xornet.xoriant.com')

    data = json.load(open('db.json'))

    username_path = chrome.find_element_by_name('UserName')
    username_path.send_keys(data["username"])

    password_path = chrome.find_element_by_name('Password')
    password_path.send_keys(data["password"])

    submit_btn = chrome.find_element_by_id('submitButton')
    submit_btn.click()

def birthday_finder():

    list_of_user_profile_links = []
    next_page = True

    while next_page:

        for counter in range(1,6):

            a_tag_xpath = '//*[@id="block-system-main"]/div/div/div/div[1]/div/div[5]/div/div/div/div/div/div/div[1]/div[' + str(counter) + ']/div[2]/span/a'
            check_a_selection = chrome.find_element_by_xpath(a_tag_xpath)
            user_profile_link = check_a_selection.get_attribute('href')

            dob_div_xpath = '//*[@id="block-system-main"]/div/div/div/div[1]/div/div[5]/div/div/div/div/div/div/div[1]/div['+ str(counter) + ']/div[3]/div'
            dob = chrome.find_element_by_xpath(dob_div_xpath).text
            #print (_,_[:2])

            if (int(dob[:2]) == int(date.today().day)):
                list_of_user_profile_links.append(user_profile_link)
                #print (list_of_user_profile_links)
            else:
                next_page = False
                break

            if counter == 5:
                current_page_li = chrome.find_element_by_xpath('//*[@id="block-system-main"]/div/div/div/div[1]/div/div[5]/div/div/div/div/div/div/div[2]/ul/li[2]')
                current_page_number = (current_page_li.text[0])
                chrome.get('https://xornet.xoriant.com/front?page='+current_page_number)

    return (list_of_user_profile_links)

def get_email(list_of_user_profile_links):

    mailing_list = []

    for user_url in list_of_user_profile_links:

        chrome.get(user_url)
        mail_address_div = chrome.find_element_by_xpath('//*[@id="block-system-main"]/div/div/div[3]/div/div[2]/div/div[2]/div/div/div/div/div/p/a')
        #print (mail_address_div.text)
        mailing_list.append(mail_address_div.text)

    return mailing_list

def fire_wishes(mailing_list):

    outlook = wc.Dispatch('Outlook.Application')

    for i in mailing_list:
        '''i = i.split()[0] +" "+ i.split()[len(i.split())-1]
        print (i)
        recipient = outlook.Session.CreateRecipient(i)
        recipient.Resolve()
        print('Resolved OK: ', recipient.Resolved)
        print('Is it a sendable? (address): ', recipient.Sendable)
        print('Name: ', recipient.Name)

        ae = recipient.AddressEntry
        email_address = None

        if 'EX' == ae.Type:
            eu = ae.GetExchangeUser()
            email_address = eu.PrimarySmtpAddress

        if 'SMTP' == ae.Type:
            email_address = ae.Address

        print('Email address: ', email_address)
        '''
        msg = outlook.CreateItem(0)
        msg.To = i
        msg.Subject = 'Happy Birthday!'
        msg.Body = "Hey "+i.split('.')[0]+"!\n\nOn your special day, I wish you good luck. I hope this wonderful day will fill up your heart with joy and blessings. Have a fantastic birthday, celebrate the happiness on every day of your life.\n Happy Birthday :) \n\n -Ammar Amjawala "
        msg.Send()
        print("Mail to "+ i +" has been sent successfully!!!")

if __name__ == '__main__':

    login_xornet()
    list_of_user_profile_links = birthday_finder()
    mailing_list = get_email(list_of_user_profile_links)
    fire_wishes(mailing_list)
    chrome.close()

    print ('\nBot has done its work...\n\nNow you do yours!')
