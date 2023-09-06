from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.select import Select
import time
import random
import configparser
from selenium.webdriver.chrome.options import Options
options = Options()
options.page_load_strategy = 'none'
options.add_experimental_option('excludeSwitches', ['enable-logging'])

import os


thisfolder = os.path.dirname(os.path.abspath(__file__))
initfile = os.path.join(thisfolder, 'config.txt')
# print thisfolder
print(initfile)

config = configparser.ConfigParser()
# Open config file for variables
config.read(initfile)
# _link = config.get('EPS-VARIABLES', 'Link')
user = config.get('EPS-VARIABLES', 'User')
pw = config.get('EPS-VARIABLES', 'PW')
partname_var = config.get('EPS-VARIABLES', 'PartName')
codename = config.get('EPS-VARIABLES', 'CodeName')
tester = config.get('EPS-VARIABLES', 'Tester')
date = config.get('EPS-VARIABLES', 'Date')
starttime = config.get('EPS-VARIABLES', 'StartTime')
endtime = config.get('EPS-VARIABLES', 'EndTime')
comment = "Please reserve {} for {} ENG activity. Thanks".format(tester, codename)
testerbook = '{}.1 (PROD)'.format(tester)
print(codename)


def delay():
    random_wait_time = random.randrange(2, 3)
    time.sleep(random_wait_time)


# Open browser window
wd = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wd.implicitly_wait(10)
wd.get('https://apex.ams.com/ords/appsprod/f?p=EPS_START:1:104349989965242:::::')
wd.maximize_window()
calamba_button = wd.find_element(By.XPATH, '//*[@id="apex_layout_81395722633610733"]/tbody/tr[2]/td[1]/span/a[2]')
calamba_button.click()
delay()
# Login
username = wd.find_element(By.XPATH, '//*[@id="P9999_USERNAME"]')
username.send_keys(user)
password = wd.find_element(By.XPATH, '//*[@id="P9999_PASSWORD"]')
password.send_keys(pw)
plantID = wd.find_element(By.XPATH, '//*[@id="P9999_PLANT_ID"]')
Select(plantID).select_by_value("1200")
delay()
login_button = wd.find_element(By.XPATH, '//*[@id="apex_layout_1044612172276942210"]/tbody/tr[5]/td/input')
login_button.click()
delay()

# Pick tester from drop down menu
def tester_pick():
    list_ = wd.find_elements(By.CLASS_NAME, "t13data")
    hardwarelist = [x.text for i, x in enumerate(list_) if i == 3 or (i-3) % 7 == 0]
    tester_found = 0
    for coltester, testername in enumerate(hardwarelist):
        if testername != 'TESTER':
            checkbox = '//*[@id="report_R1177667959440791831"]/tbody/tr[' + str(2+coltester) + ']/td[1]/input'
            wd.find_element(By.XPATH, checkbox).click()
        else:
            tester_found = 1
    if tester_found == 0:
        print("No Available Tester in SetupSheet")
        return False
    else:
        return True

# Book reservation
def book_res(partname_var, date, comment, starttime, endtime, tester, testerbook):
    new_reservation_tab = wd.find_element(By.XPATH, '//*[@id="t13PageTabs"]/table/tbody/tr/td[11]/a').click()
    delay()
    partname = wd.find_element(By.XPATH, '//*[@id="P600_PARTNAME"]')
    partname.clear()
    partname.send_keys(partname_var)
    partname.send_keys("\n")
    delay()
    #check if no data found
    try:
        wd.find_element(By.XPATH, '//*[@id="R1175262859552566590"]/tbody/tr/td/span')
        print("No data found in setup sheet")
        return
    except:
        setup_sheet = wd.find_element(By.XPATH, '//*[@id="report_R1175262859552566590"]/tbody/tr[2]/td[2]/a').click()
        delay()
        next_page = wd.find_element(By.XPATH, '//*[@id="R1175109359040223378"]/thead/tr/th[2]/span[3]/a[2]').click()
        delay()
        select_group = wd.find_element(By.XPATH, '//*[@id="P611_GRP_ID"]')
        Select(select_group).select_by_value("57")
        delay()
        begin_date = wd.find_element(By.XPATH, '//*[@id="P611_BEGIN_D|input"]').send_keys(date)
        delay()
        begin_time = wd.find_element(By.XPATH, '//*[@id="P611_BEGIN_T"]')
        Select(begin_time).select_by_value(starttime)
        time.sleep(2)
        end_time = wd.find_element(By.XPATH, '//*[@id="P611_END_T"]')
        time.sleep(2)
        Select(end_time).select_by_value(endtime)
        docking_state = wd.find_element(By.XPATH, '//*[@id="P611_DOK_ID"]/div[1]/label').click()
        delay()
        comment = wd.find_element(By.XPATH, '//*[@id="P611_NOTE"]').send_keys(comment)
        next_page2 = wd.find_element(By.XPATH, '//*[@id="R1176551356029713586"]/thead/tr/th[2]/span[3]/a[2]').click()
        time.sleep(4)
        iframe = wd.find_element(By.XPATH,
                                 '//*[@id="wwvFlowForm"]/div[4]/table/tbody/tr/td[1]/table[2]/tbody/tr/td[3]/iframe')
        wd.switch_to.frame(iframe)
        delay()
        # Call tester pick function
        if tester_pick() == True:
            delete_checkbox = wd.find_element(By.XPATH,
                                              '//*[@id="R1177667959440791831"]/tbody/tr/td/table/tbody/tr/td[1]/span[1]/a[2]').click()
            delay()
            back = wd.switch_to.default_content()
            delete_ok = wd.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/button[2]').click()
            delay()
            iframe = wd.find_element(By.XPATH,
                                     '//*[@id="wwvFlowForm"]/div[4]/table/tbody/tr/td[1]/table[2]/tbody/tr/td[3]/iframe')
            wd.switch_to.frame(iframe)
            wd.implicitly_wait(30)
            time.sleep(3)
            tester = wd.find_element(By.XPATH,
                                     '//*[@id="report_R1177667959440791831"]/tbody/tr[2]/td[6]/span/select').send_keys(
                testerbook)
            time.sleep(3)
            iframe = wd.find_element(By.XPATH,
                                     '//*[@id="wwvFlowForm"]/div[4]/table/tbody/tr/td[1]/table[2]/tbody/tr/td[3]/iframe')
            wd.switch_to.frame(iframe)
            time.sleep(3)
            tester = wd.find_element(By.XPATH,
                                     '//*[@id="report_R1177667959440791831"]/tbody/tr[2]/td[6]/span/select').send_keys(
                testerbook)
            time.sleep(5)
            iframe = wd.find_element(By.XPATH,
                                     '//*[@id="wwvFlowForm"]/div[4]/table/tbody/tr/td[1]/table[2]/tbody/tr/td[3]/iframe')
            wd.switch_to.frame(iframe)
            Next_Reservation = wd.find_element(By.XPATH, '//*[@id="R1177667959440791831"]/thead/tr/th[2]/span[3]/a[2]')
            Next_Reservation.click()
            time.sleep(3)
            finish = wd.find_element(By.XPATH, '//*[@id="R1177888456551851881"]/thead/tr/th[2]/span[3]/a[2]').click()
            time.sleep(3)
        else:
            print("Please check setupsheet or configuration file!")

# For loop for every date filed
for dates in date.split(","):
    try:
        book_res(partname_var, dates, comment, starttime, endtime, tester, testerbook)
        print("No Program error found")
    except:
        print("Error Encountered upon execution! Please check browser or configuration file!")


time.sleep(10)
# close window
wd.close()