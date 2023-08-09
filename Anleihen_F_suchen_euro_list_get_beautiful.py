
import msvcrt
import time
import csv
import numpy as np
import pandas as pd 
import random
import xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import  TimeoutException
import contextlib
from selenium.webdriver import Remote
from selenium.webdriver.support.expected_conditions import staleness_of
from POMProjectFolder.Stuttgart.Pages.locators import Locators

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
  "download.default_directory": "/path/to/download/dir",
  "download.prompt_for_download": False,
})
chromedriver_location = "C:\\Users\\lenzre\\Documents\\programs\\chromedriver.exe"
chrome_options.add_argument("--disable-gpu")#OLDER OPTIONS
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])#OLDER OPTIONS
chrome_options.add_argument("--headless")
chrome_options.add_argument("--remote-debugging-port=9222")
driver = webdriver.Chrome(chromedriver_location, options=chrome_options)

driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "/path/to/download/dir"}}
command_result = driver.execute("send_command", params)


#USED 5
def click_operation_id_A(fname):
    try:
        element=driver.find_element_by_id ( fname)
        driver.execute_script("arguments[0].click();", element)
    except NoSuchElementException as exception:
        print ('Click on ' +  fname + '  ID A not successful')
    except TimeoutException:
        pass

#USED 18
def click_operation_Xpath(fname):
    WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname)))
        driver.execute_script("arguments[0].click();", element)
    except TimeoutException:
        pass
    except NoSuchElementException:
        print ('Click on ' + fname+ ' XPath not successful')
#USED 1
def click_operation_B_Xpath(fname):
    WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).click()
        driver.find_element_by_xpath(fname).sendKeys('\uE035')
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' B XPath  not successful')
    except TimeoutException:
        pass

#USED 3
def click_operation_css(fname):
    WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        element= driver.find_element_by_css_selector(fname)
        driver.execute_script("arguments[0].click();", element)
        A=True
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' with css  not successful')
        A=False
    return A

#USED 6
def write_operation_xpath(fname, Input):
    WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        driver.find_element_by_xpath(fname).send_keys(Input)
    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with xpath celector')
        A=False
    except TimeoutException:
        pass

#USED 4
def write_operation_b_xpath(fname, Input):
    WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        element=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).send_keys(Input)
        #driver.find_element_by_xpath(fname).send_keys(Input)
    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with xpath celector')
        A=False

#USED 8
def clear_operation_A_xpath(fname):
    WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        print(" field name "+fname)
        element=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    except NoSuchElementException as exception:
        print ('Clear  ' +fname+ '   not successful')
    except TimeoutException:
        pass

driver.execute_script("document.body.style.zoom='40%'")# works with Chrome
driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
driver.get("https://www.boerse-stuttgart.de/de-de/tools/produktsuche/anleihen-finder/")
#Click Cookie acceptor 
click_operation_B_Xpath(Locators.cookie_acceptor_full_XP)
#Make "tradable unit" button visible
driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
print("Click on Zusaetzliche Filter")
driver.execute_script("document.body.style.zoom='40%'")# 
click_operation_css(Locators.Zusatzliche_Filter_css)
WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
print("Click on checkbox for handelbare einheit")
write_operation_b_xpath(Locators.Zusatzliche_Filter_Textfield_Full_XPath,"Handelbare Einheit")
#Click "Bond tradable unit"

click_operation_id_A(Locators.Handelbare_einh_Button_id)
driver.execute_script("document.body.style.zoom='40%'")# 
click_operation_Xpath(Locators.Zusatzliche_Filter_Anwenden_Full_XPath)
click_operation_Xpath(Locators.Handelbare_einh_Button_xp)
clear_operation_A_xpath(Locators.min_Littera_xp)
write_operation_b_xpath(Locators.min_Littera_xp,"1")
time.sleep(random.randint(2,6))
clear_operation_A_xpath(Locators.max_Littera_xp)
write_operation_b_xpath(Locators.max_Littera_xp,"50.000,000")
click_operation_Xpath(Locators.Anwenden_littera_full_XPath)
# Bond Due date Between two dates
click_operation_Xpath(Locators.Falligkeit_click_xp)
time.sleep(random.randint(2,6))
clear_operation_A_xpath(Locators.Smaller_date_xp)
write_operation_xpath(Locators.Smaller_date_xp,"2.2.2027")
time.sleep(random.randint(2,6))
click_operation_Xpath(Locators.Bigger_date_xp)
clear_operation_A_xpath(Locators.Bigger_date_2_xp)
write_operation_xpath(Locators.Bigger_date_2_xp,"1.4.2035")
click_operation_Xpath(Locators.Anwenden_Falligkeit_full_XPath)
driver.implicitly_wait(10)
#Gelisteter Zeitraum
click_operation_css(Locators.Zusatzliche_Filter_css)
WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
write_operation_b_xpath(Locators.Zusatzliche_Filter_Textfield_Full_XPath,"Gelist")
time.sleep(random.randint(2,6))
#T‰ss‰ pit‰‰ olla Xpath
click_operation_id_A(Locators.Gelisteter_Zeitraum_preselector_Button_id)
click_operation_Xpath(Locators.Zusatzliche_Filter_Anwenden_Full_XPath)
click_operation_Xpath(Locators.Gelisteter_Zeitraum_Button_xp)
time.sleep(random.randint(2,6))
clear_operation_A_xpath(Locators.Gelisteter_Zeitraum_smaller_date_xp)
write_operation_xpath(Locators.Gelisteter_Zeitraum_smaller_date_xp,"1.1.2020")
click_operation_Xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp)
time.sleep(random.randint(2,6))
clear_operation_A_xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp)
write_operation_xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp,"1.8.2023")
time.sleep(random.randint(2,6))
click_operation_Xpath(Locators.Anwenden_Gelisteter_Zeitraum_full_XPath)
#Select Unternehmensanleihe
click_operation_id_A(Locators.Preselector_Anleihen_Typ_ID)
click_operation_id_A(Locators.Select_Unternehmensanleihe_Checkbox_ID)
click_operation_Xpath(Locators.Anwenden_Unternehmensanleihe_Select_full_XPath)
# Choose Yield between maximum and minimum
click_operation_Xpath(Locators.Rendite_button_xp)
clear_operation_A_xpath(Locators.Rendite_lower_margin_xp)
write_operation_xpath(Locators.Rendite_lower_margin_xp, "3%")
clear_operation_A_xpath(Locators.Rendite_upper_margin_xp)
write_operation_xpath(Locators.Rendite_upper_margin_xp, "16%")
time.sleep(random.randint(2,6))
click_operation_Xpath(Locators.Rendite_anwenden_xp)
#Choosing the acceptable currency Euro
click_operation_Xpath(Locators.Wahrung_click_xp)
time.sleep(random.randint(2,6))
click_operation_id_A(Locators.Wahrung_click_new_ID)
print("Click Anwenden of euro ")
time.sleep(random.randint(2,6))
click_operation_css(Locators.Wahrung_anwenden_css)
#Sort by yield
time.sleep(random.randint(2,6))
driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
#show_more_hits
click_operation_Xpath(Locators.Weitere_Treffer_anzeigen_xp)
_j=1
_n=5
while _j<_n:
    driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
    time.sleep(random.randint(2,6))
    click_operation_Xpath(Locators.Weitere_Treffer_anzeigen_xp)
    _j+=1
    time.sleep(random.randint(2,6))
# Write Results to Excel file
html_source =driver.page_source
soup=BeautifulSoup(html_source, 'html.parser')
table=soup.table
table_rows = table.find_all('tr')
tables=[]
for tr in table_rows:
    td = tr.find_all('td')
    row=[]
    for i in td:
        word= i.text.strip()
        word=word.replace(".000", "000") #littera remove seperators
        word=word.replace(",000", "") #littera remove zeros after comma
        row.append(word)
    tables.append(row)
print(tables)
df = pd.DataFrame(tables)
df.to_excel(excel_writer = "c:\\temp\Lib\POMProjectFolder\Stuttgart\Test\\test.xlsx")


