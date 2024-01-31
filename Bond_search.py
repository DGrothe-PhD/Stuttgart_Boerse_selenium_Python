import config
import web_actions
import os
#import msvcrt
import time
import datetime
import csv
import numpy as np
import pandas as pd 
import random
import xlsxwriter
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from locators import Locators
from datetime import date
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl import load_workbook
import clipboard as clp

config_webdriver_manager.driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "/path/to/download/dir"}}
command_result = config_webdriver_manager.driver.execute("send_command", params)

web_actions.zoom40()
config_webdriver_manager.driver.get("https://www.boerse-stuttgart.de/de-de/tools/produktsuche/anleihen-finder/")
#Click Cookie acceptor 
web_actions.click_operation_B_Xpath(Locators.cookie_acceptor_full_XP)
#Make "tradable unit" button visible
config_webdriver_manager.driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")

print("Click on Zusaetzliche Filter")
config_webdriver_manager.driver.execute_script("document.body.style.zoom='40%'")# 
web_actions.click_operation_css(Locators.Zusatzliche_Filter_css)
web_actions.relax()
print("Click on checkbox for handelbare einheit")
web_actions.write_operation_b_xpath(Locators.Zusatzliche_Filter_Textfield_Full_XPath,"Handelbare Einheit")
#Click "Bond tradable unit"
web_actions.click_operation_id_A(Locators.Handelbare_einh_Button_id)
config_webdriver_manager.driver.execute_script("document.body.style.zoom='40%'")# 
web_actions.click_operation_Xpath(Locators.Zusatzliche_Filter_Anwenden_Full_XPath)
web_actions.click_operation_Xpath(Locators.Handelbare_einh_Button_xp)
web_actions.clear_operation_A_xpath(Locators.min_Littera_xp)
web_actions.write_operation_b_xpath(Locators.min_Littera_xp,"1")
time.sleep(random.randint(2,6))
web_actions.clear_operation_A_xpath(Locators.max_Littera_xp)
web_actions.write_operation_b_xpath(Locators.max_Littera_xp,"50.000,000")
web_actions.click_operation_Xpath(Locators.Anwenden_littera_full_XPath)
#Choosing the acceptable currency Euro
web_actions.click_operation_Xpath(Locators.Wahrung_click_xp)
web_actions.relax()
web_actions.write_operation_b_xpath(Locators.Zusatzliche_Filter_Textfield_Wahrung_Full_XPath,"Euro")
time.sleep(random.randint(2,6))
#web_actions_webd.click_operation_id_A(Locators.Wahrung_click_new_ID) #unreliable
web_actions.click_operation_Xpath(Locators.Wahrung_checkbox_Euro_full_XPath)
print("Milestone Click Anwenden of euro through")
time.sleep(random.randint(2,6))
web_actions.click_operation_css(Locators.Wahrung_anwenden_css)
# Bond Due date Between two dates
web_actions.click_operation_Xpath(Locators.Falligkeit_click_xp)
time.sleep(random.randint(2,6))
web_actions.clear_operation_A_xpath(Locators.Smaller_date_xp)
Smaller_date=  datetime.date.today()+ relativedelta(years=5)
Smaller_date_format = Smaller_date.strftime("%d.%m.%Y")
web_actions.write_operation_xpath(Locators.Smaller_date_xp,Smaller_date_format)
time.sleep(random.randint(2,6))
web_actions.click_operation_Xpath(Locators.Bigger_date_xp)
web_actions.clear_operation_A_xpath(Locators.Bigger_date_2_xp)
Bigger_date=  datetime.date.today()+ relativedelta(years=10)
Bigger_date_format = Bigger_date.strftime("%d.%m.%Y")
web_actions.write_operation_xpath(Locators.Bigger_date_2_xp,Bigger_date_format)
web_actions.click_operation_Xpath(Locators.Anwenden_Falligkeit_full_XPath)
config_webdriver_manager.driver.implicitly_wait(10)
#Gelisteter Zeitraum Bond Listed date from to date 
web_actions.click_operation_css(Locators.Zusatzliche_Filter_css)
web_actions.relax()
web_actions.write_operation_b_xpath(Locators.Zusatzliche_Filter_Textfield_Full_XPath,"Gelist")
time.sleep(random.randint(2,6))
web_actions.click_operation_id_A(Locators.Gelisteter_Zeitraum_preselector_Button_id)
web_actions.click_operation_Xpath(Locators.Zusatzliche_Filter_Anwenden_Full_XPath)
web_actions.click_operation_Xpath(Locators.Gelisteter_Zeitraum_Button_xp)
time.sleep(random.randint(2,6))
web_actions.clear_operation_A_xpath(Locators.Gelisteter_Zeitraum_smaller_date_xp)
Smaller_date_listed=  datetime.date.today()- relativedelta(years=3)
Smaller_date_format_listed = Smaller_date_listed.strftime("%d.%m.%Y")
web_actions.write_operation_xpath(Locators.Gelisteter_Zeitraum_smaller_date_xp,Smaller_date_format_listed)
web_actions.click_operation_Xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp)
time.sleep(random.randint(2,6))
web_actions.clear_operation_A_xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp)
Bigger_date_listed= datetime.date.today().strftime("%d.%m.%Y")
web_actions.write_operation_xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp, Bigger_date_listed)
time.sleep(random.randint(2,6))
web_actions.click_operation_Xpath(Locators.Anwenden_Gelisteter_Zeitraum_full_XPath)
#Select Unternehmensanleihe = Bond
web_actions.click_operation_id_A(Locators.Preselector_Anleihen_Typ_ID)
web_actions.click_operation_id_A(Locators.Select_Unternehmensanleihe_Checkbox_ID)
web_actions.click_operation_Xpath(Locators.Anwenden_Unternehmensanleihe_Select_full_XPath)
# Choose Yield between maximum and minimum
web_actions.click_operation_Xpath(Locators.Rendite_button_xp)
web_actions.clear_operation_A_xpath(Locators.Rendite_lower_margin_xp)
web_actions.write_operation_xpath(Locators.Rendite_lower_margin_xp, "3%")
web_actions.clear_operation_A_xpath(Locators.Rendite_upper_margin_xp)
web_actions.write_operation_xpath(Locators.Rendite_upper_margin_xp, "16%")
time.sleep(random.randint(2,6))
web_actions.click_operation_Xpath(Locators.Rendite_anwenden_xp)
#Sort by yield
time.sleep(random.randint(2,6))
config_webdriver_manager.driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
#show_more_hits
web_actions.click_operation_Xpath(Locators.Weitere_Treffer_anzeigen_xp)
_j=1
_n=5
while _j<_n:
    config_webdriver_manager.driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
    time.sleep(random.randint(2,6))
    web_actions.click_operation_Xpath(Locators.Weitere_Treffer_anzeigen_xp)
    _j+=1
    time.sleep(random.randint(2,6))

# Write Results to Excel file
html_source =config_webdriver_manager.driver.page_source
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

#out= df.to_json(orient='records')[1:-1].replace('},{', '} {')
#with open('c:\\temp\\tutorial\\df_to_json.txt', 'w') as f:
#    f.write(out)


wb = load_workbook('c:\\temp\\tutorial\\sample.xlsx')

sheet = wb.worksheets[0]
# this statement inserts a column before column 2

for x in range(0, 10, 1):
    sheet.insert_cols(0)

web_actions.df_to_excel(df, wb.active)
wb.save("c:\\temp\\tutorial\\sample.xlsx")


config_webdriver_manager.driver.quit()

