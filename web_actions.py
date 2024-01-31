import config
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import  TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def do_something(fname):
    print(config.driver)
    print(fname)


def do_something_else():
    print("Click")
    print(config.driver)
    
#USED 5
def click_operation_id_A(fname):
    try:
        element=config.driver.find_element(By.ID, fname)
        config.driver.execute_script("arguments[0].click();", element)
    except NoSuchElementException as exception:
        print ('Click on ' +  fname + '  ID A not successful')
    except TimeoutException:
        pass
#USED 1
def click_operation_B_Xpath(fname):
    WebDriverWait(config.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(config.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).click()
        config.driver.find_element("xpath", fname).sendKeys('\uE035')
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' B XPath  not successful')
    except TimeoutException:
        pass
#USED 3
def click_operation_css(fname):
    WebDriverWait(config.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        element= config.driver.find_element(By.CSS_SELECTOR, fname)
        config.driver.execute_script("arguments[0].click();", element)
        A=True
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' with css  not successful')
        A=False
#USED 4
def write_operation_b_xpath(fname, Input):
    WebDriverWait(config.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        element=WebDriverWait(config.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).send_keys(Input)
        #driver.find_element("xpath", fname).send_keys(Input)
    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with xpath celector')
        A=False
#USED 18
def click_operation_Xpath(fname):
    WebDriverWait(config.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(config.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname)))
        config.driver.execute_script("arguments[0].click();", element)
    except TimeoutException:
        pass
    except NoSuchElementException:
        print ('Click on ' + fname+ ' XPath not successful')
#USED 8
def clear_operation_A_xpath(fname):
    WebDriverWait(config.driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        print(" field name "+fname)
        element=WebDriverWait(config.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    except NoSuchElementException as exception:
        print ('Clear  ' +fname+ '   not successful')
    except TimeoutException:
        pass
#USED 6
def write_operation_xpath(fname, Input):
    WebDriverWait(config.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        config.driver.find_element("xpath", fname).send_keys(Input)
    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with xpath celector')
        A=False
    except TimeoutException:
        pass

# for writing the results to Excel
def df_to_excel(df, ws, header=True, index=True, startrow=0, startcol=0):
    """Write DataFrame df to openpyxl worksheet ws"""

    rows = dataframe_to_rows(df, header=header, index=index)

    for r_idx, row in enumerate(rows, startrow + 1):
        for c_idx, value in enumerate(row, startcol + 1):
             ws.cell(row=r_idx, column=c_idx).value = value


#RELAX
def relax():
    WebDriverWait(config.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
#ZOOM 40%
def zoom40():
    config.driver.execute_script("document.body.style.zoom='40%'")# works with Chrome
    config.driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
