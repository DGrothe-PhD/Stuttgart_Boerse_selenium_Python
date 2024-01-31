from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager


chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
  "download.default_directory": "/path/to/download/dir",
  "download.prompt_for_download": False,
})
#active_userprofile = os.environ['USERPROFILE']
#chromedriver_location = active_userprofile+"\\Documents\\programs\\chromedriver.exe"
chrome_options.add_argument("--disable-gpu")#OLDER OPTIONS
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])#OLDER OPTIONS
chrome_options.add_argument("--headless")
chrome_options.add_argument("--remote-debugging-port=9222")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("start-maximized")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_experimental_option(
    "prefs", {"profile.default_content_setting_values.notifications": 1})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
