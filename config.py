from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
chromedriver_location = "C:\\Users\\lenzre\\Documents\\programs\\chromedriver.exe"
service = Service(executable_path = chromedriver_location)
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
driver = webdriver.Chrome(service = service, options=chrome_options)

