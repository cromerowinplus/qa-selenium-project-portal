from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def inicializar_driver():
    options = Options()
    prefs = {"profile.default_content_setting_values.geolocation": 1}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--incognito")  
    service = Service(ChromeDriverManager().install())  # << Automático
    return webdriver.Chrome(service=service, options=options)