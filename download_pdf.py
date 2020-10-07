#! python3
# Download all pdf files of the monnth from SAT

import sys, datetime, time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Credentials
rfc = 'HEGG7005106N4'
password = 'GUST1506'
year = 2018
month = 10
day = 0
currentYear = datetime.datetime.now().year


# Open browser
browser = webdriver.Chrome()

def loadLoggin (): 
    """ Load loggin page from SAT"""
    # Load page
    browser.get ('https://portalcfdi.facturaelectronica.sat.gob.mx/')
    print ('Loading loggin page...')
    print ('Please confirm the captcha...')

    # Move and write credentials
    inputRfc = browser.find_element_by_css_selector("#rfc")
    inputRfc.send_keys(rfc)
    inputPassword = browser.find_element_by_css_selector("#password")
    inputPassword.send_keys(password)

# Whait to autenticate
while True: 
    loadLoggin()
    autenticate = input ('¿Captcha correct? (y/n)')
    if autenticate.lower()[0:1] == 'y':
        break        
    else: 
        # Try again the chaptcha
        print ('Realoading page...')
        loadLoggin()

# To to seccion
inputRecibidas = browser.find_element_by_css_selector('[title="Facturas Recibidas"]')
inputRecibidas.click()

# Filter cfdis
## Filter by date
inputFechas = browser.find_element_by_css_selector('#ctl00_MainContent_RdoFechas')
inputFechas.click()

## Select year
dropdownYear = WebDriverWait (browser, 20).until( # whait to load elements
        EC.element_to_be_clickable ((By.CSS_SELECTOR, '#DdlAnio'))
    )
for year in range (0, currentYear-year): 
    dropdownYear.send_keys(Keys.UP)

## Select month
dropdownMonth = browser.find_element_by_css_selector('#ctl00_MainContent_CldFecha_DdlMes')
for month in range (0, month -1): 
    dropdownMonth.send_keys(Keys.DOWN)

## Select day
dropdownDay = browser.find_element_by_css_selector('#ctl00_MainContent_CldFecha_DdlDia')
for day in range (0, day): 
    dropdownDay.send_keys(Keys.DOWN)
dropdownDay.send_keys(Keys.ENTER)

# Download files
xmlDownloadBtns = WebDriverWait (browser, 20).until( # whait to load elements
        EC.element_to_be_clickable ((By.CSS_SELECTOR, '#ctl00_MainContent_tblResult table tbody tr td'))
    )
xmlDownloadBtns = browser.find_elements_by_css_selector('#ctl00_MainContent_tblResult table tbody tr td:nth-child(4)')
for xmlDownloadBtn in xmlDownloadBtns: 
    time.sleep(1)
    xmlDownloadBtn.click()
