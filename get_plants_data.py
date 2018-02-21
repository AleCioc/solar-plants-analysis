#from lxml import html
#import requests
#import shutil
#
#url = "https://smit.comaitorino.it/2.0/pdi/ftv/forecastfv.php#"
#r = requests.get(url, stream=True)
#if r.status_code == 200:
#    with open("./test.csv", 'wb') as f:
#        r.raw.decode_content = True
#        shutil.copyfileobj(r.raw, f)  
#
#
#page = requests.get('https://smit.comaitorino.it/?asja')
#tree = html.fromstring(page.content)
#username = tree.xpath('//input[@id = "l"]')
#password = tree.xpath('//input[@id = "txtpw"]')
#button = tree.xpath('//input[@id = "e"]')
#button2 = tree.xpath('//a[@class = "button"]')

import time

import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
matplotlib.style.use('ggplot')


import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile

profile = FirefoxProfile()
profile.set_preference("browser.download.manager.showWhenStarting", False);
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream")
profile.set_preference("browser.helperApps.alwaysAsk.force", False)
profile.set_preference("browser.download.manager.showAlertOnComplete", False)
profile.set_preference("browser.download.manager.closeWhenDone", False)
profile.set_preference("browser.download.folderList", 2); 
                      
driver = webdriver.Firefox(firefox_profile=profile)
driver.get("https://smit.comaitorino.it/?asja")
    
def find_by_xpath(locator):
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, locator)))
    return element

find_by_xpath('//input[@id = "l"]').send_keys("")          
find_by_xpath('//input[@id = "txtpw"]').send_keys("")          
find_by_xpath('//input[@id = "e"]').click()
find_by_xpath('//a[@onclick = "openForecastFV()"]').click()
#driver.find_element_by_css_selector("input[onclick*='openForecastFV']").click()

#xpath = '//a[@onclick = "data_download("download_form","forecast","2017-3-01 00:00:00","2017-4-01 00:00:00")"]'
#find_by_xpath(xpath).click()

plants_list = [
                'Ramacca',
                'Capitano',
                'San Giorgio',
                'Capolongo',
                'Contessa',
                'Pianezza'
              ]

import os

for plant in plants_list:

    if not os.path.exists(plant):
        os.makedirs(plant)
    
    driver.get("https://smit.comaitorino.it/2.0/pdi/ftv/forecastfv.php")    
    #find_by_xpath('//a[@class = "button noprint"]').click()
    el = driver.find_element_by_id('plant_select')
    for option in el.find_elements_by_tag_name('option'):
        if option.text == "Asja - " + plant:
            option.click()
            break
    time.sleep(5)
    find_by_xpath('//a[@class = "button noprint"]').click()

    for year in range(2013, 2018):
        months = [
                    'Gen ' + str(year),
                    'Feb ' + str(year),
                    'Mar ' + str(year),
                    'Apr ' + str(year),
                    'Mag ' + str(year),
                    'Giu ' + str(year),
                    'Lug ' + str(year),
                    'Ago ' + str(year),
                    'Set ' + str(year),
                    'Ott ' + str(year),
                    'Nov ' + str(year),
                    'Dic ' + str(year)
                ]
    
        for month in months:
#            driver.get("https://smit.comaitorino.it/2.0/pdi/ftv/forecastfv.php")    
#            find_by_xpath('//a[@class = "button noprint"]').click()
            el = driver.find_element_by_id('my')
            for option in el.find_elements_by_tag_name('option'):
                if option.text == month:
                    option.click()
                    break
            time.sleep(5)
            find_by_xpath('//a[@class = "button noprint"]').click()

    for filename in os.listdir():
        if filename.startswith("asja"):
            os.rename("./" + filename, "./" + plant + "/" + filename)
