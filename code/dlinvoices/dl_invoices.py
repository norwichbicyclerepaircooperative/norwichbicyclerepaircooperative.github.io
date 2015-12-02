"""
Stupid selenium script to download invoices from Walkers Cycles' website,
which is annoyingly difficult.

Don't know if all the sleeps / waits are required or of appropriate length,
some of them are definitely required.

Next time this is needed, they'll have changed the website so this doesn't
work anymore.
"""

import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os

username = "xxxxx"
password = "xxxxx"

# Download excel files automatically to working dir
fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList",2)
fp.set_preference("browser.download.manager.showWhenStarting",False)
fp.set_preference("browser.download.dir", os.getcwd())
fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream, application/ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

driver = webdriver.Firefox(firefox_profile=fp)
# Log in:
driver.get("http://www.walkerscycles.co.uk")
elem = driver.find_element_by_name("ctl00$ContentPlaceHolder1$txtAccountNo")
elem.send_keys(username)
elem = driver.find_element_by_name("ctl00$ContentPlaceHolder1$txtPassword")
elem.send_keys(password)
elem = driver.find_element_by_name("ctl00$ContentPlaceHolder1$btnLogin")
elem.send_keys(Keys.RETURN)
driver.implicitly_wait(3)
driver.get("http://www.walkerscycles.co.uk/orderhistory")
inv_urls = [e.get_attribute("href") for e in driver.find_elements_by_link_text("Invoiced")]

for inv_url in inv_urls:
    driver.get(inv_url)
    time.sleep(3)
    e = driver.find_element_by_id("ctl00_ContentPlaceHolder1_uiReport_ctl05_ctl04_ctl00_ButtonLink")
    e.click()
    time.sleep(3)
    excel_e = driver.find_element_by_link_text("Excel")
    excel_e.click()
    time.sleep(3)

driver.close()
