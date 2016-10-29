"""
Stupid selenium script to download invoices from Walkers Cycles' website,
which is annoyingly fiddly.

Don't know if all the sleeps / waits are required or of appropriate length,
some of them are definitely required.

Next time this is needed, they'll have changed the website so this doesn't
work anymore.
"""

import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions

import time
import os

from walkers_credentials import USERNAME, PASSWORD


def get_loggedin_driver(download_dir=None):
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting",False)
    if download_dir:
        fp.set_preference("browser.download.dir", download_dir)
        fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream, application/ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    driver = webdriver.Firefox(firefox_profile=fp)
    driver.get("http://www.walkerscycles.co.uk")
    elem = driver.find_element_by_name("ctl00$ContentPlaceHolder1$txtAccountNo")
    elem.send_keys(USERNAME)
    elem = driver.find_element_by_name("ctl00$ContentPlaceHolder1$txtPassword")
    elem.send_keys(PASSWORD)
    elem = driver.find_element_by_name("ctl00$ContentPlaceHolder1$btnLogin")
    elem.click()
    WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located((By.CLASS_NAME, "welcomeinner")));
    return driver


# Get list of invoices
driver = get_loggedin_driver()
driver.get("http://www.walkerscycles.co.uk/orderhistory")

invoice_info = [
    (
        el.find_element_by_xpath('../../td[1]').get_attribute('innerHTML'),
        el.get_attribute("href")
    )
    for el in driver.find_elements_by_link_text("Invoiced")
]
driver.quit()


# Download the invoices
for order_no, href in invoice_info:
    assert order_no.isalnum()
    save_dir = os.path.join(os.getcwd(), "invoices", order_no)
    if os.path.isdir(save_dir):
        if len(os.listdir(save_dir)):
            continue
    else:
        os.mkdir(save_dir)
    driver = get_loggedin_driver(save_dir)
    driver.get(href)
    WebDriverWait(driver, 30).until(expected_conditions.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_uiReport_ctl05_ctl04_ctl00_ButtonImgDown")));
    time.sleep(5)
    e = driver.find_element_by_id("ctl00_ContentPlaceHolder1_uiReport_ctl05_ctl04_ctl00_ButtonImgDown")
    e.click()
    WebDriverWait(driver, 30).until(expected_conditions.element_to_be_clickable((By.LINK_TEXT, "Excel")));
    excel_e = driver.find_element_by_link_text("Excel")
    time.sleep(1)
    excel_e.click()
    time.sleep(6)
    driver.quit()


