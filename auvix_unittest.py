# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re, os
import ConfigParser as configparser


def config_read( cfgFName ):
    cfg = configparser.ConfigParser() # inline_comment_prefixes=('#'))
    if  os.path.exists('confidential.cfg'):     
        cfg.read('confidential.cfg')
    if  os.path.exists(cfgFName):     
        cfg.read( cfgFName)
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    return cfg

    

class Auvix(unittest.TestCase):
    def setUp(self):
        ffprofile = webdriver.FirefoxProfile( u"C:\\Users\\Администратор\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\hip4h4sx.selenium")
#       ffprofile = webdriver.FirefoxProfile()
        ffprofile.set_preference("browser.download.dir", os.getcwd()+'\\tmp')
        ffprofile.set_preference("browser.download.folderList",2);
        ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk", 
            ",application/octet-stream" + 
            ",application/vnd.ms-excel" + 
            ",application/vnd.msexcel" + 
            ",application/x-excel" + 
            ",application/x-msexcel" + 
            ",application/zip" + 
            ",application/xls" + 
            ",application/vnd.ms-excel" +
            ",application/vnd.ms-excel.addin.macroenabled.12" +
            ",application/vnd.ms-excel.sheet.macroenabled.12" +
            ",application/vnd.ms-excel.template.macroenabled.12" +
            ",application/vnd.ms-excelsheet.binary.macroenabled.12" +
            ",application/vnd.ms-fontobject" +
            ",application/vnd.ms-htmlhelp" +
            ",application/vnd.ms-ims" +
            ",application/vnd.ms-lrm" +
            ",application/vnd.ms-officetheme" +
            ",application/vnd.ms-pki.seccat" +
            ",application/vnd.ms-pki.stl" +
            ",application/vnd.ms-word.document.macroenabled.12" +
            ",application/vnd.ms-word.template.macroenabed.12" +
            ",application/vnd.ms-works" +
            ",application/vnd.ms-wpl" +
            ",application/vnd.ms-xpsdocument" +
            ",application/vnd.openofficeorg.extension" +
            ",application/vnd.openxmformats-officedocument.wordprocessingml.document" +
            ",application/vnd.openxmlformats-officedocument.presentationml.presentation" +
            ",application/vnd.openxmlformats-officedocument.presentationml.slide" +
            ",application/vnd.openxmlformats-officedocument.presentationml.slideshw" +
            ",application/vnd.openxmlformats-officedocument.presentationml.template" +
            ",application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" +
            ",application/vnd.openxmlformats-officedocument.spreadsheetml.template" +
            ",application/vnd.openxmlformats-officedocument.wordprocessingml.template" +
            ",application/x-ms-application" +
            ",application/x-ms-wmd" +
            ",application/x-ms-wmz" +
            ",application/x-ms-xbap" +
            ",application/x-msaccess" +
            ",application/x-msbinder" +
            ",application/x-mscardfile" +
            ",application/x-msclip" +
            ",application/x-msdownload" +
            ",application/x-msmediaview" +
            ",application/x-msmetafile" +
            ",application/x-mspublisher" +
            ",application/x-msschedule" +
            ",application/x-msterminal" +
            ",application/x-mswrite" +
            ",application/xml" +
            ",application/xml-dtd" +
            ",application/xop+xml" +
            ",application/xslt+xml" +
            ",application/xspf+xml" +
            ",application/xv+xml" +
            ",application/excel")

        self.driver = webdriver.Firefox(ffprofile)
        self.driver.implicitly_wait(20)
        self.base_url = "http://b2b.auvix.ru"
        self.verificationErrors = []
        self.accept_next_alert = True

    
    def test_auvix(self):
        cfg = config_read('cfg_auvix.cfg')
        login = cfg.get('download', 'логин')
        pw    = cfg.get('download', 'пароль')
        pause = cfg.getint('download', 'time_sleep')

        driver = self.driver
        driver.get(self.base_url + "/")
        driver.find_element_by_name("USER_LOGIN").click()
        driver.find_element_by_name("USER_LOGIN").clear()
        driver.find_element_by_name("USER_LOGIN").send_keys(login)
        driver.find_element_by_name("USER_PASSWORD").click()
        driver.find_element_by_name("USER_PASSWORD").clear()
        driver.find_element_by_name("USER_PASSWORD").send_keys(pw)
        driver.find_element_by_name("Login").click()
        driver.find_element_by_link_text(u"Прайс-листы").click()
#        driver.find_element_by_link_text(u"Прайс-лист для дилеров").click()
#        driver.find_element_by_link_text(u"Прайс-лист для дилеров").click()
#        driver.find_element_by_link_text(u"Прайс-лист для дилеров").click()
#        driver.get(self.base_url + "/prices/Price_AUVIX_dealer.xls")
#        time.sleep(22)
#        driver.get(self.base_url + "/prices/Price_AUVIX_dealer_csv.csv")
#        time.sleep(33)
        driver.get(self.base_url + "/prices/Price_AUVIX_dealer_csv.zip")
        print pause
        time.sleep(pause)
        print pause

    
    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException, e: return False
        return True
    
    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException, e: return False
        return True
    
    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True
    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
