'''
関東地方のスクレイピングをする
folder_pathにスクレイピングでダウンロードしたファイルを置く
'''
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from config import Config
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

class TokyoScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.tokyo_url = self.config.tokyo_url
        
    def scraping(self):
        #実際にスクレイピングを行う
        options = Options()

        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")  
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        prefs = {
            "download.default_directory": self.folder_path,  
            "download.prompt_for_download": False,
            "directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )

        # 東京電力のサイトにアクセス
        driver.get(self.tokyo_url)
        sleep(10)

        elem = driver.find_element(By.XPATH, "//input[@id='johokokai']")
        elem.click()
        sleep(10)
        
        elem = driver.find_element(By.XPATH, "//a[text()='同時同量公開一覧']")
        elem.click()
        sleep(10)
        
        select_element = driver.find_element("id", "DTO-LVK4RS001_INFO_KUBUN_CD")
        select = Select(select_element)
        select.select_by_value("0120")
        sleep(2)
        
        select_element = driver.find_element('id', "DTO-LVK4RS001_VOLT_SHUBT_CD")
        select = Select(select_element)
        select.select_by_value('0')
        sleep(2)
        
        elem = driver.find_element(By.NAME, 'j_idt24')
        elem.click()
        sleep(3)
        
        elem = driver.find_element(By.NAME, "DTO-LVK4RS001G02_SELECT_INDEX_LIST[0]")
        elem.click()
        sleep(3)
        
        elem = driver.find_element(By.NAME, "j_idt61")
        elem.click()
        sleep(10)
        
        driver.quit()