'''
関東地方のスクレイピングをする
folder_pathにスクレイピングでダウンロードしたファイルを置く
'''
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
from time import sleep
import threading
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from config import Config
from selenium.webdriver.support.ui import Select

class TokyoScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.tokyo_url = self.config.tokyo_url

        
    def scraping(self):
        #実際にスクレイピングを行う
        options = webdriver.ChromeOptions()

        # ★ プロキシ設定
        options.add_argument('--proxy-server=http://proxy:8080')
        options.add_argument('--disable-gpu')
        options.add_argument('--ignore-certificate-errors')

        # ★ ダウンロード先指定（prefsの設定）
        prefs = {
            "download.default_directory": self.folder_path,
            "download.prompt_for_download": False,  # ダウンロード時の確認ダイアログを無効化
            "directory_upgrade": True
        }
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(options=options)

        # 証明書許可のために、一定時間後に指定位置をクリックする処理を別スレッドで実行
        def click_cert_ok():
            sleep(2)
            pyautogui.moveTo(585, 378, duration=0.5)
            pyautogui.click()
        threading.Thread(target=click_cert_ok).start()

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