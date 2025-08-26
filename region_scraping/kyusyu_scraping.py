'''
九州地方のスクレイピングをする
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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class KyusyuScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.kyusyu_url = self.config.kyusyu_url

        
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

        # 九州電力のサイトにアクセス
        driver.get(self.kyusyu_url)
        sleep(10)

        elem = driver.find_element(By.ID, "balancingMenuButton")
        elem.click()
        sleep(3)
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//iframe[contains(@src, "InitBalancingRef.do")]'))
        )
        driver.switch_to.frame(iframe)

        elem = driver.find_elements(By.TAG_NAME, 'select')
        elem = elem[1]

        select = Select(elem)
        select.select_by_value("0120")

        elem = driver.find_element(By.ID, 'showBtn')
        elem.click()
        sleep(5)

        elem = driver.find_element(By.ID, 'grid')
        elem = elem.find_elements(By.XPATH, './*/*/*')
        elem = elem[1]
        elem1 = elem.find_elements(By.TAG_NAME, 'tr')
        elem1 = elem1[1]
        elem2 = elem1.find_element(By.TAG_NAME, 'a')
        elem2.click()
        sleep(10)
        
        driver.quit()