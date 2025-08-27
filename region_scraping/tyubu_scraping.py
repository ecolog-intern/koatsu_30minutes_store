'''
中部地方のスクレイピングをする
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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

class TyubuScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.tyubu_url = self.config.tyubu_url
        
        
    def scraping(self):
        #実際にスクレイピングを行う
        options = Options()

        #options.add_argument("--headless") #EC2にデプロイするときに外す
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
        
        # 証明書の許可処理: 指定した位置にマウスを移動し、クリックを実施
        pyautogui.moveTo(585, 378, duration=0.5)
        def click_cert_ok():
            sleep(2)
            pyautogui.moveTo(585, 378, duration=0.5)
            pyautogui.click()
        
        click_thread = threading.Thread(target=click_cert_ok)
        click_thread.start()
        
        # 中部電力のサイトへアクセス
        driver.get(self.tyubu_url)
        sleep(10)
        

        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "menu")))
        link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "日毎３０分電力量"))
        )
        link.click()
        sleep(3)
        driver.switch_to.default_content()


        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "main")))
        elem = driver.find_elements(By.TAG_NAME, 'a')
        elem = elem[0]
        elem.click()
        sleep(10)
        
        driver.quit()
