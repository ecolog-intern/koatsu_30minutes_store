'''
中国地方のスクレイピングをする
folder_pathにスクレイピングでダウンロードしたファイルを置く
'''
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import sys
import os
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from config import Config
import requests

class TyugokuScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.tyugoku_url = self.config.tyugoku_url
        self.yesterday_str = self.config.yesterday_str
        
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
        
        # 中国地方のサイトにアクセス
        driver.get(self.tyugoku_url)
        sleep(10)
        
        # メニュー内のリンクを取得し、2番目のリンクをクリック（インデックス1）
        pull_menu = driver.find_element(By.ID, "pullMenuDiv")
        menu_links = pull_menu.find_elements(By.TAG_NAME, 'a')
        menu_links[1].click()
        sleep(3)
        
        # 昨日の日付（YYYYMMDD形式）を対象とする
        target_date = self.yesterday_str
        
        # コンテンツテーブル内の各行から、対象ファイルのリンクを検索
        content_table = driver.find_element(By.ID, 'contentBody1DetailsElement')
        table_rows = content_table.find_elements(By.TAG_NAME, 'tr')
        for row in table_rows:
            # '特高・高圧日毎３０分電力量' と対象の日付が含まれる行を検出
            if '特高・高圧日毎３０分電力量' in row.text and target_date in row.text:
                file_link = row.find_element(By.TAG_NAME, 'a')
                driver.execute_script("arguments[0].scrollIntoView(true);", file_link)
                driver.execute_script("arguments[0].click();", file_link)  # ZIPファイルのダウンロードをトリガー
                break
        
        sleep(10)
        driver.close()
