'''
中国地方のスクレイピングをする
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

class TyugokuScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.tyugoku_url = self.config.tyugoku_url
        self.yesterday_str = self.config.yesterday_str
        
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
        
        # 証明書確認のため、指定位置にマウスを移動してクリックする関数を定義
        pyautogui.moveTo(585, 378, duration=0.5)
        def click_cert_ok():
            sleep(2)
            pyautogui.moveTo(585, 378, duration=0.5)
            pyautogui.click()
        
        click_thread = threading.Thread(target=click_cert_ok)
        click_thread.start()
        
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
