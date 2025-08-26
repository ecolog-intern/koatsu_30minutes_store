#このパッケージの実行はこのファイルで行う
import os
import warnings
from config import Config
import pandas as pd
from folder_file_making import FolderFileMaking
from scraping import Scraping
from zip_thawing import ZipThawing
from excel_writing import ExcelWriting
from mail_sending import MailSending

def main():
    #スクレイピングの設定
    os.environ["http_proxy"] = "http://proxy:8080"
    os.environ["https_proxy"] = "http://proxy:8080"
    os.environ["no_proxy"] = "127.0.0.1,localhost"
    warnings.simplefilter("ignore")
    
    #カレントディクショナリに変更
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    #configから変数を抽出
    config = Config()
    #consumer_filtering = pd.read_excel(config.consumer_filtering_file, index_col=0)
    base_folder = config.base_folder
    
    all_area = ['中国','関西','東北','北陸','九州','関東','中部']

    # 各フィルタリングエントリを処理
    for area in all_area:
        for attempt in range(3):
            try:
            # 対象のフォルダおよびExcelファイルを作成（存在しない場合は新規作成）
                folder_making = FolderFileMaking(base_folder, area)
                yesterday_path = folder_making.folder_making()


                scraping = Scraping(area, yesterday_path)
                scraping.scraping()
                zipthwing = ZipThawing(yesterday_path)
                excel_file = zipthwing.zip_thawing()
                
                break
            
            except Exception as e:
                if attempt == 2:
                    
                    log_path = config.error_log
                    
                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write("エラー発生: " + str(e) + "\n")
        

    #完了したらメールを送信する
    mail_address = config.mailaddress
    mail_body = '''
本日の高圧30分値の実行は完了しました。
    '''
    for each_mail in mail_address:
        mail = MailSending(
            to=each_mail,
            subject="高圧30分値に関して",
            body=mail_body
        )
        mail.send_mail()

            
if __name__ == '__main__':
    main()
    