# main.py
import os
from config import Config
from folder_file_making import FolderFileMaking
from scraping import Scraping
from zip_thawing import ZipThawing
from post_s3 import Post_S3

def main():
    #カレントディクショナリに変更
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    #configから変数を抽出
    config = Config()
    base_folder = config.base_folder # 'ファイル保存用までのpath'
    
    all_area = ['中国','関西','東北','北陸','九州','関東','中部']

    # 各フィルタリングエントリを処理
    for area in all_area:
        for attempt in range(3): # エラー対応で数回回すことも考える
            try:
            # 対象のフォルダおよびExcelファイルを作成
                folder_making = FolderFileMaking(base_folder, area)
                yesterday_path = folder_making.folder_making()
                scraping = Scraping(area, yesterday_path)
                scraping.scraping()
                zipthwing = ZipThawing(yesterday_path)
                zipthwing.zip_thawing()

                # yesterday_path 内に作成されたエクセルファイルを　config.s3_object_path へあげる
                uploader = Post_S3()
                uploader.upload_file(yesterday_path, area)
                break
            
            except Exception as e:
                if attempt == 2:
                    log_path = config.error_log
                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write("エラー発生: " + str(e) + "\n")
            
if __name__ == '__main__':
    main()
    