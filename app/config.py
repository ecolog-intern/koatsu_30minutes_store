from datetime import datetime, timedelta
from make_pem import pem_files
import os

class Config:
    def __init__(self):
        '''
        フォルダパスの一覧
        '''
        #顧客一覧
        self.consumer_filtering_file = 'consumer_region_filtering.xlsx'
        
        #作成したファイルの掃き出し先
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.base_folder = os.path.join(self.base_dir, "download_folder")

        #エラーログ
        self.error_log = "error_log.txt"
        
        '''
        電力会社url
        '''
        self.tyugoku_url = 'https://takusouhp.energia.co.jp/COMM/xhtml/COMMLOP.xhtml'
        self.tohoku_url = 'https://takuso2-web.takuso.tohoku-epco.co.jp/G83_PPS/'
        self.kansai_url = 'https://www4.kepco.co.jp/'
        #関西エラー対応分
        self.kansai_error_url = 'https://www4.kepco.co.jp/takusouinfo/H24DF700A04.do'
        self.tokyo_url = 'https://pu00.www6.tepco.co.jp/org_web/LVA2RG/pgsslogin.faces'
        self.tyubu_url = 'https://epcdss-www.chuden.co.jp/46264/'
        self.hokuriku_url = 'https://wsweb4.rikuden.co.jp/tfx/tfxo110/tfxo110s010'
        self.kyusyu_url = 'https://nsc-www.network.kyuden.co.jp/BP_WEB_SERVER/'

        '''
        AWSのS3
        '''
        self.aws_access_key_id = os.getenv("AWS_ACCESS_KEY_ID")
        self.aws_secret_access_key = os.getenv("AWS_SECRET_ACCESS_KEY")
        self.region_name = os.getenv("AWS_REGION")
        self.s3_bucket_name = os.getenv("AWS_STORAGE_BUCKET_NAME")

        '''
        スクレイピングの証明書
        '''
        # スクレイピングのためのpemファイル作成
        self.cert_path, self.key_path, self.ca_path = pem_files()

        '''
        日付系
        '''
        today = datetime.today()
        yesterday = today - timedelta(days=1)
        self.yesterday_str = yesterday.strftime("%Y%m%d")
        self.yesterday_month_str = yesterday.strftime("%Y%m")
        self.month_day = yesterday.strftime('%m%d')  
        self.year2digit = yesterday.strftime('%y') 
        self.yesterday_tsuki_nichi = f"{yesterday.month}月{yesterday.day}日"
    