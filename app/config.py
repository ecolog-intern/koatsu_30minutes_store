# config.py
from datetime import datetime, timedelta
import os
import base64
import boto3

class Config:
    def __init__(self):
        '''
        フォルダパスの一覧
        '''
        #作成したファイルの掃き出し先
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.download_folder = os.path.join(self.base_dir, "download_folder")

        '''
        電力会社url
        '''
        self.tyugoku_domain = os.getenv("CHUGOKU_DOMAIN")
        self.tyugoku_url = os.getenv("CHUGOKU_URL")
        self.tohoku_domain = os.getenv("TOHOKU_DOMAIN")
        self.tohoku_url = os.getenv("TOHOKU_URL")
        self.kansai_domain = os.getenv("KANSAI_DOMAIN")
        self.kansai_url = os.getenv("KANSAI_URL")
        self.kanto_domain = os.getenv("KANTO_DOMAIN")
        self.kanto_url = os.getenv("KANTO_URL")
        self.tyubu_domain = os.getenv("TYUBU_DOMAIN")
        self.tyubu_url = os.getenv("TYUBU_URL")
        self.hokuriku_domain = os.getenv("HOKURIKU_DOMAIN")
        self.hokuriku_url = os.getenv("HOKURIKU_URL")
        self.kyusyu_domain = os.getenv("KYUSYU_DOMAIN")
        self.kyusyu_url = os.getenv("KYUSYU_URL")
        self.hokkaido_domain = os.getenv("HOKKAIDO_DOMAIN")
        self.hokkaido_url = os.getenv("HOKKAIDO_URL")
        self.shikoku_domain = os.getenv("SHIKOKU_DOMAIN")
        self.shikoku_url = os.getenv("SHIKOKU_URL")

        '''
        AWS情報
        '''
        self.aws_access_key_id = os.getenv("AWS_ACCESS_KEY_ID")
        self.aws_secret_access_key = os.getenv("AWS_SECRET_ACCESS_KEY")
        self.aws_region = os.getenv("region")
        self.bucket_name = os.getenv("bucket_name")
        self.project = os.getenv("project")
        self.s3_client = boto3.client(
            "s3",
            aws_access_key_id= self.aws_access_key_id,
            aws_secret_access_key=self.aws_secret_access_key,
            region_name=self.aws_region,
        )

        '''
        スクレイピング証明書類
        '''
        self.client_cert_base64 = os.getenv("CLIENT_CERT_BASE64")
        self.client_key_base64 = os.getenv("CLIENT_KEY_BASE64")
        self.cert_content = base64.b64decode(self.client_cert_base64)
        self.key_content = base64.b64decode(self.client_key_base64)

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