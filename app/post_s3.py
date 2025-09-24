import os
import boto3
from botocore.exceptions import ClientError, NoCredentialsError
from config import Config

class Post_S3(Config):
    def __init__(self):
        super().__init__()
        self.s3 = boto3.client(
            "s3",
            aws_access_key_id=self.aws_access_key_id,
            aws_secret_access_key=self.aws_secret_access_key,
            region_name=self.region_name
        )

    def upload_file(self, local_folder: str, area: str):
        """
        local_folder 例:
        /home/ubuntu/koatsu_30minutes_store/download_folder/中国/202508/20250801
        area = "関東"

        → S3: フォルダ保存用/関東/202508/20250801/20250801.xlsx
        """
        if not os.path.exists(local_folder):
            print(f"❌ Folder not found: {local_folder}")
            return

        # local_folder の末尾が day_folder（例: 20250801）
        day_folder = os.path.basename(local_folder)
        # その1つ上が month_folder（例: 202508）
        month_folder = os.path.basename(os.path.dirname(local_folder))

        file_name = f"{day_folder}.xlsx"
        local_file = os.path.join(local_folder, file_name)

        if os.path.exists(local_file):
            # --- S3キーを組み立て ---
            s3_key = f"フォルダ保存用/{area}/{month_folder}/{day_folder}/{file_name}"

            try:
                self.s3.upload_file(local_file, self.s3_bucket_name, s3_key)
                print(f"✅ Uploaded {local_file} → s3://{self.s3_bucket_name}/{s3_key}")
            except FileNotFoundError:
                print(f"❌ Local file not found: {local_file}")
            except NoCredentialsError:
                print("❌ AWS credentials not found (IAMロールを確認してください)")
            except ClientError as e:
                print(f"❌ Failed to upload: {e}")
        else:
            print(f"⚠ {file_name} not found in {local_folder}")
