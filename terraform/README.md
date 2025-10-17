# Koatsu Scraper - Terraform デプロイメント

このTerraformコードは、koatsu-scraperをAWS ECS Fargateで実行するために必要なすべてのリソースを自動的にセットアップします。

## 含まれるリソース

- **ECS Cluster**: Fargateタスクを実行するクラスター
- **ECS Task Definition**: コンテナの設定とリソース割り当て
- **IAM Roles**: タスク実行ロールとタスクロール
- **AWS Secrets Manager**: AWS認証情報、証明書、スクレイピングURL等の環境変数の安全な保管
- **CloudWatch Logs**: タスクログの記録
- **Security Group**: ECSタスク用のネットワークセキュリティ
- **EventBridge Rule**: 定期実行スケジュール（デフォルト: 毎日午前3時JST）

## 前提条件

1. **Terraform**がインストールされていること（v1.0以降）
2. **AWS CLI**が設定されていること
3. **ECRにDockerイメージ**がプッシュされていること
4. 適切な**AWSクレデンシャル**が設定されていること

## セットアップ手順

### 1. terraform.tfvarsファイルを作成

```bash
cd terraform
cp terraform.tfvars.example terraform.tfvars
```

### 2. terraform.tfvarsを編集

必要な値を入力してください:

```hcl
# ECRイメージURL（重要！）
ecr_image_url = "123456789012.dkr.ecr.ap-northeast-1.amazonaws.com/koatsu-scraper:latest"

# AWS認証情報（.envファイルから値をコピー）
aws_access_key_id     = "YOUR_AWS_ACCESS_KEY_ID"
aws_secret_access_key = "YOUR_AWS_SECRET_ACCESS_KEY"

# S3設定
bucket_name = "koatsu-30minutes-test"
project     = "koatsu-30minutes"

# クライアント証明書（Base64エンコード済み、.envから値をコピー）
client_cert_base64 = "YOUR_BASE64_ENCODED_CERTIFICATE"
client_key_base64  = "YOUR_BASE64_ENCODED_KEY"

# スクレイピングURL設定
# デフォルト値が設定されているため、変更が必要な場合のみ指定

# スケジュール設定（オプション）
schedule_expression = "cron(0 18 * * ? *)"  # 毎日午前3時JST
schedule_enabled    = true

# リソース設定（オプション）
task_cpu    = "1024"  # vCPU: 0.25, 0.5, 1, 2, 4
task_memory = "2048"  # MB
```

### 3. Terraformを初期化

```bash
terraform init
```

### 4. プランを確認

```bash
terraform plan
```

作成されるリソースを確認してください。

### 5. デプロイ実行

```bash
terraform apply
```

`yes`と入力して実行を確認します。

## デプロイ後の操作

### タスクを手動で実行

Terraform applyの出力に表示されるコマンドを使用:

```bash
# 出力されたコマンド例
aws ecs run-task \
  --cluster koatsu-scraper-cluster \
  --task-definition koatsu-scraper-task \
  --launch-type FARGATE \
  --network-configuration "awsvpcConfiguration={subnets=[subnet-xxx],securityGroups=[sg-xxx],assignPublicIp=ENABLED}" \
  --region ap-northeast-1
```

または、Terraform outputから取得:

```bash
terraform output run_task_command
```

### ログを確認

```bash
# 出力されたコマンドを使用
terraform output logs_command

# または直接
aws logs tail /ecs/koatsu-scraper --follow --region ap-northeast-1
```

### AWS Consoleで確認

1. **ECS**: https://console.aws.amazon.com/ecs/
   - Cluster: `koatsu-scraper-cluster`
   - Task Definition: `koatsu-scraper-task`

2. **CloudWatch Logs**: https://console.aws.amazon.com/cloudwatch/
   - Log Group: `/ecs/koatsu-scraper`

3. **EventBridge**: https://console.aws.amazon.com/events/
   - Rule: `koatsu-scraper-schedule`

4. **Secrets Manager**: https://console.aws.amazon.com/secretsmanager/
   - Secret: `koatsu-scraper-env`

## スケジュール設定

デフォルトでは毎日午前3時（JST）に実行されます。変更する場合:

```hcl
# terraform.tfvars
schedule_expression = "cron(0 18 * * ? *)"  # 3 AM JST = 18:00 UTC
```

EventBridgeのcron式の例:
- `cron(0 18 * * ? *)` - 毎日午前3時JST（18:00 UTC）
- `cron(0 0 * * ? *)` - 毎日午前9時JST（0:00 UTC）
- `cron(0 12 * * ? *)` - 毎日午後9時JST（12:00 UTC）
- `rate(1 hour)` - 1時間ごと
- `rate(30 minutes)` - 30分ごと

スケジュールを無効化する場合:

```hcl
schedule_enabled = false
```

## リソースの更新

### ECRイメージを更新した場合

```bash
# 1. 新しいイメージをECRにプッシュ
docker tag koatsu-scraper:latest 123456789012.dkr.ecr.ap-northeast-1.amazonaws.com/koatsu-scraper:latest
docker push 123456789012.dkr.ecr.ap-northeast-1.amazonaws.com/koatsu-scraper:latest

# 2. タスク定義を更新（イメージURLが同じでも強制的に新しいリビジョンを作成）
terraform apply -replace=aws_ecs_task_definition.koatsu_scraper
```

### データベース認証情報を更新

```bash
# terraform.tfvarsを編集
vim terraform.tfvars

# 適用
terraform apply
```

### リソースサイズを変更

```hcl
# terraform.tfvars
task_cpu    = "2048"  # 2 vCPU
task_memory = "4096"  # 4GB
```

```bash
terraform apply
```

## トラブルシューティング

### タスクが起動しない

1. CloudWatch Logsでエラーを確認:
   ```bash
   aws logs tail /ecs/koatsu-scraper --follow
   ```

2. ECS Consoleでタスクの停止理由を確認:
   - ECS > Clusters > koatsu-scraper-cluster > Tasks

3. セキュリティグループを確認:
   - アウトバウンドルールで443ポートが許可されているか

### メモリ不足エラー

```hcl
# terraform.tfvars
task_memory = "4096"  # または 8192
```

### 環境変数の確認

Secrets Managerの値を確認:
```bash
aws secretsmanager get-secret-value --secret-id koatsu-scraper-env --region ap-northeast-1
```

## クリーンアップ

すべてのリソースを削除する場合:

```bash
terraform destroy
```

`yes`と入力して削除を確認します。

## 料金の目安

- **ECS Fargate** (1 vCPU, 2GB): 約$0.05/時間
- **CloudWatch Logs**: $0.50/GB (取り込み) + $0.03/GB (保存)
- **Secrets Manager**: $0.40/月/シークレット + $0.05/10,000 API呼び出し
- **EventBridge**: 無料（100万イベント/月まで）

1日1回実行（約30分）の場合: **月額 約$1-2**

## セキュリティのベストプラクティス

1. **terraform.tfvarsをgitignoreに追加**:
   ```bash
   echo "terraform/terraform.tfvars" >> .gitignore
   ```

2. **State fileをリモートに保存** (本番環境推奨):
   ```hcl
   # backend.tf
   terraform {
     backend "s3" {
       bucket = "your-terraform-state-bucket"
       key    = "koatsu-scraper/terraform.tfstate"
       region = "ap-northeast-1"
     }
   }
   ```

3. **IAMロールの権限を最小化**: 必要最小限の権限のみを付与

## サポート

問題が発生した場合:
1. CloudWatch Logsでエラーメッセージを確認
2. `terraform plan`で変更内容を確認
3. AWS Consoleで各リソースの状態を確認
