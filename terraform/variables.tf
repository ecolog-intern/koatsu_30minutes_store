variable "aws_region" {
  description = "AWS region"
  type        = string
  default     = "ap-northeast-1"
}

variable "ecr_image_url" {
  description = "ECR image URL (e.g., 123456789012.dkr.ecr.ap-northeast-1.amazonaws.com/koatsu-scraper:latest)"
  type        = string
}

variable "task_cpu" {
  description = "CPU units for the task (256, 512, 1024, 2048, 4096)"
  type        = string
  default     = "1024"
}

variable "task_memory" {
  description = "Memory for the task in MB (512, 1024, 2048, 3072, 4096, 5120, 6144, 7168, 8192)"
  type        = string
  default     = "2048"
}

# AWS関連
variable "aws_access_key_id" {
  description = "AWS Access Key ID"
  type        = string
  sensitive   = true
}

variable "aws_secret_access_key" {
  description = "AWS Secret Access Key"
  type        = string
  sensitive   = true
}

variable "bucket_name" {
  description = "S3 bucket name"
  type        = string
}

variable "project" {
  description = "Project name"
  type        = string
  default     = "koatsu-30minutes"
}

# 証明書（Base64エンコード済み）
variable "client_cert_base64" {
  description = "Client certificate (Base64 encoded)"
  type        = string
  sensitive   = true
}

variable "client_key_base64" {
  description = "Client key (Base64 encoded)"
  type        = string
  sensitive   = true
}

# スクレイピングURL - 中国
variable "chugoku_domain" {
  description = "Chugoku domain"
  type        = string
  default     = "https://takusouhp.energia.co.jp"
}

variable "chugoku_url" {
  description = "Chugoku URL"
  type        = string
  default     = "https://takusouhp.energia.co.jp/COMM/xhtml/COMMLOP.xhtml"
}

# スクレイピングURL - 東北
variable "tohoku_domain" {
  description = "Tohoku domain"
  type        = string
  default     = "https://takuso2-web.takuso.tohoku-epco.co.jp"
}

variable "tohoku_url" {
  description = "Tohoku URL"
  type        = string
  default     = "https://takuso2-web.takuso.tohoku-epco.co.jp/G83_PPS/"
}

# スクレイピングURL - 関東
variable "kanto_domain" {
  description = "Kanto domain"
  type        = string
  default     = "https://pu00.www6.tepco.co.jp"
}

variable "kanto_url" {
  description = "Kanto URL"
  type        = string
  default     = "https://pu00.www6.tepco.co.jp/org_web/LVA2RG/pgsslogin.faces"
}

# スクレイピングURL - 中部
variable "tyubu_domain" {
  description = "Tyubu domain"
  type        = string
  default     = "https://epcdss-www.chuden.co.jp"
}

variable "tyubu_url" {
  description = "Tyubu URL"
  type        = string
  default     = "https://epcdss-www.chuden.co.jp/46264/"
}

# スクレイピングURL - 北陸
variable "hokuriku_domain" {
  description = "Hokuriku domain"
  type        = string
  default     = "https://wsweb4.rikuden.co.jp"
}

variable "hokuriku_url" {
  description = "Hokuriku URL"
  type        = string
  default     = "https://wsweb4.rikuden.co.jp/tfx/tfxo110/tfxo110s010"
}

# スクレイピングURL - 九州
variable "kyusyu_domain" {
  description = "Kyusyu domain"
  type        = string
  default     = "https://nsc-www.network.kyuden.co.jp"
}

variable "kyusyu_url" {
  description = "Kyusyu URL"
  type        = string
  default     = "https://nsc-www.network.kyuden.co.jp/BP_WEB_SERVER/"
}

# スクレイピングURL - 関西
variable "kansai_domain" {
  description = "Kansai domain"
  type        = string
  default     = "https://www4.kepco.co.jp"
}

variable "kansai_url" {
  description = "Kansai URL"
  type        = string
  default     = "https://www4.kepco.co.jp/takusou/index.html"
}

# スクレイピングURL - 北海道
variable "hokkaido_domain" {
  description = "Hokkaido domain"
  type        = string
  default     = "https://nsc.hepco.co.jp"
}

variable "hokkaido_url" {
  description = "Hokkaido URL"
  type        = string
  default     = "https://nsc.hepco.co.jp/LNXWPWSS06OH"
}

# スクレイピングURL - 四国
variable "shikoku_domain" {
  description = "Shikoku domain"
  type        = string
  default     = "https://wsc3.yonden.co.jp"
}

variable "shikoku_url" {
  description = "Shikoku URL"
  type        = string
  default     = "https://wsc3.yonden.co.jp/PPS/"
}

variable "schedule_expression" {
  description = "EventBridge schedule expression (e.g., cron(0 18 * * ? *) for 3 AM JST)"
  type        = string
  default     = "cron(0 18 * * ? *)"
}

variable "schedule_enabled" {
  description = "Enable scheduled execution"
  type        = bool
  default     = true
}
