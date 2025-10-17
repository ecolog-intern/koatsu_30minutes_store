terraform {
  required_version = ">= 1.0"
  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 5.0"
    }
  }
}

provider "aws" {
  region  = var.aws_region
  profile = "jpex"
}

# CloudWatch Logs
resource "aws_cloudwatch_log_group" "koatsu_scraper" {
  name              = "/ecs/koatsu-scraper"
  retention_in_days = 7
}

# IAM Role for ECS Task Execution
resource "aws_iam_role" "ecs_task_execution_role" {
  name = "koatsu-scraper-execution-role"

  assume_role_policy = jsonencode({
    Version = "2012-10-17"
    Statement = [
      {
        Action = "sts:AssumeRole"
        Effect = "Allow"
        Principal = {
          Service = "ecs-tasks.amazonaws.com"
        }
      }
    ]
  })
}

resource "aws_iam_role_policy_attachment" "ecs_task_execution_role_policy" {
  role       = aws_iam_role.ecs_task_execution_role.name
  policy_arn = "arn:aws:iam::aws:policy/service-role/AmazonECSTaskExecutionRolePolicy"
}

# Secrets Manager access policy
resource "aws_iam_role_policy" "secrets_manager_access" {
  name = "secrets-manager-access"
  role = aws_iam_role.ecs_task_execution_role.id

  policy = jsonencode({
    Version = "2012-10-17"
    Statement = [
      {
        Effect = "Allow"
        Action = [
          "secretsmanager:GetSecretValue"
        ]
        Resource = aws_secretsmanager_secret.koatsu_env.arn
      }
    ]
  })
}

# IAM Role for ECS Task
resource "aws_iam_role" "ecs_task_role" {
  name = "koatsu-scraper-task-role"

  assume_role_policy = jsonencode({
    Version = "2012-10-17"
    Statement = [
      {
        Action = "sts:AssumeRole"
        Effect = "Allow"
        Principal = {
          Service = "ecs-tasks.amazonaws.com"
        }
      }
    ]
  })
}

# Secrets Manager
resource "aws_secretsmanager_secret" "koatsu_env" {
  name                    = "koatsu-scraper-env"
  description             = "Environment variables for koatsu scraper"
  recovery_window_in_days = 0 # 即座に削除（開発用）
}

resource "aws_secretsmanager_secret_version" "koatsu_env" {
  secret_id = aws_secretsmanager_secret.koatsu_env.id
  secret_string = jsonencode({
    # AWS
    AWS_ACCESS_KEY_ID     = var.aws_access_key_id
    AWS_SECRET_ACCESS_KEY = var.aws_secret_access_key
    bucket_name           = var.bucket_name
    project               = var.project
    region                = var.aws_region

    # Client Certificate
    CLIENT_CERT_BASE64 = var.client_cert_base64
    CLIENT_KEY_BASE64  = var.client_key_base64

    # Scraping URLs - 中国
    CHUGOKU_DOMAIN = var.chugoku_domain
    CHUGOKU_URL    = var.chugoku_url

    # Scraping URLs - 東北
    TOHOKU_DOMAIN = var.tohoku_domain
    TOHOKU_URL    = var.tohoku_url

    # Scraping URLs - 関東
    KANTO_DOMAIN = var.kanto_domain
    KANTO_URL    = var.kanto_url

    # Scraping URLs - 中部
    TYUBU_DOMAIN = var.tyubu_domain
    TYUBU_URL    = var.tyubu_url

    # Scraping URLs - 北陸
    HOKURIKU_DOMAIN = var.hokuriku_domain
    HOKURIKU_URL    = var.hokuriku_url

    # Scraping URLs - 九州
    KYUSYU_DOMAIN = var.kyusyu_domain
    KYUSYU_URL    = var.kyusyu_url

    # Scraping URLs - 関西
    KANSAI_DOMAIN = var.kansai_domain
    KANSAI_URL    = var.kansai_url

    # Scraping URLs - 北海道
    HOKKAIDO_DOMAIN = var.hokkaido_domain
    HOKKAIDO_URL    = var.hokkaido_url

    # Scraping URLs - 四国
    SHIKOKU_DOMAIN = var.shikoku_domain
    SHIKOKU_URL    = var.shikoku_url
  })
}

# VPC - デフォルトVPCを使用（既存のVPCがある場合は適宜変更）
data "aws_vpc" "default" {
  default = true
}

data "aws_subnets" "default" {
  filter {
    name   = "vpc-id"
    values = [data.aws_vpc.default.id]
  }
}

# Security Group for ECS Task
resource "aws_security_group" "ecs_task" {
  name        = "koatsu-scraper-ecs-task"
  description = "Security group for koatsu scraper ECS task"
  vpc_id      = data.aws_vpc.default.id

  egress {
    from_port   = 0
    to_port     = 0
    protocol    = "-1"
    cidr_blocks = ["0.0.0.0/0"]
  }

  tags = {
    Name = "koatsu-scraper-ecs-task"
  }
}

# ECS Cluster
resource "aws_ecs_cluster" "koatsu_scraper" {
  name = "koatsu-scraper-cluster"

  setting {
    name  = "containerInsights"
    value = "enabled"
  }
}

# ECS Task Definition
resource "aws_ecs_task_definition" "koatsu_scraper" {
  family                   = "koatsu-scraper-task"
  requires_compatibilities = ["FARGATE"]
  network_mode             = "awsvpc"
  cpu                      = var.task_cpu
  memory                   = var.task_memory
  execution_role_arn       = aws_iam_role.ecs_task_execution_role.arn
  task_role_arn            = aws_iam_role.ecs_task_role.arn

  container_definitions = jsonencode([
    {
      name      = "koatsu-scraper"
      image     = var.ecr_image_url
      essential = true

      environment = [
        {
          name  = "TZ"
          value = "Asia/Tokyo"
        },
        {
          name  = "PYTHONUNBUFFERED"
          value = "1"
        },
        {
          name  = "OPENSSL_CONF"
          value = "/etc/ssl/openssl.cnf"
        },
        {
          name  = "SSL_CERT_FILE"
          value = "/etc/ssl/certs/ca-certificates.crt"
        },
        {
          name  = "OPENSSL_ALLOW_UNSAFE_LEGACY_RENEGOTIATION"
          value = "1"
        }
      ]

      secrets = [
        # AWS
        {
          name      = "AWS_ACCESS_KEY_ID"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:AWS_ACCESS_KEY_ID::"
        },
        {
          name      = "AWS_SECRET_ACCESS_KEY"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:AWS_SECRET_ACCESS_KEY::"
        },
        {
          name      = "bucket_name"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:bucket_name::"
        },
        {
          name      = "project"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:project::"
        },
        {
          name      = "region"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:region::"
        },
        # Client Certificate
        {
          name      = "CLIENT_CERT_BASE64"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:CLIENT_CERT_BASE64::"
        },
        {
          name      = "CLIENT_KEY_BASE64"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:CLIENT_KEY_BASE64::"
        },
        # Scraping URLs - 中国
        {
          name      = "CHUGOKU_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:CHUGOKU_DOMAIN::"
        },
        {
          name      = "CHUGOKU_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:CHUGOKU_URL::"
        },
        # Scraping URLs - 東北
        {
          name      = "TOHOKU_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:TOHOKU_DOMAIN::"
        },
        {
          name      = "TOHOKU_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:TOHOKU_URL::"
        },
        # Scraping URLs - 関東
        {
          name      = "KANTO_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:KANTO_DOMAIN::"
        },
        {
          name      = "KANTO_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:KANTO_URL::"
        },
        # Scraping URLs - 中部
        {
          name      = "TYUBU_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:TYUBU_DOMAIN::"
        },
        {
          name      = "TYUBU_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:TYUBU_URL::"
        },
        # Scraping URLs - 北陸
        {
          name      = "HOKURIKU_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:HOKURIKU_DOMAIN::"
        },
        {
          name      = "HOKURIKU_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:HOKURIKU_URL::"
        },
        # Scraping URLs - 九州
        {
          name      = "KYUSYU_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:KYUSYU_DOMAIN::"
        },
        {
          name      = "KYUSYU_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:KYUSYU_URL::"
        },
        # Scraping URLs - 関西
        {
          name      = "KANSAI_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:KANSAI_DOMAIN::"
        },
        {
          name      = "KANSAI_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:KANSAI_URL::"
        },
        # Scraping URLs - 北海道
        {
          name      = "HOKKAIDO_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:HOKKAIDO_DOMAIN::"
        },
        {
          name      = "HOKKAIDO_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:HOKKAIDO_URL::"
        },
        # Scraping URLs - 四国
        {
          name      = "SHIKOKU_DOMAIN"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:SHIKOKU_DOMAIN::"
        },
        {
          name      = "SHIKOKU_URL"
          valueFrom = "${aws_secretsmanager_secret.koatsu_env.arn}:SHIKOKU_URL::"
        }
      ]

      logConfiguration = {
        logDriver = "awslogs"
        options = {
          "awslogs-group"         = aws_cloudwatch_log_group.koatsu_scraper.name
          "awslogs-region"        = var.aws_region
          "awslogs-stream-prefix" = "ecs"
        }
      }

      healthCheck = {
        command     = ["CMD-SHELL", "ps aux | grep python || exit 1"]
        interval    = 30
        timeout     = 5
        retries     = 3
        startPeriod = 60
      }
    }
  ])
}

# EventBridge Role
resource "aws_iam_role" "eventbridge_role" {
  name = "koatsu-scraper-eventbridge-role"

  assume_role_policy = jsonencode({
    Version = "2012-10-17"
    Statement = [
      {
        Action = "sts:AssumeRole"
        Effect = "Allow"
        Principal = {
          Service = "events.amazonaws.com"
        }
      }
    ]
  })
}

resource "aws_iam_role_policy" "eventbridge_ecs_policy" {
  name = "eventbridge-ecs-policy"
  role = aws_iam_role.eventbridge_role.id

  policy = jsonencode({
    Version = "2012-10-17"
    Statement = [
      {
        Effect = "Allow"
        Action = [
          "ecs:RunTask"
        ]
        Resource = aws_ecs_task_definition.koatsu_scraper.arn
      },
      {
        Effect = "Allow"
        Action = [
          "iam:PassRole"
        ]
        Resource = [
          aws_iam_role.ecs_task_execution_role.arn,
          aws_iam_role.ecs_task_role.arn
        ]
      }
    ]
  })
}

# EventBridge Rule - 毎日午前3時（JST）= 18:00 UTC
resource "aws_cloudwatch_event_rule" "koatsu_scraper_schedule" {
  name                = "koatsu-scraper-schedule"
  description         = "Run koatsu scraper daily at 3 AM JST"
  schedule_expression = var.schedule_expression
  state               = var.schedule_enabled ? "ENABLED" : "DISABLED"
}

resource "aws_cloudwatch_event_target" "koatsu_scraper_schedule" {
  rule      = aws_cloudwatch_event_rule.koatsu_scraper_schedule.name
  target_id = "koatsu-scraper-task"
  arn       = aws_ecs_cluster.koatsu_scraper.arn
  role_arn  = aws_iam_role.eventbridge_role.arn

  ecs_target {
    task_count          = 1
    task_definition_arn = aws_ecs_task_definition.koatsu_scraper.arn
    launch_type         = "FARGATE"
    platform_version    = "LATEST"

    network_configuration {
      subnets          = data.aws_subnets.default.ids
      security_groups  = [aws_security_group.ecs_task.id]
      assign_public_ip = true
    }
  }
}
