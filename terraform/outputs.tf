output "ecs_cluster_name" {
  description = "ECS Cluster name"
  value       = aws_ecs_cluster.koatsu_scraper.name
}

output "ecs_cluster_arn" {
  description = "ECS Cluster ARN"
  value       = aws_ecs_cluster.koatsu_scraper.arn
}

output "ecs_task_definition_arn" {
  description = "ECS Task Definition ARN"
  value       = aws_ecs_task_definition.koatsu_scraper.arn
}

output "cloudwatch_log_group" {
  description = "CloudWatch Log Group name"
  value       = aws_cloudwatch_log_group.koatsu_scraper.name
}

output "security_group_id" {
  description = "Security Group ID for ECS Task"
  value       = aws_security_group.ecs_task.id
}

output "eventbridge_rule_name" {
  description = "EventBridge Rule name"
  value       = aws_cloudwatch_event_rule.koatsu_scraper_schedule.name
}

output "secrets_manager_arn" {
  description = "Secrets Manager ARN"
  value       = aws_secretsmanager_secret.koatsu_env.arn
}

output "run_task_command" {
  description = "AWS CLI command to manually run the task"
  value = <<-EOT
    aws ecs run-task \
      --cluster ${aws_ecs_cluster.koatsu_scraper.name} \
      --task-definition ${aws_ecs_task_definition.koatsu_scraper.family} \
      --launch-type FARGATE \
      --network-configuration "awsvpcConfiguration={subnets=[${join(",", data.aws_subnets.default.ids)}],securityGroups=[${aws_security_group.ecs_task.id}],assignPublicIp=ENABLED}" \
      --region ${var.aws_region}
  EOT
}

output "logs_command" {
  description = "AWS CLI command to view logs"
  value       = "aws logs tail ${aws_cloudwatch_log_group.koatsu_scraper.name} --follow --region ${var.aws_region}"
}
