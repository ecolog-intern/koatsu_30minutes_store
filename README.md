
EC2にデプロイしたら, 出力先フォルダを作成する
ex) フォルダ保存用/関西

・スクレイピング時に使用する証明書
1.secret managerに以下のjsonを登録
{
  "CLIENT_CERT": "-----BEGIN CERTIFICATE-----\n...snip...\n-----END CERTIFICATE-----",
  "CLIENT_KEY": "-----BEGIN PRIVATE KEY-----\n...snip...\n-----END PRIVATE KEY-----",
  "CA_CHAIN": "-----BEGIN CERTIFICATE-----\n...snip...\n-----END CERTIFICATE-----"
}

2.ECSのタスク定義で環境変数に入れれば使えるっぽい
"containerDefinitions": [
  {
    "name": "コンテナ名",
    "image": "イメージ名？",
    "secrets": [
      {
        "name": "CLIENT_CERT",
        "valueFrom": "arn:aws:secretsmanager:ap-northeast-1:123456789012:secret:my-cert-abc123:CLIENT_CERT::"
      },
      {
        "name": "CLIENT_KEY",
        "valueFrom": "arn:aws:secretsmanager:ap-northeast-1:123456789012:secret:my-cert-abc123:CLIENT_KEY::"
      },
      {
        "name": "CA_CHAIN",
        "valueFrom": "arn:aws:secretsmanager:ap-northeast-1:123456789012:secret:my-cert-abc123:CA_CHAIN::"
      }
    ]
  }
]

具体的な情報は.env内からコピペしてね
