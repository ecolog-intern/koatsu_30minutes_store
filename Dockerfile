FROM python:3.11-slim

# システムパッケージの更新と必要な依存関係のインストール
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 作業ディレクトリの設定
WORKDIR /app

# requirements.txtをコピーして依存関係をインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Playwrightのインストール
RUN pip install playwright python-dotenv
RUN playwright install --with-deps chromium firefox

# OpenSSL設定を修正（legacy SSL renegotiation対応）
RUN sed -i '/\[openssl_init\]/a ssl_conf = ssl_sect' /etc/ssl/openssl.cnf && \
    echo "\n[ssl_sect]\nsystem_default = system_default_sect\n\n[system_default_sect]\nOptions = UnsafeLegacyServerConnect" >> /etc/ssl/openssl.cnf

# 環境変数でOpenSSL設定を強制
ENV OPENSSL_CONF=/etc/ssl/openssl.cnf
ENV SSL_CERT_FILE=/etc/ssl/certs/ca-certificates.crt
ENV OPENSSL_ALLOW_UNSAFE_LEGACY_RENEGOTIATION=1

# アプリケーションファイルをコピー
# COPY app/ ./app/
COPY ./app ./
ENV PYTHONPATH=/app

# appのmain.pyを実行
# デバッグ用: bashで起動する場合はこちらをコメントアウト解除
# CMD ["bash"]
# CMD ["tail", "-f", "/dev/null"]
CMD ["python", "main.py"]
