# make_pem.py
import os

def pem_files():
    """Secrets Managerから環境変数経由でPEMを/tmpに保存し、pathを返す"""
    cert_path = "/tmp/client_cert.pem"
    key_path = "/tmp/client_key.pem"
    ca_path = "/tmp/ca_chain.pem"

    with open(cert_path, "w") as f:
        f.write(os.environ["CLIENT_CERT"].replace("\\n", "\n"))

    with open(key_path, "w") as f:
        f.write(os.environ["CLIENT_KEY"].replace("\\n", "\n"))

    with open(ca_path, "w") as f:
        f.write(os.environ["CA_CHAIN"].replace("\\n", "\n"))

    return cert_path, key_path, ca_path
