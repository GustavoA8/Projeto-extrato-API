import requests
import base64
import ssl
import OpenSSL

CLIENT_ID = "2a461a72-04ed-444e-9897-65ac0456d674"
CLIENT_SECRET = "65f68416-a46c-4970-b512-db66bb718858"

CERTIFICADO = "certificado.crt"
CHAVE = "chave.key"

URL_TOKEN = "https://openapisandbox.prebanco.com.br/auth/server-mtls/v2/token"

# ===== BASIC AUTH =====
basic = base64.b64encode(
    f"{CLIENT_ID}:{CLIENT_SECRET}".encode()
).decode()

headers = {
    "Content-Type": "application/x-www-form-urlencoded",
    "Authorization": f"Basic {basic}"
}

data = {
    "grant_type": "client_credentials",
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET
}

r = requests.post(
    URL_TOKEN,
    headers=headers,
    data=data,
    cert=(CERTIFICADO, CHAVE),
    verify=True
)

print("Status:", r.status_code)
print("Resposta:", r.text)