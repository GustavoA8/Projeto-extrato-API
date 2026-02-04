from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import serialization
from datetime import datetime, timedelta

# 1. Gerar chave privada
chave_privada = rsa.generate_private_key(
    public_exponent=65537,
    key_size=2048,
)

# 2. Dados do certificado (sandbox pode ser fictício)
subject = issuer = x509.Name([
    x509.NameAttribute(NameOID.COUNTRY_NAME, u"BR"),
    x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, u"Sao Paulo"),
    x509.NameAttribute(NameOID.LOCALITY_NAME, u"Sao Paulo"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, u"Minha Empresa Teste"),
    x509.NameAttribute(NameOID.COMMON_NAME, u"teste.bradesco.sandbox"),
])

# 3. Criar certificado
certificado = (
    x509.CertificateBuilder()
    .subject_name(subject)
    .issuer_name(issuer)
    .public_key(chave_privada.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(datetime.utcnow())
    .not_valid_after(datetime.utcnow() + timedelta(days=365))
    .sign(chave_privada, hashes.SHA256())
)

# 4. Salvar chave privada
with open("chave.key", "wb") as f:
    f.write(
        chave_privada.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )

# 5. Salvar certificado no padrão Bradesco
with open("certificado.crt", "wb") as f:
    f.write(certificado.public_bytes(serialization.Encoding.PEM))

print("Arquivos gerados com sucesso!")
print("-> chave.key")
print("-> certificado.crt")
