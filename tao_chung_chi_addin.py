"""Sinh chứng chỉ tự ký cho backend add-in chạy ở https://localhost:8765.

Chứng chỉ phải có ĐỦ các phần mở rộng dưới đây thì WebView2 (nhân Chromium, khắt khe
hơn schannel nhiều) mới chấp nhận:
  - subjectAltName        : Chromium BỎ QUA hoàn toàn Common Name, chỉ đọc SAN.
  - basicConstraints CA   : bắt buộc để Windows chấp nhận chứng chỉ tự ký này khi nó
                            nằm trong kho Trusted Root (nó vừa là root vừa là lá).
  - keyUsage              : phải có keyCertSign vì nó tự ký cho chính mình.
  - extendedKeyUsage      : serverAuth.

Sinh lại chứng chỉ sẽ tạo khoá và thumbprint MỚI, làm mất hiệu lực bản đã cài vào
Trusted Root trước đó. Vì vậy script mặc định KHÔNG ghi đè; muốn sinh lại phải thêm
tham số --force, và sau đó bắt buộc chạy lại CAI_CHUNG_CHI.bat.
"""
import sys
from datetime import datetime, timedelta, timezone
from ipaddress import IPv4Address
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

from paths import CERT_DIR

ROOT = CERT_DIR
CRT = ROOT / "localhost.crt"
KEY = ROOT / "localhost.key"
VALID_DAYS = 825          # mốc quen thuộc cho chứng chỉ cục bộ; Chromium miễn trừ
                          # giới hạn 398 ngày cho chứng chỉ chuỗi về root cài tại máy.


def _inspect(path: Path):
    """Trả (còn hạn, đủ phần mở rộng, chứng chỉ) của tệp .crt đang có."""
    try:
        cert = x509.load_pem_x509_certificate(path.read_bytes())
    except Exception:
        return False, False, None
    now = datetime.now(timezone.utc)
    alive = cert.not_valid_before_utc <= now <= cert.not_valid_after_utc

    def has(cls):
        try:
            cert.extensions.get_extension_for_class(cls)
            return True
        except x509.ExtensionNotFound:
            return False

    complete = all(has(c) for c in (x509.SubjectAlternativeName,
                                    x509.BasicConstraints,
                                    x509.KeyUsage,
                                    x509.ExtendedKeyUsage))
    return alive, complete, cert


def create() -> None:
    ROOT.mkdir(exist_ok=True)
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Tro ly Email (local)"),
    ])
    now = datetime.now(timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(minutes=5))    # trừ hao lệch giờ máy
        .not_valid_after(now + timedelta(days=VALID_DAYS))
        .add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName("localhost"),
                x509.IPAddress(IPv4Address("127.0.0.1")),
            ]), critical=False)
        .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True, key_encipherment=True, key_cert_sign=True,
                content_commitment=False, data_encipherment=False, key_agreement=False,
                crl_sign=False, encipher_only=False, decipher_only=False,
            ), critical=True)
        .add_extension(
            x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
        .add_extension(
            x509.SubjectKeyIdentifier.from_public_key(key.public_key()), critical=False)
        .sign(key, hashes.SHA256())
    )

    KEY.write_bytes(key.private_bytes(serialization.Encoding.PEM,
                                      serialization.PrivateFormat.TraditionalOpenSSL,
                                      serialization.NoEncryption()))
    CRT.write_bytes(cert.public_bytes(serialization.Encoding.PEM))

    print("Da tao chung chi moi:")
    print(f"  Tep        : {CRT}")
    print(f"  Thumbprint : {cert.fingerprint(hashes.SHA1()).hex().upper()}")
    print(f"  Het han    : {cert.not_valid_after_utc:%d-%m-%Y}")
    print()
    print(">>> BUOC TIEP THEO BAT BUOC: chay CAI_CHUNG_CHI.bat de cai vao Trusted Root.")
    print("    Khong lam buoc nay thi Outlook se khong tai duoc icon ribbon va")
    print("    task pane se hien trang tron khong bao loi.")


def main() -> None:
    force = "--force" in sys.argv

    if CRT.exists() and KEY.exists() and not force:
        alive, complete, cert = _inspect(CRT)
        if cert is not None and alive and complete:
            print("Da co chung chi hop le, KHONG sinh lai.")
            print(f"  Thumbprint : {cert.fingerprint(hashes.SHA1()).hex().upper()}")
            print(f"  Het han    : {cert.not_valid_after_utc:%d-%m-%Y}")
            print()
            print("Sinh lai se tao khoa moi va lam mat hieu luc ban da cai vao")
            print("Trusted Root. Neu that su muon sinh lai:")
            print("    .venv\\Scripts\\python.exe tao_chung_chi_addin.py --force")
            print()
            print("Neu Outlook van bao loi chung chi, hay chay CAI_CHUNG_CHI.bat.")
            return
        if cert is None:
            print("Chung chi hien tai khong doc duoc -> sinh lai.")
        elif not alive:
            print("Chung chi hien tai da het han -> sinh lai.")
        else:
            print("Chung chi hien tai thieu phan mo rong (basicConstraints/keyUsage/EKU)")
            print("nen WebView2 co the tu choi -> sinh lai.")

    create()


if __name__ == "__main__":
    main()
