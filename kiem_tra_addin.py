"""Kiểm tra mọi điều kiện cần để add-in Outlook chạy được, in bảng Đạt/Không đạt.

Chạy tệp này ĐẦU TIÊN mỗi khi add-in không hoạt động. Mỗi mục Không đạt đều kèm câu
lệnh cụ thể để sửa, nên không phải mò.
"""
import json
import socket
import ssl
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

from paths import ADDIN_RES_DIR, CERT_DIR, MANIFEST_OUT, RES_DIR

BASE_DIR = RES_DIR                     # dùng để đối chiếu URL tài nguyên trong manifest
CRT = CERT_DIR / "localhost.crt"
KEY = CERT_DIR / "localhost.key"
# Kiểm tra bản manifest NGƯỜI DÙNG thực sự thêm vào Outlook nếu đã có, không thì bản gốc.
MANIFEST = MANIFEST_OUT if MANIFEST_OUT.is_file() else ADDIN_RES_DIR / "manifest.xml"
HOST, PORT = "localhost", 8765
BASE_URL = f"https://{HOST}:{PORT}"

NS = "{http://schemas.microsoft.com/office/appforoffice/1.1}"
VO = "{http://schemas.microsoft.com/office/mailappversionoverrides}"
BT = "{http://schemas.microsoft.com/office/officeappbasictypes/1.0}"

results = []


def check(name: str, ok: bool, detail: str = "", fix: str = "",
          required: bool = True) -> bool:
    """required=False: mục nên có nhưng thiếu vẫn chạy được -> chỉ cảnh báo."""
    results.append((name, ok, detail, fix, required))
    tag = "DAT " if ok else ("HONG" if required else "LUU Y")
    print(f"  [{tag}] {name}")
    if detail:
        print(f"         {detail}")
    if not ok and fix:
        for line in fix.splitlines():
            print(f"         -> {line}")
    return ok


# ---------------------------------------------------------------- chứng chỉ

def check_cert():
    print("\n1. CHUNG CHI")
    if not check("Tep chung chi ton tai", CRT.exists() and KEY.exists(),
                 str(CRT), "Chay TAO_CHUNG_CHI_ADDIN.bat"):
        return

    from cryptography import x509
    from cryptography.hazmat.primitives import hashes
    cert = x509.load_pem_x509_certificate(CRT.read_bytes())
    now = datetime.now(timezone.utc)
    tp = cert.fingerprint(hashes.SHA1()).hex().upper()

    check("Chung chi con han", cert.not_valid_before_utc <= now <= cert.not_valid_after_utc,
          f"het han {cert.not_valid_after_utc:%d-%m-%Y}, thumbprint {tp}",
          ".venv\\Scripts\\python.exe tao_chung_chi_addin.py --force\nroi chay CAI_CHUNG_CHI.bat")

    def has(cls):
        try:
            cert.extensions.get_extension_for_class(cls)
            return True
        except x509.ExtensionNotFound:
            return False

    missing = [n for n, c in (("subjectAltName", x509.SubjectAlternativeName),
                              ("basicConstraints", x509.BasicConstraints),
                              ("keyUsage", x509.KeyUsage),
                              ("extendedKeyUsage", x509.ExtendedKeyUsage)) if not has(c)]
    check("Du phan mo rong cho WebView2", not missing,
          "thieu: " + ", ".join(missing) if missing else "co du 4 phan mo rong",
          ".venv\\Scripts\\python.exe tao_chung_chi_addin.py --force\nroi chay CAI_CHUNG_CHI.bat")

    # Kho Trusted Root của người dùng
    import subprocess
    res = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
         "Get-ChildItem Cert:\\CurrentUser\\Root -ErrorAction SilentlyContinue | "
         "Where-Object { $_.Subject -like '*CN=localhost*' } | "
         "Select-Object -ExpandProperty Thumbprint"],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    installed = [l.strip().upper() for l in (res.stdout or "").splitlines() if l.strip()]
    check("Chung chi nam trong Trusted Root", tp in installed,
          f"kho co {len(installed)} chung chi localhost"
          + ("" if tp in installed else " nhung KHONG khop thumbprint hien tai"),
          "Chay CAI_CHUNG_CHI.bat")


# ------------------------------------------------------------------ backend

def check_backend():
    print("\n2. BACKEND")
    listening = False
    try:
        with socket.create_connection((HOST, PORT), timeout=3):
            listening = True
    except OSError:
        pass
    if not check(f"Cong {PORT} dang lang nghe", listening, "",
                 "Chay CHAY_ADDIN_BACKEND.bat va de cua so do mo"):
        return

    # Đây là phép thử quyết định: dùng ĐÚNG bộ xác thực mặc định như Outlook/WebView2.
    ok, detail, fix = False, "", "Chay CAI_CHUNG_CHI.bat"
    try:
        with urllib.request.urlopen(f"{BASE_URL}/addin/assets/icon-32.png",
                                    context=ssl.create_default_context(), timeout=10) as r:
            ok, detail = True, f"HTTP {r.status}, {len(r.read())} bytes"
    except urllib.error.URLError as e:
        reason = getattr(e, "reason", e)
        if isinstance(reason, ssl.SSLCertVerificationError):
            detail = f"Windows tu choi chung chi: {reason.verify_message}"
        else:
            detail = str(reason)
    except Exception as e:
        detail = str(e)
    check("Icon ribbon tai duoc qua HTTPS tin cay", ok, detail, fix)

    # Backend tự trả lời (bỏ qua xác thực, chỉ để lấy trạng thái ứng dụng)
    try:
        ctx = ssl._create_unverified_context()
        with urllib.request.urlopen(f"{BASE_URL}/taskpane.html", context=ctx, timeout=10) as r:
            html = r.read().decode("utf-8")
        import re as _re
        m = _re.search(r'name="addin-token" content="([^"]+)"', html)
        token = m.group(1) if m else ""
        check("Token duoc chen vao taskpane.html",
              bool(token) and not token.startswith("{{"),
              f"dai {len(token)} ky tu" if token else "khong tim thay",
              "Kiem tra addin_server.py va addin/taskpane.html")
        if token:
            req = urllib.request.Request(f"{BASE_URL}/api/health",
                                         headers={"X-Addin-Token": token})
            with urllib.request.urlopen(req, context=ctx, timeout=15) as r:
                data = json.loads(r.read().decode("utf-8")).get("data", {})
            check("API key Gemini da cau hinh", bool(data.get("api_key_configured")),
                  f"model: {data.get('model')}",
                  "Mo ung dung desktop (CHAY_UNGDUNG.bat) -> Cai dat -> nhap API key")
            kb = data.get("knowledge", {})
            # Hai muc duoi khong chan add-in chay, chi lam thu soan ra kem chat luong.
            check("Kho tri thuc da co", bool(kb.get("dir_exists")),
                  f"{kb.get('file_count', 0)} tep, {kb.get('total_chars', 0)} ky tu",
                  "Copy-Item -Recurse kien_thuc_mau kien_thuc\n"
                  "roi SUA LAI cho dung thong tin cua ban (ban mau la danh tinh gia)",
                  required=False)
            check("Ho so van phong da hoc", bool(kb.get("has_learned")),
                  f"{kb.get('sample_count', 0)} mau",
                  "Mo ung dung desktop -> Cai dat -> Phan tich van phong",
                  required=False)
    except Exception as e:
        check("Goi duoc API backend", False, str(e), "Khoi dong lai CHAY_ADDIN_BACKEND.bat")


# ----------------------------------------------------------------- manifest

def check_manifest():
    print("\n3. MANIFEST")
    if not check("Tep manifest ton tai", MANIFEST.exists(), str(MANIFEST)):
        return
    try:
        root = ET.parse(MANIFEST).getroot()
    except ET.ParseError as e:
        check("Manifest la XML hop le", False, str(e))
        return
    check("Manifest la XML hop le", True)

    children = [c.tag.replace(NS, "").replace(VO, "") for c in root]
    check("VersionOverrides nam cuoi", children and children[-1] == "VersionOverrides",
          f"phan tu cuoi: {children[-1] if children else '(trong)'}",
          "Di chuyen <VersionOverrides> xuong cuoi <OfficeApp>")

    perm = root.find(f"{NS}Permissions")
    check("Quyen la ReadItem", perm is not None and perm.text == "ReadItem",
          perm.text if perm is not None else "khong khai bao")

    dff = root.find(f".//{VO}DesktopFormFactor")
    check("FunctionFile la con dau tien cua DesktopFormFactor",
          dff is not None and len(dff) > 0 and list(dff)[0].tag == f"{VO}FunctionFile",
          "", "Dua <FunctionFile> len dau <DesktopFormFactor>")

    defined = {el.get("id") for tag in (f"{BT}Image", f"{BT}Url", f"{BT}String")
               for el in root.iter(tag) if el.get("id")}
    used = {el.get("resid") for el in root.iter() if el.get("resid")}
    check("Moi resid da duoc dinh nghia", not (used - defined),
          f"{len(used)} resid" if not (used - defined) else f"thieu: {used - defined}")

    missing = []
    for el in root.iter():
        v = el.get("DefaultValue") or ""
        if v.startswith(f"{BASE_URL}/") and not (BASE_DIR / v[len(BASE_URL) + 1:]).is_file():
            missing.append(v[len(BASE_URL) + 1:])
    check("Moi URL tro toi tep co that", not missing,
          "" if not missing else "thieu: " + ", ".join(missing))

    gid = root.find(f"{NS}Id")
    old = "1f8f5c4a-5df6-4c77-9c69-9e7c8c3a2a11"
    check("Id khac ban v1.3 cu", gid is not None and gid.text != old,
          gid.text if gid is not None else "",
          "Doi <Id> sang GUID moi de Outlook khong dung cache cua ban cu")


def main() -> int:
    print("=" * 66)
    print("  KIEM TRA DIEU KIEN CHAY ADD-IN OUTLOOK")
    print("=" * 66)
    check_cert()
    check_backend()
    check_manifest()

    failed = [r for r in results if not r[1] and r[4]]
    warned = [r for r in results if not r[1] and not r[4]]

    print("\n" + "=" * 66)
    if failed:
        print(f"  {len(failed)}/{len(results)} MUC BAT BUOC KHONG DAT:")
        for r in failed:
            print(f"    - {r[0]}")
    else:
        print("  MOI DIEU KIEN BAT BUOC DEU DAT.")
        print("  Neu van khong thay nut ribbon: go add-in trong Outlook,")
        print("  dong han Outlook, xoa %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef,")
        print("  roi cai lai addin\\manifest.xml.")
    if warned:
        print(f"\n  {len(warned)} muc nen bo sung (khong chan add-in chay):")
        for r in warned:
            print(f"    - {r[0]}")
    print("=" * 66)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
