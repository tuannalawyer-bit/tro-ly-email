"""Cài chứng chỉ localhost vào kho Trusted Root của người dùng hiện tại.

Vì sao bắt buộc: task pane của add-in chạy trong WebView2, mà WebView2 tin kho chứng
chỉ của Windows. Nếu chứng chỉ chưa được tin cậy thì Outlook không tải nổi icon ribbon
(nút "Soạn trả lời AI" không xuất hiện) và task pane hiện trắng trơn không báo lỗi.

Dùng certutil -user nên KHÔNG cần quyền Administrator. Windows có thể hiện một hộp
thoại cảnh báo bảo mật — đó là hành vi đúng khi thêm một root CA, hãy bấm Yes.
"""
import json
import ssl
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes

from paths import CERT_DIR

CRT = CERT_DIR / "localhost.crt"
PROBE_URL = "https://localhost:8765/addin/assets/icon-32.png"


def run(args, **kw):
    return subprocess.run(args, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", **kw)


def powershell(script: str):
    return run(["powershell.exe", "-NoProfile", "-NonInteractive",
                "-ExecutionPolicy", "Bypass", "-Command", script])


def thumbprint_of(path: Path) -> str:
    cert = x509.load_pem_x509_certificate(path.read_bytes())
    return cert.fingerprint(hashes.SHA1()).hex().upper()


def existing_localhost_certs():
    """Liệt kê các chứng chỉ CN=localhost đang nằm trong Trusted Root của user."""
    res = powershell(
        "Get-ChildItem Cert:\\CurrentUser\\Root -ErrorAction SilentlyContinue | "
        "Where-Object { $_.Subject -like '*CN=localhost*' } | "
        "Select-Object Thumbprint, Subject, NotAfter | ConvertTo-Json -Compress")
    out = (res.stdout or "").strip()
    if not out:
        return []
    try:
        data = json.loads(out)
    except json.JSONDecodeError:
        return []
    return data if isinstance(data, list) else [data]


def probe_trust() -> tuple:
    """Gọi HTTPS thật bằng bộ xác thực mặc định — đúng cách Outlook/WebView2 làm."""
    try:
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(PROBE_URL, context=ctx, timeout=10) as r:
            return True, f"HTTP {r.status}, {len(r.read())} bytes"
    except urllib.error.URLError as e:
        reason = getattr(e, "reason", e)
        if isinstance(reason, ssl.SSLCertVerificationError):
            return False, f"chung chi KHONG duoc tin cay: {reason.verify_message}"
        return False, f"khong ket noi duoc: {reason}"
    except Exception as e:
        return False, str(e)


def main() -> int:
    print("=" * 66)
    print("  Cai chung chi localhost vao Trusted Root")
    print("=" * 66)

    if not CRT.exists():
        print(f"\n[LOI] Khong tim thay {CRT}")
        print("      Chay TAO_CHUNG_CHI_ADDIN.bat truoc.")
        return 1

    want = thumbprint_of(CRT)
    print(f"\nChung chi can cai : {CRT.name}")
    print(f"Thumbprint        : {want}")

    # Dọn các chứng chỉ localhost cũ. Mỗi lần sinh lại chứng chỉ là một thumbprint
    # mới, không dọn thì kho Root tích tụ dần các chứng chỉ chết.
    already = False
    for c in existing_localhost_certs():
        tp = (c.get("Thumbprint") or "").upper()
        if tp == want:
            already = True
            continue
        print(f"\nGo chung chi localhost cu (thumbprint {tp})...")
        res = run(["certutil", "-user", "-delstore", "Root", tp])
        print("  " + ("da go." if res.returncode == 0
                      else f"khong go duoc: {(res.stdout or res.stderr).strip()[:120]}"))

    if already:
        print("\nChung chi nay DA nam trong Trusted Root.")
    else:
        print("\nDang cai vao Trusted Root (CurrentUser)...")
        print("  Neu Windows hien hop thoai canh bao bao mat, hay bam YES.")
        res = run(["certutil", "-user", "-addstore", "Root", str(CRT)])
        if res.returncode != 0:
            print("\n[LOI] certutil that bai:")
            print((res.stdout or "") + (res.stderr or ""))
            print("\nCach cai thu cong: nhay dup certs\\localhost.crt ->")
            print("  Install Certificate -> Current User -> Place all certificates")
            print("  in the following store -> Browse -> Trusted Root Certification")
            print("  Authorities -> OK -> Next -> Finish -> Yes")
            return 1
        print("  Da cai xong.")

    # Kiểm chứng bằng kết nối thật, không tin vào mã trả về của certutil.
    print("\n" + "-" * 66)
    print("Kiem chung bang ket noi HTTPS that toi backend...")
    ok, detail = probe_trust()
    if ok:
        print(f"  DAT — {detail}")
        print("\nWindows da tin chung chi. Outlook se tai duoc icon ribbon")
        print("va task pane se hien noi dung.")
        print("\n>>> BUOC TIEP THEO: go add-in cu trong Outlook, dong han Outlook,")
        print("    xoa %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef, roi cai lai manifest.")
        return 0

    if "khong ket noi duoc" in detail:
        print(f"  Chua kiem chung duoc — {detail}")
        print("\n  Backend chua chay nen khong thu duoc. Chung chi VAN da duoc cai.")
        print("  Hay chay CHAY_ADDIN_BACKEND.bat roi chay KIEM_TRA_ADDIN.bat de xac nhan.")
        return 0

    print(f"  KHONG DAT — {detail}")
    print("\n  Chung chi da vao kho nhung van khong duoc tin cay. Nguyen nhan thuong gap:")
    print("   - Chung chi thieu basicConstraints CA=true.")
    print("     Sinh lai: .venv\\Scripts\\python.exe tao_chung_chi_addin.py --force")
    print("     roi chay lai CAI_CHUNG_CHI.bat.")
    print("   - Backend dang chay bang chung chi CU. Khoi dong lai CHAY_ADDIN_BACKEND.bat.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
