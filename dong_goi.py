"""Dựng bộ cài một tệp: CaiTroLyEmail.exe

Ba chặng:
  1. PyInstaller --onedir  -> dist/TroLyEmail/        (bản chạy nhanh, 6,9 giây)
  2. Nén thư mục đó        -> build/payload.zip
  3. PyInstaller --onefile bộ cài nhỏ, rồi NỐI payload.zip vào đuôi exe

Vì sao không --onefile thẳng cho ứng dụng: đo trên máy đích, onefile mất 20,5 giây MỖI
lần chạy do bung ~800 tệp ra %TEMP%; onedir chỉ mất 6,9 giây. Bộ cài SFX cho người dùng
đúng một tệp mà vẫn giữ tốc độ của onedir.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
BUILD = ROOT / "build"
APP_NAME = "TroLyEmail"
SETUP_NAME = "CaiTroLyEmail"
ICON = ROOT / "addin" / "assets" / "app.ico"

sys.path.insert(0, str(ROOT))
from cai_dat_sfx import VERSION_FILE  # noqa: E402
from config import VERSION  # noqa: E402

# Tài nguyên chỉ đọc nhúng vào gói. Dữ liệu ghi được KHÔNG nhúng: nó sống ở
# %LOCALAPPDATA%\TroLyEmail (xem paths.py).
DATA = [("frontend", "frontend"), ("addin", "addin"), ("kien_thuc_mau", "kien_thuc_mau")]

# Tài liệu tra cứu khi người dùng gặp trục trặc, đi kèm cho tiện.
DOCS = ["HUONG_DAN_CAI_DAT.md", "HUONG_DAN_SU_DUNG.md", "HUONG_DAN_KIEN_THUC.md",
        "HUONG_DAN_ADDIN_LOCAL.md"]

# Hướng dẫn cài đặt phải nằm CẠNH bộ cài chứ không chỉ nằm trong nó: người nhận cần đọc
# trước khi chạy, nhất là đoạn nói SmartScreen sẽ chặn và vì sao.
HANDOFF_DOC = "HUONG_DAN_CAI_DAT.md"

# Module pywin32 nạp ĐỘNG lúc chạy nên PyInstaller không dò ra được.
# win32timezone: pywintypes cần khi đổi ngày giờ của COM sang datetime. Thiếu nó thì
# MỌI email đọc từ Outlook đều hỏng với "No module named 'win32timezone'" — lỗi này đã
# thực sự xảy ra ở bản dựng đầu tiên, không phải phòng xa.
HIDDEN = ["win32timezone"]


def run(cmd: list) -> None:
    print("  $", " ".join(str(c) for c in cmd[:4]), "…")
    subprocess.run(cmd, cwd=str(ROOT), check=True)


def pyinstaller(*args: str) -> list:
    return [sys.executable, "-m", "PyInstaller", "--noconfirm",
            "--distpath", str(DIST), "--workpath", str(BUILD),
            "--specpath", str(BUILD), *args]


def build_app() -> Path:
    print("\n[1/3] Dung ung dung (--onedir)")
    args = ["--onedir", "--windowed", "--name", APP_NAME, "--icon", str(ICON)]
    for mod in HIDDEN:
        args += ["--hidden-import", mod]
    for src, dest in DATA:
        args += ["--add-data", f"{ROOT / src}{';'}{dest}"]
    for doc in DOCS:
        if (ROOT / doc).is_file():
            args += ["--add-data", f"{ROOT / doc}{';'}."]
    run(pyinstaller(*args, str(ROOT / "main.py")))
    app = DIST / APP_NAME
    if not (app / f"{APP_NAME}.exe").is_file():
        raise SystemExit("[LOI] Khong thay exe ung dung sau khi dung.")
    # Dấu phiên bản để bộ cài biết máy đang có bản nào mà quyết định có nâng cấp không.
    (app / VERSION_FILE).write_text(VERSION, encoding="utf-8")
    return app


def zip_app(app: Path) -> Path:
    print("\n[2/3] Nen goi ung dung")
    BUILD.mkdir(parents=True, exist_ok=True)
    base = BUILD / "payload"
    t0 = time.time()
    shutil.make_archive(str(base), "zip", root_dir=str(app.parent), base_dir=app.name)
    z = base.with_suffix(".zip")
    print(f"  payload.zip: {z.stat().st_size / 1024 / 1024:.1f} MB ({time.time() - t0:.0f}s)")
    return z


def build_setup(payload: Path) -> Path:
    print("\n[3/3] Dung bo cai roi noi goi vao duoi")
    run(pyinstaller("--onefile", "--windowed", "--name", SETUP_NAME,
                    "--icon", str(ICON), str(ROOT / "cai_dat_sfx.py")))
    stub = DIST / f"{SETUP_NAME}.exe"
    if not stub.is_file():
        raise SystemExit("[LOI] Khong thay exe bo cai sau khi dung.")

    final = DIST / f"{SETUP_NAME}.exe"
    data = stub.read_bytes() + payload.read_bytes()
    final.write_bytes(data)
    return final


def make_handoff(setup: Path) -> Path:
    """Gom thứ cần đưa cho người khác vào một thư mục, bỏ hết rác của quá trình dựng."""
    print("\n[4/4] Gom bo ban giao")
    out = DIST / "BanGiao"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    shutil.move(str(setup), str(out / setup.name))   # dời, khỏi để lại bản 50 MB thừa
    for doc in DOCS:
        if (ROOT / doc).is_file():
            shutil.copy2(ROOT / doc, out / doc)
    if (ROOT / "addin" / "manifest.xml").is_file():
        shutil.copy2(ROOT / "addin" / "manifest.xml", out / "manifest.xml")
    if (ROOT / "kien_thuc_mau").is_dir():
        shutil.copytree(ROOT / "kien_thuc_mau", out / "kien_thuc", dirs_exist_ok=True)
    return out


def clean_intermediates(app: Path, payload: Path) -> None:
    """Xoá bản onedir và payload.zip: cộng lại hơn 200 MB mà dựng lại lúc nào cũng được."""
    for path in (app, payload, BUILD):
        try:
            shutil.rmtree(path) if path.is_dir() else path.unlink()
        except OSError as e:
            print(f"  (bo qua) khong xoa duoc {path.name}: {e}")


def main() -> int:
    if not ICON.is_file():
        print("[LOI] Thieu addin/assets/app.ico — chay tao_icon.py truoc.")
        return 1

    print("=" * 66)
    print("  DONG GOI TRO LY EMAIL THANH MOT TEP .EXE")
    print("=" * 66)
    t0 = time.time()
    try:
        app = build_app()
        payload = zip_app(app)
        final = build_setup(payload)
        out = make_handoff(final)
        clean_intermediates(app, payload)
    except subprocess.CalledProcessError as e:
        print(f"\n[LOI] PyInstaller that bai (ma {e.returncode}).")
        print("      Nguyen nhan hay gap NHAT: ban dong goi dang chay va khoa tep trong")
        print("      dist\\. Tat no truoc (menu khay -> Thoat, hoac Stop-Process TroLyEmail).")
        print("      Ngoai ra co che tien trinh con cua PyInstaller doi khi chet bat chot;")
        print("      truong hop do chay lai DONG_GOI.bat mot lan nua la qua.")
        return 1
    except SystemExit as e:
        print(f"\n{e}")
        return 1

    size = (out / final.name).stat().st_size / 1024 / 1024
    print()
    print("=" * 66)
    print(f"  XONG sau {time.time() - t0:.0f}s")
    print("=" * 66)
    print(f"  Bo ban giao: {out}")
    for f in sorted(out.iterdir()):
        print(f"    - {f.name}  ({f.stat().st_size / 1024 / 1024:.1f} MB)")
    print()
    print(f"  Chep ca thu muc {out.name} cho nguoi dung. Lan dau chay CaiTroLyEmail.exe")
    print("  se bung vao %LOCALAPPDATA%\\TroLyEmail roi tu thiet lap.")
    print(f"  (kich thuoc bo cai: {size:.1f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
