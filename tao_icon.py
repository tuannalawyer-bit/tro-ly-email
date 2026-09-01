"""Sinh addin/assets/app.ico từ các tệp PNG sẵn có. Chạy một lần, kết quả được commit.

PyInstaller chỉ nhận .ico cho tham số --icon, mà dự án không có Pillow. Dùng
System.Drawing qua pythonnet — đúng cách đã kiểm chứng khi làm icon khay ở v1.9.0,
không thêm phụ thuộc nào.
"""
from __future__ import annotations

import struct
import sys
from pathlib import Path

import clr  # noqa: F401

clr.AddReference("System.Drawing")
from System.Drawing import Bitmap  # noqa: E402
from System.Drawing.Imaging import ImageFormat  # noqa: E402
from System.IO import MemoryStream  # noqa: E402

ASSETS = Path(__file__).resolve().parent / "addin" / "assets"
OUT = ASSETS / "app.ico"
SIZES = (16, 32, 48, 64, 128, 256)


def frame(size: int) -> bytes:
    """Ảnh PNG ở kích thước yêu cầu, lấy từ PNG gần nhất rồi co giãn.

    Nhúng thẳng PNG vào .ico thay vì chuyển sang DIB: định dạng ICO cho phép từ Vista,
    Windows và PyInstaller đều đọc được, và tránh phải tự dựng BITMAPINFOHEADER kèm
    mặt nạ AND. Không đi đường Icon.Save vì Icon dựng từ handle không giữ dữ liệu gốc
    nên ghi ra tệp rỗng.
    """
    src = min(ASSETS.glob("icon-*.png"),
              key=lambda p: abs(int(p.stem.split("-")[1]) - size))
    scaled = Bitmap(Bitmap(str(src)), size, size)
    stream = MemoryStream()
    scaled.Save(stream, ImageFormat.Png)
    return bytes(stream.ToArray())


def build() -> Path:
    frames = [(s, frame(s)) for s in SIZES]
    head = struct.pack("<HHH", 0, 1, len(frames))
    offset = 6 + 16 * len(frames)
    entries, blobs = b"", b""
    for size, blob in frames:
        # 256 được ghi là 0 theo đặc tả ICO.
        dim = 0 if size >= 256 else size
        entries += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(blob), offset)
        blobs += blob
        offset += len(blob)
    OUT.write_bytes(head + entries + blobs)
    return OUT


if __name__ == "__main__":
    p = build()
    print(f"Da tao {p} ({p.stat().st_size:,} byte, {len(SIZES)} kich thuoc)")
    sys.exit(0)
