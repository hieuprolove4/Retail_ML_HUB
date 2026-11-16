r"""
build_submission_zip.py
Đóng gói toàn bộ đồ án hoàn chỉnh (mã nguồn, Web Studio, Dataset, Saved Models, Báo cáo Word/Excel/Flowcharts)
thành file nộp bài D:\HK7\Machine_Learning\doan\ML_Project_Nhom.zip, loại bỏ .venv/, __pycache__/, .git/, *.tmp.
"""

import os
import sys
import zipfile

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT_DIR = os.path.dirname(PROJECT_DIR)
OUT_ZIP = os.path.join(PARENT_DIR, "ML_Project_Nhom.zip")

EXCLUDE_DIRS = {".venv", "venv", "__pycache__", ".git", ".vscode", ".idea", ".ipynb_checkpoints"}
EXCLUDE_EXTS = {".pyc", ".pyo", ".tmp", ".log"}

TOP_DOC_FILES = [
    os.path.join(PARENT_DIR, "KHOA HỆ THỐNG THÔNG TIN (1).docx"),
    os.path.join(PARENT_DIR, "Bao_Cao_Tong_Hop_Du_An_Retail_ML_Hub.xlsx"),
]


def main():
    tmp_zip = OUT_ZIP + ".tmp"
    count = 0
    with zipfile.ZipFile(tmp_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(PROJECT_DIR):
            dirs[:] = sorted(d for d in dirs if d not in EXCLUDE_DIRS)
            for fn in sorted(files):
                if any(fn.lower().endswith(ext) for ext in EXCLUDE_EXTS):
                    continue
                if fn.lower().endswith(".zip") or "_backup" in fn.lower():
                    continue
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, PROJECT_DIR).replace("\\", "/")
                zf.write(full, rel)
                count += 1

        # Đưa 2 bản báo cáo chính đặt ở thư mục gốc file zip (cũng có bản đồng bộ trong reports/)
        for extra in TOP_DOC_FILES:
            if os.path.exists(extra):
                zf.write(extra, os.path.basename(extra))
                count += 1

    os.replace(tmp_zip, OUT_ZIP)

    with zipfile.ZipFile(OUT_ZIP, "r") as zchk:
        bad = zchk.testzip()
        assert bad is None, f"Tệp ZIP bị lỗi thành phần: {bad}"
        total = len(zchk.namelist())

    size_mb = os.path.getsize(OUT_ZIP) / (1024 * 1024)
    print(f"--> Đã đóng gói thành công: {OUT_ZIP}")
    print(f"    Tổng số tệp: {total} | Dung lượng: {size_mb:.2f} MB | Kiểm tra testzip(): OK")


if __name__ == "__main__":
    main()
