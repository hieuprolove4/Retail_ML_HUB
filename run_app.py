"""
Quick Launcher: run_app.py
Chạy trực tiếp từ PyCharm / VS Code (nút Run) hoặc nhấp đúp vào start_app.bat.
Tự động mở trình duyệt tại http://127.0.0.1:8000 và khởi động FastAPI server.
"""

import sys
from pathlib import Path

# Thêm thư mục gốc của dự án vào sys.path để import an toàn
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# Nếu đang chạy trực tiếp run_app.py bằng Python ngoài nhưng có .venv, tự kích hoạt .venv
if __name__ == "__main__":
    venv_python = BASE_DIR / ".venv" / "Scripts" / "python.exe"
    if venv_python.exists() and Path(sys.executable).resolve() != venv_python.resolve():
        import subprocess
        sys.exit(subprocess.call([str(venv_python), str(Path(__file__).resolve())] + sys.argv[1:]))

import threading
import time
import webbrowser
import uvicorn


def _open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    print("=" * 68)
    print("   RETAIL ML HUB - HỆ THỐNG DỰ BÁO NHU CẦU & TỐI ƯU TỒN KHO")
    print("   Đang mở trình duyệt tại: http://127.0.0.1:8000")
    print("=" * 68)
    threading.Thread(target=_open_browser, daemon=True).start()
    uvicorn.run("backend.api:app", host="127.0.0.1", port=8000, reload=False)
