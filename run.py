"""
Entry point — Tìm port trống, chạy Streamlit, mở browser.

Dùng cho:
  - Chạy trực tiếp:  python run.py
  - Chạy từ .exe:   PyInstaller gọi run.py làm entry point.
"""

from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path


def find_free_port(start: int = 8500, end: int = 9000) -> int:
    """
    Tìm port TCP trống trong khoảng [start, end].
    Trả về port đầu tiên tìm được.
    """
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            result = sock.connect_ex(("127.0.0.1", port))
            if result != 0:
                return port
    # Nếu không tìm được port nào, mặc định trả về 8500
    # (sẽ fail ở bước spawn, nhưng vẫn có port để báo lỗi)
    return start


def resolve_app_path() -> str:
    """
    Tìm đường dẫn đến app.py.
    Khi chạy từ .exe (PyInstaller), sys._MEIPASS trỏ đến thư mục tạm.
    """
    # Nếu chạy từ .exe
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)  # type: ignore[arg-type]
    else:
        base = Path(__file__).parent

    app_path = base / "app.py"
    if app_path.exists():
        return str(app_path.resolve())

    # Fallback: thử tìm trong thư mục hiện tại
    fallback = Path.cwd() / "app.py"
    if fallback.exists():
        return str(fallback.resolve())

    raise FileNotFoundError(f"Không tìm thấy app.py tại {app_path} hoặc {fallback}")


def wait_for_server(url: str, timeout: int = 10, interval: float = 0.5) -> bool:
    """
    Chờ server Streamlit sẵn sàng bằng cách poll HTTP.
    Trả về True nếu server up trong thời gian timeout.
    """
    import urllib.request
    import urllib.error

    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            urllib.request.urlopen(url, timeout=2)
            return True
        except (urllib.error.URLError, ConnectionResetError, OSError):
            time.sleep(interval)
    return False


def main() -> None:
    """
    Hàm chính:
      1. Tìm port trống.
      2. Spawn Streamlit subprocess.
      3. Đợi server sẵn sàng.
      4. Mở browser.
      5. Giữ process cha sống đến khi user tắt.
    """
    # Cho phép ghi đè port qua biến môi trường (debug)
    port = int(os.environ.get("STREAMLIT_PORT", "0")) or find_free_port()
    app_path = resolve_app_path()

    print(f"Khởi động Streamlit tại port {port}...")
    print(f"App path: {app_path}")

    # Build command
    cmd = [
        sys.executable or "python",
        "-m",
        "streamlit",
        "run",
        app_path,
        "--server.port", str(port),
        "--server.headless", "true",
    ]

    # Nếu chạy từ .exe, thêm --server.address 127.0.0.1 để chỉ listen local
    if getattr(sys, "frozen", False):
        cmd.extend(["--server.address", "127.0.0.1"])

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            bufsize=1,
            universal_newlines=True,
        )
    except FileNotFoundError:
        print("Lỗi: Không tìm thấy Python hoặc streamlit. Hãy chạy 'pip install -r requirements.txt' trước.")
        sys.exit(1)

    url = f"http://localhost:{port}"

    # Đợi server sẵn sàng
    print("Đợi server sẵn sàng...")
    ready = wait_for_server(url)

    if ready:
        print(f"Server đã sẵn sàng tại: {url}")
        webbrowser.open(url)
    else:
        print(f"Cảnh báo: Server chưa sẵn sàng sau 10 giây. Thử mở {url}...")
        webbrowser.open(url)

    # Giữ process cha sống — đọc stdout và in ra terminal
    try:
        if proc.stdout:
            for line in iter(proc.stdout.readline, ""):
                print(line, end="")
    except KeyboardInterrupt:
        print("\nNhận tín hiệu tắt. Đang dừng...")
    finally:
        proc.terminate()
        proc.wait(timeout=5)
        print("Đã dừng server.")


if __name__ == "__main__":
    main()
