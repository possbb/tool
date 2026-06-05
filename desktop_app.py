import socket
import threading
import time
import webbrowser

from app import app


def find_free_port(start=5000, end=5099):
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.2)
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError("找不到可用端口，请关闭其他占用 5000-5099 端口的程序后重试。")


def open_browser(url):
    time.sleep(1.2)
    webbrowser.open(url)


def main():
    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    threading.Thread(target=open_browser, args=(url,), daemon=True).start()
    print("时间列自动转行工具已启动")
    print(f"如果浏览器没有自动打开，请手动访问：{url}")
    print("关闭这个窗口即可退出程序。")
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
