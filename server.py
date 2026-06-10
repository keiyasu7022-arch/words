#!/usr/bin/env python3
import json, os, glob, sys, webbrowser, threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '-q'])
    import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = 8765

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        if self.path == '/api/excel-files':
            files = glob.glob(os.path.join(BASE_DIR, '*.xlsx')) + \
                    glob.glob(os.path.join(BASE_DIR, '*.xls'))
            names = sorted(os.path.basename(f) for f in files)
            self._send_json(names)

        elif self.path.startswith('/api/load?file='):
            raw = self.path.split('=', 1)[1]
            filename = urllib.parse.unquote(raw)
            filepath = os.path.join(BASE_DIR, filename)
            if not os.path.isfile(filepath):
                self.send_error(404, 'File not found')
                return
            words = []
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                ja = str(row[0] if row[0] is not None else '').strip()
                en = str(row[1] if row[1] is not None else '').strip()
                if ja and en:
                    words.append({'ja': ja, 'en': en})
            wb.close()
            self._send_json({'filename': filename, 'words': words})

        else:
            super().do_GET()

    def _send_json(self, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *_):
        pass  # ログを静かにする

def _open_browser():
    import time; time.sleep(0.6)
    webbrowser.open(f'http://localhost:{PORT}')

if __name__ == '__main__':
    # スマホ用にLAN IPも表示
    import socket
    try:
        lan_ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        lan_ip = '(取得失敗)'

    print(f'✅ 単語帳サーバー起動中')
    print(f'   PC:    http://localhost:{PORT}')
    print(f'   スマホ: http://{lan_ip}:{PORT}  ← 同じWi-Fiなら使えます')
    print(f'終了: Ctrl+C')

    threading.Thread(target=_open_browser, daemon=True).start()
    try:
        HTTPServer(('', PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print('\n停止しました')
