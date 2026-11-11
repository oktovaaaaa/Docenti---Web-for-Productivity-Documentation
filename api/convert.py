import io
import os
import sys
import urllib.parse
from http.server import BaseHTTPRequestHandler

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF_DIR = os.path.join(ROOT_DIR, "pdf_to_word")
if PDF_DIR not in sys.path:
    sys.path.insert(0, PDF_DIR)

try:
    from pdf2docx import Converter
    HAS_ENGINE = True
except Exception:
    HAS_ENGINE = False

class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(b'{"status":"ok","engine":"pdf2docx"}')

    def do_POST(self):
        if not HAS_ENGINE:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(b'{"error":"Engine pdf2docx gagal dimuat"}')
            return

        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_response(400)
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(b'{"error":"Berkas kosong"}')
                return

            pdf_bytes = self.rfile.read(content_length)

            out_stream = io.BytesIO()
            cv = Converter(stream=pdf_bytes)
            cv.convert(out_stream)
            cv.close()
            out_stream.seek(0)
            docx_data = out_stream.getvalue()

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Content-Disposition", 'attachment; filename="dokumen.docx"')
            self.send_header("Content-Length", str(len(docx_data)))
            self.end_headers()
            self.wfile.write(docx_data)
        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            err_msg = str(e).replace('"', '\\"')
            self.wfile.write(f'{{"error":"{err_msg}"}}'.encode("utf-8"))

