import os
import sys
import io
import json
import logging
import urllib.parse
from flask import Flask, request, send_file, jsonify, send_from_directory

# Setup logging
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(levelname)s: %(message)s")

# Ensure BASE_DIR is in path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

try:
    from pdf_to_word import convert_pdf_to_docx
    HAS_PDF2DOCX = True
    logging.info("Modul pdf_to_word berhasil dimuat!")
except Exception as e:
    HAS_PDF2DOCX = False
    logging.error("Gagal memuat modul pdf_to_word: %s", e)

try:
    from pdf_compress import compress_pdf
    HAS_COMPRESS = True
except Exception as e:
    HAS_COMPRESS = False

app = Flask(__name__, static_folder=BASE_DIR, static_url_path="")

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Expose-Headers"] = "Content-Disposition, Content-Type, X-Original-Size, X-Compressed-Size, X-Compress-Method, X-Compress-Quality"
    return response

@app.route("/", methods=["GET"])
def serve_index():
    return send_from_directory(BASE_DIR, "index.html")

@app.route("/<path:filename>", methods=["GET"])
def serve_static(filename):
    return send_from_directory(BASE_DIR, filename)

@app.route("/api/health", methods=["GET", "OPTIONS"])
def health():
    if request.method == "OPTIONS":
        return "", 200
    return jsonify({
        "status": "ok",
        "engine": "pdf_to_word",
        "compress_ready": HAS_COMPRESS,
        "ready": HAS_PDF2DOCX,
        "message": "Docenti PDF-to-Word Engine Aktif & Siap Digunakan"
    })

@app.route("/api/convert", methods=["POST", "OPTIONS"])
def convert():
    if request.method == "OPTIONS":
        return "", 200

    if not HAS_PDF2DOCX:
        return jsonify({"error": "Engine pdf_to_word belum terpasang atau gagal dimuat."}), 500

    pdf_bytes = None
    filename = "dokumen.docx"

    # 1. Cek dari multipart/form-data
    if "file" in request.files:
        uploaded_file = request.files["file"]
        orig_name = uploaded_file.filename or "dokumen.pdf"
        base_name = os.path.splitext(orig_name)[0]
        filename = f"{base_name}.docx"
        pdf_bytes = uploaded_file.read()
    elif "pdf" in request.files:
        uploaded_file = request.files["pdf"]
        orig_name = uploaded_file.filename or "dokumen.pdf"
        base_name = os.path.splitext(orig_name)[0]
        filename = f"{base_name}.docx"
        pdf_bytes = uploaded_file.read()
    # 2. Cek dari raw body binary
    elif request.data and len(request.data) > 0:
        pdf_bytes = request.data
        raw_name = request.headers.get("X-File-Name", "")
        if raw_name:
            orig_name = urllib.parse.unquote(raw_name)
            base_name = os.path.splitext(orig_name)[0]
            filename = f"{base_name}.docx"

    if not pdf_bytes or len(pdf_bytes) == 0:
        return jsonify({"error": "Berkas PDF tidak ditemukan atau kosong."}), 400

    try:
        logging.info("Memulai konversi PDF (%d bytes)...", len(pdf_bytes))
        docx_data = convert_pdf_to_docx(pdf_bytes)
        logging.info("Konversi selesai, ukuran DOCX: %d bytes", len(docx_data))

        out_stream = io.BytesIO(docx_data)
        return send_file(
            out_stream,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        logging.exception("Kesalahan saat konversi PDF:")
        return jsonify({"error": f"Gagal mengonversi PDF: {str(e)}"}), 500

@app.route("/api/compress", methods=["POST", "OPTIONS"])
def compress():
    if request.method == "OPTIONS":
        return "", 200

    if not HAS_COMPRESS:
        return jsonify({"error": "Engine kompres PDF belum siap."}), 500

    uploaded_file = request.files.get("file") or request.files.get("pdf")
    if uploaded_file is not None:
        orig_name = uploaded_file.filename or "dokumen.pdf"
        pdf_bytes = uploaded_file.read()
    else:
        pdf_bytes = request.data
        orig_name = urllib.parse.unquote(request.headers.get("X-File-Name", "") or "dokumen.pdf")

    if not pdf_bytes:
        return jsonify({"error": "Berkas PDF tidak ditemukan atau kosong."}), 400

    quality = (request.form.get("quality") or request.args.get("quality") or "medium").strip()
    keep_text_raw = (request.form.get("keep_text") or request.args.get("keep_text") or "1").strip().lower()
    keep_text = keep_text_raw not in ("0", "false", "no", "off")

    try:
        result, info = compress_pdf(pdf_bytes, quality=quality, keep_text=keep_text)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Gagal mengompres PDF: {str(e)}"}), 500

    base_name = os.path.splitext(orig_name)[0]
    response = send_file(
        io.BytesIO(result),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"{base_name}_compressed.pdf",
    )
    response.headers["X-Original-Size"] = str(info["original_size"])
    response.headers["X-Compressed-Size"] = str(info["compressed_size"])
    response.headers["X-Compress-Method"] = info["method"]
    response.headers["X-Compress-Quality"] = info["quality"]
    return response

def main():
    import webbrowser
    import threading

    port = 5000
    host = "127.0.0.1"
    url = f"http://localhost:{port}/"

    print("=" * 60)
    print("DOCENTI - SERVER LOKAL & ENGINE PDF TO WORD (PDF2DOCX)")
    print("=" * 60)
    print(f"Status Engine: {'SIAP' if HAS_PDF2DOCX else 'GAGAL'}")
    print(f"Buka browser di: {url}")
    print("Tekan Ctrl+C untuk menghentikan server.")
    print("=" * 60)

    def open_browser():
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Timer(1.2, open_browser).start()
    app.run(host=host, port=port, debug=False, use_reloader=False)

if __name__ == "__main__":
    main()
