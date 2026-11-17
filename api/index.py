import os
import sys
import io
import urllib.parse
from flask import Flask, request, send_file, jsonify

# Pastikan BASE_DIR ada di sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

try:
    from pdf_to_word import convert_pdf_to_docx
    HAS_ENGINE = True
except Exception as e:
    HAS_ENGINE = False
    ENGINE_ERROR = str(e)

try:
    from pdf_compress import compress_pdf
    HAS_COMPRESS = True
except Exception as e:
    HAS_COMPRESS = False

app = Flask(__name__)

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Expose-Headers"] = "Content-Disposition, Content-Type, X-Original-Size, X-Compressed-Size, X-Compress-Method, X-Compress-Quality"
    return response

@app.route("/api/health", methods=["GET", "OPTIONS"])
def health():
    if request.method == "OPTIONS":
        return "", 200
    return jsonify({
        "status": "ok",
        "engine": "pdf_to_word",
        "compress_ready": HAS_COMPRESS,
        "ready": HAS_ENGINE,
        "message": "Docenti Cloud Engine Aktif (Vercel Serverless)" if HAS_ENGINE else f"Engine error: {ENGINE_ERROR}"
    })

@app.route("/api/convert", methods=["POST", "OPTIONS"])
def convert():
    if request.method == "OPTIONS":
        return "", 200

    if not HAS_ENGINE:
        return jsonify({"error": f"Engine pdf_to_word belum siap di cloud: {ENGINE_ERROR}"}), 500

    pdf_bytes = None
    filename = "dokumen.docx"

    # 1. Dari multipart/form-data
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
    # 2. Dari raw body binary
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
        docx_data = convert_pdf_to_docx(pdf_bytes)
        out_stream = io.BytesIO(docx_data)
        return send_file(
            out_stream,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
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
