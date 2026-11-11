"""Docenti PDF to Word Module - Powered by pdf2docx.
Menyediakan antarmuka langsung untuk konversi berkas PDF ke DOCX presisi tinggi.
"""
import io
import os
import sys

# Tambahkan direktori modul ini ke sys.path
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if _CURRENT_DIR not in sys.path:
    sys.path.insert(0, _CURRENT_DIR)

from .pdf2docx import Converter, parse

def convert_pdf_to_docx(pdf_input, docx_output=None, **kwargs):
    """
    Konversi PDF (path berkas, bytes, atau stream) ke DOCX.
    Jika docx_output tidak ditentukan, fungsi ini mengembalikan bytes DOCX.
    """
    if isinstance(pdf_input, (bytes, bytearray)):
        cv = Converter(stream=pdf_input)
    elif hasattr(pdf_input, "read"):
        cv = Converter(stream=pdf_input.read())
    else:
        cv = Converter(pdf_file=str(pdf_input))

    try:
        if docx_output is None or isinstance(docx_output, io.BytesIO):
            out_stream = docx_output if docx_output is not None else io.BytesIO()
            cv.convert(out_stream, **kwargs)
            if docx_output is None:
                out_stream.seek(0)
                return out_stream.getvalue()
            return out_stream
        else:
            cv.convert(docx_output, **kwargs)
            return docx_output
    finally:
        cv.close()

__all__ = ["Converter", "parse", "convert_pdf_to_docx"]

