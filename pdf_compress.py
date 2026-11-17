"""Docenti PDF Compression Engine.

Port dari modul kompresi SafePDF (SafePDF/ops/pdf_compress.py, lisensi GPL-3.0)
yang disesuaikan untuk Docenti:

* Memakai PyMuPDF (sudah menjadi dependensi Docenti) sebagai pengganti
  PyPDF2 + pypdfium2 + Pillow, sehingga tidak perlu menambah library baru.
* Memakai profil kualitas yang sama dengan SafePDF (DPI + kualitas JPEG).
* Strategi SafePDF: buat beberapa kandidat hasil kompresi lalu pilih yang
  ukurannya paling kecil. Kandidatnya:
    1. "struktur"  : kompresi stream & pembersihan objek (lossless).
    2. "gambar"    : downsample + re-encode gambar di dalam PDF, teks tetap vektor.
    3. "raster"    : render ulang halaman menjadi JPEG (metode utama SafePDF).
                     Jika keep_text=True, halaman yang berisi teks digital tidak
                     dirasterisasi agar teks tetap tajam & bisa dicari/di-copy.
* Tidak ada UI Tkinter - fungsi menerima & mengembalikan bytes.
"""

import io
import logging

import pymupdf

logger = logging.getLogger("Docenti.PDFCompress")

# Profil kualitas identik dengan SafePDF.PDFCompressor.QUALITY_PROFILES
QUALITY_PROFILES = {
    "low": {"dpi": 110, "jpeg_quality": 45},
    "medium": {"dpi": 140, "jpeg_quality": 60},
    "high": {"dpi": 180, "jpeg_quality": 75},
    "very_high": {"dpi": 210, "jpeg_quality": 90},
    "ultra": {"dpi": 240, "jpeg_quality": 95},
}

# Alias nama lama dari UI Docenti
QUALITY_ALIASES = {
    "extreme": "low",
    "recommended": "medium",
    "lossless": "lossless",
}

_SAVE_OPTS = dict(
    garbage=4,
    clean=True,
    deflate=True,
    deflate_images=True,
    deflate_fonts=True,
    use_objstms=1,
)

# Jumlah karakter minimal agar sebuah halaman dianggap berisi teks digital
_TEXT_CHAR_THRESHOLD = 20


def _save(doc):
    return doc.tobytes(**_SAVE_OPTS)


def _open(pdf_bytes):
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    if doc.needs_pass:
        doc.close()
        raise ValueError("PDF terkunci kata sandi. Buka kunci terlebih dahulu.")
    return doc


def _compress_structure(pdf_bytes):
    """Setara SafePDF._compress_streams_only: kompres stream tanpa mengubah isi."""
    doc = _open(pdf_bytes)
    try:
        return _save(doc)
    finally:
        doc.close()


def _compress_images(pdf_bytes, dpi, jpeg_quality):
    """Downsample & re-encode gambar yang tertanam; teks/vektor tidak disentuh."""
    doc = _open(pdf_bytes)
    try:
        doc.rewrite_images(
            dpi_threshold=dpi + 10,
            dpi_target=dpi,
            quality=jpeg_quality,
            lossy=True,
            lossless=True,
            bitonal=True,
            color=True,
            gray=True,
        )
        return _save(doc)
    finally:
        doc.close()


def _page_has_text(page):
    try:
        return len(page.get_text("text").strip()) >= _TEXT_CHAR_THRESHOLD
    except Exception:
        return False


def _compress_by_rasterizing(pdf_bytes, dpi, jpeg_quality, keep_text):
    """Setara SafePDF._compress_by_rasterizing: bangun ulang PDF dari render JPEG.

    Jika keep_text=True, halaman berteks digital disalin apa adanya (vektor)
    sehingga hanya halaman hasil scan/foto yang dirasterisasi.
    Mengembalikan None jika tidak ada halaman yang dirasterisasi.
    """
    src = _open(pdf_bytes)
    out = pymupdf.open()
    rasterized = 0
    try:
        scale = max(dpi, 72) / 72.0
        matrix = pymupdf.Matrix(scale, scale)
        for i, page in enumerate(src):
            if keep_text and _page_has_text(page):
                out.insert_pdf(src, from_page=i, to_page=i)
                continue

            pix = page.get_pixmap(matrix=matrix, alpha=False, colorspace=pymupdf.csRGB)
            jpeg = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
            pix = None

            rect = page.rect
            new_page = out.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(new_page.rect, stream=jpeg)
            rasterized += 1

        if rasterized == 0:
            return None
        return _save(out)
    finally:
        out.close()
        src.close()


def compress_pdf(pdf_bytes, quality="medium", keep_text=True):
    """Kompres PDF dan kembalikan (bytes_hasil, info).

    Args:
        pdf_bytes: isi berkas PDF.
        quality: "low" | "medium" | "high" | "very_high" | "ultra" | "lossless"
                 (alias lama: "extreme", "recommended").
        keep_text: jika True, halaman berteks digital tidak pernah dirasterisasi.

    Returns:
        tuple(bytes, dict) - dict berisi original_size, compressed_size,
        method, quality, reduced (bool).
    """
    if not pdf_bytes:
        raise ValueError("Berkas PDF kosong.")

    quality = QUALITY_ALIASES.get(quality, quality)
    if quality != "lossless" and quality not in QUALITY_PROFILES:
        quality = "medium"

    original_size = len(pdf_bytes)
    candidates = []

    def _try(name, fn, *args):
        try:
            data = fn(*args)
            if data:
                candidates.append((name, data))
                logger.info("Kandidat %s: %d bytes", name, len(data))
        except ValueError:
            raise
        except Exception:
            logger.warning("Kandidat %s gagal", name, exc_info=True)

    _try("struktur", _compress_structure, pdf_bytes)

    if quality != "lossless":
        profile = QUALITY_PROFILES[quality]
        dpi, jq = profile["dpi"], profile["jpeg_quality"]
        _try("gambar", _compress_images, pdf_bytes, dpi, jq)
        _try("raster", _compress_by_rasterizing, pdf_bytes, dpi, jq, keep_text)

    # Validasi kandidat: harus bisa dibuka & jumlah halamannya sama
    try:
        with _open(pdf_bytes) as d:
            expected_pages = d.page_count
    except ValueError:
        raise
    except Exception:
        expected_pages = None

    valid = []
    for name, data in candidates:
        try:
            with pymupdf.open(stream=data, filetype="pdf") as d:
                if expected_pages is None or d.page_count == expected_pages:
                    valid.append((name, data))
        except Exception:
            logger.warning("Kandidat %s menghasilkan PDF tidak valid", name)

    best_name, best = "asli", pdf_bytes
    for name, data in valid:
        if len(data) < len(best):
            best_name, best = name, data

    info = {
        "original_size": original_size,
        "compressed_size": len(best),
        "method": best_name,
        "quality": quality,
        "reduced": len(best) < original_size,
    }
    return best, info


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Pemakaian: python pdf_compress.py input.pdf output.pdf [quality] [keep_text=1]")
        sys.exit(1)
    with open(sys.argv[1], "rb") as f:
        data = f.read()
    q = sys.argv[3] if len(sys.argv) > 3 else "medium"
    kt = (sys.argv[4] != "0") if len(sys.argv) > 4 else True
    result, meta = compress_pdf(data, q, kt)
    with open(sys.argv[2], "wb") as f:
        f.write(result)
    print(meta)
