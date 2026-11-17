/*
 * Docenti - Worker kompres PDF di browser.
 *
 * Menjalankan engine kompres yang SAMA dengan server (pdf_compress.py, port dari
 * SafePDF) memakai PyMuPDF versi WebAssembly lewat Pyodide. Berkas PDF tidak
 * diunggah ke mana pun, sehingga tidak terkena batas upload Vercel (4,5 MB).
 *
 * Pesan masuk : { id, bytes: ArrayBuffer, quality: string, keepText: boolean }
 * Pesan keluar: { id, type: 'status', stage: 'loading' | 'compressing' }
 *               { id, type: 'done', bytes: ArrayBuffer, info: {...} }
 *               { id, type: 'error', message: string }
 */
const PYODIDE_URL = 'https://cdn.jsdelivr.net/pyodide/v0.28.3/full/';

let enginePromise = null;

function loadEngine() {
  if (!enginePromise) {
    enginePromise = (async () => {
      importScripts(PYODIDE_URL + 'pyodide.js');
      const py = await loadPyodide({ indexURL: PYODIDE_URL });
      await py.loadPackage('pymupdf');

      const res = await fetch('pdf_compress.py', { cache: 'no-cache' });
      if (!res.ok) throw new Error('pdf_compress.py tidak ditemukan (HTTP ' + res.status + ')');
      py.FS.mkdirTree('/docenti');
      py.FS.writeFile('/docenti/pdf_compress.py', await res.text());
      py.runPython("import sys; sys.path.insert(0, '/docenti'); import pdf_compress");
      return py;
    })();
    // Jika gagal (misal koneksi putus), izinkan percobaan ulang berikutnya
    enginePromise.catch(() => { enginePromise = null; });
  }
  return enginePromise;
}

self.onmessage = async (event) => {
  const { id, bytes, quality, keepText } = event.data;
  try {
    if (!enginePromise) self.postMessage({ id, type: 'status', stage: 'loading' });
    const py = await loadEngine();

    self.postMessage({ id, type: 'status', stage: 'compressing' });
    py.FS.writeFile('/tmp/in.pdf', new Uint8Array(bytes));

    const infoJson = py.runPython(`
import json, os, pdf_compress
with open('/tmp/in.pdf', 'rb') as _f:
    _data = _f.read()
os.remove('/tmp/in.pdf')
_out, _info = pdf_compress.compress_pdf(_data, ${JSON.stringify(String(quality))}, ${keepText ? 'True' : 'False'})
with open('/tmp/out.pdf', 'wb') as _f:
    _f.write(_out)
del _data, _out
json.dumps(_info)
`);

    const out = py.FS.readFile('/tmp/out.pdf');
    py.FS.unlink('/tmp/out.pdf');
    const buffer = out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
    self.postMessage({ id, type: 'done', bytes: buffer, info: JSON.parse(infoJson) }, [buffer]);
  } catch (err) {
    const raw = String((err && err.message) || err);
    let message = (raw.split('\n').filter(Boolean).pop() || raw).replace(/^\w+Error: /, '');
    if (/MemoryError|out of memory|Out of bounds memory/i.test(raw)) {
      message = 'Memori browser tidak cukup untuk berkas sebesar ini. Tutup tab lain lalu coba lagi, atau gunakan jalankan_docenti.bat.';
    } else if (/importScripts|NetworkError|Failed to fetch|load/i.test(raw) && !enginePromise) {
      message = 'Engine kompres gagal diunduh. Periksa koneksi internet lalu coba lagi.';
    }
    self.postMessage({ id, type: 'error', message });
  }
};
