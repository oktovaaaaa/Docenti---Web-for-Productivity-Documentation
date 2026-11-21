/* ============================================================================
 * Docenti - Konversi ke PDF: Word, PowerPoint, Excel, HTML
 * ----------------------------------------------------------------------------
 * Semua proses berjalan di browser (tidak ada berkas yang diunggah):
 *   1. Dokumen dirender menjadi HTML di dalam iframe tersembunyi
 *      - Word (.docx)        -> docx-preview
 *      - PowerPoint (.pptx)  -> pptx-preview
 *      - Excel (.xlsx/.csv)  -> SheetJS (tabel HTML)
 *      - HTML                -> iframe ber-sandbox (skrip tidak dijalankan)
 *   2. Setiap halaman "difoto" dengan html2canvas (dipotong di batas baris/paragraf)
 *   3. Gambar halaman disusun menjadi PDF dengan pdf-lib
 *
 * File ini memakai helper global dari index.html: currentMode, showStatus,
 * triggerDownload, radioValue, stripExt, formatSize, escapeHtml, svgIcon,
 * enableSortable, applyOrder, canvasToBlobAsync, showHeroStage,
 * renderActiveToolWorkspace, PDFLib, JSZip.
 * ========================================================================== */

const O2P_TOOLS = {
  word_to_pdf: {
    kind: 'word', app: 'Microsoft Word', exts: ['docx'], legacy: ['doc', 'dot', 'rtf', 'odt'],
    listTitle: 'Dokumen Word Terpilih', formatLabel: 'dokumen Word (.docx)',
    libs: ['html2canvas', 'docx'],
    note: 'Dokumen dirender di browser, jadi tata letaknya bisa sedikit berbeda dari Microsoft Word, terutama jika font yang dipakai tidak terpasang di perangkat ini. Hasil PDF berupa gambar per halaman, sehingga teksnya tidak bisa diblok atau dicari.'
  },
  ppt_to_pdf: {
    kind: 'ppt', app: 'Microsoft PowerPoint', exts: ['pptx'], legacy: ['ppt', 'pps', 'pot', 'odp'],
    listTitle: 'Presentasi Terpilih', formatLabel: 'presentasi PowerPoint (.pptx)',
    libs: ['html2canvas', 'pptx'],
    note: 'Setiap slide menjadi satu halaman PDF. Animasi, transisi, dan video tidak ikut, dan beberapa efek (bayangan, 3D, SmartArt rumit) bisa tampil lebih sederhana. Hasil berupa gambar per slide, sehingga teksnya tidak bisa diblok atau dicari.'
  },
  excel_to_pdf: {
    kind: 'excel', app: 'Microsoft Excel', exts: ['xlsx', 'xlsm', 'xls', 'xlsb', 'ods', 'csv'], legacy: [],
    listTitle: 'Lembar Kerja Terpilih', formatLabel: 'lembar kerja Excel (.xlsx, .xls, .csv, .ods)',
    libs: ['html2canvas', 'xlsx'],
    note: 'Setiap lembar kerja dimulai di halaman baru dan diperkecil agar muat selebar kertas. Nilai dan sel gabungan ikut, tetapi warna sel, grafik, dan format bersyarat tidak ikut. Hasil berupa gambar per halaman.'
  },
  html_to_pdf: {
    kind: 'html', app: 'editor teks', exts: ['html', 'htm'], legacy: [],
    listTitle: 'Halaman HTML Terpilih', formatLabel: 'berkas HTML (.html, .htm)',
    libs: ['html2canvas'],
    note: 'Demi keamanan, skrip (JavaScript) di dalam HTML tidak dijalankan. Gambar atau CSS dengan alamat relatif tidak bisa dimuat, jadi gunakan URL lengkap (https://...). Hasil berupa gambar per halaman.'
  }
};

const O2P_BADGE_COLORS = {
  word: ['#eff6ff', '#2563eb'], ppt: ['#fff1ec', '#ea580c'], excel: ['#ecfdf5', '#059669'], html: ['#fffbeb', '#d97706']
};
const O2P_PAGE_SIZES = { a4: [595.28, 841.89], letter: [612, 792], legal: [612, 1008] };
const O2P_MARGINS = { none: 0, small: 20, large: 42 };
const O2P_DPI = { standard: 150, high: 200, max: 300 };
const O2P_PT_PER_PX = 0.75;          // 1 px CSS = 0.75 pt
const O2P_MAX_CHUNK_PX = 12000;      // tinggi kanvas maksimum sekali render
const O2P_MAX_CHUNK_AREA = 60e6;     // luas kanvas maksimum sekali render
const O2P_EXCEL_MAX_ROWS = 20000;
const O2P_EXCEL_MAX_COLS = 300;
const O2P_BLANK_DOC = '<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0;padding:0;background:#fff;}</style></head><body></body></html>';

const O2P_LIBS = {
  html2canvas: { global: 'html2canvas', local: 'html2canvas.min.js', cdn: 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js' },
  docx: { global: 'docx', local: 'docx-preview.min.js', cdn: 'https://cdn.jsdelivr.net/npm/docx-preview@0.4.1/dist/docx-preview.min.js' },
  pptx: { global: 'pptxPreview', local: 'pptx-preview.min.js', cdn: 'https://cdn.jsdelivr.net/npm/pptx-preview@1.0.7/dist/pptx-preview.umd.js' },
  xlsx: { global: 'XLSX', local: 'xlsx.full.min.js', cdn: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js' }
};

let o2pFiles = { word_to_pdf: [], ppt_to_pdf: [], excel_to_pdf: [], html_to_pdf: [] };
let o2pOutNames = {};
let o2pSeq = 0;
let o2pPasteOpen = false;
let o2pBusy = false;
const o2pLibPromises = {};

// ---------------------------------------------------------------------------
// Utilitas umum
// ---------------------------------------------------------------------------
function o2pIsTool(mode) { return Object.prototype.hasOwnProperty.call(O2P_TOOLS, mode); }
function o2pExt(name) { const m = /\.([^.]+)$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
function o2pCurrentList() { return o2pFiles[currentMode] || []; }
function o2pDelay(ms) { return new Promise(r => setTimeout(r, ms)); }
function o2pNextFrame() { return new Promise(r => requestAnimationFrame(() => r())); }
function o2pHasFiles(mode) {
  return !!(o2pFiles[mode] && (o2pFiles[mode].length > 0 || (mode === 'html_to_pdf' && o2pPasteOpen)));
}
function o2pClearAll() {
  Object.keys(o2pFiles).forEach(k => { o2pFiles[k] = []; });
  o2pPasteOpen = false;
  const ta = document.getElementById('o2pHtmlCode');
  if (ta) ta.value = '';
}

function o2pErrMsg(err) {
  const raw = String((err && (err.message || err)) || '');
  if (/password|encrypt/i.test(raw)) return 'Berkas diproteksi kata sandi. Buka proteksinya terlebih dahulu.';
  if (/central directory|corrupted zip|zip file|not a zip|unsupported file/i.test(raw)) return 'Berkas rusak atau formatnya tidak dikenali.';
  return raw || 'Terjadi kesalahan yang tidak diketahui';
}

// Memuat pustaka hanya saat dibutuhkan: berkas lokal dulu, lalu CDN sebagai cadangan
function o2pInjectScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => { s.remove(); reject(new Error('Gagal memuat ' + src)); };
    document.head.appendChild(s);
  });
}
function o2pLoadLib(key) {
  const lib = O2P_LIBS[key];
  if (window[lib.global]) return Promise.resolve(window[lib.global]);
  if (!o2pLibPromises[key]) {
    o2pLibPromises[key] = (async () => {
      try { await o2pInjectScript(lib.local); } catch (_) {}
      if (!window[lib.global]) await o2pInjectScript(lib.cdn);
      if (!window[lib.global]) throw new Error('Pustaka ' + key + ' tidak tersedia');
      return window[lib.global];
    })().catch(err => {
      delete o2pLibPromises[key];
      console.error(err);
      throw new Error('Komponen konversi gagal dimuat. Periksa koneksi internet lalu coba lagi.');
    });
  }
  return o2pLibPromises[key];
}

// ---------------------------------------------------------------------------
// Daftar berkas & tampilan
// ---------------------------------------------------------------------------
function o2pAddFiles(files) {
  const cfg = O2P_TOOLS[currentMode];
  const list = o2pFiles[currentMode];
  const accepted = [], legacy = [], other = [];
  files.forEach(f => {
    const ext = o2pExt(f.name);
    if (cfg.exts.includes(ext)) accepted.push(f);
    else if (cfg.legacy.includes(ext)) legacy.push(f);
    else other.push(f);
  });
  const legacyMsg = legacy.length
    ? 'Format .' + o2pExt(legacy[0].name) + ' belum didukung. Buka berkas di ' + cfg.app + ' lalu simpan ulang sebagai .' + cfg.exts[0] + '.'
    : '';
  if (!accepted.length) {
    showStatus(legacyMsg || ('Silakan pilih ' + cfg.formatLabel), 'error');
    return false;
  }
  accepted.forEach(file => list.push({ id: ++o2pSeq, file, status: 'ready', note: '' }));
  if (list.length === accepted.length) o2pOutNames[currentMode] = (stripExt(accepted[0].name) || 'dokumen') + '.pdf';
  if (legacy.length) showStatus(legacy.length + ' berkas dilewati. ' + legacyMsg, 'error');
  else if (other.length) showStatus(other.length + ' berkas dilewati karena bukan ' + cfg.formatLabel, 'error');
  return true;
}

function o2pOpenPaste() {
  if (currentMode !== 'html_to_pdf') return;
  o2pPasteOpen = true;
  renderActiveToolWorkspace();
  setTimeout(() => { const ta = document.getElementById('o2pHtmlCode'); if (ta) ta.focus(); }, 60);
}

function o2pAddPastedHtml(silent) {
  const ta = document.getElementById('o2pHtmlCode');
  const code = ta ? ta.value.trim() : '';
  if (!code) { if (!silent) showStatus('Tempel kode HTML terlebih dahulu', 'error'); return false; }
  const list = o2pFiles.html_to_pdf;
  const n = list.filter(it => it.pasted).length + 1;
  const file = new File([code], 'kode_html' + (n > 1 ? '_' + n : '') + '.html', { type: 'text/html' });
  list.push({ id: ++o2pSeq, file, pasted: true, status: 'ready', note: '' });
  if (list.length === 1) o2pOutNames.html_to_pdf = 'halaman_html.pdf';
  ta.value = '';
  o2pRenderSection();
  if (!silent) showStatus('Kode HTML ditambahkan ke daftar', 'success');
  return true;
}

function o2pMove(i, d) {
  if (o2pBusy) return;
  const l = o2pCurrentList(), j = i + d;
  if (j < 0 || j >= l.length) return;
  [l[i], l[j]] = [l[j], l[i]];
  o2pRenderSection();
}
function o2pRemove(i) {
  if (o2pBusy) return;
  o2pCurrentList().splice(i, 1);
  o2pRenderSection();
}

function o2pStatusHtml(it) {
  if (it.status === 'working') return ' &middot; <span class="text-blue-600 font-semibold">Memproses...</span>';
  if (it.status === 'done') return ' &middot; <span class="text-emerald-600 font-semibold">Selesai' + (it.note ? ' &middot; ' + escapeHtml(it.note) : '') + '</span>';
  if (it.status === 'error') return ' &middot; <span class="text-red-600 font-semibold">Gagal: ' + escapeHtml(it.note || '') + '</span>';
  return '';
}
function o2pSetStatus(it, status, note) {
  it.status = status;
  it.note = note || '';
  const el = document.getElementById('o2pMeta_' + it.id);
  if (el) el.innerHTML = formatSize(it.file.size) + o2pStatusHtml(it);
}

function o2pRenderSection() {
  const cfg = O2P_TOOLS[currentMode];
  if (!cfg) return;
  const list = o2pCurrentList();
  const isHtml = cfg.kind === 'html';
  if (!list.length && !(isHtml && o2pPasteOpen)) { showHeroStage(); return; }

  document.getElementById('o2pListTitle').textContent = cfg.listTitle;
  document.getElementById('o2pCountBadge').textContent = list.length + ' berkas';
  document.getElementById('o2pTotalSize').textContent = 'Total: ' + formatSize(list.reduce((s, it) => s + it.file.size, 0));
  document.getElementById('o2pNoteText').textContent = cfg.note;
  document.getElementById('o2pHintText').textContent = list.length > 1
    ? 'Seret ikon titik-titik untuk mengubah urutan. Urutan ini menentukan urutan halaman pada PDF gabungan.'
    : (list.length ? 'Atur opsi di bawah, lalu klik tombol konversi.' : 'Tempel kode HTML di bawah atau tambahkan berkas .html.');
  document.querySelectorAll('[data-o2p-kinds]').forEach(el => {
    el.classList.toggle('hidden', !el.getAttribute('data-o2p-kinds').split(' ').includes(cfg.kind));
  });
  // Opsi "Otomatis" hanya untuk Excel
  if (isHtml) {
    const auto = document.querySelector('input[name="o2pOrientation"][value="auto"]');
    if (auto && auto.checked) document.querySelector('input[name="o2pOrientation"][value="portrait"]').checked = true;
  }
  document.getElementById('o2pOutputModeGroup').classList.toggle('hidden', list.length < 2);
  const outInput = document.getElementById('o2pOutputFileName');
  outInput.value = o2pOutNames[currentMode] || (list[0] ? stripExt(list[0].file.name) + '.pdf' : 'hasil.pdf');
  outInput.dataset.mode = currentMode;
  document.getElementById('o2pConvertBtn').disabled = o2pBusy;

  const ul = document.getElementById('o2pList');
  ul.classList.toggle('hidden', !list.length);
  const [bg, fg] = O2P_BADGE_COLORS[cfg.kind];
  const dis = o2pBusy ? ' disabled' : '';
  ul.innerHTML = list.map((it, i) => {
    const name = escapeHtml(it.file.name);
    const handle = list.length > 1
      ? '<span class="drag-handle" title="Tarik untuk memindahkan"><svg class="w-4 h-4" viewBox="0 0 24 24" fill="currentColor"><circle cx="9" cy="6" r="1.6"/><circle cx="15" cy="6" r="1.6"/><circle cx="9" cy="12" r="1.6"/><circle cx="15" cy="12" r="1.6"/><circle cx="9" cy="18" r="1.6"/><circle cx="15" cy="18" r="1.6"/></svg></span>'
      : '';
    return '<li class="file-row sortable-item" draggable="' + (list.length > 1 && !o2pBusy) + '" data-sort-index="' + i + '">' +
      '<div class="flex items-center gap-2.5 min-w-0">' + handle +
        '<span class="file-badge" style="background:' + bg + ';color:' + fg + '">' + escapeHtml(o2pExt(it.file.name).toUpperCase().slice(0, 4) || 'FILE') + '</span>' +
        '<div class="min-w-0"><p class="font-semibold text-sm text-slate-800 truncate" title="' + name + '">' + name + '</p>' +
        '<p class="text-xs text-slate-400" id="o2pMeta_' + it.id + '">' + formatSize(it.file.size) + o2pStatusHtml(it) + '</p></div>' +
      '</div>' +
      '<div class="flex items-center gap-1 shrink-0">' +
        '<span class="order-chip mr-1" title="Urutan">' + (i + 1) + '</span>' +
        '<button type="button" onclick="o2pMove(' + i + ', -1)" draggable="false" class="icon-btn" title="Naikkan"' + (i === 0 ? ' disabled' : dis) + '><svg class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg></button>' +
        '<button type="button" onclick="o2pMove(' + i + ', 1)" draggable="false" class="icon-btn" title="Turunkan"' + (i === list.length - 1 ? ' disabled' : dis) + '><svg class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg></button>' +
        '<button type="button" onclick="o2pRemove(' + i + ')" draggable="false" class="icon-btn danger" title="Hapus berkas"' + dis + '>' + svgIcon('x', 'w-4 h-4', 2.2) + '</button>' +
      '</div></li>';
  }).join('');

  enableSortable(ul, {
    item: 'li',
    onReorder: (order) => { if (!o2pBusy) applyOrder(o2pCurrentList(), order); o2pRenderSection(); },
    onCancel: () => o2pRenderSection()
  });
}

function o2pReadOptions(kind) {
  let orientation = radioValue('o2pOrientation', kind === 'excel' ? 'auto' : 'portrait');
  if (kind !== 'excel' && orientation === 'auto') orientation = 'portrait';
  return {
    dpi: O2P_DPI[radioValue('o2pQuality', 'high')] || 200,
    size: radioValue('o2pPageSize', 'a4'),
    orientation,
    margin: radioValue('o2pMargin', 'small'),
    viewport: radioValue('o2pViewport', 'paper'),
    sheets: radioValue('o2pSheets', 'all'),
    grid: radioValue('o2pGrid', 'show') === 'show'
  };
}

// ---------------------------------------------------------------------------
// Proses konversi utama
// ---------------------------------------------------------------------------
async function o2pConvert() {
  if (o2pBusy) return;
  const mode = currentMode, cfg = O2P_TOOLS[mode];
  if (!cfg) return;
  if (cfg.kind === 'html') o2pAddPastedHtml(true); // kode yang masih di kotak tempel ikut dikonversi
  const list = o2pFiles[mode].slice();
  if (!list.length) { showStatus('Tambahkan minimal 1 berkas terlebih dahulu', 'error'); return; }

  const opts = o2pReadOptions(cfg.kind);
  const separate = radioValue('o2pOutputMode', 'merge') === 'separate' && list.length > 1;
  const outName = (document.getElementById('o2pOutputFileName').value || 'hasil.pdf').trim();
  o2pOutNames[mode] = outName;
  const btnText = document.getElementById('o2pConvertBtnText');

  o2pBusy = true;
  list.forEach(it => { it.status = 'ready'; it.note = ''; });
  o2pRenderSection();
  btnText.textContent = 'Memproses...';
  try {
    showStatus('Menyiapkan komponen konversi...');
    for (const key of cfg.libs) await o2pLoadLib(key);

    const merged = separate ? null : await PDFLib.PDFDocument.create();
    const zip = separate ? new JSZip() : null;
    const used = new Set();
    let ok = 0, pages = 0;
    const errors = [];

    for (let i = 0; i < list.length; i++) {
      const item = list[i];
      o2pSetStatus(item, 'working');
      const prefix = list.length > 1 ? '(' + (i + 1) + '/' + list.length + ') ' : '';
      const progress = (msg) => showStatus(prefix + item.file.name + ': ' + msg);
      try {
        const doc = await PDFLib.PDFDocument.create();
        await o2pConvertFile(item.file, cfg.kind, doc, opts, progress);
        const count = doc.getPageCount();
        if (!count) throw new Error('Dokumen tidak memiliki halaman yang bisa dikonversi');
        if (separate) {
          const base = stripExt(item.file.name) || ('dokumen_' + (i + 1));
          let name = base + '.pdf', n = 2;
          while (used.has(name)) name = base + '_' + (n++) + '.pdf';
          used.add(name);
          zip.file(name, await doc.save());
        } else {
          const copied = await merged.copyPages(doc, doc.getPageIndices());
          copied.forEach(p => merged.addPage(p));
        }
        pages += count;
        ok++;
        o2pSetStatus(item, 'done', count + ' halaman');
      } catch (err) {
        console.error(err);
        const msg = o2pErrMsg(err);
        errors.push(item.file.name + ': ' + msg);
        o2pSetStatus(item, 'error', msg);
      }
    }

    if (!ok) throw new Error(list.length === 1 ? errors[0].replace(/^[^:]*: /, '') : 'Semua berkas gagal dikonversi');

    if (separate) {
      showStatus('Menyiapkan berkas ZIP...');
      triggerDownload(await zip.generateAsync({ type: 'blob' }), stripExt(outName) + '.zip', '.zip');
    } else {
      showStatus('Menyimpan dokumen PDF...');
      const bytes = await merged.save();
      triggerDownload(new Blob([bytes], { type: 'application/pdf' }), outName, '.pdf');
    }
    if (errors.length) showStatus(ok + ' berkas berhasil (' + pages + ' halaman), ' + errors.length + ' gagal. ' + errors[0], 'error');
    else showStatus(ok + ' berkas berhasil diubah menjadi PDF (' + pages + ' halaman)!', 'success');
  } catch (err) {
    console.error(err);
    showStatus('Gagal mengonversi: ' + o2pErrMsg(err), 'error');
  } finally {
    o2pBusy = false;
    btnText.textContent = 'Konversi ke PDF & Unduh';
    if (currentMode === mode) o2pRenderSection();
  }
}

function o2pConvertFile(file, kind, pdfDoc, opts, progress) {
  if (kind === 'word') return o2pWordToPdf(file, pdfDoc, opts, progress);
  if (kind === 'ppt') return o2pPptToPdf(file, pdfDoc, opts, progress);
  if (kind === 'excel') return o2pExcelToPdf(file, pdfDoc, opts, progress);
  return o2pHtmlToPdf(file, pdfDoc, opts, progress);
}

// ---------------------------------------------------------------------------
// Mesin render: iframe tersembunyi -> html2canvas -> halaman PDF
// ---------------------------------------------------------------------------
function o2pCreateFrame(width, opts = {}) {
  return new Promise((resolve) => {
    const iframe = document.createElement('iframe');
    iframe.setAttribute('aria-hidden', 'true');
    iframe.tabIndex = -1;
    iframe.style.cssText = 'position:fixed;left:-100000px;top:0;width:' + Math.ceil(width) + 'px;height:900px;border:0;background:#fff;pointer-events:none;';
    if (opts.sandbox != null) iframe.setAttribute('sandbox', opts.sandbox);
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(iframe); } };
    // Halaman dengan sumber eksternal yang lambat tetap diproses setelah batas waktu
    const timer = setTimeout(finish, opts.timeout || 20000);
    iframe.onload = () => { clearTimeout(timer); finish(); };
    iframe.srcdoc = opts.html != null ? opts.html : O2P_BLANK_DOC;
    document.body.appendChild(iframe);
  });
}

// Tunggu gambar & font selesai dimuat sebelum difoto
async function o2pSettle(frame, extraMs) {
  const doc = frame.contentDocument;
  if (!doc) return;
  const imgs = Array.from(doc.images || []);
  imgs.forEach(img => { if (img.loading === 'lazy') img.loading = 'eager'; });
  await Promise.race([
    Promise.all(imgs.map(img => img.complete ? null : new Promise(r => {
      img.addEventListener('load', r, { once: true });
      img.addEventListener('error', r, { once: true });
    }))),
    o2pDelay(10000)
  ]);
  try { if (doc.fonts && doc.fonts.ready) await Promise.race([doc.fonts.ready, o2pDelay(5000)]); } catch (_) {}
  await o2pDelay(extraMs || 50);
  await o2pNextFrame();
}

// Kandidat titik potong halaman: batas atas/bawah elemen blok (paragraf, baris tabel, gambar, ...)
const O2P_BREAK_SEL = 'p,li,tr,h1,h2,h3,h4,h5,h6,img,svg,canvas,table,pre,blockquote,figure,hr,section,article,header,footer,ul,ol,dl,dt,dd,div';
function o2pCollectBreaks(root, originTop, withForced) {
  let els = root.querySelectorAll(O2P_BREAK_SEL);
  if (els.length > 30000) els = root.querySelectorAll('p,li,tr,h1,h2,h3,h4,h5,h6,table,pre,blockquote,figure,hr');
  const win = root.ownerDocument.defaultView;
  const cand = [], forced = [];
  for (const el of els) {
    const r = el.getBoundingClientRect();
    if (!r.height) continue;
    cand.push(r.top - originTop, r.bottom - originTop);
    if (withForced) {
      const cs = win.getComputedStyle(el);
      if (cs.breakBefore === 'page' || cs.pageBreakBefore === 'always') forced.push(r.top - originTop);
      if (cs.breakAfter === 'page' || cs.pageBreakAfter === 'always') forced.push(r.bottom - originTop);
    }
  }
  return {
    candidates: Float64Array.from(cand).sort(),
    forced: Array.from(new Set(forced.map(v => Math.round(v * 10) / 10))).sort((a, b) => a - b)
  };
}

function o2pLastAtMost(sorted, v) {
  let lo = 0, hi = sorted.length - 1, ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (sorted[mid] <= v) { ans = mid; lo = mid + 1; } else hi = mid - 1;
  }
  return ans < 0 ? null : sorted[ans];
}

// Bagi tinggi total menjadi potongan halaman, sebisa mungkin di batas elemen
function o2pPlanSlices(total, firstMax, nextMax, brk) {
  const slices = [];
  const cand = brk && brk.candidates && brk.candidates.length ? brk.candidates : null;
  const forced = brk && brk.forced ? brk.forced : [];
  let pos = 0, first = true, guard = 0;
  while (pos < total - 1 && guard++ < 10000) {
    const max = Math.max(20, first ? firstMax : nextMax);
    let end = pos + max;
    const f = forced.find(v => v > pos + 2 && v < end);
    if (f !== undefined) {
      end = f;
    } else if (end >= total) {
      end = total;
    } else if (cand) {
      const c = o2pLastAtMost(cand, end);
      if (c !== null && c > pos + max * 0.55) end = c;
    }
    slices.push([pos, end]);
    pos = end;
    first = false;
  }
  if (!slices.length) slices.push([0, Math.max(1, total)]);
  return slices;
}

function o2pScale(dpi, ptPerCss, cssWidth) {
  let s = dpi / 72 * ptPerCss;
  if (cssWidth) s = Math.min(s, 16000 / cssWidth);
  return Math.max(0.5, Math.min(4, s));
}

function o2pH2c(el, opts) {
  const win = el.ownerDocument.defaultView;
  return window.html2canvas(el, Object.assign({
    backgroundColor: '#ffffff',
    useCORS: true,
    allowTaint: false,
    logging: false,
    imageTimeout: 15000,
    windowWidth: win.innerWidth,
    windowHeight: win.innerHeight,
    scrollX: 0,
    scrollY: 0
  }, opts));
}

function o2pCrop(src, sy, sh) {
  const y = Math.max(0, Math.min(src.height - 1, Math.round(sy)));
  const h = Math.max(1, Math.min(src.height - y, Math.round(sh)));
  const c = document.createElement('canvas');
  c.width = src.width;
  c.height = h;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(src, 0, y, src.width, h, 0, 0, src.width, h);
  return c;
}

// Render potongan [a,b] dari sebuah elemen; beberapa halaman sekaligus per kanvas agar cepat
async function o2pRenderSlices(el, slices, scale, cssWidth, onPage) {
  const maxPx = Math.max(256, Math.min(O2P_MAX_CHUNK_PX, Math.floor(O2P_MAX_CHUNK_AREA / Math.max(1, cssWidth * scale))));
  const maxCss = maxPx / scale;
  let i = 0;
  while (i < slices.length) {
    let j = i;
    const start = slices[i][0];
    while (j + 1 < slices.length && slices[j + 1][1] - start <= maxCss) j++;
    const end = slices[j][1];
    const chunk = await o2pH2c(el, { scale, x: 0, y: start, width: Math.ceil(cssWidth), height: Math.max(1, Math.ceil(end - start)) });
    for (let k = i; k <= j; k++) {
      const [a, b] = slices[k];
      const page = o2pCrop(chunk, (a - start) * scale, (b - a) * scale);
      await onPage(page, k, a, b);
      page.width = page.height = 0;
    }
    chunk.width = chunk.height = 0;
    i = j + 1;
  }
}

async function o2pAddCanvasPage(pdfDoc, canvas, pageW, pageH, x, yTop, drawW, drawH) {
  const blob = await canvasToBlobAsync(canvas, 'image/jpeg', 0.9);
  const img = await pdfDoc.embedJpg(new Uint8Array(await blob.arrayBuffer()));
  const page = pdfDoc.addPage([pageW, pageH]);
  page.drawImage(img, { x, y: pageH - yTop - drawH, width: drawW, height: drawH });
}

function o2pPaper(opts, landscape) {
  let [pw, ph] = O2P_PAGE_SIZES[opts.size] || O2P_PAGE_SIZES.a4;
  if (landscape) [pw, ph] = [ph, pw];
  return [pw, ph];
}

// ---------------------------------------------------------------------------
// WORD (.docx)
// ---------------------------------------------------------------------------
async function o2pWordToPdf(file, pdfDoc, opts, progress) {
  progress('membaca dokumen...');
  const buf = await file.arrayBuffer();
  const frame = await o2pCreateFrame(2400);
  try {
    const doc = frame.contentDocument;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    await window.docx.renderAsync(buf, host, doc.head, {
      className: 'docx', inWrapper: false, ignoreWidth: false, ignoreHeight: false, ignoreFonts: false,
      breakPages: true, ignoreLastRenderedPageBreak: false, experimental: true, useBase64URL: true,
      renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true,
      renderComments: false, renderChanges: false
    });
    const sections = Array.from(host.querySelectorAll('section.docx'));
    if (!sections.length) throw new Error('Isi dokumen tidak dapat dibaca');
    frame.style.height = Math.ceil(host.getBoundingClientRect().height + 10) + 'px';
    await o2pSettle(frame, 120);

    const win = frame.contentWindow;
    for (let si = 0; si < sections.length; si++) {
      progress('merender halaman ' + (si + 1) + ' dari ' + sections.length + '...');
      const sec = sections[si];
      const rect = sec.getBoundingClientRect();
      const cs = win.getComputedStyle(sec);
      const pageW = rect.width;
      let pageH = parseFloat(cs.minHeight) || 0;
      if (pageH < 50) pageH = Math.max(rect.height, pageW * 1.4142);
      const padT = parseFloat(cs.paddingTop) || 0;
      const padB = parseFloat(cs.paddingBottom) || 0;
      const total = rect.height;
      // Isi lebih panjang dari satu halaman: lanjutkan ke halaman berikutnya dengan margin atas yang sama
      const slices = total <= pageH + 1
        ? [[0, total]]
        : o2pPlanSlices(total, pageH - padB, Math.max(pageH * 0.5, pageH - padT - padB), o2pCollectBreaks(sec, rect.top, false));
      const pw = pageW * O2P_PT_PER_PX, ph = pageH * O2P_PT_PER_PX;
      const scale = o2pScale(opts.dpi, O2P_PT_PER_PX, pageW);
      await o2pRenderSlices(sec, slices, scale, pageW, (canvas, k, a, b) => {
        const yTop = k === 0 ? 0 : padT * O2P_PT_PER_PX;
        return o2pAddCanvasPage(pdfDoc, canvas, pw, ph, 0, yTop, pw, (b - a) * O2P_PT_PER_PX);
      });
    }
  } finally {
    frame.remove();
  }
}

// ---------------------------------------------------------------------------
// POWERPOINT (.pptx)
// ---------------------------------------------------------------------------
async function o2pPptToPdf(file, pdfDoc, opts, progress) {
  progress('membaca presentasi...');
  const buf = await file.arrayBuffer();
  const W = 960;
  const frame = await o2pCreateFrame(W + 40);
  try {
    const doc = frame.contentDocument;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    const previewer = window.pptxPreview.init(host, { width: W, mode: 'list' });
    await previewer.preview(buf);
    const slides = Array.from(host.querySelectorAll('.pptx-preview-slide-wrapper'));
    if (!slides.length) throw new Error('Presentasi tidak memiliki slide yang bisa dibaca');
    frame.style.height = Math.ceil(host.getBoundingClientRect().height + 10) + 'px';
    // Grafik (echarts) punya animasi awal: beri jeda agar tergambar penuh
    await o2pSettle(frame, host.querySelector('.chart-node, canvas') ? 1300 : 150);

    for (let i = 0; i < slides.length; i++) {
      progress('merender slide ' + (i + 1) + ' dari ' + slides.length + '...');
      const rect = slides[i].getBoundingClientRect();
      const pw = 720, ph = 720 * rect.height / rect.width;
      const ptPerCss = pw / rect.width;
      const scale = o2pScale(opts.dpi, ptPerCss, rect.width);
      await o2pRenderSlices(slides[i], [[0, rect.height]], scale, rect.width,
        (canvas) => o2pAddCanvasPage(pdfDoc, canvas, pw, ph, 0, 0, pw, ph));
    }
    try { previewer.destroy(); } catch (_) {}
  } finally {
    frame.remove();
  }
}

// ---------------------------------------------------------------------------
// EXCEL (.xlsx, .xls, .csv, .ods)
// ---------------------------------------------------------------------------
const O2P_SHEET_CSS =
  'body{font-family:Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-size:14.5px;color:#111;}' +
  '.sheet{display:inline-block;padding:2px;background:#fff;}' +
  '.sheet-title{font-weight:700;font-size:15px;padding:4px 2px 8px;}' +
  'table{border-collapse:collapse;}' +
  'td,th{border:1px solid #c7ccd1;padding:2px 6px;vertical-align:bottom;white-space:nowrap;text-align:left;line-height:1.3;}' +
  'td[data-t="n"],td[data-t="d"]{text-align:right;}' +
  'td[data-t="b"],td[data-t="e"]{text-align:center;}' +
  'table.is-fixed td{white-space:pre-wrap;overflow-wrap:anywhere;overflow:hidden;}' +
  '.no-grid td,.no-grid th{border-color:transparent;}';

// Rentang sel yang benar-benar berisi data (abaikan area kosong yang hanya berformat)
function o2pUsedRange(XLSX, ws) {
  let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
  for (const key in ws) {
    if (key[0] === '!') continue;
    const cell = ws[key];
    if (!cell || (cell.v === undefined && cell.f === undefined) || cell.v === '') continue;
    const a = XLSX.utils.decode_cell(key);
    if (a.r < r0) r0 = a.r;
    if (a.c < c0) c0 = a.c;
    if (a.r > r1) r1 = a.r;
    if (a.c > c1) c1 = a.c;
  }
  if (r1 < 0) return null;
  (ws['!merges'] || []).forEach(m => {
    if (m.s.r <= r1 && m.s.c <= c1 && m.s.r >= r0 && m.s.c >= c0) {
      r1 = Math.max(r1, m.e.r);
      c1 = Math.max(c1, m.e.c);
    }
  });
  const truncated = (r1 - r0 + 1) > O2P_EXCEL_MAX_ROWS || (c1 - c0 + 1) > O2P_EXCEL_MAX_COLS;
  r1 = Math.min(r1, r0 + O2P_EXCEL_MAX_ROWS - 1);
  c1 = Math.min(c1, c0 + O2P_EXCEL_MAX_COLS - 1);
  return { s: { r: r0, c: c0 }, e: { r: r1, c: c1 }, truncated };
}

// Kunci lebar kolom (lebar asli Excel atau lebar isi, maks. 480 px) agar baris bisa dibagi per halaman
function o2pFreezeColumns(XLSX, ws, table, range) {
  const doc = table.ownerDocument;
  const n = range.e.c - range.s.c + 1;
  const natural = new Array(n).fill(0);
  table.querySelectorAll('td').forEach(td => {
    if ((td.colSpan || 1) !== 1) return;
    const m = /-([A-Z]+)\d+$/.exec(td.id || '');
    if (!m) return;
    const c = XLSX.utils.decode_col(m[1]) - range.s.c;
    if (c >= 0 && c < n) natural[c] = Math.max(natural[c], td.getBoundingClientRect().width);
  });
  const cols = ws['!cols'] || [];
  const widths = natural.map((w, i) => {
    const info = cols[range.s.c + i];
    let setW = 0;
    if (info) setW = info.wpx || (info.wch ? info.wch * 7 + 5 : 0) || (info.width ? info.width * 7 + 5 : 0);
    return Math.round(Math.min(480, Math.max(w, setW || 0, 28)));
  });
  const cg = doc.createElement('colgroup');
  widths.forEach(w => { const col = doc.createElement('col'); col.style.width = w + 'px'; cg.appendChild(col); });
  table.insertBefore(cg, table.firstChild);
  table.style.tableLayout = 'fixed';
  table.style.width = widths.reduce((s, w) => s + w, 0) + 'px';
  table.classList.add('is-fixed');
}

async function o2pExcelToPdf(file, pdfDoc, opts, progress) {
  progress('membaca lembar kerja...');
  const XLSX = window.XLSX;
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
  let names = wb.SheetNames.filter((n, i) => !(meta[i] && meta[i].Hidden));
  if (opts.sheets === 'first') names = names.slice(0, 1);

  const frame = await o2pCreateFrame(4000);
  let rendered = 0, truncated = false;
  try {
    const doc = frame.contentDocument;
    const style = doc.createElement('style');
    style.textContent = O2P_SHEET_CSS;
    doc.head.appendChild(style);
    const win = frame.contentWindow;

    for (const name of names) {
      const ws = wb.Sheets[name];
      const range = ws && o2pUsedRange(XLSX, ws);
      if (!range) continue;
      if (range.truncated) truncated = true;
      ws['!ref'] = XLSX.utils.encode_range(range);
      progress('merender lembar "' + name + '"...');

      const wrap = doc.createElement('div');
      wrap.className = 'sheet' + (opts.grid ? '' : ' no-grid');
      if (names.length > 1) {
        const title = doc.createElement('div');
        title.className = 'sheet-title';
        title.textContent = name;
        wrap.appendChild(title);
      }
      const holder = doc.createElement('div');
      holder.innerHTML = XLSX.utils.sheet_to_html(ws, { header: '', footer: '' });
      const table = holder.querySelector('table');
      if (!table) continue;
      wrap.appendChild(table);
      doc.body.appendChild(wrap);
      o2pFreezeColumns(XLSX, ws, table, range);

      const contentW = wrap.getBoundingClientRect().width;
      frame.style.width = Math.ceil(Math.max(800, contentW + 40)) + 'px';
      frame.style.height = Math.ceil(wrap.getBoundingClientRect().height + 20) + 'px';
      await o2pSettle(frame);

      const m = O2P_MARGINS[opts.margin] || 0;
      const portrait = o2pPaper(opts, false);
      const landscape = opts.orientation === 'landscape' ||
        (opts.orientation === 'auto' && contentW * O2P_PT_PER_PX > (portrait[0] - 2 * m) * 1.05);
      const [pw, ph] = o2pPaper(opts, landscape);
      const s = Math.min(O2P_PT_PER_PX, (pw - 2 * m) / contentW); // pt per px CSS
      const usableH = (ph - 2 * m) / s;
      const scale = o2pScale(opts.dpi, s, contentW);

      const wrapTop = wrap.getBoundingClientRect().top;
      const total = wrap.getBoundingClientRect().height;
      const rows = Array.from(table.rows);
      const rowTop = rows.map(r => r.getBoundingClientRect().top - wrapTop);
      const rowBot = rows.map(r => r.getBoundingClientRect().bottom - wrapTop);
      const slices = o2pPlanSlices(total, usableH, usableH, { candidates: Float64Array.from(rowBot).sort(), forced: [] });

      // Render per kelompok halaman: baris di luar kelompok disembunyikan agar render tetap ringan
      const maxPx = Math.max(256, Math.min(O2P_MAX_CHUNK_PX, Math.floor(O2P_MAX_CHUNK_AREA / Math.max(1, contentW * scale))));
      const maxCss = maxPx / scale;
      const titleEl = wrap.querySelector('.sheet-title');
      let i = 0;
      while (i < slices.length) {
        let j = i;
        const start = slices[i][0];
        while (j + 1 < slices.length && slices[j + 1][1] - start <= maxCss) j++;
        const end = slices[j][1];
        progress('merender lembar "' + name + '" (halaman ' + (i + 1) + '-' + (j + 1) + ' dari ' + slices.length + ')...');

        let firstVisible = -1;
        rows.forEach((r, idx) => {
          const show = rowBot[idx] > start + 0.5 && rowTop[idx] < end - 0.5;
          r.style.display = show ? '' : 'none';
          if (show && firstVisible < 0) firstVisible = idx;
        });
        if (titleEl) titleEl.style.display = start <= 0.5 ? '' : 'none';
        let delta = 0;
        if (start > 0.5 && firstVisible >= 0) {
          delta = (rows[firstVisible].getBoundingClientRect().top - wrap.getBoundingClientRect().top) - rowTop[firstVisible];
        }
        const chunk = await o2pH2c(wrap, { scale, windowWidth: win.innerWidth, windowHeight: win.innerHeight });
        for (let k = i; k <= j; k++) {
          const [a, b] = slices[k];
          const page = o2pCrop(chunk, (a + delta) * scale, (b - a) * scale);
          await o2pAddCanvasPage(pdfDoc, page, pw, ph, m, m, contentW * s, (b - a) * s);
          page.width = page.height = 0;
        }
        chunk.width = chunk.height = 0;
        i = j + 1;
      }
      rows.forEach(r => { r.style.display = ''; });
      wrap.remove();
      rendered++;
    }
  } finally {
    frame.remove();
  }
  if (!rendered) throw new Error('Lembar kerja kosong, tidak ada data untuk dikonversi');
  if (truncated) console.warn('Lembar kerja sangat besar: hanya ' + O2P_EXCEL_MAX_ROWS + ' baris / ' + O2P_EXCEL_MAX_COLS + ' kolom pertama yang dikonversi.');
}

// ---------------------------------------------------------------------------
// HTML (.html, .htm, atau kode yang ditempel)
// ---------------------------------------------------------------------------
async function o2pHtmlToPdf(file, pdfDoc, opts, progress) {
  progress('memuat halaman HTML...');
  const html = await file.text();
  const [pw, ph] = o2pPaper(opts, opts.orientation === 'landscape');
  const m = O2P_MARGINS[opts.margin] || 0;
  const usableW = pw - 2 * m;
  const viewport = opts.viewport === 'paper' ? Math.round(usableW / O2P_PT_PER_PX) : (parseInt(opts.viewport, 10) || 1280);
  // sandbox tanpa allow-scripts: skrip di dalam HTML tidak dijalankan
  const frame = await o2pCreateFrame(viewport, { html, sandbox: 'allow-same-origin', timeout: 20000 });
  try {
    const doc = frame.contentDocument;
    if (!doc || !doc.documentElement) throw new Error('Halaman HTML tidak dapat dibaca');
    const root = doc.documentElement;
    const measure = () => Math.max(root.scrollHeight, doc.body ? doc.body.scrollHeight : 0, 1);
    frame.style.height = measure() + 'px';
    await o2pSettle(frame, 150);
    frame.style.height = measure() + 'px';
    await o2pNextFrame();
    const total = measure();
    if (!doc.body || (!doc.body.innerText.trim() && !doc.body.querySelector('img,svg,canvas,table,video,iframe,object,embed,input'))) {
      throw new Error('Halaman HTML kosong (konten yang dibuat dengan JavaScript tidak ikut dirender)');
    }

    progress('merender halaman...');
    const s = usableW / viewport;
    const usableH = (ph - 2 * m) / s;
    const slices = o2pPlanSlices(total, usableH, usableH, o2pCollectBreaks(doc.body, -(frame.contentWindow.scrollY || 0), true));
    const scale = o2pScale(opts.dpi, s, viewport);
    await o2pRenderSlices(root, slices, scale, viewport,
      (canvas, k, a, b) => o2pAddCanvasPage(pdfDoc, canvas, pw, ph, m, m, usableW, (b - a) * s));
  } finally {
    frame.remove();
  }
}
