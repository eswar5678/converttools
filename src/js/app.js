/* ============================================================
   ConvertTools — all logic runs 100% client-side.
   Libraries used (bundled locally in /vendor):
     pdf-lib        → PDF creation & merging
     pdf.js         → PDF reading / rendering
     JSZip          → zipping multi-page image output
     mammoth        → .docx → HTML
     docx (UMD)     → building .docx files
     html2pdf.js    → HTML → PDF (html2canvas + jsPDF)
   ============================================================ */
(function () {
  'use strict';

  /* ---------- pdf.js worker ----------
     Self-contained build: worker ships inline as text inside
     <script type="text/plain" id="ct-pdf-worker"> and is spun up from a
     Blob URL. Multi-file layout: loaded from vendor/. */
  if (window.pdfjsLib) {
    const workerTag = document.getElementById('ct-pdf-worker');
    if (workerTag) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(
        new Blob([workerTag.textContent], { type: 'application/javascript' })
      );
    } else {
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
    }
  }

  /* ============================================================
     Small utilities
     ============================================================ */
  const $ = (sel, root) => (root || document).querySelector(sel);

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  /* Escape anything placed into HTML — file names can contain any characters */
  const esc = s => String(s).replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function toast(msg, isErr) {
    const root = $('#toast-root');
    const el = document.createElement('div');
    el.className = 'toast' + (isErr ? ' err' : '');
    el.textContent = msg;
    root.appendChild(el);
    setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; }, 2800);
    setTimeout(() => el.remove(), 3200);
  }

  /* Each conversion result gets one object URL that lives until the
     modal is closed/reset — so both the Download button and the
     fallback "open file" link can reuse it. */
  function revokeResultUrl() {
    if (state && state.resultUrl) {
      URL.revokeObjectURL(state.resultUrl);
      state.resultUrl = null;
    }
  }

  function sanitizeFileName(raw, fallback, ext) {
    let name = (raw || '').trim().replace(/[\\/:*?"<>|]/g, '');
    if (!name) name = fallback;
    if (ext && !name.toLowerCase().endsWith(ext.toLowerCase())) name += ext;
    return name;
  }

  function blobToUint8(dataUri) {
    const b64 = dataUri.substring(dataUri.indexOf(',') + 1);
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function loadImageDims(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
      img.onerror = () => reject(new Error('Could not read an embedded image.'));
      img.src = src;
    });
  }

  function waitForImages(root) {
    const imgs = [...root.querySelectorAll('img')];
    return Promise.all(imgs.map(img => img.complete
      ? Promise.resolve()
      : new Promise(res => {
          img.onload = res; img.onerror = res;
          setTimeout(res, 10000);
        })
    ));
  }

  function friendlyPdfLoadError(fileName, e) {
    if (/encrypt|password/i.test(e.message || '')) {
      return new Error(`"${fileName}" is password-protected. Please remove the password and try again.`);
    }
    return new Error(`"${fileName}" doesn't look like a valid PDF file.`);
  }

  /* pdf.js with eval disabled — keeps the strict Content-Security-Policy intact */
  function loadPdf(arrayBuffer) {
    return pdfjsLib.getDocument({ data: arrayBuffer, isEvalSupported: false }).promise;
  }

  /* ============================================================
     HTML → DOCX converter (shared by "Merge Word Docs")
     Parses the clean HTML that mammoth produces and rebuilds it
     as real Word elements with the `docx` library.
     ============================================================ */
  async function htmlToDocxChildren(html, numberingConfigs, counters) {
    const D = docx;
    const parsed = new DOMParser().parseFromString('<body>' + html + '</body>', 'text/html');
    const root = parsed.body;
    const out = [];

    async function makeImageRun(imgEl) {
      const src = imgEl.getAttribute('src') || '';
      if (!src.startsWith('data:')) return null;
      try {
        const mime = src.substring(5, src.indexOf(';')).toLowerCase();
        const type = mime.includes('png') ? 'png' : mime.includes('gif') ? 'gif' : mime.includes('bmp') ? 'bmp' : 'jpg';
        const bytes = blobToUint8(src);
        const dims = await loadImageDims(src);
        const maxW = 460;
        const scale = Math.min(1, maxW / dims.w);
        return new D.ImageRun({
          data: bytes,
          type: type,
          transformation: { width: Math.round(dims.w * scale), height: Math.round(dims.h * scale) }
        });
      } catch (e) { return null; }
    }

    async function inlineRuns(el, style) {
      const runs = [];
      for (const node of el.childNodes) {
        if (node.nodeType === 3) {
          if (node.textContent.length) runs.push(new D.TextRun(Object.assign({ text: node.textContent }, style)));
          continue;
        }
        if (node.nodeType !== 1) continue;
        const tag = node.tagName.toLowerCase();
        if (tag === 'ul' || tag === 'ol') continue;           // handled at block level
        if (tag === 'br') { runs.push(new D.TextRun(Object.assign({ break: 1 }, style))); continue; }
        if (tag === 'img') { const ir = await makeImageRun(node); if (ir) runs.push(ir); continue; }
        const next = Object.assign({}, style);
        if (tag === 'b' || tag === 'strong') next.bold = true;
        if (tag === 'i' || tag === 'em') next.italics = true;
        if (tag === 'u' || tag === 'ins') next.underline = {};
        if (tag === 's' || tag === 'del' || tag === 'strike') next.strike = true;
        if (tag === 'sup') next.superScript = true;
        if (tag === 'sub') next.subScript = true;
        if (tag === 'a') { next.color = '1155CC'; next.underline = {}; }
        runs.push(...await inlineRuns(node, next));
      }
      return runs;
    }

    async function walkBlock(el, listLevel) {
      const tag = el.tagName ? el.tagName.toLowerCase() : '';
      const headMatch = /^h([1-6])$/.exec(tag);
      if (headMatch) {
        const lvl = 'HEADING_' + headMatch[1];
        out.push(new D.Paragraph({
          heading: D.HeadingLevel[lvl],
          spacing: { before: 160, after: 100 },
          children: await inlineRuns(el, {})
        }));
      } else if (tag === 'p') {
        const runs = await inlineRuns(el, {});
        if (runs.length) out.push(new D.Paragraph({ spacing: { after: 160 }, children: runs }));
      } else if (tag === 'ul' || tag === 'ol') {
        let ref = null;
        if (tag === 'ol') {
          ref = 'ct-num-' + (++counters.num);
          numberingConfigs.push({
            reference: ref,
            levels: [0, 1, 2].map(i => ({
              level: i,
              format: D.LevelFormat.DECIMAL,
              text: '%' + (i + 1) + '.',
              alignment: D.AlignmentType.START,
              style: { paragraph: { indent: { left: 720 * (i + 1), hanging: 320 } } }
            }))
          });
        }
        for (const li of el.querySelectorAll(':scope > li')) {
          out.push(new D.Paragraph(Object.assign(
            tag === 'ul'
              ? { bullet: { level: listLevel } }
              : { numbering: { reference: ref, level: listLevel } },
            { spacing: { after: 80 }, children: await inlineRuns(li, {}) }
          )));
          for (const sub of li.querySelectorAll(':scope > ul, :scope > ol')) {
            await walkBlock(sub, Math.min(listLevel + 1, 2));
          }
        }
        out.push(new D.Paragraph({ spacing: { after: 80 }, children: [] }));
      } else if (tag === 'table') {
        const rows = [];
        for (const tr of el.querySelectorAll(':scope tr')) {
          const cells = [];
          for (const td of tr.querySelectorAll(':scope > th, :scope > td')) {
            const isHeader = td.tagName.toLowerCase() === 'th';
            const cellChildren = [];
            const blocks = td.children.length ? [...td.children] : [td];
            for (const b of blocks) {
              const runs = await inlineRuns(b, isHeader ? { bold: true } : {});
              cellChildren.push(new D.Paragraph({ children: runs.length ? runs : [new D.TextRun('')] }));
            }
            cells.push(new D.TableCell({ children: cellChildren }));
          }
          if (cells.length) rows.push(new D.TableRow({ children: cells }));
        }
        if (rows.length) {
          out.push(new D.Table({
            width: { size: 100, type: D.WidthType.PERCENTAGE },
            rows: rows
          }));
          out.push(new D.Paragraph({ spacing: { after: 120 }, children: [] }));
        }
      } else if (tag === 'img') {
        const ir = await makeImageRun(el);
        if (ir) out.push(new D.Paragraph({ children: [ir] }));
      } else if (tag === 'hr') {
        out.push(new D.Paragraph({
          children: [],
          border: { bottom: { color: '999999', space: 1, style: D.BorderStyle.SINGLE, size: 6 } }
        }));
      } else if (/^(div|section|article|blockquote|figure|figcaption|main|aside|header|footer)$/.test(tag)) {
        for (const child of el.children) await walkBlock(child, listLevel || 0);
      } else if (tag) {
        const runs = await inlineRuns(el, {});
        if (runs.length) out.push(new D.Paragraph({ spacing: { after: 160 }, children: runs }));
      }
    }

    for (const child of root.children) await walkBlock(child, 0);
    if (out.length === 0) {
      out.push(new D.Paragraph({ children: [new D.TextRun({ text: root.textContent.trim() || '' })] }));
    }
    return out;
  }

  /* ============================================================
     PDF text helpers (shared by PDF→DOCX and PDF→TXT)
     ============================================================ */
  function groupTextLines(items) {
    const rows = new Map();
    const ys = [];
    const TOL = 2.5;
    for (const it of items) {
      if (!it.str || !it.str.trim()) continue;
      const y = it.transform[5];
      let key = null;
      for (const yy of ys) { if (Math.abs(yy - y) <= TOL) { key = yy; break; } }
      if (key === null) { key = y; ys.push(y); rows.set(y, []); }
      rows.get(key).push(it);
    }
    return [...rows.entries()]
      .sort((a, b) => b[0] - a[0])                    // PDF y grows upward → top first
      .map(([y, its]) => {
        its.sort((a, b) => a.transform[4] - b.transform[4]);
        const text = its.map(i => i.str).join('').replace(/\s+/g, ' ').trim();
        const h = Math.max(...its.map(i => i.height || Math.abs(i.transform[3]) || 10));
        return { y, text, h };
      })
      .filter(l => l.text.length);
  }

  function median(nums) {
    if (!nums.length) return 0;
    const s = [...nums].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
  }

  /* ============================================================
     TOOL IMPLEMENTATIONS
     Each returns { blob, defaultName, note? }
     ============================================================ */

  /* ---- 1 · Merge PDF ---- */
  async function runMergePdf(files, opts, onProgress) {
    const { PDFDocument } = PDFLib;
    const total = files.reduce((s, f) => s + f.size, 0);
    if (total > 250 * 1024 * 1024) throw new Error('These files are over 250 MB combined — too large to merge safely in a browser tab.');

    const out = await PDFDocument.create();
    for (let i = 0; i < files.length; i++) {
      onProgress(`Adding ${files[i].name}  (${i + 1} of ${files.length})…`);
      let src;
      try {
        src = await PDFDocument.load(await files[i].arrayBuffer(), { ignoreEncryption: true, updateMetadata: false });
      } catch (e) { throw friendlyPdfLoadError(files[i].name, e); }
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach(p => out.addPage(p));
    }
    if (out.getPageCount() === 0) throw new Error('No pages could be read from these files.');
    onProgress('Saving merged PDF…');
    out.setTitle('Merged PDF');
    out.setProducer('ConvertTools');
    out.setCreationDate(new Date());
    const bytes = await out.save();
    return { blob: new Blob([bytes], { type: 'application/pdf' }), defaultName: 'merged.pdf' };
  }

  /* ---- 2 · Merge Word docs (.docx) ---- */
  async function runMergeDocs(files, opts, onProgress) {
    const D = docx;
    const numberingConfigs = [];
    const counters = { num: 0 };
    const children = [];

    for (let i = 0; i < files.length; i++) {
      onProgress(`Reading ${files[i].name}  (${i + 1} of ${files.length})…`);
      let result;
      try {
        result = await mammoth.convertToHtml({ arrayBuffer: await files[i].arrayBuffer() });
      } catch (e) {
        throw new Error(`Could not read "${files[i].name}". Only modern .docx files are supported (old .doc files are not).`);
      }
      const part = await htmlToDocxChildren(result.value, numberingConfigs, counters);
      if (i > 0) {
        // each document starts on a fresh page
        children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
      }
      children.push(...part);
    }
    onProgress('Building merged document…');
    const doc = new D.Document({
      creator: 'ConvertTools',
      title: 'Merged Document',
      numbering: { config: numberingConfigs },
      sections: [{ properties: {}, children }]
    });
    const blob = await D.Packer.toBlob(doc);
    return {
      blob,
      defaultName: 'merged.docx',
      note: 'Documents are joined with a page break between each one. Very complex styling may be simplified — check the result in Word.'
    };
  }

  /* ---- 3 · Photos → PDF ---- */
  async function embedPdfImage(pdf, file) {
    const buf = await file.arrayBuffer();
    if (file.type === 'image/jpeg' || /\.jpe?g$/i.test(file.name)) return pdf.embedJpg(buf);
    if (file.type === 'image/png' || /\.png$/i.test(file.name)) return pdf.embedPng(buf);
    const bitmap = await createImageBitmap(new Blob([buf], { type: file.type }));
    const c = document.createElement('canvas');
    c.width = bitmap.width; c.height = bitmap.height;
    c.getContext('2d').drawImage(bitmap, 0, 0);
    const png = await new Promise(r => c.toBlob(r, 'image/png'));
    return pdf.embedPng(await png.arrayBuffer());
  }

  async function runImagesToPdf(files, opts, onProgress) {
    const { PDFDocument } = PDFLib;
    const PAGE_SIZES = { a4: [595.28, 841.89], letter: [612, 792] };
    const MARGINS = { none: 0, small: 28, large: 57 };
    const margin = MARGINS[opts.margin] || 0;

    const pdf = await PDFDocument.create();
    for (let i = 0; i < files.length; i++) {
      onProgress(`Embedding ${files[i].name}  (${i + 1} of ${files.length})…`);
      let img;
      try { img = await embedPdfImage(pdf, files[i]); }
      catch (e) { throw new Error(`"${files[i].name}" could not be read as an image.`); }

      let pageW, pageH;
      if (opts.pageSize === 'fit') {
        const landscape = opts.orientation === 'landscape';
        pageW = (landscape ? Math.max(img.width, img.height) : img.width) + margin * 2;
        pageH = (landscape ? Math.min(img.width, img.height) : img.height) + margin * 2;
        if (landscape && img.width < img.height) { pageW = img.height + margin * 2; pageH = img.width + margin * 2; }
      } else {
        [pageW, pageH] = PAGE_SIZES[opts.pageSize];
        const imgLandscape = img.width > img.height;
        const wantLandscape = opts.orientation === 'landscape' ||
          (opts.orientation === 'auto' && imgLandscape);
        if (wantLandscape) [pageW, pageH] = [pageH, pageW];
      }

      const page = pdf.addPage([pageW, pageH]);
      const maxW = pageW - margin * 2, maxH = pageH - margin * 2;
      const s = opts.pageSize === 'fit' ? 1 : Math.min(maxW / img.width, maxH / img.height, 1);
      let dw = img.width * s, dh = img.height * s;
      if (opts.pageSize === 'fit' && opts.orientation === 'landscape' && img.width < img.height) { /* keep natural */ }
      page.drawImage(img, { x: (pageW - dw) / 2, y: (pageH - dh) / 2, width: dw, height: dh });
    }
    onProgress('Saving PDF…');
    pdf.setTitle('Images PDF');
    pdf.setProducer('ConvertTools');
    const bytes = await pdf.save();
    return { blob: new Blob([bytes], { type: 'application/pdf' }), defaultName: 'images.pdf' };
  }

  /* ---- 4 · Word (.docx) → PDF ---- */
  async function runDocToPdf(files, opts, onProgress) {
    const file = files[0];
    onProgress('Reading ' + file.name + '…');
    let result;
    try {
      result = await mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() });
    } catch (e) {
      throw new Error('Could not read this file. Only modern .docx files are supported — if it is an old .doc file, open it in Word/LibreOffice and "Save As" .docx first.');
    }
    if (!result.value.trim()) throw new Error('This document appears to be empty.');

    const host = $('#render-host');
    host.innerHTML = '<div class="pdf-page">' + result.value + '</div>';
    onProgress('Rendering pages…');
    await waitForImages(host);

    try {
      const blob = await html2pdf()
        .set({
          margin: 0,
          image: { type: 'jpeg', quality: 0.95 },
          html2canvas: { scale: 2, useCORS: true, allowTaint: true, backgroundColor: '#ffffff', logging: false },
          jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
          pagebreak: { mode: ['css', 'legacy'] }
        })
        .from(host.firstElementChild)
        .outputPdf('blob');
      return { blob, defaultName: 'converted.pdf' };
    } catch (e) {
      throw new Error('Rendering failed. The document may be too large or complex for in-browser conversion.');
    } finally {
      host.innerHTML = '';
    }
  }

  /* ---- 5 · PDF → Word (.docx) ---- */
  async function runPdfToDoc(files, opts, onProgress) {
    const D = docx;
    let pdf;
    try {
      pdf = await loadPdf(await files[0].arrayBuffer());
    } catch (e) { throw friendlyPdfLoadError(files[0].name, e); }

    const children = [];
    let charCount = 0;
    for (let p = 1; p <= pdf.numPages; p++) {
      onProgress(`Reading page ${p} of ${pdf.numPages}…`);
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      const lines = groupTextLines(tc.items);
      const med = median(lines.map(l => l.h));
      let prevY = null;
      for (const line of lines) {
        charCount += line.text.length;
        let heading;
        if (med && line.text.length < 120) {
          if (line.h >= med * 1.75) heading = D.HeadingLevel.HEADING_1;
          else if (line.h >= med * 1.32) heading = D.HeadingLevel.HEADING_2;
        }
        const gap = prevY !== null ? Math.abs(prevY - line.y) : 0;
        children.push(new D.Paragraph({
          heading,
          spacing: { after: heading ? 140 : (gap > line.h * 1.8 ? 120 : 40) },
          children: [new D.TextRun(line.text)]
        }));
        prevY = line.y;
      }
      if (p < pdf.numPages) children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
    }
    if (charCount === 0) {
      throw new Error('No text could be found in this PDF — it looks like a scanned/image-only document. Try the "PDF → Photos" tool instead.');
    }
    onProgress('Building Word document…');
    const doc = new D.Document({
      creator: 'ConvertTools',
      title: 'Converted Document',
      sections: [{ properties: {}, children }]
    });
    const blob = await D.Packer.toBlob(doc);
    return {
      blob,
      defaultName: 'converted.docx',
      note: 'Text and headings were extracted for editing. Images, columns, tables and exact fonts are not reproduced — that requires server-side conversion.'
    };
  }

  /* ---- 6 · PDF → Photos ---- */
  async function runPdfToImages(files, opts, onProgress) {
    let pdf;
    try {
      pdf = await loadPdf(await files[0].arrayBuffer());
    } catch (e) { throw friendlyPdfLoadError(files[0].name, e); }

    const count = pdf.numPages;
    if (count > 40) throw new Error(`This PDF has ${count} pages — too many to render at once in a browser (limit: 40).`);
    const ext = opts.format === 'jpeg' ? 'jpg' : 'png';
    const mime = opts.format === 'jpeg' ? 'image/jpeg' : 'image/png';
    const scale = opts.dpi / 72;

    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const blobs = [];
    for (let p = 1; p <= count; p++) {
      onProgress(`Rendering page ${p} of ${count}…`);
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale });
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport }).promise;
      const blob = await new Promise(r => canvas.toBlob(r, mime, 0.92));
      blobs.push({ name: `page-${p}.${ext}`, blob });
    }

    if (blobs.length === 1) {
      return { blob: blobs[0].blob, defaultName: blobs[0].name };
    }
    onProgress('Zipping images…');
    const zip = new JSZip();
    blobs.forEach(b => zip.file(b.name, b.blob));
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    return { blob: zipBlob, defaultName: 'pages.zip' };
  }

  /* ---- 7 · TXT → PDF ---- */
  async function runTxtToPdf(files, opts, onProgress) {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    onProgress('Reading text file…');
    const raw = (await files[0].text()).replace(/\t/g, '    ').replace(/\r\n?/g, '\n');
    if (!raw.trim()) throw new Error('This text file is empty.');

    const pdf = await PDFDocument.create();
    const font = await pdf.embedFont(StandardFonts.Helvetica);
    const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
    const size = parseInt(opts.fontSize, 10) || 12;
    const lineH = size * 1.45;
    const pageW = 595.28, pageH = 841.89, m = 56.7;
    const maxW = pageW - m * 2;

    function widthOf(str) {
      try { return font.widthOfTextAtSize(str, size); }
      catch (e) {
        throw new Error('This text uses characters the built-in PDF font cannot draw (only Latin/Western characters are supported — e.g. no Hindi, Telugu or emoji).');
      }
    }

    onProgress('Laying out pages…');
    const lines = [];
    for (const rawLine of raw.split('\n')) {
      if (!rawLine.trim()) { lines.push(''); continue; }
      let line = '';
      for (const word of rawLine.split(' ')) {
        const test = line ? line + ' ' + word : word;
        if (widthOf(test) > maxW) {
          if (line) { lines.push(line); line = word; }
          else { // single very long word → hard-split
            let chunk = '';
            for (const ch of word) {
              if (widthOf(chunk + ch) > maxW) { lines.push(chunk); chunk = ch; } else chunk += ch;
            }
            line = chunk;
          }
        } else line = test;
      }
      if (line) lines.push(line);
    }

    const perPage = Math.floor((pageH - m * 2) / lineH);
    const pageCount = Math.max(1, Math.ceil(lines.length / perPage));
    for (let p = 0; p < pageCount; p++) {
      const page = pdf.addPage([pageW, pageH]);
      let y = pageH - m - lineH;
      for (let i = p * perPage; i < Math.min((p + 1) * perPage, lines.length); i++) {
        if (lines[i]) page.drawText(lines[i], { x: m, y, size, font, color: rgb(0.12, 0.12, 0.16) });
        y -= lineH;
      }
      const label = `Page ${p + 1} of ${pageCount}`;
      page.drawText(label, { x: (pageW - bold.widthOfTextAtSize(label, 9)) / 2, y: m / 2, size: 9, font: bold, color: rgb(0.5, 0.5, 0.56) });
    }
    onProgress('Saving PDF…');
    pdf.setTitle('Text PDF');
    pdf.setProducer('ConvertTools');
    const bytes = await pdf.save();
    return { blob: new Blob([bytes], { type: 'application/pdf' }), defaultName: 'text.pdf' };
  }

  /* ---- 8 · PDF → TXT ---- */
  async function runPdfToTxt(files, opts, onProgress) {
    let pdf;
    try {
      pdf = await loadPdf(await files[0].arrayBuffer());
    } catch (e) { throw friendlyPdfLoadError(files[0].name, e); }

    const chunks = [];
    let charCount = 0;
    for (let p = 1; p <= pdf.numPages; p++) {
      onProgress(`Reading page ${p} of ${pdf.numPages}…`);
      const page = await pdf.getPage(p);
      const lines = groupTextLines((await page.getTextContent()).items);
      chunks.push(lines.map(l => l.text).join('\n'));
      charCount += lines.reduce((s, l) => s + l.text.length, 0);
    }
    if (charCount === 0) {
      throw new Error('No text found — this looks like a scanned/image-only PDF. Use the "PDF → Photos" tool instead.');
    }
    const text = chunks.join('\n\n');
    return { blob: new Blob([text], { type: 'text/plain;charset=utf-8' }), defaultName: 'extracted.txt' };
  }

  /* ============================================================
     Tool registry
     ============================================================ */
  const svgWrap = inner =>
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;

  const ICONS = {
    layers: svgWrap('<path d="m12 2 10 5-10 5L2 7z"/><path d="m2 12 10 5 10-5"/><path d="m2 17 10 5 10-5"/>'),
    copy: svgWrap('<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>'),
    image: svgWrap('<rect width="18" height="18" x="3" y="3" rx="2" ry="2"/><circle cx="9" cy="9" r="2"/><path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"/>'),
    fileText: svgWrap('<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>'),
    swap: svgWrap('<path d="m17 11 4 4-4 4"/><path d="M21 15H7"/><path d="m7 21-4-4 4-4"/><path d="M3 9h14"/>'),
    images: svgWrap('<path d="M18 22H4a2 2 0 0 1-2-2V6"/><path d="m22 13-1.296-1.296a2.41 2.41 0 0 0-3.408 0L11 18"/><circle cx="12" cy="8" r="2"/><rect width="16" height="16" x="6" y="2" rx="2"/>'),
    filePlus: svgWrap('<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M12 12v6"/><path d="M9 15h6"/>'),
    text: svgWrap('<polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/>')
  };

  const isPdf = f => /\.pdf$/i.test(f.name) || f.type === 'application/pdf';
  const isDocx = f => /\.docx$/i.test(f.name);
  const isOldDoc = f => /\.doc$/i.test(f.name);
  const isImage = f => /^image\//.test(f.type) || /\.(png|jpe?g|gif|webp|bmp|avif)$/i.test(f.name);
  const isTxt = f => /\.(txt|md|log|text|csv)$/i.test(f.name) || f.type === 'text/plain';

  const TOOLS = [
    {
      id: 'merge-pdf', title: 'Merge PDF', icon: ICONS.layers,
      grad: ['#f43f5e', '#f97316'],
      desc: 'Combine multiple PDF files into a single document, in any order.',
      accept: '.pdf,application/pdf', multiple: true, minFiles: 2,
      check: f => isPdf(f) || 'Only PDF files are accepted here.',
      run: runMergePdf,
      info: 'Files are merged in the order shown — use the ▲▼ arrows to rearrange. Encryption is ignored where possible, but password-locked PDFs must be unlocked first.'
    },
    {
      id: 'merge-docs', title: 'Merge Word Docs', icon: ICONS.copy,
      grad: ['#3b82f6', '#06b6d4'],
      desc: 'Join several .docx documents into one Word file with page breaks.',
      accept: '.docx', multiple: true, minFiles: 2,
      check: f => isDocx(f) || (isOldDoc(f)
        ? 'Old .doc (Word 97–2003) files can\'t be read in a browser — please "Save As" .docx first.'
        : 'Only .docx files are accepted here.'),
      run: runMergeDocs,
      info: 'Each document starts on a new page. Headings, bold/italic, lists, tables and images are carried over; very complex styling may be simplified.'
    },
    {
      id: 'img-to-pdf', title: 'Photos → PDF', icon: ICONS.image,
      grad: ['#ec4899', '#a855f7'],
      desc: 'Turn JPG, PNG, WebP or GIF images into one PDF — one page per photo.',
      accept: 'image/*', multiple: true, minFiles: 1,
      check: f => isImage(f) || 'Only image files are accepted here.',
      options: [
        { key: 'pageSize', label: 'Page size', type: 'select', default: 'a4', choices: { fit: 'Fit each image', a4: 'A4', letter: 'US Letter' } },
        { key: 'orientation', label: 'Orientation', type: 'select', default: 'auto', choices: { auto: 'Auto', portrait: 'Portrait', landscape: 'Landscape' } },
        { key: 'margin', label: 'Margin', type: 'select', default: 'small', choices: { none: 'None', small: 'Small', large: 'Large' } }
      ],
      run: runImagesToPdf
    },
    {
      id: 'doc-to-pdf', title: 'Word → PDF', icon: ICONS.fileText,
      grad: ['#14b8a6', '#22c55e'],
      desc: 'Convert a .docx document into a clean, shareable PDF.',
      accept: '.docx', multiple: false, minFiles: 1,
      check: f => isDocx(f) || (isOldDoc(f)
        ? 'Old .doc files can\'t be converted in a browser — please "Save As" .docx first.'
        : 'Only .docx files are accepted here.'),
      run: runDocToPdf,
      limits: 'Browser-only conversion is approximate: text, headings, lists, tables and images carry over, but exact fonts, headers/footers and complex layouts may shift. Old .doc format is not supported.'
    },
    {
      id: 'pdf-to-doc', title: 'PDF → Word', icon: ICONS.swap,
      grad: ['#6366f1', '#8b5cf6'],
      desc: 'Extract text from a PDF into an editable .docx document.',
      accept: '.pdf,application/pdf', multiple: false, minFiles: 1,
      check: f => isPdf(f) || 'Only PDF files are accepted here.',
      run: runPdfToDoc,
      limits: 'Produces an editable Word file with text and detected headings. Images, columns, tables and exact fonts are not reproduced. Scanned (image-only) PDFs have no text to extract.'
    },
    {
      id: 'pdf-to-img', title: 'PDF → Photos', icon: ICONS.images,
      grad: ['#f59e0b', '#ef4444'],
      desc: 'Render every page of a PDF as a PNG or JPG image (zipped if many).',
      accept: '.pdf,application/pdf', multiple: false, minFiles: 1,
      check: f => isPdf(f) || 'Only PDF files are accepted here.',
      options: [
        { key: 'format', label: 'Image format', type: 'select', default: 'png', choices: { png: 'PNG (best quality)', jpeg: 'JPG (smaller size)' } },
        { key: 'dpi', label: 'Resolution', type: 'select', default: 150, choices: { 72: '72 DPI (screen)', 150: '150 DPI (standard)', 300: '300 DPI (print)' } }
      ],
      run: runPdfToImages
    },
    {
      id: 'txt-to-pdf', title: 'TXT → PDF', icon: ICONS.filePlus,
      grad: ['#84cc16', '#10b981'],
      desc: 'Turn a plain-text or markdown file into a neatly paginated PDF.',
      accept: '.txt,.md,.log,.csv,text/plain', multiple: false, minFiles: 1,
      check: f => isTxt(f) || 'Only plain-text files (.txt, .md, …) are accepted here.',
      options: [
        { key: 'fontSize', label: 'Font size', type: 'select', default: 12, choices: { 10: 'Small (10pt)', 12: 'Normal (12pt)', 14: 'Large (14pt)' } }
      ],
      run: runTxtToPdf,
      info: 'Text is wrapped and paginated on A4 with page numbers. Only Latin characters are supported by the built-in font.'
    },
    {
      id: 'pdf-to-txt', title: 'PDF → TXT', icon: ICONS.text,
      grad: ['#0ea5e9', '#6366f1'],
      desc: 'Pull all the text out of a PDF into a plain .txt file.',
      accept: '.pdf,application/pdf', multiple: false, minFiles: 1,
      check: f => isPdf(f) || 'Only PDF files are accepted here.',
      run: runPdfToTxt
    }
  ];

  /* ============================================================
     Tool cards
     ============================================================ */
  const grid = $('#tools-grid');
  for (const tool of TOOLS) {
    const card = document.createElement('button');
    card.className = 'tool-card';
    card.style.setProperty('--g1', tool.grad[0]);
    card.style.setProperty('--g2', tool.grad[1]);
    card.setAttribute('aria-label', tool.title);
    card.innerHTML = `
      <div class="tool-icon">${tool.icon}</div>
      <h3>${tool.title}</h3>
      <p>${tool.desc}</p>
      <div class="tool-go">Open tool →</div>`;
    card.addEventListener('click', () => openTool(tool));
    grid.appendChild(card);
  }

  /* ============================================================
     Modal framework
     ============================================================ */
  const modalRoot = $('#modal-root');
  let state = null;

  function openTool(tool) {
    state = { tool, files: [], busy: false, result: null, uid: 0 };

    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal" role="dialog" aria-modal="true" aria-label="${tool.title}"
           style="--g1:${tool.grad[0]};--g2:${tool.grad[1]}">
        <div class="modal-head">
          <div class="tool-icon">${tool.icon}</div>
          <h2>${tool.title}</h2>
          <button class="modal-close" aria-label="Close">✕</button>
        </div>
        <div class="modal-body">
          <div class="dropzone" tabindex="0" role="button" aria-label="Add files">
            <div class="dropzone-icon">📂</div>
            <strong>Drop ${tool.multiple ? 'files' : 'a file'} here, or click to browse</strong>
            <span>Accepts: ${tool.accept.replace(/,/g, ', ')}</span>
          </div>
          <input type="file" hidden ${tool.multiple ? 'multiple' : ''} accept="${tool.accept}">
          <div class="file-list"></div>
          <div class="options-area"></div>
          ${tool.limits ? `<div class="limits-note">⚠️ <strong>Good to know:</strong> ${tool.limits}</div>` : ''}
          ${tool.info ? `<div class="info-note">💡 ${tool.info}</div>` : ''}
          <button class="convert-btn" disabled>Convert</button>
          <div class="progress-wrap">
            <div class="progress-bar"><div class="progress-fill"></div></div>
            <div class="progress-text">Working…</div>
          </div>
          <div class="result-box"></div>
          <div class="result-error"></div>
        </div>
      </div>`;
    modalRoot.appendChild(overlay);
    document.body.style.overflow = 'hidden';

    const close = () => {
      revokeResultUrl();
      overlay.remove();
      document.body.style.overflow = '';
      state = null;
      document.removeEventListener('keydown', onKey);
    };
    const onKey = e => { if (e.key === 'Escape') close(); };
    document.addEventListener('keydown', onKey);
    overlay.addEventListener('click', e => { if (e.target === overlay) close(); });
    $('.modal-close', overlay).addEventListener('click', close);

    const dz = $('.dropzone', overlay);
    const input = $('input[type=file]', overlay);
    dz.addEventListener('click', () => input.click());
    dz.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', () => { addFiles(input.files); input.value = ''; });
    ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('dragover'); }));
    dz.addEventListener('drop', e => addFiles(e.dataTransfer.files));

    renderOptions(tool, overlay);
    $('.convert-btn', overlay).addEventListener('click', () => doConvert(overlay));
  }

  function addFiles(fileList) {
    if (!state) return;
    const { tool } = state;
    let added = 0;
    for (const f of fileList) {
      const verdict = tool.check(f);
      if (verdict !== true) { toast(verdict, true); continue; }
      if (!tool.multiple && state.files.length >= 1) { state.files = []; toast('This tool accepts one file — replaced the previous one.'); }
      state.files.push({ id: ++state.uid, file: f });
      added++;
    }
    if (added) renderFiles();
  }

  function renderFiles() {
    if (!state) return;
    const { tool } = state;
    const overlay = modalRoot.querySelector('.modal-overlay');
    const list = $('.file-list', overlay);
    list.innerHTML = '';
    state.files.forEach((entry, idx) => {
      const f = entry.file;
      const ext = (f.name.match(/\.(\w+)$/) || [null, 'file'])[1].slice(0, 4);
      const row = document.createElement('div');
      row.className = 'file-item';
      row.innerHTML = `
        <div class="f-icon">${esc(ext)}</div>
        <div class="f-meta">
          <div class="f-name" title="${esc(f.name)}">${esc(f.name)}</div>
          <div class="f-size">${formatSize(f.size)}</div>
        </div>
        ${tool.multiple && state.files.length > 1 ? `
          <div class="f-order">
            <button class="f-btn f-up" title="Move up" ${idx === 0 ? 'disabled' : ''}>▲</button>
            <button class="f-btn f-down" title="Move down" ${idx === state.files.length - 1 ? 'disabled' : ''}>▼</button>
          </div>` : ''}
        <button class="f-remove" title="Remove">✕</button>`;
      $('.f-remove', row).addEventListener('click', () => {
        state.files = state.files.filter(x => x.id !== entry.id);
        renderFiles();
      });
      const up = $('.f-up', row), down = $('.f-down', row);
      if (up) up.addEventListener('click', () => { [state.files[idx - 1], state.files[idx]] = [state.files[idx], state.files[idx - 1]]; renderFiles(); });
      if (down) down.addEventListener('click', () => { [state.files[idx + 1], state.files[idx]] = [state.files[idx], state.files[idx + 1]]; renderFiles(); });
      list.appendChild(row);
    });

    // "add more" affordance
    const old = $('.add-more', overlay);
    if (old) old.remove();
    if (tool.multiple && state.files.length) {
      const btn = document.createElement('button');
      btn.className = 'add-more';
      btn.textContent = '＋ Add more files';
      btn.addEventListener('click', () => $('input[type=file]', overlay).click());
      list.after(btn);
    }

    const convertBtn = $('.convert-btn', overlay);
    convertBtn.disabled = state.files.length < (tool.minFiles || 1) || state.busy;
    if (tool.minFiles > 1 && state.files.length === 1) {
      convertBtn.textContent = `Add at least ${tool.minFiles} files`;
      convertBtn.disabled = true;
    } else {
      convertBtn.textContent = 'Convert';
    }
  }

  function renderOptions(tool, overlay) {
    const area = $('.options-area', overlay);
    if (!tool.options || !tool.options.length) return;
    const wrap = document.createElement('div');
    wrap.className = 'options';
    for (const opt of tool.options) {
      const field = document.createElement('div');
      field.className = 'option-field' + (tool.options.length === 1 ? ' full' : '');
      const id = `opt-${opt.key}`;
      if (opt.type === 'select') {
        field.innerHTML = `<label for="${id}">${opt.label}</label>
          <select id="${id}" data-key="${opt.key}">
            ${Object.entries(opt.choices).map(([v, t]) =>
              `<option value="${v}" ${String(opt.default) === v ? 'selected' : ''}>${t}</option>`).join('')}
          </select>`;
      }
      wrap.appendChild(field);
    }
    area.appendChild(wrap);
  }

  function readOptions(overlay) {
    const opts = {};
    overlay.querySelectorAll('[data-key]').forEach(el => {
      opts[el.dataset.key] = el.tagName === 'SELECT' && /^\d+$/.test(el.value) ? Number(el.value) : el.value;
    });
    return opts;
  }

  async function doConvert(overlay) {
    if (!state || state.busy) return;
    const { tool } = state;
    state.busy = true;

    const btn = $('.convert-btn', overlay);
    const progress = $('.progress-wrap', overlay);
    const progressText = $('.progress-text', overlay);
    const errBox = $('.result-error', overlay);
    const resultBox = $('.result-box', overlay);
    btn.disabled = true;
    btn.textContent = 'Working…';
    errBox.classList.remove('active');
    resultBox.classList.remove('active');
    progress.classList.add('active');
    progressText.textContent = 'Starting…';

    try {
      const res = await tool.run(
        state.files.map(e => e.file),
        readOptions(overlay),
        msg => { progressText.textContent = msg; }
      );
      state.result = res;
      progress.classList.remove('active');
      showResult(overlay, tool, res);
    } catch (e) {
      progress.classList.remove('active');
      errBox.innerHTML = `<strong>😕 Something went wrong</strong>${e.message || 'Unexpected error.'}`;
      errBox.classList.add('active');
      btn.disabled = false;
      btn.textContent = 'Try again';
      state.busy = false;
    }
  }

  function showResult(overlay, tool, res) {
    const resultBox = $('.result-box', overlay);
    const btn = $('.convert-btn', overlay);
    btn.style.display = 'none';

    revokeResultUrl();
    state.resultUrl = URL.createObjectURL(res.blob);
    const ext = res.defaultName.slice(res.defaultName.lastIndexOf('.'));

    resultBox.innerHTML = `
      <div class="result-success">
        <div class="result-check">✓</div>
        <div>
          <strong>Done! Your file is ready.</strong>
          <span>${formatSize(res.blob.size)} · processed entirely on your device</span>
        </div>
      </div>
      ${res.note ? `<div class="limits-note" style="margin-top:12px">📌 ${res.note}</div>` : ''}
      <div class="filename-row">
        <label for="dl-name">File name</label>
        <input class="filename-input" id="dl-name" type="text" value="${res.defaultName}" spellcheck="false"
               autocomplete="off" autocapitalize="off" autocorrect="off">
      </div>
      <button class="download-btn">⬇ Download</button>
      <a class="download-link" href="#">Download not starting? Click here instead</a>
      <p class="result-hint">Some embedded page previews block downloads. If the button does nothing,
        right-click the link above and choose <em>“Save link as…”</em> — or open this site in a full browser tab.</p>
      <button class="reset-btn">↺ Convert another file</button>`;
    resultBox.classList.add('active');

    const nameInput = $('#dl-name', overlay);
    const fallback = $('.download-link', overlay);
    fallback.href = state.resultUrl;
    fallback.setAttribute('download', res.defaultName);
    fallback.setAttribute('target', '_blank');
    fallback.setAttribute('rel', 'noopener');

    // Pre-select just the base name so typing replaces it, keeping the extension visible
    const dot = res.defaultName.lastIndexOf('.');
    nameInput.focus();
    nameInput.setSelectionRange(0, dot > 0 ? dot : res.defaultName.length);

    nameInput.addEventListener('input', () => {
      fallback.setAttribute('download', sanitizeFileName(nameInput.value, res.defaultName, ext));
    });

    $('.download-btn', overlay).addEventListener('click', () => {
      const name = sanitizeFileName(nameInput.value, res.defaultName, ext);
      fallback.setAttribute('download', name);
      const a = document.createElement('a');
      a.href = state.resultUrl;
      a.download = name;
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      a.remove();
      toast(`Downloading "${name}" ⬇`);
    });

    $('.reset-btn', overlay).addEventListener('click', () => {
      revokeResultUrl();
      state.files = [];
      state.busy = false;
      state.result = null;
      resultBox.classList.remove('active');
      resultBox.innerHTML = '';
      btn.style.display = '';
      btn.disabled = true;
      btn.textContent = 'Convert';
      renderFiles();
    });
  }
})();
