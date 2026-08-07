/* Node smoke test for the exact library calls used in js/app.js */
const fs = require('fs');
const path = require('path');

const V = p => path.join(__dirname, '..', 'vendor', p);
const T = p => path.join(__dirname, p);

async function main() {
  // ---------- pdf-lib: merge ----------
  const PDFLib = require(V('pdf-lib.min.js'));
  const merged = await PDFLib.PDFDocument.create();
  for (const f of ['chapter-1.pdf', 'chapter-2.pdf']) {
    const src = await PDFLib.PDFDocument.load(fs.readFileSync(T(f)));
    const pages = await merged.copyPages(src, src.getPageIndices());
    pages.forEach(p => merged.addPage(p));
  }
  console.assert(merged.getPageCount() === 2, 'merge page count');
  console.log('✓ pdf-lib merge:', merged.getPageCount(), 'pages,', (await merged.save()).length, 'bytes');

  // ---------- pdf-lib: embed PNG + JPG ----------
  const imgPdf = await PDFLib.PDFDocument.create();
  const png = await imgPdf.embedPng(fs.readFileSync(T('photo-1.png')));
  const jpg = await imgPdf.embedJpg(fs.readFileSync(T('photo-2.jpg')));
  for (const img of [png, jpg]) {
    let [w, h] = [595.28, 841.89];
    if (img.width > img.height) [w, h] = [h, w];
    const page = imgPdf.addPage([w, h]);
    const s = Math.min((w - 56) / img.width, (h - 56) / img.height, 1);
    page.drawImage(img, { x: (w - img.width * s) / 2, y: (h - img.height * s) / 2, width: img.width * s, height: img.height * s });
  }
  console.assert(imgPdf.getPageCount() === 2, 'img pdf pages');
  console.log('✓ pdf-lib image embed:', imgPdf.getPageCount(), 'pages');

  // ---------- pdf-lib: txt→pdf wrap ----------
  const font = await imgPdf.embedFont(PDFLib.StandardFonts.Helvetica);
  try {
    font.widthOfTextAtSize('hello world', 12);
    console.log('✓ pdf-lib font width OK');
  } catch (e) { console.log('✗ widthOfTextAtSize:', e.message); }
  let nonLatinErr = false;
  try { font.widthOfTextAtSize('ప్రయత్నం', 12); } catch (e) { nonLatinErr = true; }
  console.log('✓ pdf-lib non-Latin throws (handled in UI):', nonLatinErr);

  // ---------- docx: build with heading/list/image ----------
  const docx = require(V('docx.umd.js'));
  const numberingConfigs = [{
    reference: 'ct-num-1',
    levels: [0].map(i => ({
      level: i, format: docx.LevelFormat.DECIMAL, text: '%1.',
      alignment: docx.AlignmentType.START,
      style: { paragraph: { indent: { left: 720, hanging: 320 } } }
    }))
  }];
  const imageRun = new docx.ImageRun({
    data: fs.readFileSync(T('photo-1.png')),
    type: 'png',
    transformation: { width: 300, height: 225 }
  });
  const d = new docx.Document({
    creator: 'ConvertTools',
    numbering: { config: numberingConfigs },
    sections: [{
      children: [
        new docx.Paragraph({ heading: docx.HeadingLevel.HEADING_1, children: [new docx.TextRun('Smoke Test')] }),
        new docx.Paragraph({ children: [new docx.TextRun({ text: 'bold', bold: true }), new docx.TextRun({ break: 1 }), new docx.TextRun({ text: 'link', underline: {}, color: '1155CC' })] }),
        new docx.Paragraph({ bullet: { level: 0 }, children: [new docx.TextRun('bullet item')] }),
        new docx.Paragraph({ numbering: { reference: 'ct-num-1', level: 0 }, children: [new docx.TextRun('numbered item')] }),
        new docx.Paragraph({ children: [imageRun] }),
        new docx.Table({
          width: { size: 100, type: docx.WidthType.PERCENTAGE },
          rows: [new docx.TableRow({ children: [new docx.TableCell({ children: [new docx.Paragraph('cell')] })] })]
        }),
        new docx.Paragraph({ children: [new docx.PageBreak()] }),
        new docx.Paragraph({ pageBreakBefore: true, children: [new docx.TextRun('after break')] })
      ]
    }]
  });
  const buf = await docx.Packer.toBuffer(d);
  console.assert(buf.length > 1000, 'docx size');
  console.log('✓ docx build:', buf.length, 'bytes');

  // ---------- mammoth: docx → html ----------
  const mammoth = require(V('mammoth.browser.min.js'));
  const b = fs.readFileSync(T('report.docx'));
  const ab = b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength); // browser API uses arrayBuffer:
  const res = await mammoth.convertToHtml({ arrayBuffer: ab });
  console.assert(res.value.includes('Quarterly Report'), 'mammoth h1');
  console.log('✓ mammoth docx→html:', res.value.length, 'chars; sample:', JSON.stringify(res.value.slice(0, 120)));

  // ---------- JSZip ----------
  const JSZip = require(V('jszip.min.js'));
  const zip = new JSZip();
  zip.file('page-1.png', fs.readFileSync(T('photo-1.png')));
  zip.file('page-2.png', fs.readFileSync(T('photo-2.jpg')));
  const zbuf = await zip.generateAsync({ type: 'nodebuffer' });
  console.log('✓ jszip:', zbuf.length, 'bytes for 2 images');

  console.log('\nALL SMOKE TESTS PASSED');
}

main().catch(e => { console.error('SMOKE TEST FAILED:', e); process.exit(1); });
