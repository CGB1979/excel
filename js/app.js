const state = { files: [], records: [], headers: [], output: null, outputBytes: null };
const $ = id => document.getElementById(id);

function norm(v) {
  return String(v ?? '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function findCol(headers, names) {
  const wanted = names.map(norm);
  return (Array.isArray(headers) ? headers : []).findIndex(h => {
    const header = norm(h);
    return wanted.some(name => header === name || header.includes(name));
  });
}

function setProgress(n, text) {
  if ($('progressBar')) $('progressBar').style.width = `${n}%`;
  if ($('progressText')) $('progressText').textContent = text;
}

function cellValue(v) {
  if (v == null) return '';
  if (typeof v === 'object' && v.result != null) return v.result;
  return String(v);
}

async function readFile(file) {
  const wb = new ExcelJS.Workbook();
  const buf = await file.arrayBuffer();
  await wb.xlsx.load(buf);

  let records = [];
  let headers = [];

  wb.eachSheet(ws => {
    if (!headers.length) {
      const r = ws.getRow(1);
      const values = r && Array.isArray(r.values) ? r.values : [];
      headers = values.slice(1).map(cellValue);
    }

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const vals = row && Array.isArray(row.values) ? row.values.slice(1) : [];
      if (vals.some(v => cellValue(v).trim() !== '')) {
        records.push({
          values: headers.map((_, i) => cellValue(vals[i] ?? '')),
          source: file.name,
          sheet: ws.name
        });
      }
    }
  });

  if (!headers.length) throw new Error(`El archivo "${file.name}" no contiene una cabecera válida en la fila 1.`);
  return { name: file.name, rows: records.length, records, headers };
}

function renderFiles() {
  const box = $('fileList');
  if (!box) return;
  box.innerHTML = state.files.map(f => `<div class="file-row"><div class="file-name">${escapeHtml(f.name)}</div><div class="file-meta">${f.rows} registros</div></div>`).join('');
  box.classList.toggle('hidden', !state.files.length);
  $('fileCount').textContent = state.files.length;
  $('rowCount').textContent = state.records.length;
  $('statusPill').textContent = state.files.length ? `${state.files.length} archivo${state.files.length === 1 ? '' : 's'}` : 'Sin archivos';
  $('statusPill').classList.toggle('ready', !!state.files.length);
  $('exportBtn').disabled = !state.records.length;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}

function resetAppState() {
  state.files = [];
  state.records = [];
  state.headers = [];
  state.output = null;
  state.outputBytes = null;
  const input = $('fileInput');
  if (input) input.value = '';
  const error = $('loadError');
  if (error) { error.textContent = ''; error.classList.add('hidden'); }
  const result = $('resultBox');
  if (result) result.classList.add('hidden');
  const progress = $('progressArea');
  if (progress) progress.classList.add('hidden');
  renderFiles();
}

$('fileInput').addEventListener('change', async e => {
  state.files = [];
  state.records = [];
  state.headers = [];
  state.output = null;
  state.outputBytes = null;
  $('loadError').classList.add('hidden');
  $('resultBox').classList.add('hidden');

  try {
    const selected = Array.from(e.target.files || []);
    if (!selected.length) { renderFiles(); return; }

    for (const file of selected) {
      const r = await readFile(file);
      if (!state.headers.length) state.headers = r.headers;
      state.files.push({ name: r.name, rows: r.rows });
      state.records.push(...r.records);
    }
    renderFiles();
  } catch (err) {
    console.error(err);
    $('loadError').textContent = 'No se pudo leer el archivo Excel: ' + (err.message || err);
    $('loadError').classList.remove('hidden');
    state.files = [];
    state.records = [];
    state.headers = [];
    renderFiles();
  }
});

function barcodeDataUrl(value) {
  const c = document.createElement('canvas');
  c.width = 285;
  c.height = 40;
  JsBarcode(c, String(value), {
    format: 'CODE128',
    width: 1.45,
    height: 32,
    margin: 0,
    displayValue: false,
    lineColor: '#000000'
  });
  return c.toDataURL('image/png');
}

function dataUrlToUint8Array(dataUrl) {
  const comma = dataUrl.indexOf(',');
  if (comma < 0) throw new Error('Imagen de código de barras inválida.');
  const binary = atob(dataUrl.slice(comma + 1));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function xmlEscape(value) {
  return String(value).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&apos;' }[c]));
}

function nextRelationshipId(relsXml) {
  const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m => Number(m[1]));
  return `rId${ids.length ? Math.max(...ids) + 1 : 1}`;
}

function ensureDrawingRelationship(sheetRelsXml, drawingRid) {
  const rel = `<Relationship Id="${drawingRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>`;
  if (sheetRelsXml.includes('relationships/drawing')) return sheetRelsXml;
  return sheetRelsXml.replace('</Relationships>', `${rel}</Relationships>`);
}

function ensureDrawingNode(sheetXml, drawingRid) {
  if (sheetXml.includes('<drawing ')) return sheetXml;
  return sheetXml.replace('</worksheet>', `<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${drawingRid}"/></worksheet>`);
}

function ensureContentTypes(contentTypesXml) {
  if (!contentTypesXml.includes('Extension="png"')) {
    contentTypesXml = contentTypesXml.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
  }
  if (!contentTypesXml.includes('/xl/drawings/drawing1.xml')) {
    contentTypesXml = contentTypesXml.replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
  }
  return contentTypesXml;
}

function buildDrawingXml(images) {
  const NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const cx = 285 * 9525;
  const cy = 40 * 9525;

  const anchors = images.map((img, i) => {
    const row = img.row;
    const col = img.col;
    const rid = `rId${i + 1}`;
    const name = `CodigoBarra_${row + 1}`;
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="${xmlEscape(name)}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:a="${A}" r:embed="${rid}"/><a:stretch xmlns:a="${A}"><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm xmlns:a="${A}"><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom xmlns:a="${A}" prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${NS}" xmlns:a="${A}" xmlns:r="${R}">${anchors}</xdr:wsDr>`;
}

function buildDrawingRels(images) {
  const rels = images.map((img, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${i + 1}.png"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

async function addBarcodeImagesToXlsx(buffer, images) {
  const zip = await JSZip.loadAsync(buffer);
  const sheetPath = 'xl/worksheets/sheet1.xml';
  const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const contentTypesPath = '[Content_Types].xml';

  let sheetXml = await zip.file(sheetPath).async('string');
  let relsXml = zip.file(relsPath) ? await zip.file(relsPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let contentTypesXml = await zip.file(contentTypesPath).async('string');

  const drawingRid = nextRelationshipId(relsXml);
  relsXml = ensureDrawingRelationship(relsXml, drawingRid);
  sheetXml = ensureDrawingNode(sheetXml, drawingRid);
  contentTypesXml = ensureContentTypes(contentTypesXml);

  images.forEach((img, i) => {
    zip.file(`xl/media/image${i + 1}.png`, img.bytes);
  });
  zip.file('xl/drawings/drawing1.xml', buildDrawingXml(images));
  zip.file('xl/drawings/_rels/drawing1.xml.rels', buildDrawingRels(images));
  zip.file(sheetPath, sheetXml);
  zip.file(relsPath, relsXml);
  zip.file(contentTypesPath, contentTypesXml);

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

async function buildWorkbook() {
  if (!Array.isArray(state.headers) || !state.headers.length) throw new Error('No hay cabeceras cargadas. Volvé a seleccionar el Excel.');
  if (!Array.isArray(state.records) || !state.records.length) throw new Error('No hay registros cargados. Volvé a seleccionar el Excel.');

  const chasisCol = findCol(state.headers, ['chasis']);
  if (chasisCol < 0) throw new Error('No se encontró ninguna columna cuyo encabezado contenga "Chasis".');

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Datos');
  const headers = [...state.headers, 'Código de barras'];
  ws.addRow(headers);
  ws.autoFilter = { from: 1, to: headers.length, fromRow: 1, toRow: Math.max(1, state.records.length + 1) };

  const barcodeCol = headers.length;
  ws.getColumn(barcodeCol).width = 41;
  ws.getRow(1).height = 22;
  ws.getRow(1).font = { bold: true };

  const images = [];
  for (let i = 0; i < state.records.length; i++) {
    const rec = state.records[i];
    const row = ws.addRow([...rec.values, '']);
    row.height = 30;
    const chasis = String(rec.values[chasisCol] ?? '').trim();
    if (chasis) {
      const data = barcodeDataUrl(chasis);
      images.push({ row: i + 1, col: barcodeCol - 1, bytes: dataUrlToUint8Array(data), chasis });
    }
    setProgress(Math.round(((i + 1) / state.records.length) * 80), `Generando código ${i + 1} de ${state.records.length}`);
    await new Promise(r => setTimeout(r, 0));
  }

  setProgress(85, 'Preparando archivo Excel...');
  const raw = await wb.xlsx.writeBuffer();
  setProgress(92, 'Insertando códigos de barras...');
  return addBarcodeImagesToXlsx(raw, images);
}

$('exportBtn').addEventListener('click', async () => {
  try {
    $('exportBtn').disabled = true;
    $('progressArea').classList.remove('hidden');
    $('resultBox').classList.add('hidden');
    $('loadError').classList.add('hidden');
    setProgress(5, 'Preparando Excel...');
    state.outputBytes = await buildWorkbook();
    setProgress(100, 'Listo');
    $('resultText').textContent = `${state.records.length} registros preparados. El archivo XLSX incluye los códigos de barras.`;
    $('resultBox').classList.remove('hidden');
  } catch (err) {
    console.error(err);
    $('loadError').textContent = 'No se pudo generar el Excel: ' + (err.message || err);
    $('loadError').classList.remove('hidden');
  } finally {
    $('exportBtn').disabled = false;
  }
});

$('downloadBtn').addEventListener('click', () => {
  if (!state.outputBytes) return;
  const blob = new Blob([state.outputBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'excel_unificado_v1.2.3.xlsx';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
});

window.addEventListener('DOMContentLoaded', resetAppState);
