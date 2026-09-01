const state = {
    files: [],
    records: [],
    headers: [],
    output: null,
    outputBytes: null
};

const $ = id => document.getElementById(id);

function norm(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');
}

function findCol(headers, names) {
    const list = Array.isArray(headers) ? headers : [];
    const wanted = names.map(norm);

    return list.findIndex(h => {
        const header = norm(h);
        return wanted.some(name => header === name || header.includes(name));
    });
}

function setProgress(n, text) {
    const bar = $('progressBar');
    const label = $('progressText');
    if (bar) bar.style.width = `${n}%`;
    if (label) label.textContent = text;
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
            headers = Array.from(r.values || [])
                .slice(1)
                .map(cellValue);
        }

        for (let r = 2; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            const vals = Array.from(row.values || []).slice(1);

            if (vals.some(v => cellValue(v).trim() !== '')) {
                records.push({
                    values: headers.map((_, i) => cellValue(vals[i] ?? '')),
                    source: file.name,
                    sheet: ws.name
                });
            }
        }
    });

    return {
        name: file.name,
        rows: records.length,
        records,
        headers
    };
}

function renderFiles() {
    const box = $('fileList');

    if (box) {
        box.innerHTML = state.files.map(f => `
            <div class="file-row">
                <div class="file-name">${escapeHtml(f.name)}</div>
                <div class="file-meta">${f.rows} registros</div>
            </div>
        `).join('');

        box.classList.toggle('hidden', !state.files.length);
    }

    $('fileCount').textContent = state.files.length;
    $('rowCount').textContent = state.records.length;
    $('statusPill').textContent = state.files.length
        ? `${state.files.length} archivo${state.files.length === 1 ? '' : 's'}`
        : 'Sin archivos';

    $('statusPill').classList.toggle('ready', !!state.files.length);
    $('exportBtn').disabled = !state.records.length;
}

function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    }[c]));
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
    if (error) {
        error.textContent = '';
        error.classList.add('hidden');
    }

    const result = $('resultBox');
    if (result) result.classList.add('hidden');

    const progress = $('progressArea');
    if (progress) progress.classList.add('hidden');

    setProgress(0, 'Preparando...');
    renderFiles();
}

window.addEventListener('DOMContentLoaded', resetAppState);

$('fileInput').addEventListener('change', async e => {
    state.files = [];
    state.records = [];
    state.headers = [];
    state.output = null;
    state.outputBytes = null;

    $('loadError').classList.add('hidden');

    try {
        const selectedFiles = Array.from(e.target.files || []);

        if (!selectedFiles.length) {
            renderFiles();
            return;
        }

        for (const file of selectedFiles) {
            const r = await readFile(file);

            if (!state.headers.length) {
                state.headers = r.headers;
            }

            state.files.push({
                name: r.name,
                rows: r.rows
            });

            state.records.push(...r.records);
        }

        renderFiles();

    } catch (err) {
        console.error(err);

        $('loadError').textContent =
            'No se pudo leer el archivo Excel: ' +
            (err.message || err);

        $('loadError').classList.remove('hidden');

        state.files = [];
        state.records = [];
        state.headers = [];
        state.output = null;
        state.outputBytes = null;

        renderFiles();
    }
});

function barcodeDataUrl(value) {
    if (typeof JsBarcode !== 'function') {
        throw new Error('No se pudo cargar la librería de códigos de barras (JsBarcode).');
    }

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

function excelColumnLetter(number) {
    let n = number;
    let result = '';

    while (n > 0) {
        const rem = (n - 1) % 26;
        result = String.fromCharCode(65 + rem) + result;
        n = Math.floor((n - 1) / 26);
    }

    return result;
}

async function buildWorkbook() {
    if (!Array.isArray(state.headers) || !state.headers.length) {
        throw new Error('No hay cabeceras cargadas. Volvé a seleccionar el Excel.');
    }

    if (!Array.isArray(state.records) || !state.records.length) {
        throw new Error('No hay registros cargados. Volvé a seleccionar el Excel.');
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Datos');

    const headers = [...state.headers, 'Código de barras'];
    ws.addRow(headers);

    ws.autoFilter = {
        from: 1,
        to: headers.length,
        fromRow: 1,
        toRow: Math.max(1, state.records.length + 1)
    };

    const chasisCol = findCol(state.headers, ['chasis', 'chassis']);

    if (chasisCol < 0) {
        throw new Error(
            'No se encontró ninguna columna cuyo encabezado contenga "Chasis".'
        );
    }

    const barcodeCol = headers.length;

    // 285 px ≈ 40.7 caracteres de ancho en Excel.
    ws.getColumn(barcodeCol).width = 40.7;

    // 40 px ≈ 30 puntos.
    ws.getRow(1).height = 22;
    ws.getRow(1).font = { bold: true };

    const barcodeColumnLetter = excelColumnLetter(barcodeCol);

    for (let i = 0; i < state.records.length; i++) {
        const rec = state.records[i];

        if (!Array.isArray(rec.values)) {
            throw new Error(`El registro ${i + 1} no tiene valores válidos.`);
        }

        const row = ws.addRow([...rec.values, '']);

        // 40 px aproximadamente.
        row.height = 30;

        const chasis = String(rec.values[chasisCol] ?? '').trim();

        if (chasis) {
            const data = barcodeDataUrl(chasis);

            const imageId = wb.addImage({
                base64: data,
                extension: 'png'
            });

            /*
             * Usamos rango de celda en lugar de tl/br.
             * ExcelJS documenta este formato y evita varios
             * problemas de anclaje de imágenes presentes en 4.4.0.
             */
            const cellAddress = `${barcodeColumnLetter}${i + 2}`;
            ws.addImage(imageId, cellAddress);
        }

        setProgress(
            Math.round(((i + 1) / state.records.length) * 100),
            `Generando código ${i + 1} de ${state.records.length}`
        );

        await new Promise(r => setTimeout(r, 0));
    }

    return wb;
}

$('exportBtn').addEventListener('click', async () => {
    try {
        $('exportBtn').disabled = true;
        $('progressArea').classList.remove('hidden');
        $('resultBox').classList.add('hidden');
        $('loadError').classList.add('hidden');

        setProgress(5, 'Preparando Excel...');

        state.output = await buildWorkbook();

        const buf = await state.output.xlsx.writeBuffer();

        state.outputBytes = buf;

        setProgress(100, 'Listo');

        $('resultText').textContent =
            `${state.records.length} registros preparados. El archivo XLSX incluye los códigos de barras.`;

        $('resultBox').classList.remove('hidden');

    } catch (err) {
        console.error(err);

        $('loadError').textContent =
            'No se pudo generar el Excel: ' +
            (err.message || err);

        $('loadError').classList.remove('hidden');

    } finally {
        $('exportBtn').disabled = false;
    }
});

$('downloadBtn').addEventListener('click', () => {
    if (!state.outputBytes) return;

    const blob = new Blob(
        [state.outputBytes],
        {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    );

    const a = document.createElement('a');

    a.href = URL.createObjectURL(blob);
    a.download = 'excel_unificado_v1.2.1.xlsx';
    a.click();

    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
});
