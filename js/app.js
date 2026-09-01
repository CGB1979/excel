```javascript
const state = {
    files: [],
    records: [],
    headers: [],
    output: null,
    outputBytes: null
};

const $ = id => document.getElementById(id);

/*
 * Normaliza un texto para poder compararlo:
 * - Convierte a string
 * - Elimina espacios al principio/final
 * - Pasa a minúsculas
 * - Elimina acentos
 */
function norm(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');
}

/*
 * Busca una columna por coincidencia exacta.
 * Se mantiene por compatibilidad con otras partes del programa.
 */
function findCol(headers, names) {
    const wanted = names.map(norm);
    return headers.findIndex(h => wanted.includes(norm(h)));
}

/*
 * Busca específicamente la columna de CHASIS.
 *
 * Ejemplos que reconoce:
 *   Chasis
 *   CHASIS
 *   Numero de Chasis
 *   Número de Chasis
 *   Nro de Chasis
 *   N° Chasis
 *   Chasis Vehiculo
 *   Identificacion de Chasis
 *
 * La comparación ignora mayúsculas, minúsculas y acentos.
 */
function findChasisCol(headers) {
    const normalizedHeaders = headers.map(norm);

    // 1. Primero buscamos "chasis" exactamente.
    const exactIndex = normalizedHeaders.findIndex(h => h === 'chasis');

    if (exactIndex >= 0) {
        return exactIndex;
    }

    // 2. Después buscamos cualquier encabezado que CONTENGA "chasis".
    const candidates = [];

    normalizedHeaders.forEach((header, index) => {
        if (header.includes('chasis')) {
            candidates.push({
                index,
                header
            });
        }
    });

    if (!candidates.length) {
        return -1;
    }

    // 3. Si hay varias coincidencias, priorizamos:
    //    a) encabezados que terminan en "chasis"
    const endsWithChasis = candidates.find(item =>
        item.header.endsWith('chasis')
    );

    if (endsWithChasis) {
        return endsWithChasis.index;
    }

    // 4. Si no, utilizamos la primera coincidencia.
    return candidates[0].index;
}

function setProgress(n, text) {
    $('progressBar').style.width = `${n}%`;
    $('progressText').textContent = text;
}

function cellValue(v) {
    if (v == null) return '';

    if (typeof v === 'object' && v.result != null) {
        return v.result;
    }

    return String(v);
}

async function readFile(file) {
    const wb = new ExcelJS.Workbook();

    const buf = await file.arrayBuffer();

    await wb.xlsx.load(buf);

    let records = [];
    let headers = [];

    wb.eachSheet(ws => {

        /*
         * Utilizamos la primera fila de la primera hoja
         * como encabezado.
         */
        if (!headers.length) {
            const r = ws.getRow(1);

            headers = r.values
                .slice(1)
                .map(v => cellValue(v));
        }

        /*
         * Leemos los registros desde la fila 2.
         */
        for (let r = 2; r <= ws.rowCount; r++) {

            const row = ws.getRow(r);
            const vals = row.values.slice(1);

            /*
             * Ignoramos filas completamente vacías.
             */
            if (
                vals.some(
                    v => cellValue(v).trim() !== ''
                )
            ) {
                records.push({
                    values: headers.map(
                        (_, i) => cellValue(vals[i] ?? '')
                    ),
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

    box.innerHTML = state.files
        .map(f => `
            <div class="file-row">
                <div class="file-name">
                    ${escapeHtml(f.name)}
                </div>

                <div class="file-meta">
                    ${f.rows} registros
                </div>
            </div>
        `)
        .join('');

    box.classList.toggle(
        'hidden',
        !state.files.length
    );

    $('fileCount').textContent =
        state.files.length;

    $('rowCount').textContent =
        state.records.length;

    $('statusPill').textContent =
        state.files.length
            ? `${state.files.length} archivo${state.files.length === 1 ? '' : 's'}`
            : 'Sin archivos';

    $('statusPill').classList.toggle(
        'ready',
        !!state.files.length
    );

    $('exportBtn').disabled =
        !state.records.length;
}

function escapeHtml(s) {

    return String(s).replace(
        /[&<>"']/g,
        c => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[c])
    );
}

/*
 * CARGA DE ARCHIVOS
 */
$('fileInput').addEventListener(
    'change',
    async e => {

        state.files = [];
        state.records = [];
        state.headers = [];
        state.output = null;
        state.outputBytes = null;

        $('loadError').classList.add('hidden');

        try {

            for (const file of e.target.files) {

                const r = await readFile(file);

                if (!state.headers.length) {
                    state.headers = r.headers;
                }

                state.files.push({
                    name: r.name,
                    rows: r.rows
                });

                state.records.push(
                    ...r.records
                );
            }

            renderFiles();

        } catch (err) {

            console.error(err);

            $('loadError').textContent =
                'No se pudo leer el archivo Excel: ' +
                (err.message || err);

            $('loadError').classList.remove(
                'hidden'
            );

            state.files = [];
            state.records = [];
            state.headers = [];
            state.output = null;
            state.outputBytes = null;

            renderFiles();
        }
    }
);

/*
 * GENERACIÓN DEL CÓDIGO DE BARRAS
 *
 * Tamaño solicitado:
 *   285 px de ancho
 *   40 px de alto
 */
function barcodeDataUrl(value) {

    const c = document.createElement('canvas');

    c.width = 285;
    c.height = 40;

    JsBarcode(
        c,
        String(value),
        {
            format: 'CODE128',
            width: 1.45,
            height: 32,
            margin: 0,
            displayValue: false,
            lineColor: '#000000'
        }
    );

    return c.toDataURL('image/png');
}

/*
 * GENERACIÓN DEL EXCEL
 */
async function buildWorkbook() {

    const wb = new ExcelJS.Workbook();

    const ws = wb.addWorksheet('Datos');

    const headers = [
        ...state.headers,
        'Código de barras'
    ];

    ws.addRow(headers);

    /*
     * Autofiltro.
     */
    ws.autoFilter = {
        from: 1,
        to: headers.length,
        fromRow: 1,
        toRow: Math.max(
            1,
            state.records.length + 1
        )
    };

    /*
     * BUSQUEDA ROBUSTA DEL CHASIS
     */
    const chasisCol =
        findChasisCol(state.headers);

    if (chasisCol < 0) {

        throw new Error(
            'No se encontró ninguna columna cuyo encabezado contenga "Chasis".'
        );
    }

    /*
     * Columna del código de barras.
     */
    const barcodeCol = headers.length;

    /*
     * Ancho aproximado de la columna.
     *
     * ExcelJS utiliza unidades de caracteres,
     * no píxeles directamente.
     */
    ws.getColumn(barcodeCol).width = 38.0;

    /*
     * Encabezado.
     */
    ws.getRow(1).height = 22;

    ws.getRow(1).font = {
        bold: true
    };

    /*
     * Registros.
     */
    for (
        let i = 0;
        i < state.records.length;
        i++
    ) {

        const rec = state.records[i];

        const row = ws.addRow([
            ...rec.values,
            ''
        ]);

        /*
         * Altura aproximada para el código
         * de 40 px.
         */
        row.height = 30;

        /*
         * Obtiene el valor del CHASIS.
         */
        const chasis =
            rec.values[chasisCol] || '';

        /*
         * Genera el código únicamente
         * si hay un valor de chasis.
         */
        if (chasis) {

            const data =
                barcodeDataUrl(chasis);

            const imageId =
                wb.addImage({
                    base64: data,
                    extension: 'png'
                });

            /*
             * Imagen anclada a la celda
             * de código de barras.
             */
            ws.addImage(
                imageId,
                {
                    tl: {
                        col: barcodeCol - 1,
                        row: i + 1
                    },

                    br: {
                        col: barcodeCol,
                        row: i + 2
                    }
                }
            );
        }

        setProgress(
            Math.round(
                ((i + 1) /
                    state.records.length) *
                100
            ),
            `Generando código ${i + 1} de ${state.records.length}`
        );

        await new Promise(
            r => setTimeout(r, 0)
        );
    }

    return wb;
}

/*
 * EXPORTAR
 */
$('exportBtn').addEventListener(
    'click',
    async () => {

        try {

            $('exportBtn').disabled = true;

            $('progressArea')
                .classList
                .remove('hidden');

            $('resultBox')
                .classList
                .add('hidden');

            $('loadError')
                .classList
                .add('hidden');

            setProgress(
                5,
                'Preparando Excel...'
            );

            state.output =
                await buildWorkbook();

            const buf =
                await state.output.xlsx.writeBuffer();

            state.outputBytes = buf;

            setProgress(
                100,
                'Listo'
            );

            $('resultText').textContent =
                `${state.records.length} registros preparados. El archivo XLSX incluye los códigos de barras.`;

            $('resultBox')
                .classList
                .remove('hidden');

        } catch (err) {

            console.error(err);

            $('loadError').textContent =
                'No se pudo generar el Excel: ' +
                (err.message || err);

            $('loadError')
                .classList
                .remove('hidden');

        } finally {

            $('exportBtn').disabled = false;
        }
    }
);

/*
 * DESCARGAR
 */
$('downloadBtn').addEventListener(
    'click',
    () => {

        if (!state.outputBytes) {
            return;
        }

        const blob = new Blob(
            [state.outputBytes],
            {
                type:
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
        );

        const a =
            document.createElement('a');

        a.href =
            URL.createObjectURL(blob);

        a.download =
            'excel_unificado_v1.2.0.xlsx';

        a.click();

        setTimeout(
            () =>
                URL.revokeObjectURL(
                    a.href
                ),
            1000
        );
    }
);
```
