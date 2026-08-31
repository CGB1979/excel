# Convertidor de Excel — 1.0.0.c

Proyecto independiente. Selecciona uno o varios Excel y genera un XLSX unificado. Las columnas se mapean por cabecera, no por posición. El Chasis se configura en `js/configuracionColumnas.js`. Los duplicados con ubicación completa idéntica se unifican; si difiere Playa, Bloque, Carril, Posición o Ubicación, se conservan y se resaltan en naranja suave. `Información` se conserva y se combina al unificar. El resultado incluye autofiltro y códigos CODE128 como imágenes embebidas.

Dependencias de navegador: SheetJS, ExcelJS y JsBarcode vía jsDelivr.
