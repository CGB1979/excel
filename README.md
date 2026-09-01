# Excel Main v1.2.1

Esta versión parte de la base visual v1.1.0 y agrega carga, lectura, compilación y exportación XLSX con CODE128.

## Estado
- Carga uno o varios XLSX/XLSM.
- Lee hojas y registros.
- Detecta la columna Chasis por cabecera.
- Genera CODE128 como PNG de 285 x 40 px.
- Inserta las imágenes en la columna Código de barras.
- Exporta XLSX.

## VBA / XLSM
La documentación oficial de SheetJS y XlsxWriter confirma que VBA no puede vivir en XLSX: se almacena en `xl/vbaProject.bin` dentro de XLSM. Por eso la aplicación todavía no fabrica por sí sola el proyecto VBA binario. El código preparado está en `VBA/ThisWorkbook.txt` y debe incorporarse a un XLSM real.

Para la versión final con selección de fila, la macro debe estar en `ThisWorkbook`. Las imágenes deben llamarse `BARCODE_ROW_<fila>` y quedar ocultas por defecto. El evento `Workbook_SheetSelectionChange` mostrará únicamente la imagen de la fila seleccionada.
