# Convertidor de Excel

Versión 1.0.0.h. Unifica archivos Excel y genera códigos de barras CODE128 para cada fila con Chasis.

La columna Código de barras se configura a 285 px de ancho y 40 px de alto. Las imágenes se incrustan directamente en el OOXML del XLSX, evitando las limitaciones de anclaje de ExcelJS 4.4.0.

Importante: un XLSX estándar no permite mostrar/ocultar una imagen en función de la fila seleccionada. Para ese comportamiento hace falta VBA/macros en un XLSM.
