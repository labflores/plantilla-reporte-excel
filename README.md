# Plantilla de reporte diario en Excel con VBA

Plantilla genérica de un reporte diario que se arma automáticamente con macros de Excel. Las macros copian rangos desde varios libros de origen, generan un Excel con los datos y lo exportan a PDF.

Todos los datos de esta plantilla están vacíos o son de ejemplo. No contiene información real.

## Archivos
- `Config_POSICION_DIARIA_generico.xlsm`: libro con las macros y la hoja `Config`, donde se define qué copiar.
- `POSICION_DIARIA_generica.xlsx`: plantilla de la hoja de reporte, con bancos, cuentas y entidades genéricas.

## Cómo usarlo
1. Copiá la plantilla de origen en la ruta indicada en `Config!A6`, o cambiá esa ruta por la de tu archivo.
2. Abrí `Config_POSICION_DIARIA_generico.xlsm` y habilitá las macros. Funciona en Excel para Windows.
3. Completá la fecha en `Config!B1` y la carpeta de salida en `Config!B3`.
4. Ejecutá la macro `GenerarExcelPosicion`. Crea un Excel con la fecha en el nombre y lo deja abierto.
5. Agregá a mano la proyección en el Excel generado.
6. Ejecutá la macro `PasarAPDF` para exportar el reporte a PDF.

## Qué hace cada macro
- `GenerarExcelPosicion`: copia cada hoja y rango configurado en `Config`, convierte fórmulas en valores y guarda el Excel.
- `PasarAPDF`: exporta a PDF el Excel generado, respetando el área de impresión.

## Tecnologías
Excel, VBA.
