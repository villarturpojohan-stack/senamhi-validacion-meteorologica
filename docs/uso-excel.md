# Versión Excel / VBA

`excel/SENAMHI_Macro.bas` contiene el módulo VBA exportado. `excel/Scrip.xlsm` es el libro histórico con proyecto VBA incorporado y las hojas Hoja5, Config y Maestro. Se conservaron copias exactas. No se ha comprobado que el módulo exportado sea idéntico al código incrustado en el libro, ni se han ejecutado las macros durante esta revisión.

## Funcionamiento

`SENAMHI_LimpiarYConsolidar` lee la hoja activa, encuentra la primera fecha en la columna A, valida los valores y agrega registros a Maestro. Config guarda la estación activa. `SENAMHI_CambiarEstacionActiva` borra esa selección.

La tabla debe ocupar cinco columnas: fecha, temperatura máxima, temperatura mínima, humedad relativa y precipitación. La macro omite un registro si ya existe el mismo nombre de estación y fecha en Maestro. Los códigos deben conservarse como texto para no perder ceros iniciales.

## Probar en una copia de trabajo

1. Crear un libro nuevo en Excel de escritorio y guardarlo como libro habilitado para macros (.xlsm).
2. Abrir el editor VBA con Alt+F11 e importar `excel/SENAMHI_Macro.bas` mediante Archivo > Importar archivo.
3. Importar un CSV de ejemplo en una hoja comenzando en A1. Comprobar que las fechas y las cuatro variables queden en columnas separadas y que el punto decimal se interprete correctamente.
4. Con esa hoja activa, ejecutar `SENAMHI_LimpiarYConsolidar` desde Alt+F8. Si no reconoce el nombre, ingresar FONAGRO (CHINCHA) y el código 000791 del encabezado.
5. Revisar Maestro y el resumen. Para otra tabla, limpiar completamente la hoja de entrada antes de importar nuevos datos.

El libro histórico sirve como referencia. La guía anterior permite comenzar con un Maestro vacío.

## Alcance observado en el código

- Revisa S/D, T, valores ≤ −900, precipitación negativa y humedad fuera de 0–100 %.
- No incluye la comparación de temperatura máxima contra mínima que existe en Python.
- Espera el nombre de estación en una celda contigua a su etiqueta. Con los ejemplos puede necesitar selección manual.
- Omite duplicados por nombre y fecha sin comparar sus valores ni conservarlos en Maestro.
- La lectura termina en la primera fila sin fecha o cuando las fechas retroceden. Usar una tabla ordenada y sin filas vacías intermedias.
- Los días faltantes del resumen corresponden a la tabla leída en esa ejecución, no a todo el Maestro acumulado.
- La interpretación de fechas y números depende de la configuración regional y de la importación.
