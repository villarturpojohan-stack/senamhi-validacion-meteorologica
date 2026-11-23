# Probar Python con las descargas de FONAGRO

1. En GitHub, seleccionar **Code > Download ZIP**. Extraer el ZIP completo.
2. Dentro de la carpeta extraída, crear una carpeta llamada `datos_locales`.
3. Copiar dentro de ella la carpeta completa `Estación  FONAGRO (CHINCHA)`, con sus años y archivos CSV. No es necesario cambiarle el nombre.

La estructura debe quedar así:

```text
senamhi-validacion-meteorologica-main/
    ejecutar_python.py
    requirements.txt
    python/
        senamhi_consolidar.py
    datos_locales/
        Estación  FONAGRO (CHINCHA)/
            2017/
                table.csv
            2018/
                table.csv
                table (1).csv
            ...
```

4. Abrir una terminal en la carpeta que contiene `ejecutar_python.py` y ejecutar:

```shell
python -m pip install -r requirements.txt
python ejecutar_python.py --entrada "datos_locales" --salida "resultados_fonagro"
```

Si Windows reconoce `py` en lugar de `python`, usar `py` al comienzo de ambos comandos.

5. Abrir el Excel generado dentro de `resultados_fonagro`. Revisar las columnas de observaciones y el resumen de la terminal. Cerrar ese Excel antes de volver a ejecutar el script, porque la nueva ejecución escribe sobre el archivo de salida del mismo nombre.

## Qué revisar

- El registro de 10/10/2018 debe tener la precipitación vacía y la observación `Precipitacion negativa: -887.7`, si se incluyó la descarga correspondiente.
- S/D y T deben seguir distinguiéndose en observaciones. Una traza queda vacía en la columna numérica; no representa precipitación cero.
- Los días ausentes y los duplicados omitidos se indican en la terminal. La versión histórica conserva la primera fila de cada fecha y no resuelve conflictos entre duplicados.

Con los 106 CSV de FONAGRO revisados al preparar el proyecto, el lector encontró 3197 registros y 3169 fechas únicas. Esos conteos pueden cambiar si se utiliza otra selección de descargas. Con los cuatro CSV pequeños incluidos en `ejemplo`, la prueba generó 103 fechas únicas.

## CSV y Excel

Los archivos `table.csv` pueden mostrar el icono de Excel y abrirse con ese programa, pero siguen siendo CSV. Son la entrada esperada. Si se guardaron como `.xlsx`, el programa histórico no los leerá: utilizar las descargas CSV originales con la metadata del SENAMHI.

El parámetro `--entrada` apunta a `datos_locales`, la carpeta que contiene las estaciones, no directamente a una carpeta de año. Las carpetas locales y los resultados de esta prueba están excluidos de Git mediante `.gitignore`.
