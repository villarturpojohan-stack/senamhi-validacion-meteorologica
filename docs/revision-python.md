# Revisión de la versión histórica de Python

Proyecto de aprendizaje en Python para consolidar descargas CSV diarias y marcar datos que requieren revisión. Utiliza pandas y openpyxl. El código original se conserva en `python/senamhi_consolidar.py`.

## Qué hacía el ejercicio

El programa recorre una carpeta por estación y sus subcarpetas de años. Lee las tablas descargadas, interpreta fechas, convierte valores numéricos y genera un Excel con temperaturas máxima y mínima, humedad relativa, precipitación y observaciones por variable.

Marca S/D, trazas de precipitación, valores menores o iguales a −900, precipitación negativa, humedad fuera de 0–100 % y temperatura máxima menor que mínima. Deja vacíos los valores inválidos y las trazas; conserva las temperaturas cuando detecta una inconsistencia entre máxima y mínima para permitir revisión manual. La leyenda de los archivos define T como precipitación inferior a 0,1 mm/día. Una traza no se convierte en cero.

Ordena las fechas, conserva la primera aparición de cada fecha y comunica en consola cuántas fechas faltan dentro del intervalo observado. Los días ausentes no se incorporan como filas al Excel.

## Hallazgo comprobado

El archivo `ejemplo/FONAGRO/2018/table (1).csv`, línea 17, contiene:

```text
2018-10-10,24.2,17.4,84.4,-887.7
```

Corresponde a FONAGRO (CHINCHA), código 000791. La precipitación registrada es −887,7 mm/día y el programa la marca como negativa. El archivo permite identificar un valor inválido, pero no determinar si se originó en el sensor, la transcripción, una codificación o el procesamiento posterior.

## Revisión del material conservado

Revisión local realizada el 23 de noviembre de 2026. Los conteos corresponden a los archivos disponibles, no a todo el archivo histórico de SENAMHI.

| Carpeta local | CSV | Filas de datos | Fechas únicas | Filas adicionales con fecha repetida | Fechas con registros validados distintos |
|---|---:|---:|---:|---:|---:|
| Estación FONAGRO (CHINCHA) | 106 | 3197 | 3169 | 28 | 14 |
| Socsi CANETE | 5 | 154 | 151 | 3 | 2 |

La carpeta FONAGRO contiene fechas entre el 1 de enero de 2017 y el 1 de agosto de 2026, con 331 días sin fila dentro de ese intervalo. Estos días ausentes se distinguen de las filas existentes que contienen S/D.

Los cinco CSV guardados en `Socsi CANETE` también declaran FONAGRO (CHINCHA), código 000791, en su metadata. Por ello no se presentan como observaciones de Socsi.

## Limitaciones del código original

- Detecta el código 000791, pero no extrae el nombre en estos archivos porque `Estación : FONAGRO (CHINCHA)` aparece en una sola celda. El lector busca el nombre en la celda siguiente. El nombre del Excel procede entonces de la carpeta.
- Procesa por carpeta y no verifica que los archivos correspondan a una misma estación. La identificación y agrupación por metadata deben mejorarse antes de usarlo con descargas mezcladas.
- `drop_duplicates(subset="Fecha", keep="first")` omite registros repetidos sin guardar un listado de conflictos. No puede afirmarse que el Excel conserva todos los registros originales.
- Las columnas meteorológicas se interpretan por posición. Se requiere el formato de los CSV incluidos como ejemplo.
- La comprobación de temperatura se limita a valores centinela y a la comparación entre máxima y mínima. No es un control climatológico completo.
- La ruta de entrada está fijada en `BASE_DIR` y debe editarse antes de ejecutar el original.

## Reproducir un ejemplo

Se incluyen cuatro descargas de 2018 con su metadata y leyenda originales. Para probar el programa:

1. Instalar Python y las dependencias con `python -m pip install -r requirements.txt`.
2. Abrir `python/senamhi_consolidar.py` y ajustar `BASE_DIR` a la ruta absoluta de la carpeta `ejemplo` de este proyecto.
3. Ejecutar `python python/senamhi_consolidar.py` desde la carpeta del proyecto.
4. Revisar el Excel generado dentro de `ejemplo/_consolidado` y los mensajes de la consola.

La revisión ejecutó las funciones de lectura y validación contra los 111 CSV disponibles. No se volvió a generar el Excel histórico ni se realizó una descarga nueva. Entorno de revisión: Python 3.14.4, pandas 3.0.5 y openpyxl 3.1.5.

## Procedencia y alcance

Los CSV conservados identifican como fuente a SENAMHI / DRD e incluyen la advertencia «Datos sin control de calidad». Se mantienen los originales para poder contrastar las observaciones. Este ejercicio automatiza comprobaciones iniciales y requiere revisión especializada para resolver conflictos, evaluar series o decidir imputaciones.

## Mejoras pendientes

Leer el nombre dentro de la misma celda, agrupar por código de estación, conservar archivo y fila de origen, exportar todos los duplicados con sus conflictos y registrar las fechas ausentes en una hoja de revisión. Son mejoras propuestas; no se atribuyen a la versión histórica.
