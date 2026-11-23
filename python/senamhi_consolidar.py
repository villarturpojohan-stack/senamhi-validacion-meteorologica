"""
SENAMHI - Consolidacion y validacion automatica de datos de estaciones
=======================================================================
Que hace:
  1. Busca, dentro de BASE_DIR, una subcarpeta por cada estacion (el nombre
     de la carpeta no importa, pero se recomienda que sea el nombre de la
     estacion, ej: "Estacion FONAGRO (CHINCHA)").
  2. Dentro de cada carpeta de estacion, busca TODOS los .csv que haya,
     sin importar si estan sueltos o repartidos en subcarpetas por anio
     (busqueda recursiva).
  3. Lee cada CSV tal cual lo entrega el SENAMHI (con toda la metadata
     arriba), detecta donde empiezan los datos, y valida cada valor:
       - S/D                    -> se deja vacio, observacion "S/D"
       - T                      -> se deja vacio, observacion "T (traza...)"
       - <= -900 (tipo -999)    -> se deja vacio, observacion "Valor invalido"
       - Precipitacion negativa -> se deja vacio, observacion correspondiente
       - Humedad fuera de 0-100 -> se deja vacio, observacion correspondiente
  4. Junta todos los anios de esa estacion en una sola tabla continua,
     ordenada por fecha, sin fechas duplicadas, y revisa si faltan dias
     en todo el rango.
  5. Guarda un Excel por estacion (en la carpeta OUTPUT_DIR) con el mismo
     formato que ya usabas: datos juntos, columna en blanco, observaciones
     agrupadas, y colores (rojo = error real, celeste = traza).

Como usarlo:
  1. Instala las librerias necesarias (una sola vez):
       pip install pandas openpyxl
  2. Cambia la variable BASE_DIR de aqui abajo por la ruta real de tu
     carpeta "intento de scrip estaciones senamhi" (o como la hayas llamado).
  3. Corre el script:
       python senamhi_consolidar.py
  4. Revisa la carpeta "_consolidado" que se crea dentro de BASE_DIR:
     ahi va a aparecer un Excel por cada estacion.
"""

import os
import glob
import csv
from datetime import timedelta

import pandas as pd
from openpyxl.styles import Font, PatternFill

# ============================================================
# AJUSTA ESTA RUTA a tu carpeta real antes de correr el script
# ============================================================
BASE_DIR = r"C:\Users\villa\Documents\Cuenca_Ica_proyecto hidro qswat y pisco\intento de scrip estaciones senamhi\estaciones convencionales con recepción de datos en tiempo real"

OUTPUT_DIR_NAME = "_consolidado"

COLOR_ERROR = "FFC7CE"   # rojo clarito: error real
COLOR_TRAZA = "BDD7EE"   # celeste clarito: solo informativo (traza)


def validar_valor(valor_crudo, es_precipitacion=False, es_humedad=False):
    """
    Devuelve (valor_limpio, observacion).
    Si el valor esta bien: valor_limpio = numero, observacion = "".
    Si el valor tiene algun problema: valor_limpio = None (queda vacio en
    la tabla final) y observacion explica que traia el dato original.
    """
    txt = "" if valor_crudo is None else str(valor_crudo).strip()

    if txt.upper() == "S/D" or txt == "" or txt.lower() == "nan":
        return None, "S/D"

    if txt.upper() == "T":
        return None, "T (traza < 0.1 mm/dia)"

    try:
        numero = float(txt.replace(",", "."))
    except ValueError:
        return None, f"Valor no numerico: {txt}"

    if numero <= -900:
        return None, f"Valor invalido: {txt}"
    if es_precipitacion and numero < 0:
        return None, f"Precipitacion negativa: {txt}"
    if es_humedad and (numero < 0 or numero > 100):
        return None, f"Fuera de rango (0-100): {txt}"

    return numero, ""


def leer_archivo_senamhi(ruta_csv):
    """
    Lee un CSV crudo del SENAMHI (con metadata arriba) y devuelve
    (nombre_estacion, codigo_estacion, lista_de_registros_validados).
    """
    with open(ruta_csv, encoding="utf-8-sig", newline="") as f:
        filas_crudas = list(csv.reader(f))

    estacion, codigo = "", ""
    for fila in filas_crudas[:15]:
        for j, celda in enumerate(fila):
            if "Estaci" in celda:
                if j + 1 < len(fila) and fila[j + 1].strip():
                    estacion = fila[j + 1].strip()
                elif j + 2 < len(fila):
                    estacion = fila[j + 2].strip()
            if "digo" in celda:
                if j + 1 < len(fila) and fila[j + 1].strip():
                    codigo = fila[j + 1].strip()

    registros = []
    for fila in filas_crudas:
        if not fila or not fila[0].strip():
            continue
        try:
            fecha = pd.to_datetime(fila[0].strip()).date()
        except Exception:
            continue  # no es una fila de datos (es metadata o encabezado)

        tmax_raw = fila[1] if len(fila) > 1 else ""
        tmin_raw = fila[2] if len(fila) > 2 else ""
        hr_raw = fila[3] if len(fila) > 3 else ""
        precip_raw = fila[4] if len(fila) > 4 else ""

        tmax, obs_tmax = validar_valor(tmax_raw)
        tmin, obs_tmin = validar_valor(tmin_raw)
        hr, obs_hr = validar_valor(hr_raw, es_humedad=True)
        precip, obs_precip = validar_valor(precip_raw, es_precipitacion=True)

        # Chequeo logico: la temperatura maxima nunca deberia ser menor que
        # la minima del mismo dia. No se vacian los valores (no sabemos cual
        # de los dos esta mal), solo se marca para que lo revises a mano.
        if tmax is not None and tmin is not None and tmax < tmin:
            nota = "Temp_Max menor que Temp_Min"
            obs_tmax = f"{obs_tmax} + {nota}" if obs_tmax else nota
            obs_tmin = f"{obs_tmin} + {nota}" if obs_tmin else nota

        registros.append({
            "Fecha": fecha,
            "Anio": fecha.year,
            "Mes": fecha.month,
            "Dia": fecha.day,
            "Temp_Max": tmax, "Obs_Temp_Max": obs_tmax,
            "Temp_Min": tmin, "Obs_Temp_Min": obs_tmin,
            "Humedad_Relativa": hr, "Obs_Humedad": obs_hr,
            "Precipitacion_Total": precip, "Obs_Precipitacion": obs_precip,
        })

    return estacion, codigo, registros


def guardar_excel(df, ruta_salida):
    columnas = ["Fecha", "Anio", "Mes", "Dia", "Temp_Max", "Temp_Min",
                "Humedad_Relativa", "Precipitacion_Total",
                "Obs_Temp_Max", "Obs_Temp_Min", "Obs_Humedad", "Obs_Precipitacion"]
    df_salida = df.reindex(columns=columnas)
    df_salida.insert(8, "", "")  # columna en blanco de separacion (posicion 9)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df_salida.to_excel(writer, index=False, sheet_name="Maestro")
        ws = writer.sheets["Maestro"]

        for celda in ws[1]:
            celda.font = Font(bold=True)

        for fila_idx in range(2, ws.max_row + 1):
            ws.cell(row=fila_idx, column=1).number_format = "dd/mm/yyyy"

        fill_error = PatternFill("solid", fgColor=COLOR_ERROR)
        fill_traza = PatternFill("solid", fgColor=COLOR_TRAZA)
        columnas_obs = [10, 11, 12, 13]  # J,K,L,M (Obs_Temp_Max .. Obs_Precipitacion)

        for fila_idx in range(2, ws.max_row + 1):
            observaciones = [ws.cell(row=fila_idx, column=c).value or "" for c in columnas_obs]
            hay_error = any(o and not o.startswith("T (") for o in observaciones)
            hay_traza = any(o.startswith("T (") for o in observaciones)
            fill = fill_error if hay_error else (fill_traza if hay_traza else None)
            if fill:
                for col in range(1, 14):
                    ws.cell(row=fila_idx, column=col).fill = fill

        ws.column_dimensions["A"].width = 12
        for letra in "BCDEFGHIJKLM":
            ws.column_dimensions[letra].width = 16


def procesar_estacion(carpeta_estacion, carpeta_salida):
    archivos = glob.glob(os.path.join(carpeta_estacion, "**", "*.csv"), recursive=True)
    if not archivos:
        print("  (sin archivos .csv en esta carpeta)")
        return

    todos_los_registros = []
    estacion_nombre, estacion_codigo = "", ""

    for archivo in sorted(archivos):
        est, cod, registros = leer_archivo_senamhi(archivo)
        if est:
            estacion_nombre = est
        if cod:
            estacion_codigo = cod
        todos_los_registros.extend(registros)
        print(f"  leido: {os.path.relpath(archivo, carpeta_estacion)} -> {len(registros)} filas")

    if not todos_los_registros:
        print("  no se encontraron filas de datos validas en ningun archivo.")
        return

    df = pd.DataFrame(todos_los_registros)
    filas_antes = len(df)
    df = df.drop_duplicates(subset="Fecha", keep="first")
    duplicados = filas_antes - len(df)
    df = df.sort_values("Fecha").reset_index(drop=True)

    fecha_min, fecha_max = df["Fecha"].min(), df["Fecha"].max()
    todas_las_fechas = {fecha_min + timedelta(days=d) for d in range((fecha_max - fecha_min).days + 1)}
    faltantes = sorted(todas_las_fechas - set(df["Fecha"]))

    inconsistentes = df[
        df["Obs_Temp_Max"].fillna("").str.contains("Temp_Max menor que Temp_Min")
        | df["Obs_Temp_Min"].fillna("").str.contains("Temp_Max menor que Temp_Min")
    ]

    nombre_base = estacion_nombre if estacion_nombre else os.path.basename(carpeta_estacion)
    nombre_archivo = "".join(c for c in nombre_base if c not in '<>:"/\\|?*').strip()
    ruta_salida = os.path.join(carpeta_salida, f"{nombre_archivo}.xlsx")

    guardar_excel(df, ruta_salida)

    print(f"  -> {estacion_nombre or '(sin nombre detectado)'} ({estacion_codigo or 's/c'}): "
          f"{len(df)} dias del {fecha_min.strftime('%d/%m/%Y')} al {fecha_max.strftime('%d/%m/%Y')}, "
          f"{duplicados} duplicados omitidos, {len(faltantes)} dias faltantes, "
          f"{len(inconsistentes)} dias con Temp_Max < Temp_Min")
    if faltantes:
        muestra = ", ".join(f.strftime("%d/%m/%Y") for f in faltantes[:10])
        extra = "..." if len(faltantes) > 10 else ""
        print(f"     dias faltantes (primeros 10 de {len(faltantes)}): {muestra}{extra}")
    print(f"     guardado en: {ruta_salida}")


def main():
    if not os.path.isdir(BASE_DIR):
        print(f"No encuentro la carpeta BASE_DIR: {BASE_DIR}")
        print("Edita la variable BASE_DIR al inicio del script con la ruta correcta.")
        return

    carpeta_salida = os.path.join(BASE_DIR, OUTPUT_DIR_NAME)
    os.makedirs(carpeta_salida, exist_ok=True)

    subcarpetas = [
        os.path.join(BASE_DIR, d) for d in sorted(os.listdir(BASE_DIR))
        if os.path.isdir(os.path.join(BASE_DIR, d)) and not d.startswith("_")
    ]

    if not subcarpetas:
        print("No encontre subcarpetas de estaciones dentro de BASE_DIR.")
        return

    for carpeta in subcarpetas:
        print(f"\nProcesando carpeta: {os.path.basename(carpeta)}")
        procesar_estacion(carpeta, carpeta_salida)

    print(f"\nListo. Archivos consolidados guardados en: {carpeta_salida}")


if __name__ == "__main__":
    main()
