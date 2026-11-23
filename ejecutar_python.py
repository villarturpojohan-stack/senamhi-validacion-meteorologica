"""Lanzador del portafolio: usa las funciones históricas sin modificarlas."""
import argparse
import importlib.util
from pathlib import Path


def main():
    raiz = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="Consolidar CSV con el script histórico SENAMHI")
    parser.add_argument("--entrada", type=Path, default=raiz / "ejemplo")
    parser.add_argument("--salida", type=Path, default=raiz / "resultados")
    args = parser.parse_args()
    entrada, salida = args.entrada.resolve(), args.salida.resolve()
    if not entrada.is_dir():
        parser.error(f"No existe la entrada: {entrada}")
    if salida == entrada or entrada in salida.parents:
        parser.error("La salida debe quedar fuera de la carpeta de entrada.")
    carpetas = [p for p in sorted(entrada.iterdir())
                if p.is_dir() and not p.name.startswith("_") and list(p.rglob("*.csv"))]
    if not carpetas:
        parser.error("La entrada necesita subcarpetas de estaciones con CSV.")
    spec = importlib.util.spec_from_file_location("senamhi_historico", raiz / "python" / "senamhi_consolidar.py")
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    salida.mkdir(parents=True, exist_ok=True)
    for carpeta in carpetas:
        print(f"Procesando: {carpeta.name}")
        modulo.procesar_estacion(str(carpeta), str(salida))


if __name__ == "__main__":
    main()
