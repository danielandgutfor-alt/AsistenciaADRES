from pathlib import Path

from config import INPUT_DIR
from core.lector_cronograma import LectorCronograma


def main():

    print("=" * 60)
    print("        ASISTENCIA ADRES")
    print("=" * 60)

    archivo = INPUT_DIR / "Cronograma - Consolidado de asistencia.xlsx"

    lector = LectorCronograma(archivo)

    lector.cargar()

    hojas = lector.obtener_hojas()

    print()

    print(f"Se encontraron {len(hojas)} hojas:\n")

    for hoja in hojas:

        print("•", hoja)


if __name__ == "__main__":

    main()