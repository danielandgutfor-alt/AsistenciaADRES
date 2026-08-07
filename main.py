from config import INPUT_DIR

from core.parser import Parser

from core.excel import ExcelExporter

from core.zip_manager import ZipManager


def main():

    archivo = INPUT_DIR / "Cronograma - Consolidado de asistencia.xlsx"

    parser = Parser(archivo)

    df = parser.obtener_dataframe()

    exportador = ExcelExporter("salida")

    exportador.exportar(df)

    zip_manager = ZipManager("salida")

    archivo_zip = zip_manager.comprimir_dependencias()

    print()

    print("ZIP generado:")

    print(archivo_zip)

    print()

    print("=" * 80)
    print("DATAFRAME")
    print("=" * 80)

    print(df.head())

    print()

    print(df.info())

    print()

    print("Total registros:", len(df))


if __name__ == "__main__":
    main()