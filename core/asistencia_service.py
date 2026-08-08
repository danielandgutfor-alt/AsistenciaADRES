from core.parser import Parser
from core.excel import ExcelExporter
from core.zip_manager import ZipManager


class AsistenciaService:

    def procesar(
        self,
        archivo,
        carpeta_salida
    ):

        parser = Parser(archivo)

        df = parser.obtener_dataframe()

        exportador = ExcelExporter(carpeta_salida)

        exportador.exportar(df)

        zip_manager = ZipManager(carpeta_salida)

        zip_manager.comprimir_dependencias()

        return {
            "registros": len(df),
            "dependencias": df["dependencia"].nunique()
        }