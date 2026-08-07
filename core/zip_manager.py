from pathlib import Path
import zipfile


class ZipManager:

    def __init__(self, carpeta):

        self.carpeta = Path(carpeta)

    # ---------------------------------------------------------

    def comprimir_dependencias(self):

        carpeta = self.carpeta / "Dependencias"

        archivo_zip = self.carpeta / "Dependencias.zip"

        with zipfile.ZipFile(
            archivo_zip,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zipf:

            for archivo in carpeta.glob("*.xlsx"):

                zipf.write(
                    archivo,
                    arcname=archivo.name
                )

        return archivo_zip