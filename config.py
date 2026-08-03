"""
Configuración general del proyecto AsistenciaADRES.
"""

from pathlib import Path


# ===========================
# RUTAS DEL PROYECTO
# ===========================

ROOT_DIR = Path(__file__).parent

DATA_DIR = ROOT_DIR / "data"
INPUT_DIR = DATA_DIR / "entrada"
OUTPUT_DIR = DATA_DIR / "salida"
TEMP_DIR = DATA_DIR / "temp"

TEMPLATES_DIR = ROOT_DIR / "templates"
LOG_DIR = ROOT_DIR / "logs"


# ===========================
# ARCHIVOS
# ===========================

PLANTILLA_CONSOLIDADO = TEMPLATES_DIR / "Plantilla Consolidado.xlsx"


# ===========================
# HOJAS QUE NO SON CURSOS
# (las ampliaremos cuando conozcamos la estructura completa)
# ===========================

HOJAS_EXCLUIDAS = [
    "RESUMEN",
    "CRONOGRAMA",
]


# ===========================
# EXTENSIONES
# ===========================

EXTENSION_EXCEL = ".xlsx"

EXTENSION_ZIP = ".zip"