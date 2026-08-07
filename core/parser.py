from pathlib import Path

import pandas as pd


class Parser:

    def __init__(self, archivo):

        self.archivo = Path(archivo)

        self.libro = pd.ExcelFile(self.archivo)

    # ---------------------------------------------------------

    def actividades(self):

        resultado = []

        for hoja in self.libro.sheet_names:

            if hoja == "Formato Cronograma":
                continue

            if hoja == "Datos":
                continue

            if hoja.startswith("Fechas"):
                continue

            resultado.append(hoja)

        return resultado

    # ---------------------------------------------------------

    def hoja_fechas(self, actividad):

        numero = actividad.split("-")[0].strip()

        return f"Fechas - {numero}"

    # ---------------------------------------------------------

    def leer_actividad(self, actividad):

        return pd.read_excel(
            self.archivo,
            sheet_name=actividad,
            header=11
        )

    # ---------------------------------------------------------

    def leer_fechas(self, actividad):

        return pd.read_excel(
            self.archivo,
            sheet_name=self.hoja_fechas(actividad)
        )

    # ---------------------------------------------------------

    def detectar_columnas(self, actividad):

        hoja = self.leer_actividad(actividad)

        columnas = {}

        sesiones = []

        for indice, nombre in enumerate(hoja.columns):

            texto = str(nombre).strip().upper()

            if texto == "DOCUMENTO":

                columnas["documento"] = indice

            elif "NOMBRE Y APELLIDOS" in texto:

                columnas["nombre"] = indice

            elif "TIPO DE VINCULACIÓN" in texto:

                columnas["tipo"] = indice

            elif "ÁREA A LA QUE PERTENECE" in texto:

                columnas["dependencia"] = indice

            elif "NOMBRE DEL CARGO" in texto:

                columnas["cargo"] = indice

            elif "HORAS ACUMULADAS" in texto:

                columnas["horas"] = indice

            elif "SESIONES ASISTIDAS" in texto:

                columnas["sesiones_asistidas"] = indice

            elif "% ASISTENCIA" in texto:

                columnas["porcentaje"] = indice

            else:

                try:

                    numero = int(float(texto))

                    sesiones.append((numero, indice))

                except:

                    pass

        sesiones.sort(key=lambda x: x[0])

        columnas["sesiones"] = sesiones

        return columnas

    # ---------------------------------------------------------

    def obtener_fechas(self, actividad):

        hoja = self.leer_fechas(actividad)

        fechas = {}

        for _, fila in hoja.iloc[1:].iterrows():

            sesion = fila.iloc[0]

            fecha = fila.iloc[1]

            if pd.isna(sesion):
                continue

            try:
                sesion = int(sesion)
            except:
                continue

            fechas[sesion] = pd.to_datetime(fecha)

        return fechas

    # ---------------------------------------------------------

    def construir_dataframe(self):

        registros = []

        for actividad in self.actividades():

            print(f"Leyendo {actividad}")

            hoja = self.leer_actividad(actividad)

            fechas = self.obtener_fechas(actividad)

            columnas = self.detectar_columnas(actividad)

            registros.extend(

                self._leer_funcionarios(
                    actividad,
                    hoja,
                    fechas,
                    columnas
                )

            )

        return pd.DataFrame(registros)

        # ---------------------------------------------------------

       # ---------------------------------------------------------

    def _leer_funcionarios(
        self,
        actividad,
        hoja,
        fechas,
        columnas
    ):

        registros = []

        sesiones_totales = len(columnas["sesiones"])
        sesiones_realizadas = len(fechas)

        for _, fila in hoja.iterrows():

            no = fila.iloc[0]

            if pd.isna(no):
                continue

            if str(no).strip().upper().startswith("OBSERVACIONES"):
                break

            nombre = fila.iloc[columnas["nombre"]]

            if pd.isna(nombre):
                continue

            documento = fila.iloc[columnas["documento"]]

            if pd.isna(documento):
                continue

            dependencia = fila.iloc[columnas["dependencia"]]
            cargo = fila.iloc[columnas["cargo"]]
            tipo = fila.iloc[columnas["tipo"]]

            sesiones_asistidas = self._obtener_sesiones_asistidas(
                fila,
                columnas
            )

            porcentaje = self._obtener_porcentaje(
                fila,
                columnas
            )

            detalle = self._detalle_inasistencias(
                fila,
                fechas,
                columnas
            )

            sesiones_falladas = max(
                sesiones_realizadas - sesiones_asistidas,
                0
            )
            try:
                documento_limpio = str(int(float(documento)))
            except (ValueError, TypeError):
                documento_limpio = str(documento).strip()
            registros.append(
                {
                    "actividad": actividad,
                    "documento": documento_limpio,
                    "nombre": str(nombre),
                    "dependencia": str(dependencia),
                    "cargo": str(cargo),
                    "tipo_vinculacion": str(tipo),
                    "sesiones_totales": sesiones_totales,
                    "sesiones_realizadas": sesiones_realizadas,
                    "sesiones_asistidas": sesiones_asistidas,
                    "sesiones_falladas": sesiones_falladas,
                    "detalle_inasistencias": detalle,
                    "porcentaje_asistencia": porcentaje,
                }
            )

        return registros
        # ---------------------------------------------------------

    def _obtener_sesiones_asistidas(
        self,
        fila,
        columnas
    ):

        indice = columnas.get("sesiones_asistidas")

        if indice is None:
            return 0

        valor = fila.iloc[indice]

        if pd.isna(valor):
            return 0

        try:
            return int(valor)
        except Exception:
            return 0
            # ---------------------------------------------------------

    def _obtener_porcentaje(
        self,
        fila,
        columnas
    ):

        indice = columnas.get("porcentaje")

        if indice is None:
            return 0.0

        valor = fila.iloc[indice]

        if pd.isna(valor):
            return 0.0

        try:

            if isinstance(valor, str):

                valor = valor.replace("%", "").replace(",", ".").strip()

            return float(valor)

        except Exception:

            return 0.0
            # ---------------------------------------------------------

    def _detalle_inasistencias(
        self,
        fila,
        fechas,
        columnas
    ):

        detalle = []

        for sesion, columna in columnas["sesiones"]:

            #
            # Solo revisar sesiones que ya existen
            #
            if sesion not in fechas:
                continue

            valor = fila.iloc[columna]

            #
            # Celda vacía = no asistió
            #
            if pd.isna(valor):

                fecha = fechas[sesion]

                detalle.append(
                    f"Sesión #{sesion} ({fecha.strftime('%d/%m/%Y')})"
                )

        return "\n".join(detalle)
        # ---------------------------------------------------------

    def obtener_dataframe(self):

        return self.construir_dataframe() 