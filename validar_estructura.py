"""
Valida un archivo ya transformado a formato SIGA contra las reglas
de la hoja ESTRUCTURA (Estructura_Migracion_Patrimonio.xlsx) y genera
un reporte de errores (Reporte_Validacion_SIGA.xlsx) con:

  - RESUMEN: cuántos errores hay por campo y por tipo de error.
  - DETALLE: fila por fila qué está mal (limitado a MAX_DETALLE filas
    para no generar un archivo gigante; el resto queda contado en RESUMEN).

Uso:
    python validar_estructura.py archivo_siga.xlsx [--estructura ruta] [--salida ruta]

También se puede importar y usar programáticamente:
    from validar_estructura import validar_archivo
    resumen_df, detalle_df = validar_archivo("Migracion_SIGA_Prueba_SEDE.xlsx")
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CARPETA = Path(__file__).parent
ESTRUCTURA_DEFAULT = CARPETA / "Estructura_Migracion_Patrimonio.xlsx"
MAX_DETALLE = 20000  # tope de filas de error a listar en detalle


# ============================================================
# 1) CARGAR REGLAS DESDE LA HOJA ESTRUCTURA
# ============================================================

def cargar_reglas(archivo_estructura=ESTRUCTURA_DEFAULT):
    df = pd.read_excel(archivo_estructura, sheet_name="ESTRUCTURA", header=2)
    df = df.dropna(subset=["PARÁMETRO"])

    reglas = {}
    for _, fila in df.iterrows():
        param = str(fila["PARÁMETRO"]).strip()
        tipo = str(fila["TIPO DE DATO"]).strip().upper() if pd.notna(fila["TIPO DE DATO"]) else ""
        valor_max = fila["VALOR MÁXIMO"]
        condicion = str(fila["CONDICIÓN"]).strip().upper()

        # Normaliza "CARÁCTER" / "CARACTER" a un solo valor
        if "CARACT" in tipo:
            tipo = "CARACTER"
        elif "NUM" in tipo:
            tipo = "NUMERICO"
        elif "FECHA" in tipo:
            tipo = "FECHA"

        obligatorio_base = condicion.startswith("OBLIGATORIO") and "NO OBLIGATORIO" not in condicion
        # Caso especial: "OBLIGATORIO / NO OBLIGATORIO" (depende de FLAG_ITEM_ESNI)
        condicional = condicion == "OBLIGATORIO / NO OBLIGATORIO"

        reglas[param] = {
            "tipo": tipo,
            "valor_maximo": valor_max,
            "obligatorio_base": obligatorio_base,
            "condicional": condicional,
        }
    return reglas


# ============================================================
# 2) REGLAS CONDICIONALES (dependen de otros campos de la misma fila)
#    Ver columna DESCRIPCIÓN de ESTRUCTURA para el detalle de cada una.
# ============================================================

def es_obligatorio(campo, fila, reglas):
    tipo_doc_adq = str(fila.get("TIPO_DOC_ADQUISICION", "")).strip()
    flag_esni = str(fila.get("FLAG_ITEM_ESNI", "")).strip().upper()
    flag_etiqueta = str(fila.get("FLAG_ETIQUETA", "")).strip().upper()

    # Grupo NEA (045) vs O/C-Guía (031)
    if campo in ("TIPO_MOV_NEA", "TIPO_TRAN_NEA", "NRO_DOCUMENTO", "VALOR_NEA", "FECHA_NEA"):
        return tipo_doc_adq == "045"

    if campo in ("FECHA_COMPRA", "VALOR_COMPRA", "NRO_ORDEN"):
        return tipo_doc_adq == "031"

    # NRO_PECOSA: obligatorio salvo que la adquisición ya sea O/C o NEA
    if campo == "NRO_PECOSA":
        return tipo_doc_adq not in ("031", "045")

    # ESNI: MODELO / NRO_SERIE / COD_MODELO
    if campo in ("MODELO", "NRO_SERIE", "COD_MODELO"):
        return flag_esni == "S"

    # Etiquetado
    if campo == "FECHA_ETIQUET":
        return flag_etiqueta == "S"

    return reglas[campo]["obligatorio_base"]


# ============================================================
# 3) VALIDACIÓN DE UN VALOR SEGÚN SU TIPO
# ============================================================

def _vacio(valor):
    if valor is None:
        return True
    if isinstance(valor, float) and pd.isna(valor):
        return True
    if pd.isna(valor):
        return True
    if str(valor).strip() == "":
        return True
    return False


def _max_digitos(valor_maximo):
    """'16,6' -> 16+6=22 dígitos en total; '5' -> 5."""
    s = str(valor_maximo)
    partes = re.findall(r"\d+", s)
    if not partes:
        return None
    return sum(int(p) for p in partes)


def validar_valor(campo, valor, tipo, valor_maximo):
    """Devuelve None si está OK, o un mensaje de error si no."""

    if tipo == "CARACTER":
        texto = str(valor)
        maximo = _max_digitos(valor_maximo)
        if maximo and len(texto) > maximo:
            return f"excede longitud máxima ({len(texto)} > {maximo})"
        return None

    if tipo == "NUMERICO":
        texto = str(valor).strip()
        # Acepta enteros, decimales con punto, y signo
        if not re.fullmatch(r"-?\d+(\.\d+)?", texto):
            return f"no es numérico ('{valor}')"
        maximo = _max_digitos(valor_maximo)
        if maximo:
            solo_digitos = re.sub(r"[^\d]", "", texto)
            if len(solo_digitos) > maximo:
                return f"excede dígitos máximos ({len(solo_digitos)} > {maximo})"
        return None

    if tipo == "FECHA":
        fecha = pd.to_datetime(valor, errors="coerce")
        if pd.isna(fecha):
            return f"fecha inválida ('{valor}')"
        return None

    return None


# ============================================================
# 4) VALIDAR UN DATAFRAME COMPLETO
# ============================================================

def validar_dataframe(df_siga, reglas):
    errores = []

    columnas_a_validar = [c for c in df_siga.columns if c in reglas]

    for idx, fila in df_siga.iterrows():
        secuencia = fila.get("SECUENCIA", idx + 1)
        codigo_activo = fila.get("CODIGO_ACTIVO", "")

        for campo in columnas_a_validar:
            valor = fila[campo]
            regla = reglas[campo]
            obligatorio = es_obligatorio(campo, fila, reglas)

            if _vacio(valor):
                if obligatorio:
                    errores.append({
                        "SECUENCIA": secuencia,
                        "CODIGO_ACTIVO": codigo_activo,
                        "CAMPO": campo,
                        "VALOR": "",
                        "TIPO_ERROR": "OBLIGATORIO_VACIO",
                        "MENSAJE": "Campo obligatorio sin valor",
                    })
                continue  # si no es obligatorio y está vacío, no hay más que validar

            mensaje = validar_valor(campo, valor, regla["tipo"], regla["valor_maximo"])
            if mensaje:
                errores.append({
                    "SECUENCIA": secuencia,
                    "CODIGO_ACTIVO": codigo_activo,
                    "CAMPO": campo,
                    "VALOR": str(valor)[:80],
                    "TIPO_ERROR": "FORMATO_INVALIDO",
                    "MENSAJE": mensaje,
                })

    detalle_df = pd.DataFrame(errores)

    if detalle_df.empty:
        resumen_df = pd.DataFrame(columns=["CAMPO", "TIPO_ERROR", "N_ERRORES"])
    else:
        resumen_df = (
            detalle_df.groupby(["CAMPO", "TIPO_ERROR"])
            .size()
            .reset_index(name="N_ERRORES")
            .sort_values("N_ERRORES", ascending=False)
        )

    return resumen_df, detalle_df


# ============================================================
# 5) ESCRIBIR EL REPORTE EN EXCEL
# ============================================================

def escribir_reporte(resumen_df, detalle_df, total_filas, ruta_salida):
    wb = Workbook()
    ws_r = wb.active
    ws_r.title = "RESUMEN"

    HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
    HEADER_FONT = Font(bold=True, color="FFFFFF")

    ws_r.append(["Total de filas evaluadas", total_filas])
    ws_r.append(["Total de errores encontrados", len(detalle_df)])
    ws_r.append(["Filas con al menos un error",
                  detalle_df["SECUENCIA"].nunique() if not detalle_df.empty else 0])
    ws_r.append([])

    ws_r.append(["CAMPO", "TIPO_ERROR", "N_ERRORES"])
    for c in range(1, 4):
        ws_r.cell(5, c).font = HEADER_FONT
        ws_r.cell(5, c).fill = HEADER_FILL

    for _, fila in resumen_df.iterrows():
        ws_r.append([fila["CAMPO"], fila["TIPO_ERROR"], int(fila["N_ERRORES"])])

    for c, ancho in zip("ABC", (30, 22, 14)):
        ws_r.column_dimensions[c].width = ancho

    ws_d = wb.create_sheet("DETALLE")
    columnas = ["SECUENCIA", "CODIGO_ACTIVO", "CAMPO", "VALOR", "TIPO_ERROR", "MENSAJE"]
    ws_d.append(columnas)
    for c in range(1, len(columnas) + 1):
        ws_d.cell(1, c).font = HEADER_FONT
        ws_d.cell(1, c).fill = HEADER_FILL
    ws_d.freeze_panes = "A2"

    detalle_recortado = detalle_df.head(MAX_DETALLE)
    for _, fila in detalle_recortado.iterrows():
        ws_d.append([fila[c] for c in columnas])

    if len(detalle_df) > MAX_DETALLE:
        ws_d.append([])
        ws_d.append([f"... se truncó el detalle en {MAX_DETALLE} filas. "
                     f"Quedan {len(detalle_df) - MAX_DETALLE} errores adicionales sin listar; "
                     "revisa RESUMEN para verlos agrupados por campo."])

    for c, letra in enumerate(get_column_letter(i) for i in range(1, len(columnas) + 1)):
        ws_d.column_dimensions[letra].width = 22

    wb.save(ruta_salida)


# ============================================================
# 6) FUNCIÓN PRINCIPAL
# ============================================================

def validar_archivo(ruta_siga, archivo_estructura=ESTRUCTURA_DEFAULT,
                     hoja_siga=0, ruta_salida=None):
    print(f"Cargando reglas desde {archivo_estructura} ...")
    reglas = cargar_reglas(archivo_estructura)

    print(f"Leyendo archivo a validar: {ruta_siga} ...")
    # dtype=str es clave: al leer sin esto, pandas auto-convierte columnas de
    # texto cuyo contenido parece numérico (ej. '04', '031', '046') a int/float
    # y se pierden los ceros a la izquierda, aunque el .xlsx los guarda bien
    # como texto. Leer como texto refleja el valor real que se está migrando.
    df_siga = pd.read_excel(ruta_siga, sheet_name=hoja_siga, dtype=str)
    print(f"Filas a validar: {len(df_siga)}")

    print("Validando contra reglas de ESTRUCTURA...")
    resumen_df, detalle_df = validar_dataframe(df_siga, reglas)

    if ruta_salida:
        escribir_reporte(resumen_df, detalle_df, len(df_siga), ruta_salida)
        print(f"Reporte guardado en: {ruta_salida}")

    print()
    print("==========================================")
    print(f"Total de errores: {len(detalle_df)}")
    if not resumen_df.empty:
        print(resumen_df.to_string(index=False))
    print("==========================================")

    return resumen_df, detalle_df


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Valida un archivo SIGA contra ESTRUCTURA")
    parser.add_argument("archivo_siga", help="Ruta al archivo ya transformado a formato SIGA")
    parser.add_argument("--estructura", default=str(ESTRUCTURA_DEFAULT),
                         help="Ruta al archivo Estructura_Migracion_Patrimonio.xlsx")
    parser.add_argument("--salida", default=str(CARPETA / "Reporte_Validacion_SIGA.xlsx"),
                         help="Ruta del reporte de salida")
    args = parser.parse_args()

    validar_archivo(args.archivo_siga, archivo_estructura=args.estructura, ruta_salida=args.salida)
