"""
============================================================
VALIDADOR DE ESTRUCTURA - MIGRACIÓN PATRIMONIO SIAPA -> SIGA
============================================================

Lee el archivo ya transformado (el que genera tu script de
migración) y lo valida campo por campo contra las reglas de
la hoja ESTRUCTURA de Estructura_Migracion_Patrimonio.xlsx:

    - Tipo de dato (NUMÉRICO / CARACTER / FECHA)
    - Longitud máxima
    - Obligatoriedad (incluye obligatoriedad CONDICIONAL:
      NEA vs O/C, PECOSA, ESNI, etiquetado)
    - Valores permitidos cuando la ESTRUCTURA los fija
      (ESTADO, TIPO_BIEN, TIPO_ACTIVO, TIPO_DOC_ADQUISICION,
      TIPO_TRAN_NEA, TIPO_DOC_ALTA, flags S/N, ESTADO_CONSERV)

No corrige nada. Solo te dice, fila por fila y campo por
campo, qué está mal, para que decidas si corregir el script
de transformación o completar catálogos antes de recorrer
121k filas a mano.

Uso:
    python validar_estructura.py

Genera:
    Reporte_Errores_Validacion.xlsx
        - Hoja "Errores": una fila por cada error encontrado
        - Hoja "Resumen_Campo": conteo de errores por campo
        - Hoja "Resumen_Fila": cuántas filas tienen 0, 1, 2... errores
"""

import datetime
import pandas as pd
from pathlib import Path
from collections import Counter

# ============================================================
# RUTAS - AJUSTA SI ES NECESARIO
# ============================================================

CARPETA = Path(__file__).parent

ARCHIVO_SIGA = CARPETA / "Migracion_SIGA_Prueba_SEDE.xlsx"      # salida de tu script de migración
ARCHIVO_REPORTE = CARPETA / "Reporte_Errores_Validacion.xlsx"   # resúmenes (siempre cabe en Excel)
ARCHIVO_ERRORES_DETALLE = CARPETA / "Errores_Detalle.csv"       # detalle fila por fila (puede ser muy grande)

# Si el detalle de errores no cabe en una hoja de Excel (más de ~1,000,000 filas),
# igual se exporta completo en CSV; en el Excel se deja solo una muestra.
MUESTRA_MAXIMA_EN_EXCEL = 200_000

HOJA_SIGA = None  # None = primera hoja


# ============================================================
# REGLAS DE ESTRUCTURA (según ESTRUCTURA + FORMATO)
# ============================================================
#
# tipo: "NUMERICO" | "CARACTER" | "FECHA"
# max: longitud máxima (CARACTER/NUMERICO) o (enteros, decimales) para NUMERICO con decimales
# obligatorio: True/False fijo
# obligatorio_si: función(registro) -> True/False, para obligatoriedad condicional
# valores: set de valores permitidos (se valida solo si el campo no está vacío)
#
# Los campos con obligatoriedad condicional se documentan en la
# columna CONDICIÓN/DESCRIPCIÓN de ESTRUCTURA:
#   - NEA (045) exige TIPO_MOV_NEA, TIPO_TRAN_NEA, VALOR_NEA, FECHA_NEA
#   - O/C (031) exige NRO_ORDEN, VALOR_COMPRA, FECHA_COMPRA
#   - TIPO_DOC_ALTA=046 exige NRO_PECOSA
#   - FLAG_ITEM_ESNI=S exige MODELO, NRO_SERIE, COD_MODELO
#   - FLAG_ETIQUETA=S exige FECHA_ETIQUET

REGLAS = {
    "SEC_EJEC":              {"tipo": "NUMERICO", "max": 6,  "obligatorio": True},
    "TIPO_MODALIDAD":        {"tipo": "NUMERICO", "max": 1,  "obligatorio": True},
    "SECUENCIA":             {"tipo": "NUMERICO", "max": 8,  "obligatorio": True},
    "CODIGO_ACTIVO":         {"tipo": "CARACTER", "max": 12, "obligatorio": True},
    "ESTADO":                {"tipo": "CARACTER", "max": 1,  "obligatorio": True, "valores": {"1", "2"}},
    "TIPO_BIEN":              {"tipo": "CARACTER", "max": 1,  "obligatorio": True, "valores": {"B"}},
    "GRUPO_BIEN":            {"tipo": "CARACTER", "max": 2,  "obligatorio": True},
    "CLASE_BIEN":            {"tipo": "CARACTER", "max": 2,  "obligatorio": True},
    "FAMILIA_BIEN":          {"tipo": "CARACTER", "max": 4,  "obligatorio": True},
    "ITEM_BIEN":             {"tipo": "CARACTER", "max": 4,  "obligatorio": True},
    "DESCRIPCION":           {"tipo": "CARACTER", "max": 200, "obligatorio": True},
    "TIPO_ACTIVO":           {"tipo": "CARACTER", "max": 1,  "obligatorio": True, "valores": {"1", "2"}},
    "TIPO_DOC_ADQUISICION":  {"tipo": "CARACTER", "max": 3,  "obligatorio": True, "valores": {"031", "045"}},

    "TIPO_MOV_NEA":          {"tipo": "CARACTER", "max": 1,
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ADQUISICION") == "045",
                               "valores": {"I"}},
    "TIPO_TRAN_NEA":         {"tipo": "CARACTER", "max": 2,
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ADQUISICION") == "045",
                               "valores": {"3", "4", "5", "8", "9"}},
    "NRO_DOCUMENTO":         {"tipo": "CARACTER", "max": 10, "obligatorio": True},
    "VALOR_NEA":             {"tipo": "NUMERICO", "max": (16, 6),
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ADQUISICION") == "045"},
    "FECHA_NEA":             {"tipo": "FECHA",
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ADQUISICION") == "045"},

    "NRO_ORDEN":             {"tipo": "NUMERICO", "max": 7,
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ADQUISICION") == "031"},
    "VALOR_COMPRA":          {"tipo": "NUMERICO", "max": (16, 6),
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ADQUISICION") == "031"},
    "FECHA_COMPRA":          {"tipo": "FECHA",
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ADQUISICION") == "031"},

    "TIPO_DOC_ALTA":         {"tipo": "CARACTER", "max": 3, "obligatorio": True, "valores": {"046"}},
    "NRO_PECOSA":            {"tipo": "NUMERICO", "max": 5,
                               "obligatorio_si": lambda r: r.get("TIPO_DOC_ALTA") == "046"},
    "FECHA_ALTA":            {"tipo": "FECHA", "obligatorio": True},

    "SEDE":                  {"tipo": "NUMERICO", "max": 3,  "obligatorio": True},
    "CENTRO_COSTO":          {"tipo": "CARACTER", "max": 15, "obligatorio": True},
    "EMPLEADO_RESPONSABLE":  {"tipo": "CARACTER", "max": 15, "obligatorio": True},
    "EMPLEADO_FINAL":        {"tipo": "CARACTER", "max": 15, "obligatorio": True},
    "TIPO_UBICAC":           {"tipo": "NUMERICO", "max": 3,  "obligatorio": True},
    "SUBTIPO_UBICAC":        {"tipo": "CARACTER", "max": 3,  "obligatorio": True},

    "CARACTERISTICAS":       {"tipo": "CARACTER", "max": 300, "obligatorio": False},
    "MODELO":                {"tipo": "CARACTER", "max": 40,
                               "obligatorio_si": lambda r: r.get("FLAG_ITEM_ESNI") == "S"},
    "MEDIDAS":               {"tipo": "CARACTER", "max": 30, "obligatorio": False},
    "NRO_SERIE":             {"tipo": "CARACTER", "max": 20,
                               "obligatorio_si": lambda r: r.get("FLAG_ITEM_ESNI") == "S"},
    "FLAG_ITEM_ESNI":        {"tipo": "CARACTER", "max": 1, "obligatorio": True, "valores": {"S", "N"}},
    "MARCA":                 {"tipo": "NUMERICO", "max": 5,  "obligatorio": True},
    "COD_MODELO":            {"tipo": "NUMERICO", "max": 3,
                               "obligatorio_si": lambda r: r.get("FLAG_ITEM_ESNI") == "S"},
    "ESTADO_ACTUAL":         {"tipo": "CARACTER", "max": 1, "obligatorio": True, "valores": {"S", "N"}},
    "ESTADO_CONSERV":        {"tipo": "CARACTER", "max": 1, "obligatorio": True,
                               "valores": {"1", "2", "3", "4", "5"}},
    "FECHA_MOVIMTO":         {"tipo": "FECHA", "obligatorio": True},
    "PROVEEDOR":             {"tipo": "NUMERICO", "max": 5, "obligatorio": False},
    "COD_ALMACEN":           {"tipo": "CARACTER", "max": 3, "obligatorio": True},
    "SEC_ALMACEN":           {"tipo": "CARACTER", "max": 3, "obligatorio": True},
    "FLAG_ETIQUETA":         {"tipo": "CARACTER", "max": 1, "obligatorio": False, "valores": {"S", "N"}},
    "FECHA_ETIQUET":         {"tipo": "FECHA",
                               "obligatorio_si": lambda r: r.get("FLAG_ETIQUETA") == "S"},
    "VALOR_INICIAL":         {"tipo": "NUMERICO", "max": (16, 2), "obligatorio": True},
    "VALOR_DEPREC":          {"tipo": "NUMERICO", "max": (16, 2), "obligatorio": True},
    "INVENT_SCANER":         {"tipo": "CARACTER", "max": 1, "obligatorio": True, "valores": {"S", "N"}},
    "CLASIFICADOR":          {"tipo": "CARACTER", "max": 20, "obligatorio": True},
    "AÑO_EJE":               {"tipo": "NUMERICO", "max": 4, "obligatorio": True},
    "SUB_CTA":               {"tipo": "CARACTER", "max": 8, "obligatorio": True},
    "MAYOR":                 {"tipo": "CARACTER", "max": 4, "obligatorio": True},
    "CODIGO_BARRA":          {"tipo": "CARACTER", "max": 15, "obligatorio": False},
    "OBSERVACIONES":         {"tipo": "CARACTER", "max": 300, "obligatorio": False},
}


# ============================================================
# FUNCIONES DE VALIDACIÓN
# ============================================================

def esta_vacio(valor):
    """Determina si un valor cuenta como vacío para efectos de OBLIGATORIO."""
    if valor is None:
        return True
    if isinstance(valor, float) and pd.isna(valor):
        return True
    if isinstance(valor, str) and valor.strip() == "":
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return False


def es_fecha_valida(valor):
    return isinstance(valor, (pd.Timestamp, datetime.datetime, datetime.date))


def validar_registro(registro, identificador):
    """
    Valida un registro (dict o Series) contra REGLAS.
    Devuelve una lista de dicts, uno por cada error encontrado.
    """
    errores = []

    for campo, regla in REGLAS.items():

        valor = registro.get(campo, None)
        vacio = esta_vacio(valor)

        # --- Determinar si es obligatorio en este registro ---
        if "obligatorio_si" in regla:
            try:
                obligatorio = bool(regla["obligatorio_si"](registro))
            except Exception:
                obligatorio = False
        else:
            obligatorio = regla.get("obligatorio", False)

        if vacio:
            if obligatorio:
                errores.append({
                    "IDENTIFICADOR": identificador,
                    "CAMPO": campo,
                    "VALOR": "",
                    "TIPO_ERROR": "CAMPO_OBLIGATORIO_VACIO",
                    "DETALLE": "El campo es obligatorio para este registro y está vacío."
                })
            continue  # si está vacío y no es obligatorio, no hay más que validar

        tipo = regla["tipo"]

        # --- NUMÉRICO ---
        if tipo == "NUMERICO":
            try:
                num = float(valor)
            except (ValueError, TypeError):
                errores.append({
                    "IDENTIFICADOR": identificador,
                    "CAMPO": campo,
                    "VALOR": str(valor),
                    "TIPO_ERROR": "TIPO_DATO_INVALIDO",
                    "DETALLE": "Se esperaba un valor NUMÉRICO."
                })
                continue

            max_spec = regla.get("max")
            if max_spec is not None:
                if isinstance(max_spec, tuple):
                    enteros, _decimales = max_spec
                    n_digitos = len(str(int(abs(num))))
                    if n_digitos > enteros:
                        errores.append({
                            "IDENTIFICADOR": identificador,
                            "CAMPO": campo,
                            "VALOR": str(valor),
                            "TIPO_ERROR": "LONGITUD_EXCEDIDA",
                            "DETALLE": f"Máximo {enteros} dígitos enteros, tiene {n_digitos}."
                        })
                else:
                    n_digitos = len(str(int(abs(num))))
                    if n_digitos > max_spec:
                        errores.append({
                            "IDENTIFICADOR": identificador,
                            "CAMPO": campo,
                            "VALOR": str(valor),
                            "TIPO_ERROR": "LONGITUD_EXCEDIDA",
                            "DETALLE": f"Máximo {max_spec} dígitos, tiene {n_digitos}."
                        })

        # --- CARACTER ---
        elif tipo == "CARACTER":
            texto = str(valor).strip()
            max_len = regla.get("max")
            if max_len is not None and len(texto) > max_len:
                errores.append({
                    "IDENTIFICADOR": identificador,
                    "CAMPO": campo,
                    "VALOR": texto[:50],
                    "TIPO_ERROR": "LONGITUD_EXCEDIDA",
                    "DETALLE": f"Máximo {max_len} caracteres, tiene {len(texto)}."
                })

            valores_permitidos = regla.get("valores")
            if valores_permitidos and texto not in valores_permitidos:
                errores.append({
                    "IDENTIFICADOR": identificador,
                    "CAMPO": campo,
                    "VALOR": texto,
                    "TIPO_ERROR": "VALOR_NO_PERMITIDO",
                    "DETALLE": f"Valores válidos: {sorted(valores_permitidos)}."
                })

        # --- FECHA ---
        elif tipo == "FECHA":
            if not es_fecha_valida(valor):
                errores.append({
                    "IDENTIFICADOR": identificador,
                    "CAMPO": campo,
                    "VALOR": str(valor),
                    "TIPO_ERROR": "TIPO_DATO_INVALIDO",
                    "DETALLE": "Se esperaba una FECHA válida (DD/MM/AAAA)."
                })

    return errores


def validar_dataframe(df):
    """
    Recorre todo el DataFrame y devuelve:
    - df_errores: un DataFrame con todos los errores encontrados
    - resumen_campo: Counter de errores por campo
    - filas_con_errores: dict {identificador: cantidad_de_errores}
    """
    todos_los_errores = []
    filas_con_errores = {}

    columnas_disponibles = set(df.columns)
    campos_no_encontrados = [c for c in REGLAS if c not in columnas_disponibles]
    if campos_no_encontrados:
        print("AVISO: estos campos de la ESTRUCTURA no existen en el archivo a validar:")
        for c in campos_no_encontrados:
            print(f"   - {c}")

    for idx, fila in df.iterrows():
        registro = fila.to_dict()

        identificador = registro.get("CODIGO_ACTIVO", "")
        if esta_vacio(identificador):
            identificador = f"fila_excel_{idx + 2}"  # +2: encabezado + índice base 1

        errores_fila = validar_registro(registro, identificador)

        if errores_fila:
            todos_los_errores.extend(errores_fila)
            filas_con_errores[identificador] = len(errores_fila)

    df_errores = pd.DataFrame(
        todos_los_errores,
        columns=["IDENTIFICADOR", "CAMPO", "VALOR", "TIPO_ERROR", "DETALLE"]
    )

    resumen_campo = Counter(e["CAMPO"] for e in todos_los_errores)

    return df_errores, resumen_campo, filas_con_errores


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":

    print(f"Leyendo archivo a validar: {ARCHIVO_SIGA.name}")
    df = pd.read_excel(ARCHIVO_SIGA, sheet_name=HOJA_SIGA or 0)
    print(f"Registros a validar: {len(df)}")

    df_errores, resumen_campo, filas_con_errores = validar_dataframe(df)

    total_filas = len(df)
    total_filas_con_error = len(filas_con_errores)
    total_filas_ok = total_filas - total_filas_con_error

    print()
    print("==========================================")
    print("VALIDACIÓN FINALIZADA")
    print("==========================================")
    print(f"Filas totales:        {total_filas}")
    print(f"Filas SIN errores:    {total_filas_ok}")
    print(f"Filas CON errores:    {total_filas_con_error}")
    print(f"Errores totales:      {len(df_errores)}")
    print("==========================================")
    print("Errores por campo (top 15):")
    for campo, cantidad in resumen_campo.most_common(15):
        print(f"   {campo:<25} {cantidad}")
    print("==========================================")

    # --------------------------------------------------------
    # Resumen por campo
    # --------------------------------------------------------

    df_resumen_campo = pd.DataFrame(
        [{"CAMPO": campo, "CANTIDAD_ERRORES": cantidad,
          "PORCENTAJE_FILAS": round(100 * cantidad / total_filas, 2)}
         for campo, cantidad in resumen_campo.most_common()]
    )

    # --------------------------------------------------------
    # Resumen por cantidad de errores por fila
    # --------------------------------------------------------

    distribucion = Counter(filas_con_errores.values())
    filas_sin_error_count = total_filas - len(filas_con_errores)

    filas_resumen = [{"CANTIDAD_ERRORES_EN_LA_FILA": 0, "CANTIDAD_DE_FILAS": filas_sin_error_count}]
    for n_errores, n_filas in sorted(distribucion.items()):
        filas_resumen.append({"CANTIDAD_ERRORES_EN_LA_FILA": n_errores, "CANTIDAD_DE_FILAS": n_filas})

    df_resumen_fila = pd.DataFrame(filas_resumen)

    # --------------------------------------------------------
    # Exportar detalle completo (CSV, sin límite de filas)
    # --------------------------------------------------------

    df_errores.to_csv(ARCHIVO_ERRORES_DETALLE, index=False, encoding="utf-8-sig")
    print(f"Detalle completo de errores ({len(df_errores)} filas): {ARCHIVO_ERRORES_DETALLE}")

    # --------------------------------------------------------
    # Exportar reporte Excel (resúmenes + muestra si el detalle es muy grande)
    # --------------------------------------------------------

    with pd.ExcelWriter(ARCHIVO_REPORTE, engine="openpyxl") as writer:
        df_resumen_campo.to_excel(writer, sheet_name="Resumen_Campo", index=False)
        df_resumen_fila.to_excel(writer, sheet_name="Resumen_Fila", index=False)

        if len(df_errores) <= MUESTRA_MAXIMA_EN_EXCEL:
            df_errores.to_excel(writer, sheet_name="Errores", index=False)
        else:
            df_errores.head(MUESTRA_MAXIMA_EN_EXCEL).to_excel(
                writer, sheet_name="Errores_Muestra", index=False
            )
            aviso = pd.DataFrame([{
                "AVISO": (f"El detalle completo tiene {len(df_errores)} filas, supera el límite de "
                          f"Excel. Esta hoja muestra solo las primeras {MUESTRA_MAXIMA_EN_EXCEL}. "
                          f"El archivo {ARCHIVO_ERRORES_DETALLE.name} tiene el detalle completo.")
            }])
            aviso.to_excel(writer, sheet_name="Aviso", index=False)

    print(f"Reporte generado: {ARCHIVO_REPORTE}")
