import pandas as pd
from pathlib import Path
from openpyxl import load_workbook


# ============================================================
# RUTAS
# ============================================================

CARPETA = Path(__file__).parent

ARCHIVO_SIAPA = CARPETA / "siapa-17-08-2026 (6).xlsx"
ARCHIVO_SALIDA = CARPETA / "Migracion_SIGA_Prueba_SEDE.xlsx"

HOJA_SIAPA = "Hoja1"


# ============================================================
# CONFIGURACIÓN
# ============================================================

CANTIDAD_PRUEBA = 121312
SEC_EJEC = 160
TIPO_MODALIDAD = 1
TIPO_BIEN = "B"


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def limpiar(valor):
    """Convierte un valor a texto limpio."""
    if pd.isna(valor):
        return ""

    return str(valor).strip()


def convertir_fecha(valor):
    """Convierte un valor a fecha."""
    if pd.isna(valor) or valor == "":
        return None

    try:
        return pd.to_datetime(valor)
    except:
        return None


def obtener_anio(valor):
    """Obtiene el año de una fecha."""
    fecha = convertir_fecha(valor)

    if fecha is None:
        return ""

    return fecha.year


def estado_conservacion(estado):
    """
    SIGA:
    1 = Bueno
    2 = Regular
    3 = Malo
    4 = Muy Malo
    5 = Nuevo
    """

    estado = limpiar(estado).upper()

    equivalencias = {
        "BUENO": 1,
        "REGULAR": 2,
        "MALO": 3,
        "MUY MALO": 4,
        "NUEVO": 5
    }

    return equivalencias.get(estado, "")


def obtener_componentes_codigo(codigo):
    """
    Código patrimonial SIGA:
    GGCCFFFFIIII

    GG = Grupo
    CC = Clase
    FFFF = Familia
    IIII = Item
    """

    codigo = limpiar(codigo)

    codigo = codigo.replace(" ", "")

    if len(codigo) != 12:
        return "", "", "", ""

    grupo = codigo[0:2]
    clase = codigo[2:4]
    familia = codigo[4:8]
    item = codigo[8:12]

    return grupo, clase, familia, item


def tipo_documento_adquisicion(documento):
    """
    SIGA:
    031 = O/C - Guía
    045 = NEA
    """

    documento = limpiar(documento).upper()

    if documento in ["OC", "OC/GI", "O/C", "O/C-GI"]:
        return "031"

    if documento == "NEA":
        return "045"

    return documento


def tipo_documento_alta(documento):
    """
    Para la prueba se utiliza:
    046 = PECOSA

    Si posteriormente encontramos otros tipos
    en la documentación de la institución,
    se deberán homologar.
    """

    documento = limpiar(documento).upper()

    if documento != "":
        return "046"

    return ""


def obtener_sede(nombre_de, nombre_ub):
    """
    Determina la sede utilizando NombreDe y NombreUb.

    Prioridad:
    1. Se revisa NombreDe y NombreUb.
    2. Si aparece una sede conocida, se asigna esa sede.
    3. Si aparece LA RAYA, se asigna MARANGANI.
    4. Si aparece KAYRA o PERAYOC, se asigna PRINCIPAL.
    5. Si no se encuentra ninguna sede conocida,
       se asume SEDE PRINCIPAL.
    """

    nombre_de = limpiar(nombre_de).upper()
    nombre_ub = limpiar(nombre_ub).upper()

    texto = f"{nombre_de} {nombre_ub}"

    # --------------------------------------------------------
    # SEDES / FILIALES
    # --------------------------------------------------------

    if "SICUANI" in texto:
        return "SEDE SICUANI"

    if "MARANGANI" in texto:
        return "SEDE MARANGANI"

    # Centro Experimental La Raya pertenece a Marangani
    if "LA RAYA" in texto:
        return "SEDE MARANGANI"

    if "ESPINAR" in texto or "YAURI" in texto:
        return "SEDE ESPINAR"

    if "YANAOCA" in texto or "CANAS" in texto:
        return "SEDE YANAOCA"

    if "SANTO TOMAS" in texto or "SANTO TOMÁS" in texto:
        return "SEDE SANTO TOMÁS"

    if "ANDAHUAYLAS" in texto:
        return "SEDE ANDAHUAYLAS"

    if "PUERTO MALDONADO" in texto or "TAMBOPATA" in texto:
        return "SEDE PUERTO MALDONADO"

    if "SANGARARA" in texto or "SANGARARÁ" in texto:
        return "SEDE SANGARARÁ"

    # --------------------------------------------------------
    # INSTALACIONES DE CUSCO
    # --------------------------------------------------------

    if "KAYRA" in texto:
        return "SEDE PRINCIPAL"

    if "PERAYOC" in texto:
        return "SEDE PRINCIPAL"

    if "CUSCO" in texto:
        return "SEDE PRINCIPAL"

    # --------------------------------------------------------
    # REGLA FINAL
    # --------------------------------------------------------
    # Si no encontramos ninguna ubicación conocida,
    # asumimos que pertenece a la sede principal de Cusco.

    return "SEDE PRINCIPAL"

def procesar_marca(marca):
    """
    Para esta primera transformación:

    - Si existe marca, conservamos el nombre textual.
    - Si no existe marca, utilizamos 328 = SIN MARCA.

    Posteriormente se reemplazará el nombre por el
    código numérico oficial SIGA.
    """

    marca = limpiar(marca)

    if marca == "":
        return 328

    return marca


# ============================================================
# LEER SIAPA
# ============================================================

print("Leyendo archivo SIAPA...")

df_siapa = pd.read_excel(
    ARCHIVO_SIAPA,
    sheet_name=HOJA_SIAPA
)

print(f"Registros encontrados: {len(df_siapa)}")


# ============================================================
# TOMAR SOLO LOS PRIMEROS 30
# ============================================================

df_siapa = df_siapa.head(CANTIDAD_PRUEBA).copy()

print(f"Registros seleccionados para la prueba: {len(df_siapa)}")


# ============================================================
# COLUMNAS DEL FORMATO SIGA
# ============================================================

columnas_siga = [
    "SEC_EJEC",
    "TIPO_MODALIDAD",
    "SECUENCIA",
    "CODIGO_ACTIVO",
    "ESTADO",
    "TIPO_BIEN",
    "GRUPO_BIEN",
    "CLASE_BIEN",
    "FAMILIA_BIEN",
    "ITEM_BIEN",
    "DESCRIPCION",
    "TIPO_ACTIVO",
    "TIPO_DOC_ADQUISICION",
    "TIPO_MOV_NEA",
    "TIPO_TRAN_NEA",
    "NRO_DOCUMENTO",
    "VALOR_NEA",
    "FECHA_NEA",
    "NRO_ORDEN",
    "VALOR_COMPRA",
    "FECHA_COMPRA",
    "TIPO_DOC_ALTA",
    "NRO_PECOSA",
    "FECHA_ALTA",
    "SEDE",
    "CENTRO_COSTO",
    "EMPLEADO_RESPONSABLE",
    "EMPLEADO_FINAL",
    "TIPO_UBICAC",
    "SUBTIPO_UBICAC",
    "CARACTERISTICAS",
    "MODELO",
    "MEDIDAS",
    "NRO_SERIE",
    "FLAG_ITEM_ESNI",
    "MARCA",
    "COD_MODELO",
    "ESTADO_ACTUAL",
    "ESTADO_CONSERV",
    "FECHA_MOVIMTO",
    "PROVEEDOR",
    "COD_ALMACEN",
    "SEC_ALMACEN",
    "FLAG_ETIQUETA",
    "FECHA_ETIQUET",
    "VALOR_INICIAL",
    "VALOR_DEPREC",
    "INVENT_SCANER",
    "CLASIFICADOR",
    "AÑO_EJE",
    "SUB_CTA",
    "MAYOR",
    "CODIGO_BARRA",
    "OBSERVACIONES"
]


# ============================================================
# TRANSFORMACIÓN
# ============================================================

registros_siga = []


for indice, fila in df_siapa.iterrows():

    # --------------------------------------------------------
    # Datos SIAPA
    # --------------------------------------------------------

    codigo_patrimonial = limpiar(fila["CodigoPatrimonial"])

    grupo, clase, familia, item = obtener_componentes_codigo(
        codigo_patrimonial
    )

    fecha_compra = convertir_fecha(
        fila["FechaDocIngreso"]
    )

    fecha_alta = convertir_fecha(
        fila["fechapecosa"]
    )

    # --------------------------------------------------------
    # ESTADO DEL ACTIVO
    # --------------------------------------------------------

    baja = limpiar(fila["baja"]).upper()

    if baja == "SI":
        estado = "2"
    else:
        estado = "1"

    # --------------------------------------------------------
    # MARCA
    # --------------------------------------------------------

    marca_siga = procesar_marca(
        fila["marca"]
    )

    # --------------------------------------------------------
    # SEDE
    # --------------------------------------------------------
    #
    # Se obtiene a partir del nombre de la dependencia.
    # Si contiene un guion, se toma lo que sigue como sede;
    # si no, se asume SEDE CENTRAL.
    # --------------------------------------------------------

    sede = obtener_sede(
    fila["NombreDe"],
    fila["NombreUb"]
)
    # --------------------------------------------------------
    # ESNI
    # --------------------------------------------------------
    #
    # Actualmente no tenemos todavía una tabla que permita
    # determinar qué bienes son ESNI.
    #
    # Por ello, para esta primera prueba:
    #
    # FLAG_ITEM_ESNI = N
    #
    # Cuando tengamos el catálogo ESNI se podrá cambiar
    # esta parte.
    # --------------------------------------------------------

    flag_item_esni = "N"

    codigo_modelo = ""

    # --------------------------------------------------------
    # ESTADO DE CONSERVACIÓN
    # --------------------------------------------------------

    estado_conservacion_siga = estado_conservacion(
        fila["Estado"]
    )

    # --------------------------------------------------------
    # REGISTRO SIGA
    # --------------------------------------------------------

    registro = {

        "SEC_EJEC": SEC_EJEC,

        "TIPO_MODALIDAD": TIPO_MODALIDAD,

        "SECUENCIA": indice + 1,

        "CODIGO_ACTIVO": codigo_patrimonial,

        "ESTADO": estado,

        "TIPO_BIEN": TIPO_BIEN,

        "GRUPO_BIEN": grupo,

        "CLASE_BIEN": clase,

        "FAMILIA_BIEN": familia,

        "ITEM_BIEN": item,

        "DESCRIPCION": limpiar(
            fila["Denominacion"]
        ),

        "TIPO_ACTIVO": "1",

        "TIPO_DOC_ADQUISICION": tipo_documento_adquisicion(
            fila["Docingreso"]
        ),

        "TIPO_MOV_NEA": "",

        "TIPO_TRAN_NEA": "",

        "NRO_DOCUMENTO": limpiar(
            fila["NroDocIngreso"]
        ),

        "VALOR_NEA": "",

        "FECHA_NEA": None,

        "NRO_ORDEN": limpiar(
            fila["NroDocIngreso"]
        ),

        "VALOR_COMPRA": fila["ImporteUnitario"],

        "FECHA_COMPRA": fecha_compra,

        "TIPO_DOC_ALTA": tipo_documento_alta(
            fila["Docingreso"]
        ),

        "NRO_PECOSA": limpiar(
            fila["NroPecosa"]
        ),

        "FECHA_ALTA": fecha_alta,

        "SEDE": sede,

        "CENTRO_COSTO": limpiar(
            fila["idcc"]
        ),

        "EMPLEADO_RESPONSABLE": limpiar(
            fila["idpersonal"]
        ),

        "EMPLEADO_FINAL": limpiar(
            fila["idpersonal"]
        ),

        "TIPO_UBICAC": limpiar(
            fila["iddependencia"]
        ),

        "SUBTIPO_UBICAC": limpiar(
            fila["codubicacion"]
        ),

        "CARACTERISTICAS": limpiar(
            fila["Caracteristica"]
        ),

        "MODELO": limpiar(
            fila["modelo"]
        ),

        "MEDIDAS": "",

        "NRO_SERIE": limpiar(
            fila["nroserie"]
        ),

        "FLAG_ITEM_ESNI": flag_item_esni,

        "MARCA": marca_siga,

        "COD_MODELO": codigo_modelo,

        "ESTADO_ACTUAL": "S",

        "ESTADO_CONSERV": estado_conservacion_siga,

        "FECHA_MOVIMTO": fecha_alta,

        "PROVEEDOR": "",

        "COD_ALMACEN": "",

        "SEC_ALMACEN": "",

        "FLAG_ETIQUETA": "N",

        "FECHA_ETIQUET": None,

        "VALOR_INICIAL": fila["ImporteUnitario"],

        "VALOR_DEPREC": fila["DepreciacionAcumulada"],

        "INVENT_SCANER": "N",

        "CLASIFICADOR": "",

        "AÑO_EJE": obtener_anio(
            fila["FechaDocIngreso"]
        ),

        "SUB_CTA": "",

        "MAYOR": "",

        "CODIGO_BARRA": limpiar(
            fila["CodAntiguo"]
        ),

        "OBSERVACIONES": limpiar(
            fila["observacion"]
        )
    }

    registros_siga.append(registro)


# ============================================================
# CREAR DATAFRAME SIGA
# ============================================================

df_siga = pd.DataFrame(
    registros_siga,
    columns=columnas_siga
)


# ============================================================
# EXPORTAR A EXCEL
# ============================================================

print("Generando archivo SIGA...")

df_siga.to_excel(
    ARCHIVO_SALIDA,
    index=False,
    engine="openpyxl"
)


# ============================================================
# FORMATO DEL EXCEL
# ============================================================

wb = load_workbook(
    ARCHIVO_SALIDA
)

ws = wb.active


# ------------------------------------------------------------
# Formato de fechas
# ------------------------------------------------------------

columnas_fecha = [
    "FECHA_NEA",
    "FECHA_COMPRA",
    "FECHA_ALTA",
    "FECHA_MOVIMTO",
    "FECHA_ETIQUET"
]


for nombre_columna in columnas_fecha:

    numero_columna = None

    for celda in ws[1]:

        if celda.value == nombre_columna:
            numero_columna = celda.column
            break

    if numero_columna:

        for fila_excel in range(
            2,
            ws.max_row + 1
        ):

            ws.cell(
                fila_excel,
                numero_columna
            ).number_format = "DD/MM/YYYY"


# ------------------------------------------------------------
# Formato de valores monetarios
# ------------------------------------------------------------

columnas_monetarias = [
    "VALOR_NEA",
    "VALOR_COMPRA",
    "VALOR_INICIAL",
    "VALOR_DEPREC"
]


for nombre_columna in columnas_monetarias:

    numero_columna = None

    for celda in ws[1]:

        if celda.value == nombre_columna:
            numero_columna = celda.column
            break

    if numero_columna:

        for fila_excel in range(
            2,
            ws.max_row + 1
        ):

            ws.cell(
                fila_excel,
                numero_columna
            ).number_format = "0.00"


# ------------------------------------------------------------
# Ajustar ancho de columnas
# ------------------------------------------------------------

for columna in ws.columns:

    longitud_maxima = 0

    letra = columna[0].column_letter

    for celda in columna:

        if celda.value is not None:

            longitud = len(
                str(celda.value)
            )

            if longitud > longitud_maxima:
                longitud_maxima = longitud

    ws.column_dimensions[
        letra
    ].width = min(
        longitud_maxima + 2,
        50
    )


# ============================================================
# GUARDAR
# ============================================================

wb.save(
    ARCHIVO_SALIDA
)


print()
print("==========================================")
print("TRANSFORMACIÓN FINALIZADA")
print("==========================================")
print(f"Registros procesados: {len(df_siga)}")
print(f"Archivo generado: {ARCHIVO_SALIDA}")
print("==========================================")