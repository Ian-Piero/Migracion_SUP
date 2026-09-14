"""
Carga los catálogos de mapeo SIAPA -> SIGA desde
Catalogos_Migracion_SIAPA_SIGA.xlsx (una vez llenada la columna
amarilla de cada hoja) y expone diccionarios listos para usar
dentro del script de transformación principal.

Uso en tu script de migración:

    from cargar_catalogos import CATALOGOS

    codigo_sede = CATALOGOS["sede"].get(sede_texto, "")
    codigo_marca = CATALOGOS["marca"].get(marca_texto, 328)
    centro_costo = CATALOGOS["centro_costo"].get(idcc, "")
    tipo_ubicac = CATALOGOS["ubicacion"].get(iddependencia, "")
    subtipo_ubicac = CATALOGOS["subtipo_ubicacion"].get(codubicacion, "")
    cod_empleado = CATALOGOS["personal"].get(idpersonal, "")
    clasificador, sub_cta, mayor = CATALOGOS["clasificador"].get(
        (grupo_bien, clase_bien), ("", "", "")
    )
"""

import pandas as pd
from pathlib import Path

ARCHIVO_CATALOGOS = Path(__file__).parent / "Catalogos_Migracion_SIAPA_SIGA.xlsx"


def _limpio(v):
    if pd.isna(v):
        return ""
    v = str(v).strip()
    return v


def _cargar_dict_simple(hoja, col_clave, col_valor, saltar_filas_ejemplo=1):
    """Lee una hoja de 1 clave -> 1 valor, ignorando filas sin código llenado."""
    df = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name=hoja, skiprows=range(1, 1 + saltar_filas_ejemplo))
    mapa = {}
    vacios = 0
    for _, fila in df.iterrows():
        clave = _limpio(fila[col_clave])
        valor = _limpio(fila[col_valor])
        if clave == "":
            continue
        if valor == "":
            vacios += 1
            continue
        mapa[clave] = valor
    if vacios:
        print(f"  [{hoja}] {vacios} valores sin código SIGA todavía (quedarán fuera del mapeo)")
    return mapa


def _cargar_clasificador(saltar_filas_ejemplo=1):
    df = pd.read_excel(
        ARCHIVO_CATALOGOS,
        sheet_name="CATALOGO_CLASIFICADOR",
        skiprows=range(1, 1 + saltar_filas_ejemplo),
    )
    mapa = {}
    vacios = 0
    for _, fila in df.iterrows():
        grupo = _limpio(fila["GRUPO_BIEN"])
        clase = _limpio(fila["CLASE_BIEN"])
        clasificador = _limpio(fila["CLASIFICADOR (llenar)"])
        sub_cta = _limpio(fila["SUB_CTA (llenar)"])
        mayor = _limpio(fila["MAYOR (llenar)"])
        if grupo == "":
            continue
        if clasificador == "" and sub_cta == "" and mayor == "":
            vacios += 1
            continue
        mapa[(grupo, clase)] = (clasificador, sub_cta, mayor)
    if vacios:
        print(f"  [CATALOGO_CLASIFICADOR] {vacios} combinaciones grupo/clase sin llenar")
    return mapa


def cargar_todo():
    if not ARCHIVO_CATALOGOS.exists():
        raise FileNotFoundError(
            f"No encuentro {ARCHIVO_CATALOGOS}. Coloca el archivo de catálogos "
            "en la misma carpeta que este script, o ajusta ARCHIVO_CATALOGOS."
        )

    print("Cargando catálogos de mapeo SIAPA -> SIGA...")

    catalogos = {
        "sede": _cargar_dict_simple(
            "CATALOGO_SEDE", "SEDE (derivada por el script)", "CODIGO_SIGA (llenar)"
        ),
        "marca": _cargar_dict_simple(
            "CATALOGO_MARCA", "MARCA (texto en SIAPA)", "CODIGO_SIGA (llenar)"
        ),
        "centro_costo": _cargar_dict_simple(
            "CATALOGO_CENTRO_COSTO", "idcc (SIAPA)", "CENTRO_COSTO_SIGA (llenar)"
        ),
        "ubicacion": _cargar_dict_simple(
            "CATALOGO_UBICACION", "iddependencia (SIAPA)", "TIPO_UBICAC_SIGA (llenar)"
        ),
        "subtipo_ubicacion": _cargar_dict_simple(
            "CATALOGO_SUBTIPO_UBIC", "codubicacion (SIAPA)", "SUBTIPO_UBICAC_SIGA (llenar)"
        ),
        "personal": _cargar_dict_simple(
            "CATALOGO_PERSONAL", "idpersonal (SIAPA)", "COD_EMPLEADO_SIGA (llenar)"
        ),
        "clasificador": _cargar_clasificador(),
    }

    print("Catálogos cargados:")
    for nombre, mapa in catalogos.items():
        print(f"  - {nombre}: {len(mapa)} equivalencias con código SIGA")

    return catalogos


# Se carga una sola vez al importar el módulo
CATALOGOS = cargar_todo()


if __name__ == "__main__":
    # Ejecutar este archivo directamente solo imprime el resumen de carga.
    pass
