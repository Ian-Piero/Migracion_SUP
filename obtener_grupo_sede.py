
import pandas as pd
from pathlib import Path


# ============================================================
# RUTAS
# ============================================================

CARPETA = Path(__file__).parent

ARCHIVO_SIAPA = CARPETA / "siapa-17-08-2026 (6).xlsx"
ARCHIVO_SALIDA = CARPETA / "Tipos_NombreDe.xlsx"

HOJA_SIAPA = "Hoja1"


# ============================================================
# LEER ARCHIVO SIAPA
# ============================================================

df = pd.read_excel(
    ARCHIVO_SIAPA,
    sheet_name=HOJA_SIAPA
)


# ============================================================
# OBTENER VALORES ÚNICOS DE NombreDe
# ============================================================

valores_unicos = (
    df["NombreDe"]
    .fillna("")
    .astype(str)
    .str.strip()
    .unique()
)


# Ordenar alfabéticamente
valores_unicos = sorted(valores_unicos)


# ============================================================
# CREAR DATAFRAME
# ============================================================

df_unicos = pd.DataFrame({
    "NombreDe": valores_unicos
})


# ============================================================
# MOSTRAR RESULTADOS
# ============================================================

print("==========================================")
print("VALORES ÚNICOS DE NombreDe")
print("==========================================")

for indice, valor in enumerate(df_unicos["NombreDe"], start=1):

    print(f"{indice}. {valor}")


print()
print("==========================================")
print(f"TOTAL DE VALORES ÚNICOS: {len(df_unicos)}")
print("==========================================")


# ============================================================
# EXPORTAR A EXCEL
# ============================================================

df_unicos.to_excel(
    ARCHIVO_SALIDA,
    index=False
)


print()
print(f"Archivo generado: {ARCHIVO_SALIDA}")

