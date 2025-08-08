import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from datetime import datetime
import re
import os

# === INSTRUCCIONES ===
INSTRUCCIONES = """
INSTRUCCIONES DE USO:
1. Colocar en la misma carpeta los archivos A.xlsx (referencia) y B.xlsx o B.xlsm (a validar).
2. Ambos archivos deben tener:
   - Mismas columnas.
   - Mismo orden de columnas.
   - Encabezados sin modificar.
3. El validador revisará:
   - Coincidencia exacta de valores.
   - Tipos de datos esperados:
        numerico     → solo dígitos
        texto        → solo letras y espacios
        alfanumerico → letras y números
        fecha        → formato MM/DD/YYYY
        fecha_corta  → formato MM/YY
4. El resultado se guardará como "validado.xlsx".
"""

print(INSTRUCCIONES)

# === FUNCIONES DE VALIDACIÓN ===
def normalizar_columna(nombre):
    return str(nombre).strip().lower()

def es_numerico(valor):
    return valor.isdigit()

def es_texto(valor):
    return bool(re.match(r"^[a-zA-ZáéíóúÁÉÍÓÚüÜñÑ\s]+$", valor))

def es_alfanumerico(valor):
    return bool(re.match(r"^[a-zA-Z0-9\s]+$", valor))

def es_fecha(valor):
    try:
        datetime.strptime(valor, "%m/%d/%Y")
        return True
    except:
        return False

def es_fecha_corta(valor):
    return bool(re.match(r"^(0[1-9]|1[0-2])\/\d{2}$", valor))

validadores = {
    "numerico": es_numerico,
    "texto": es_texto,
    "alfanumerico": es_alfanumerico,
    "fecha": es_fecha,
    "fecha_corta": es_fecha_corta
}

# Tipos esperados por columna (puedes editar según tu necesidad)
tipos_columna = {
    "id": "numerico",
    "nombre": "texto",
    "codigo": "alfanumerico",
    "fecha": "fecha",
    "mes": "fecha_corta"
}

# === CARGA DE ARCHIVOS ===
archivo_A = "A.xlsx"
archivo_B = None
for ext in ["B.xlsx", "B.xlsm"]:
    if os.path.exists(ext):
        archivo_B = ext
        break

df_A = pd.read_excel(archivo_A, dtype=str)
df_B = pd.read_excel(archivo_B, dtype=str)

# === VALIDAR ESTRUCTURA DE COLUMNAS ===
cols_A = [normalizar_columna(c) for c in df_A.columns]
cols_B = [normalizar_columna(c) for c in df_B.columns]

if cols_A != cols_B:
    raise ValueError("❌ Las columnas no coinciden o el orden ha sido alterado.")

# === PROCESAR VALIDACIÓN ===
wb = load_workbook(archivo_B)
ws = wb.active
fill_rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

for col_idx, col_name in enumerate(df_B.columns, start=1):
    tipo_esperado = tipos_columna.get(normalizar_columna(col_name))
    for row_idx, valor in enumerate(df_B[col_name], start=2):
        valor_str = str(valor).strip() if pd.notna(valor) else ""
        error_msg = None

        if valor_str != str(df_A.iloc[row_idx - 2, col_idx - 1]).strip():
            error_msg = "Valor diferente al archivo A"

        if tipo_esperado and valor_str and not validadores[tipo_esperado](valor_str):
            error_msg = f"No cumple el formato {tipo_esperado}"

        if error_msg:
            ws.cell(row=row_idx, column=col_idx).fill = fill_rojo
            ws.cell(row=row_idx, column=col_idx).comment = Comment(error_msg, "Validador")

# === GUARDAR RESULTADO ===
wb.save("validado.xlsx")
print("✅ Validación completada. Archivo guardado como 'validado.xlsx'.")





