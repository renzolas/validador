import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from datetime import datetime
import re
import os

# ==========================================================
# INSTRUCCIONES DE USO
# ==========================================================
# 1️⃣ Guarda este script en la misma carpeta que tus archivos A y B.
# 2️⃣ A debe ser el archivo de referencia (el correcto).
# 3️⃣ B es el archivo que quieres validar.
# 4️⃣ Ambos deben ser .xlsx y tener las MISMAS columnas, en el mismo orden, sin alterar nombres.
# 5️⃣ El script verificará:
#     - Que las columnas no hayan sido cambiadas ni reordenadas.
#     - Que los valores coincidan con el archivo A.
#     - Que los datos cumplan con el tipo esperado (numérico, texto, alfanumérico, fecha, fecha corta).
# 6️⃣ El resultado se guardará como "Validado.xlsx" con celdas en rojo donde haya errores.
# ==========================================================

# === Funciones de validación ===
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

# Diccionario de validadores
validadores = {
    "numerico": es_numerico,
    "texto": es_texto,
    "alfanumerico": es_alfanumerico,
    "fecha": es_fecha,
    "fecha_corta": es_fecha_corta
}

# === Configura aquí los tipos esperados por columna ===
tipos_columna = {
    "id": "numerico",
    "nombre": "texto",
    "codigo": "alfanumerico",
    "fecha": "fecha",
    "mes": "fecha_corta"
}

# === Rutas de archivos ===
archivo_a = "A.xlsx"  # Referencia
archivo_b = "B.xlsx"  # Archivo a validar

if not os.path.exists(archivo_a) or not os.path.exists(archivo_b):
    raise FileNotFoundError("❌ No se encontraron los archivos A.xlsx y B.xlsx en la carpeta.")

# === Cargar como DataFrames ===
df_a = pd.read_excel(archivo_a, dtype=str)
df_b = pd.read_excel(archivo_b, dtype=str)

# Normalizar encabezados
df_a.columns = [normalizar_columna(col) for col in df_a.columns]
df_b.columns = [normalizar_columna(col) for col in df_b.columns]

# Validar que tengan las mismas columnas
if list(df_a.columns) != list(df_b.columns):
    raise ValueError("❌ Las columnas no coinciden o están en diferente orden. No se puede validar.")

# Cargar archivo B en openpyxl
wb = load_workbook(archivo_b)
ws = wb.active

rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

# === Validación celda por celda ===
for fila in range(df_a.shape[0]):
    for col in range(df_a.shape[1]):
        valor_a = str(df_a.iat[fila, col]).strip() if pd.notna(df_a.iat[fila, col]) else ""
        valor_b = str(df_b.iat[fila, col]).strip() if pd.notna(df_b.iat[fila, col]) else ""
        col_name = df_a.columns[col]
        celda = ws.cell(row=fila+2, column=col+1)

        tipo_esperado = tipos_columna.get(col_name)

        # Celda vacía
        if not valor_b:
            celda.fill = rojo
            celda.comment = Comment("Celda vacía", "Validador")
            continue

        # Diferente al valor esperado
        if valor_a != valor_b:
            celda.fill = rojo
            celda.comment = Comment(
                f'Valor diferente:\nEsperado: "{valor_a}"\nEncontrado: "{valor_b}"',
                "Validador"
            )

        # Validación de tipo
        if tipo_esperado in validadores:
            if not validadores[tipo_esperado](valor_b):
                celda.fill = rojo
                mensaje = f"Tipo inválido: se esperaba {tipo_esperado}"
                celda.comment = Comment(mensaje, "Validador")

# Guardar archivo validado
wb.save("Validado.xlsx")
print("✅ Validación completada. Revisa el archivo 'Validado.xlsx'")





