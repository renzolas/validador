import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from datetime import datetime
import re
import time
from google.colab import files
from tqdm import tqdm

# === INSTRUCCIONES DE USO ===
print("""
📄 INSTRUCCIONES DE USO:
1️⃣ Sube primero el archivo A (referencia) y luego el archivo B (a validar).
2️⃣ Ambos archivos deben tener:
    - Exactamente las mismas columnas.
    - El mismo orden de columnas.
    - Los mismos encabezados (sin cambiar nombres).
3️⃣ Los valores serán validados según el tipo de dato esperado:
    - numérico: solo números
    - texto: solo letras y espacios
    - alfanumérico: letras y números
    - fecha: formato MM/DD/YYYY
    - fecha corta: formato MM/YY
4️⃣ El archivo validado se descargará como 'validado.xlsx' con celdas rojas si hay errores.
""")

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

# === TIPOS ESPERADOS POR COLUMNA (modificar según tus datos) ===
tipos_columna = {
    "id": "numerico",
    "nombre": "texto",
    "codigo": "alfanumerico",
    "fecha": "fecha",
    "mes": "fecha_corta"
}

# === SUBIDA DE ARCHIVOS ===
print("\n📁 Sube primero el archivo A (referencia)")
archivo_a = files.upload()
archivo_a = list(archivo_a.keys())[0]

print("\n📁 Ahora sube el archivo B (a validar)")
archivo_b = files.upload()
archivo_b = list(archivo_b.keys())[0]

# === CARGA DE DATAFRAMES ===
df_a = pd.read_excel(archivo_a, dtype=str)
df_b = pd.read_excel(archivo_b, dtype=str)

# === NORMALIZA ENCABEZADOS ===
df_a.columns = [normalizar_columna(col) for col in df_a.columns]
df_b.columns = [normalizar_columna(col) for col in df_b.columns]

# === VALIDA COLUMNAS ===
if list(df_a.columns) != list(df_b.columns):
    raise ValueError("❌ Las columnas no coinciden o no están en el mismo orden.")

# === ABRE ARCHIVO B COMO LIBRO DE EXCEL ===
wb = load_workbook(archivo_b)
ws = wb.active

rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

# === VALIDACIÓN CELDA POR CELDA CON PROGRESO Y COUNTDOWN ===
total_celdas = df_a.shape[0] * df_a.shape[1]
print("\n🔍 Validando datos...")
for i in tqdm(range(df_a.shape[0]), desc="Progreso", unit="fila"):
    time.sleep(0.05)  # Simula tiempo de validación
    for j in range(df_a.shape[1]):
        valor_a = str(df_a.iat[i, j]).strip()
        valor_b = str(df_b.iat[i, j]).strip()
        col_name = df_a.columns[j]
        celda = ws.cell(row=i+2, column=j+1)

        tipo_esperado = tipos_columna.get(col_name)

        # Celda vacía
        if not valor_b:
            celda.fill = rojo
            celda.comment = Comment("Celda vacía", "Validador")
            continue

        # Valor distinto
        if valor_a != valor_b:
            celda.fill = rojo
            celda.comment = Comment(
                f'Valor diferente:\nEsperado: "{valor_a}"\nEncontrado: "{valor_b}"',
                "Validador"
            )

        # Tipo de dato inválido
        if tipo_esperado in validadores:
            if not validadores[tipo_esperado](valor_b):
                celda.fill = rojo
                celda.comment = Comment(f"Tipo inválido: se esperaba {tipo_esperado}", "Validador")

# === GUARDA ARCHIVO VALIDADO ===
output_name = "validado.xlsx"
wb.save(output_name)

# === COUNTDOWN FINAL ===
print("\n⏳ Finalizando y preparando descarga...")
for t in range(3, 0, -1):
    print(f"📦 Descargando en {t}...")
    time.sleep(1)

# === DESCARGA ===
files.download(output_name)
print("\n✅ Validación finalizada. Archivo 'validado.xlsx' listo.")





