import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from datetime import datetime
import re
import sys
import os

# === VALIDACIONES POR TIPO ===
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

# === TIPOS ESPERADOS POR COLUMNA ===
tipos_columna = {
    "id": "numerico",
    "nombre": "texto",
    "codigo": "alfanumerico",
    "fecha": "fecha",
    "mes": "fecha_corta"
}

def validar_excel(archivo_a, archivo_b):
    if not os.path.exists(archivo_a):
        raise FileNotFoundError(f"No se encontró el archivo A: {archivo_a}")
    if not os.path.exists(archivo_b):
        raise FileNotFoundError(f"No se encontró el archivo B: {archivo_b}")

    # === CARGA LOS ARCHIVOS COMO DATAFRAMES ===
    df_a = pd.read_excel(archivo_a, sheet_name=0, dtype=str)
    df_b = pd.read_excel(archivo_b, sheet_name=0, dtype=str)

    # === NORMALIZA ENCABEZADOS ===
    df_a.columns = [normalizar_columna(col) for col in df_a.columns]
    df_b.columns = [normalizar_columna(col) for col in df_b.columns]

    # === VALIDA COLUMNAS FALTANTES ===
    faltantes = set(df_a.columns) - set(df_b.columns)
    if faltantes:
        raise ValueError(f"❌ Faltan columnas en B: {faltantes}")

    # === REORDENA COLUMNAS DE B PARA QUE COINCIDAN CON A ===
    df_b = df_b[df_a.columns]

    # === CARGA EL ARCHIVO B COMO LIBRO DE EXCEL ===
    wb = load_workbook(archivo_b)
    ws = wb.active
    rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # === VALIDACIÓN CELDA POR CELDA ===
    for fila in range(df_a.shape[0]):
        for col in range(df_a.shape[1]):
            valor_a = str(df_a.iat[fila, col]).strip()
            valor_b = str(df_b.iat[fila, col]).strip()
            col_name = df_a.columns[col]
            col_index_in_b = list(df_b.columns).index(col_name)
            celda = ws.cell(row=fila+2, column=col_index_in_b+1)

            tipo_esperado = tipos_columna.get(col_name)

            if not valor_b:
                celda.fill = rojo
                celda.comment = Comment("Celda vacía", "Validador")
                continue

            if valor_a != valor_b:
                celda.fill = rojo
                celda.comment = Comment(
                    f'Valor diferente:\nEsperado: "{valor_a}"\nEncontrado: "{valor_b}"',
                    "Validador"
                )

            if tipo_esperado in validadores:
                if not validadores[tipo_esperado](valor_b):
                    celda.fill = rojo
                    mensaje = f"Tipo inválido: se esperaba {tipo_esperado}"
                    celda.comment = Comment(mensaje, "Validador")

    # === GUARDAR EL ARCHIVO VALIDADO COMO .xlsx ===
    salida = archivo_b.rsplit(".", 1)[0] + "_validado.xlsx"
    wb.save(salida)
    print(f"\n✅ Validación completada. Archivo guardado como: {salida}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python validador_excel.py archivo_a.xlsx archivo_b.xlsx")
        sys.exit(1)

    validar_excel(sys.argv[1], sys.argv[2])

