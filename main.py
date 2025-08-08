import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
import tempfile
import time

# --- Función de validación ---
def validar_excel(archivo_a_path, archivo_b_path):
    # Leer ambos archivos
    df_a = pd.read_excel(archivo_a_path)
    df_b = pd.read_excel(archivo_b_path)

    # Cargar archivo B con openpyxl para poder modificarlo
    wb = load_workbook(archivo_b_path)
    ws = wb.active

    # Ejemplo: comparar columna 1 de A con columna 1 de B
    col_a = df_a.columns[0]
    col_b = df_b.columns[0]
    set_a = set(df_a[col_a])

    for idx, valor in enumerate(df_b[col_b], start=2):
        if valor not in set_a:
            ws[f"A{idx}"].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            ws[f"A{idx}"].comment = Comment("No coincide con archivo A", "Validador")

    salida = "validado.xlsx"
    wb.save(salida)
    return salida

# --- Configuración de página ---
st.set_page_config(page_title="Validador de Excel", page_icon="✅", layout="centered")

# --- Mensaje de bienvenida ---
st.title("✅ Validador de Archivos Excel")
st.markdown("""
¡Bienvenido al **Validador de Excel**!  
Sube el **Archivo A** (referencia) y el **Archivo B** (validar).  
Este sistema resaltará en rojo las celdas de B que **no estén en A**.
""")

# --- Subida de archivos ---
archivo_a = st.file_uploader("📂 Sube el archivo A (referencia)", type=["xlsx", "xlsm"])
archivo_b = st.file_uploader("📂 Sube el archivo B (validar)", type=["xlsx", "xlsm"])

# --- Botón para iniciar validación ---
if archivo_a and archivo_b:
    if st.button("🚀 Iniciar validación"):
        # Guardar archivos temporales
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_a:
            tmp_a.write(archivo_a.read())
            tmp_a_path = tmp_a.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_b:
            tmp_b.write(archivo_b.read())
            tmp_b_path = tmp_b.name

        # Barra de progreso con cuenta regresiva
        progreso = st.progress(0)
        cuenta = st.empty()
        for i in range(5, 0, -1):
            progreso.progress((5 - i) * 20)
            cuenta.write(f"⏳ Validando... {i} segundos restantes")
            time.sleep(1)
        progreso.progress(100)
        cuenta.write("✅ Validación completa")

        # Ejecutar validación
        archivo_salida = validar_excel(tmp_a_path, tmp_b_path)

        # Botón para descargar
        with open(archivo_salida, "rb") as f:
            st.download_button(
                label="💾 Descargar archivo validado",
                data=f,
                file_name="validado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )




