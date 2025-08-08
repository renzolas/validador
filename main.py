import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
import io

# ======================
# CONFIGURACIÓN STREAMLIT
# ======================
st.set_page_config(page_title="Validador de Archivos", layout="centered")
st.title("📊 Validador de coincidencias entre archivos Excel")

# ======================
# FUNCIONES
# ======================

# Validar extensión
def validar_extension(nombre_archivo):
    return nombre_archivo.lower().endswith(('.xlsx', '.xlsm'))

# Función para comparar y resaltar en el archivo B
def comparar_y_resaltar(archivo_a, archivo_b):
    # Leer ambos archivos en pandas
    df_a = pd.read_excel(archivo_a, dtype=str, engine="openpyxl")
    df_b = pd.read_excel(archivo_b, dtype=str, engine="openpyxl")

    # Reemplazar NaN por cadena vacía
    df_a = df_a.fillna("")
    df_b = df_b.fillna("")

    # Cargar archivo B en openpyxl para modificarlo
    wb = load_workbook(archivo_b)
    ws = wb.active

    # Definir color de relleno para diferencias
    rojo = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    # Recorrer y comparar celda por celda
    for fila in range(len(df_b)):
        for col in range(len(df_b.columns)):
            valor_a = str(df_a.iat[fila, col]) if fila < len(df_a) else ""
            valor_b = str(df_b.iat[fila, col])

            if valor_a != valor_b:
                celda_excel = ws.cell(row=fila+2, column=col+1)  # +2 para ignorar encabezado
                celda_excel.fill = rojo
                comentario_texto = f'Se esperaba encontrar "{valor_a}" y se encontró "{valor_b}"'
                celda_excel.comment = Comment(comentario_texto, "Validador")

    # Guardar resultado en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ======================
# SUBIDA DE ARCHIVOS
# ======================
archivo_a = st.file_uploader("📂 Sube el Archivo A (referencia)", type=["xlsx", "xlsm"])
archivo_b = st.file_uploader("📂 Sube el Archivo B (comparar y modificar)", type=["xlsx", "xlsm"])

if archivo_a and archivo_b:
    if validar_extension(archivo_a.name) and validar_extension(archivo_b.name):
        if st.button("🔍 Validar Archivos"):
            resultado = comparar_y_resaltar(archivo_a, archivo_b)
            st.success("✅ Comparación completada")

            st.download_button(
                label="📥 Descargar archivo B validado",
                data=resultado,
                file_name="archivo_B_validado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ Solo se permiten archivos con extensión .xlsx o .xlsm")

