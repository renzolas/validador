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

# Función para comparar celdas y resaltar diferencias
def comparar_y_resaltar(archivo_a, archivo_b):
    # Leer archivos con openpyxl (pandas para DataFrames)
    df_a = pd.read_excel(archivo_a, dtype=str, engine="openpyxl")
    df_b = pd.read_excel(archivo_b, dtype=str, engine="openpyxl")

    # Convertir NaN a cadena vacía
    df_a = df_a.fillna("")
    df_b = df_b.fillna("")

    # Cargar libro original para modificar formato
    wb = load_workbook(archivo_a)
    ws = wb.active

    # Definir color de relleno para errores
    rojo = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    # Comparar celda por celda
    for fila in range(len(df_a)):
        for col in range(len(df_a.columns)):
            valor_a = str(df_a.iat[fila, col])
            valor_b = str(df_b.iat[fila, col]) if fila < len(df_b) else ""

            if valor_a != valor_b:
                # Resaltar celda en rojo
                celda_excel = ws.cell(row=fila+2, column=col+1)  # +2 para ignorar encabezado
                celda_excel.fill = rojo
                celda_excel.comment = Comment("No coincide con Archivo B", "Validador")

    # Guardar en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ======================
# SUBIDA DE ARCHIVOS
# ======================
archivo_a = st.file_uploader("📂 Sube el Archivo A (referencia)", type=["xlsx", "xlsm"])
archivo_b = st.file_uploader("📂 Sube el Archivo B (comparar)", type=["xlsx", "xlsm"])

if archivo_a and archivo_b:
    if validar_extension(archivo_a.name) and validar_extension(archivo_b.name):
        if st.button("🔍 Validar Archivos"):
            resultado = comparar_y_resaltar(archivo_a, archivo_b)
            st.success("✅ Comparación completada")

            st.download_button(
                label="📥 Descargar archivo con errores resaltados",
                data=resultado,
                file_name="resultado_validacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ Solo se permiten archivos con extensión .xlsx o .xlsm")




