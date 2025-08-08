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
# FUNCIONES AUXILIARES
# ======================
def validar_extension(nombre_archivo: str) -> bool:
    """Verifica que el archivo tenga extensión válida."""
    return nombre_archivo.lower().endswith(('.xlsx', '.xlsm'))

def comparar_y_resaltar(archivo_a, archivo_b, color_hex="FF9999", agregar_comentarios=True) -> io.BytesIO:
    """
    Compara dos archivos Excel y resalta las diferencias en el segundo archivo.
    
    Parámetros:
        archivo_a: Archivo de referencia.
        archivo_b: Archivo a modificar.
        color_hex: Color de fondo para diferencias en formato HEX (sin '#').
        agregar_comentarios: Si True, añade un comentario a cada celda diferente.
        
    Retorna:
        BytesIO con el archivo modificado.
    """
    # Leer ambos archivos en DataFrames
    df_a = pd.read_excel(archivo_a, dtype=str, engine="openpyxl").fillna("")
    df_b = pd.read_excel(archivo_b, dtype=str, engine="openpyxl").fillna("")

    # Cargar archivo B en openpyxl
    wb = load_workbook(archivo_b)
    ws = wb.active

    # Definir estilo
    relleno = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

    # Determinar límites máximos para comparación
    max_filas = max(len(df_a), len(df_b))
    max_cols = max(len(df_a.columns), len(df_b.columns))

    # Recorrer celdas
    for fila in range(max_filas):
        for col in range(max_cols):
            valor_a = str(df_a.iat[fila, col]) if fila < len(df_a) and col < len(df_a.columns) else ""
            valor_b = str(df_b.iat[fila, col]) if fila < len(df_b) and col < len(df_b.columns) else ""

            if valor_a != valor_b:
                celda_excel = ws.cell(row=fila + 2, column=col + 1)  # +2 asume encabezado
                celda_excel.fill = relleno

                if agregar_comentarios:
                    comentario_texto = f'Se esperaba "{valor_a}" y se encontró "{valor_b}"'
                    celda_excel.comment = Comment(comentario_texto, "Validador")

    # Guardar resultado en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ======================
# INTERFAZ STREAMLIT
# ======================
archivo_a = st.file_uploader("📂 Sube el Archivo A (referencia)", type=["xlsx", "xlsm"])
archivo_b = st.file_uploader("📂 Sube el Archivo B (comparar y modificar)", type=["xlsx", "xlsm"])

if archivo_a and archivo_b:
    if validar_extension(archivo_a.name) and validar_extension(archivo_b.name):
        if st.button("🔍 Validar Archivos"):
            resultado = comparar_y_resaltar(archivo_a, archivo_b, color_hex="FF9999", agregar_comentarios=True)
            st.success("✅ Comparación completada")

            st.download_button(
                label="📥 Descargar archivo B validado",
                data=resultado,
                file_name="archivo_B_validado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ Solo se permiten archivos con extensión .xlsx o .xlsm")

