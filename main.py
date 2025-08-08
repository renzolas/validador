import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
import streamlit as st
from io import BytesIO

def comparar_y_resaltar(archivo_a, archivo_b):
    df_a = pd.read_excel(archivo_a, dtype=str)
    df_b = pd.read_excel(archivo_b, dtype=str)

    # Rellenar NaN para evitar errores
    df_a = df_a.fillna("")
    df_b = df_b.fillna("")

    # Detectar columnas faltantes o sobrantes
    columnas_a = set(df_a.columns)
    columnas_b = set(df_b.columns)

    columnas_faltantes = columnas_a - columnas_b
    columnas_sobrantes = columnas_b - columnas_a

    # Alinear columnas para comparación
    columnas_comunes = list(columnas_a & columnas_b)
    df_a = df_a[columnas_comunes]
    df_b = df_b[columnas_comunes]

    # Cargar workbook B (porque ese será el exportado)
    wb = load_workbook(archivo_b)
    ws = wb.active

    rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # Marcar columnas faltantes y sobrantes
    for col in columnas_faltantes:
        nota = f"Columna faltante en archivo B"
        # No hay columna física en B, así que no se pinta, solo se registra en log
        st.warning(f"{nota}: {col}")

    for col in columnas_sobrantes:
        idx_col_b = list(df_b.columns).index(col) + 1
        for fila in range(2, len(df_b) + 2):  # Asume fila 1 = encabezados
            celda = ws.cell(row=fila, column=idx_col_b)
            celda.fill = rojo
            celda.comment = Comment(f"Columna adicional en archivo B", "Validador")

    # Comparar valores y detectar filas faltantes o sobrantes
    max_filas = max(len(df_a), len(df_b))
    for fila in range(max_filas):
        if fila >= len(df_a):
            # Fila sobrante en B
            for col_idx in range(1, len(df_b.columns) + 1):
                celda = ws.cell(row=fila + 2, column=col_idx)
                celda.fill = rojo
                celda.comment = Comment("Fila adicional no presente en archivo A", "Validador")
            continue

        if fila >= len(df_b):
            # Fila faltante en B → No se puede marcar porque no existe físicamente en archivo B
            st.warning(f"Fila {fila+2} faltante en archivo B")
            continue

        for col_idx, col_nombre in enumerate(df_a.columns, start=1):
            valor_a = str(df_a.iat[fila, col_idx - 1])
            valor_b = str(df_b.iat[fila, col_idx - 1])

            if valor_a != valor_b:
                celda = ws.cell(row=fila + 2, column=col_idx)
                celda.fill = rojo
                celda.comment = Comment(f"Se esperaba '{valor_a}' y se encontró '{valor_b}'", "Validador")

    # Guardar archivo en memoria
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Interfaz Streamlit
st.title("📊 Validador de Archivos Excel")
archivo_a = st.file_uploader("Subir Archivo A (referencia)", type=["xls", "xlsx", "xlsm"])
archivo_b = st.file_uploader("Subir Archivo B (validar)", type=["xls", "xlsx", "xlsm"])

if archivo_a and archivo_b:
    if st.button("Validar"):
        resultado = comparar_y_resaltar(archivo_a, archivo_b)
        st.success("Validación completada ✅")
        st.download_button(
            label="📥 Descargar archivo validado",
            data=resultado,
            file_name="archivo_B_validado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )



