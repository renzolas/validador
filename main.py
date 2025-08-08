import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
import io
import difflib

# ======================
# CONFIGURACIÓN STREAMLIT
# ======================
st.set_page_config(page_title="Validador de Archivos", layout="centered")
st.title("📊 Validador de coincidencias entre archivos Excel")

# ======================
# FUNCIONES AUXILIARES
# ======================
def validar_extension(nombre_archivo: str) -> bool:
    return nombre_archivo.lower().endswith(('.xlsx', '.xlsm'))

def comparar_y_resaltar(
    archivo_a,
    archivo_b,
    color_diferencia="FF9999",
    color_insert="FFFF99",
    agregar_comentarios=True,
    normalize_for_alignment=True,
    header=True
) -> io.BytesIO:
    """Compara archivo A (referencia) vs archivo B (a modificar) alineando filas."""
    df_a = pd.read_excel(archivo_a, dtype=str, engine="openpyxl").fillna("")
    df_b = pd.read_excel(archivo_b, dtype=str, engine="openpyxl").fillna("")

    orig_a = df_a.copy()
    orig_b = df_b.copy()

    max_cols = max(len(df_a.columns), len(df_b.columns))

    def norm_cell(val):
        s = "" if val is None else str(val)
        return s.strip().lower() if normalize_for_alignment else s

    rows_a = [
        tuple(norm_cell(df_a.iat[r, c]) if c < len(df_a.columns) else "" for c in range(max_cols))
        for r in range(len(df_a))
    ]
    rows_b = [
        tuple(norm_cell(df_b.iat[r, c]) if c < len(df_b.columns) else "" for c in range(max_cols))
        for r in range(len(df_b))
    ]

    sm = difflib.SequenceMatcher(a=rows_a, b=rows_b)
    opcodes = sm.get_opcodes()

    wb = load_workbook(archivo_b)
    ws = wb.active

    fill_diff = PatternFill(start_color=color_diferencia, end_color=color_diferencia, fill_type="solid")
    fill_insert = PatternFill(start_color=color_insert, end_color=color_insert, fill_type="solid")

    excel_offset = 2 if header else 1

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            for a_idx, b_idx in zip(range(i1, i2), range(j1, j2)):
                for col in range(max_cols):
                    val_a = orig_a.iat[a_idx, col] if col < len(orig_a.columns) else ""
                    val_b = orig_b.iat[b_idx, col] if col < len(orig_b.columns) else ""
                    if str(val_a) != str(val_b):
                        excel_row = b_idx + excel_offset
                        cell = ws.cell(row=excel_row, column=col + 1)
                        cell.fill = fill_diff
                        if agregar_comentarios:
                            cell.comment = Comment(f'Se esperaba "{val_a}" y se encontró "{val_b}"', "Validador")
        elif tag == "replace":
            len_a = i2 - i1
            len_b = j2 - j1
            min_len = min(len_a, len_b)

            for k in range(min_len):
                a_idx = i1 + k
                b_idx = j1 + k
                for col in range(max_cols):
                    val_a = orig_a.iat[a_idx, col] if col < len(orig_a.columns) else ""
                    val_b = orig_b.iat[b_idx, col] if col < len(orig_b.columns) else ""
                    if str(val_a) != str(val_b):
                        excel_row = b_idx + excel_offset
                        cell = ws.cell(row=excel_row, column=col + 1)
                        cell.fill = fill_diff
                        if agregar_comentarios:
                            cell.comment = Comment(f'Se esperaba "{val_a}" y se encontró "{val_b}"', "Validador")

            if len_b > len_a:
                for b_idx in range(j1 + min_len, j2):
                    for col in range(max_cols):
                        excel_row = b_idx + excel_offset
                        cell = ws.cell(row=excel_row, column=col + 1)
                        cell.fill = fill_insert
                        if agregar_comentarios:
                            val_b = orig_b.iat[b_idx, col] if col < len(orig_b.columns) else ""
                            cell.comment = Comment(f'Fila añadida en B. Valor: "{val_b}"', "Validador")
        elif tag == "delete":
            # Solo marcaríamos si quisieras visualmente mostrar las faltantes,
            # pero no las insertamos ni creamos hoja resumen.
            pass
        elif tag == "insert":
            for b_idx in range(j1, j2):
                for col in range(max_cols):
                    excel_row = b_idx + excel_offset
                    cell = ws.cell(row=excel_row, column=col + 1)
                    cell.fill = fill_insert
                    if agregar_comentarios:
                        val_b = orig_b.iat[b_idx, col] if col < len(orig_b.columns) else ""
                        cell.comment = Comment(f'Fila añadida en B. Valor: "{val_b}"', "Validador")

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
            try:
                resultado = comparar_y_resaltar(
                    archivo_a,
                    archivo_b,
                    color_diferencia="FF9999",
                    color_insert="FFFF99",
                    agregar_comentarios=True,
                    normalize_for_alignment=True,
                    header=True
                )
                st.success("✅ Comparación completada")
                st.download_button(
                    label="📥 Descargar archivo B validado",
                    data=resultado,
                    file_name="archivo_B_validado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Ocurrió un error durante la comparación: {e}")
    else:
        st.error("❌ Solo se permiten archivos con extensión .xlsx o .xlsm")

# Botón de refresco para reiniciar la carga de archivos
if st.button("🔄 Refrescar y cargar otros archivos"):
    st.rerun()


