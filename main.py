import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

# ==============================
# Función de validación con formato
# ==============================
def validar_excel(file_a, file_b):
    try:
        # Leer datos en DataFrames para comparación
        df_a = pd.read_excel(file_a, dtype=str)
        df_b = pd.read_excel(file_b, dtype=str)

        # Validar columnas en el mismo orden
        if list(df_a.columns) != list(df_b.columns):
            return None, "❌ Los archivos no tienen las mismas columnas o el mismo orden."

        # Cargar archivo B con openpyxl para mantener formato
        file_b.seek(0)
        wb = load_workbook(file_b)
        ws = wb.active

        # Definir formato de resaltado
        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Recorrer y comparar celda por celda
        for row in range(2, ws.max_row + 1):  # Empieza en 2 para saltar encabezado
            for col in range(1, ws.max_column + 1):
                val_a = str(df_a.iloc[row - 2, col - 1]).strip() if pd.notna(df_a.iloc[row - 2, col - 1]) else ""
                val_b = str(df_b.iloc[row - 2, col - 1]).strip() if pd.notna(df_b.iloc[row - 2, col - 1]) else ""

                if val_a != val_b:
                    cell = ws.cell(row=row, column=col)
                    cell.fill = fill
                    cell.comment = Comment("No coincide con referencia", "Validador")

        # Guardar archivo en memoria
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return output, "✅ Validación completada. Celdas diferentes resaltadas en amarillo con comentarios."

    except Exception as e:
        return None, f"⚠️ Error al procesar: {e}"

# ==============================
# Interfaz Streamlit
# ==============================
st.set_page_config(page_title="Validador de Excel", page_icon="📊")

st.title("📊 Validador de Archivos Excel")
st.write("""
### Instrucciones de uso:
1. Sube **dos archivos Excel**:  
   - **Archivo A**: referencia original.  
   - **Archivo B**: archivo a validar.  
2. Ambos deben tener **las mismas columnas en el mismo orden**.  
3. El resultado será el archivo B **idéntico** pero con:
   - Celdas diferentes resaltadas en **amarillo**.  
   - Comentario en la celda: *"No coincide con referencia"*.  
4. Se permite subir `.xlsx` o `.xlsm`.
""")

archivo_a = st.file_uploader("📁 Sube el archivo A (referencia)", type=["xlsx", "xlsm"])
archivo_b = st.file_uploader("📁 Sube el archivo B (a validar)", type=["xlsx", "xlsm"])

if archivo_a and archivo_b:
    if st.button("▶️ Validar Archivos"):
        salida, mensaje = validar_excel(archivo_a, archivo_b)
        st.write(mensaje)

        if salida:
            st.download_button(
                label="💾 Descargar archivo validado",
                data=salida,
                file_name="validado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )




