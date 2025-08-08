import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ==============================
# Función de validación
# ==============================
def validar_excel(file_a, file_b):
    try:
        # Cargar ambos archivos
        df_a = pd.read_excel(file_a, dtype=str)
        df_b = pd.read_excel(file_b, dtype=str)

        # Validar que tengan las mismas columnas en el mismo orden
        if list(df_a.columns) != list(df_b.columns):
            return None, "❌ Los archivos no tienen las mismas columnas o el mismo orden. Verifica que no se hayan modificado."

        # Ejemplo de validación: Columna 'vendor style' con formato específico
        if "vendor style" in df_b.columns:
            patron = r"^[A-Za-z0-9]+$"  # solo letras y números sin espacios
            df_b["vendor style_valido"] = df_b["vendor style"].apply(lambda x: bool(re.match(patron, str(x))))

        # Guardar archivo validado en memoria
        output = BytesIO()
        df_b.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)
        return output, "✅ Validación completada con éxito."

    except Exception as e:
        return None, f"⚠️ Error al procesar los archivos: {e}"

# ==============================
# Interfaz Streamlit
# ==============================
st.set_page_config(page_title="Validador de Excel", page_icon="📊")

st.title("📊 Validador de Archivos Excel")
st.write("""
### Instrucciones de uso:
1. Sube **dos archivos Excel**:  
   - **Archivo A**: referencia original (no modificado).  
   - **Archivo B**: archivo a validar.  
2. Ambos deben tener:
   - Las **mismas columnas** en el **mismo orden**.  
   - No deben haberse eliminado ni añadido columnas.  
3. Se permite subir archivos `.xlsx` o `.xlsm` (Excel con macros).
4. El resultado validado se descargará en formato `.xlsx` como **validado.xlsx**.
""")

# Subida de archivos
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





