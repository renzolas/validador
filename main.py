import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from datetime import datetime
import re
import tempfile
import os
import time

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

def validar_excel(archivo_a_path, archivo_b_path):
    df_a = pd.read_excel(archivo_a_path, sheet_name=0, dtype=str)
    df_b = pd.read_excel(archivo_b_path, sheet_name=0, dtype=str)

    df_a.columns = [normalizar_columna(col) for col in df_a.columns]
    df_b.columns = [normalizar_columna(col) for col in df_b.columns]

    faltantes = set(df_a.columns) - set(df_b.columns)
    if faltantes:
        st.error(f"❌ Faltan columnas en B: {faltantes}")
        return None, None

    df_b = df_b[df_a.columns]

    wb = load_workbook(archivo_b_path, keep_vba=True)
    ws = wb.active
    rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    total_celdas = df_a.shape[0] * df_a.shape[1]
    progreso = 0
    barra = st.progress(0)
    tiempo_estimado = total_celdas * 0.02  # 0.02 seg por celda aprox
    texto_tiempo = st.empty()

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
            elif valor_a != valor_b:
                celda.fill = rojo
                celda.comment = Comment(
                    f'Valor diferente:\nEsperado: "{valor_a}"\nEncontrado: "{valor_b}"',
                    "Validador"
                )
            elif tipo_esperado in validadores and not validadores[tipo_esperado](valor_b):
                celda.fill = rojo
                celda.comment = Comment(f"Tipo inválido: se esperaba {tipo_esperado}", "Validador")

            progreso += 1
            porcentaje = int((progreso / total_celdas) * 100)
            barra.progress(porcentaje)
            tiempo_restante = tiempo_estimado * (1 - progreso / total_celdas)
            texto_tiempo.text(f"⏳ Tiempo estimado restante: {tiempo_restante:.1f} segundos")

            time.sleep(0.002)  # Simulación ligera para que se vea el avance

    # Guardar con formato (.xlsx)
    salida_xlsx = os.path.splitext(archivo_b_path)[0] + "_validado.xlsx"
    wb.save(salida_xlsx)

    # Guardar sin formato como .xls
    salida_xls = os.path.splitext(archivo_b_path)[0] + "_validado.xls"
    df_b.to_excel(salida_xls, index=False)

    return salida_xlsx, salida_xls

# === STREAMLIT UI ===
st.set_page_config(page_title="Validador de Excel", page_icon="📊")
st.title("📊 Validador de Excel")
st.markdown("### ¡Bienvenido! 👋")
st.info("Esta herramienta compara dos archivos Excel (.xlsx o .xlsm), valida datos y resalta errores en **rojo** con comentarios. "
        "El archivo resultante puede descargarse en `.xlsx` (con formato) o `.xls` (sin formato, pero compatible con más sistemas).")

archivo_a = st.file_uploader("📂 Sube el archivo A (referencia)", type=["xlsx", "xlsm"])
archivo_b = st.file_uploader("📂 Sube el archivo B (validar)", type=["xlsx", "xlsm"])

if archivo_a and archivo_b:
    if st.button("🚀 Ejecutar validación"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(archivo_a.name)[1]) as tmp_a, \
             tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(archivo_b.name)[1]) as tmp_b:
            tmp_a.write(archivo_a.read())
            tmp_b.write(archivo_b.read())
            tmp_a_path = tmp_a.name
            tmp_b_path = tmp_b.name

        salida_xlsx, salida_xls = validar_excel(tmp_a_path, tmp_b_path)

        if salida_xlsx and salida_xls:
            st.success("✅ Validación completada con éxito.")
            with open(salida_xlsx, "rb") as f1:
                st.download_button("📥 Descargar en .xlsx (con formato)", f1, file_name=os.path.basename(salida_xlsx))
            with open(salida_xls, "rb") as f2:
                st.download_button("📥 Descargar en .xls (sin formato)", f2, file_name=os.path.basename(salida_xls))
else:
    st.warning("Por favor, sube **ambos archivos** antes de ejecutar la validación.")



