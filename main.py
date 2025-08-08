import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

def comparar_y_resaltar(archivo_a, archivo_b):
    # === Leer archivos como DataFrames ===
    df_a = pd.read_excel(archivo_a, dtype=str).fillna("")
    df_b = pd.read_excel(archivo_b, dtype=str).fillna("")

    # === Alinear columnas por nombre ===
    columnas_comunes = [col for col in df_a.columns if col in df_b.columns]
    columnas_solo_a = [col for col in df_a.columns if col not in df_b.columns]
    columnas_solo_b = [col for col in df_b.columns if col not in df_a.columns]

    # Crear copia de B para no modificar original
    df_b_alineado = df_b.copy()

    # Agregar columnas faltantes en B
    for col in columnas_solo_a:
        df_b_alineado[col] = ""

    # Mantener el orden de columnas como en A
    columnas_finales = columnas_comunes + columnas_solo_a
    df_b_alineado = df_b_alineado[columnas_finales]

    # === Alinear filas por clave (primera columna) ===
    clave_col = df_a.columns[0]  # Usar la primera columna como clave
    df_a_indexado = df_a.set_index(clave_col)
    df_b_indexado = df_b_alineado.set_index(clave_col)

    # Unir claves
    todas_claves = sorted(set(df_a_indexado.index) | set(df_b_indexado.index))

    # Crear libro para exportar
    df_b_export = df_b_alineado.reindex(columns=columnas_finales)
    ruta_export = "archivo_B_validado.xlsx"
    df_b_export.to_excel(ruta_export, index=False)

    # === Cargar libro con openpyxl ===
    wb = load_workbook(ruta_export)
    ws = wb.active

    # Colores
    rojo = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    # === Comparar fila a fila ===
    for clave in todas_claves:
        if clave not in df_a_indexado.index:
            # Fila sobrante en B
            fila_ws = ws.max_row + 1
            ws.append([clave] + [""] * (len(columnas_finales) - 1))
            for col in range(1, len(columnas_finales) + 1):
                ws.cell(row=fila_ws, column=col).fill = rojo
                ws.cell(row=fila_ws, column=col).comment = Comment("Fila presente en B pero no en A", "Validador")
        elif clave not in df_b_indexado.index:
            # Fila faltante en B
            fila_ws = ws.max_row + 1
            ws.append([clave] + [""] * (len(columnas_finales) - 1))
            for col in range(1, len(columnas_finales) + 1):
                ws.cell(row=fila_ws, column=col).fill = rojo
                ws.cell(row=fila_ws, column=col).comment = Comment("Fila presente en A pero no en B", "Validador")
        else:
            # Fila presente en ambos → comparar columnas
            fila_a = df_a_indexado.loc[clave]
            fila_b = df_b_indexado.loc[clave]

            # Si hay varias filas con la misma clave, convertir en lista
            if isinstance(fila_a, pd.DataFrame):
                fila_a = fila_a.iloc[0]
            if isinstance(fila_b, pd.DataFrame):
                fila_b = fila_b.iloc[0]

            # Buscar índice real en hoja Excel
            fila_ws = list(df_b_export[clave_col]).index(clave) + 2  # +2 por encabezado y base 1

            for col_idx, col_name in enumerate(columnas_finales, start=1):
                if col_name in columnas_solo_a:
                    # Columna faltante en B
                    ws.cell(row=fila_ws, column=col_idx).fill = rojo
                    ws.cell(row=fila_ws, column=col_idx).comment = Comment("Columna presente en A pero no en B", "Validador")
                elif col_name in columnas_solo_b:
                    # Columna sobrante en B (en teoría no debería pasar por alineación, pero lo dejamos por seguridad)
                    ws.cell(row=fila_ws, column=col_idx).fill = rojo
                    ws.cell(row=fila_ws, column=col_idx).comment = Comment("Columna presente en B pero no en A", "Validador")
                else:
                    val_a = str(fila_a[col_name])
                    val_b = str(fila_b[col_name])
                    if val_a != val_b:
                        ws.cell(row=fila_ws, column=col_idx).fill = rojo
                        ws.cell(row=fila_ws, column=col_idx).comment = Comment(f"Se esperaba '{val_a}' y se encontró '{val_b}'", "Validador")

    wb.save(ruta_export)
    return ruta_export

# === Ejemplo de uso ===
# resultado = comparar_y_resaltar("archivo_A.xlsx", "archivo_B.xlsx")
# print(f"Archivo exportado: {resultado}")

