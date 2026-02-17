import streamlit as st
import pandas as pd
import io
from io import StringIO

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Comparación L / R", layout="wide")
st.title("Comparación Vs CMM")

# =========================
# CARGA DE ARCHIVOS
# =========================
archivo_L = st.file_uploader(
    "Carga el archivo TXT - Lado Izquierdo",
    type=["txt"]
)

archivo_R = st.file_uploader(
    "Carga el archivo TXT - Lado Derecho",
    type=["txt"]
)

# =========================
# LECTOR ROBUSTO TXT
# =========================
def leer_txt(archivo):
    contenido = archivo.read().decode("utf-8", errors="ignore")
    lineas = contenido.splitlines()

    inicio = None
    for i, linea in enumerate(lineas):
        if (
            "Cycle Time" in linea and
            "Corr. Coef." in linea and
            "Offset" in linea and
            "T-Test" in linea and
            "F-Test" in linea
        ):
            inicio = i
            break

    if inicio is None:
        st.error("No se encontró la fila de encabezados reales")
        st.stop()

    datos = "\n".join(lineas[inicio:])

    df = pd.read_csv(
        StringIO(datos),
        sep=r"\t+",
        engine="python",
        header=0,
        on_bad_lines="skip"
    )

    return df

# =========================
# FUNCIONES DE COLOR
# =========================
def color_t_test(val):
    try:
        val = float(val)
        return (
            "background-color: #00C853; color: black; font-weight: bold"
            if val < 0.005
            else "background-color: #D50000; color: white; font-weight: bold"
        )
    except:
        return ""

def color_f_test(val):
    try:
        val = float(val)
        return (
            "background-color: #D50000; color: white; font-weight: bold"
            if val < 0.005
            else "background-color: #00C853; color: black; font-weight: bold"
        )
    except:
        return ""

def color_corr(val):
    try:
        val = float(val)
        if val >= 0.95:
            return "background-color: #00C853; color: black; font-weight: bold"
        elif 0.90 <= val <= 0.94:
            return "background-color: #FFD600; color: black; font-weight: bold"
        else:
            return "background-color: #D50000; color: white; font-weight: bold"
    except:
        return ""

def color_offset(val):
    try:
        val = float(val)
        return (
            "background-color: #AA00FF; color: white; font-weight: bold"
            if abs(val) > 0.5
            else ""
        )
    except:
        return ""


# =========================
# PROCESO PRINCIPAL
# =========================
if archivo_L and archivo_R:

    df_L = leer_txt(archivo_L)
    df_R = leer_txt(archivo_R)

    df = pd.concat([df_L, df_R], ignore_index=True)

    # =========================
    # LIMPIEZA
    # =========================
    col_cycle = df.columns[0]

    df[col_cycle] = df[col_cycle].astype(str).str.strip()
    df = df[df[col_cycle] != ""]
    df = df[~df[col_cycle].str.startswith("CT", na=False)]

    # =========================
    # EXTRAER NOMBRE / EJE
    # =========================
    df["Nombre"] = df[col_cycle].str.extract(r"(^\d+)")
    df["Eje"] = df[col_cycle].str.extract(r"\[([A-Z])\]")

    # =========================
    # FILTROS INTERACTIVOS
    # =========================
    st.sidebar.header("Filtros")

    nombres = sorted(df["Nombre"].dropna().unique())
    ejes = sorted(df["Eje"].dropna().unique())

    nombre_sel = st.sidebar.multiselect(
        "Filtrar por nombre (Cycle Time)",
        nombres,
        default=nombres
    )

    eje_sel = st.sidebar.multiselect(
        "Filtrar por eje",
        ejes,
        default=ejes
    )

    df_filtrado = df[
        df["Nombre"].isin(nombre_sel) &
        df["Eje"].isin(eje_sel)
    ]

    # =========================
    # ORDEN TIPO EXCEL
    # =========================
    def orden_excel(nombre):
        nombre = str(nombre)
        lado = 0 if "L[" in nombre else 1
        eje = 0 if "[Y]" in nombre else 1
        return lado, eje

    df_filtrado["__orden"] = df_filtrado[col_cycle].apply(orden_excel)
    df_filtrado = df_filtrado.sort_values("__orden").drop(columns="__orden")

    # =========================
    # SALIDA
    # =========================
    st.success("Filtro aplicado correctamente")
    st.write("Filas visibles:", df_filtrado.shape[0])

    styled_df = (
        df_filtrado.style
        .applymap(color_t_test, subset=["T-Test"])
        .applymap(color_f_test, subset=["F-Test"])
        .applymap(color_corr, subset=["Corr. Coef."])
        .applymap(color_offset, subset=["Offset"])
    )

    st.dataframe(styled_df, use_container_width=True)

    # =========================
    # DESCARGAR A EXCEL PRO (CON COLORES + FILTROS + AUTOAJUSTE + KPIs)
    # =========================
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name="Comparacion")

        workbook = writer.book
        worksheet = writer.sheets["Comparacion"]

        # =========================
        # FORMATOS
        # =========================
        format_green = workbook.add_format({"bg_color": "#00C853", "font_color": "black", "bold": True})
        format_red = workbook.add_format({"bg_color": "#D50000", "font_color": "white", "bold": True})
        format_yellow = workbook.add_format({"bg_color": "#FFD600", "font_color": "black", "bold": True})
        format_purple = workbook.add_format({"bg_color": "#AA00FF", "font_color": "white", "bold": True})
        format_header = workbook.add_format({"bold": True, "border": 1})
        format_border = workbook.add_format({"border": 1})

        # =========================
        # FORMATO ENCABEZADO
        # =========================
        for col_num, value in enumerate(df_filtrado.columns):
            worksheet.write(0, col_num, value, format_header)

        # =========================
        # AUTO AJUSTAR COLUMNAS
        # =========================
        for i, col in enumerate(df_filtrado.columns):
            max_len = max(
                df_filtrado[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            worksheet.set_column(i, i, max_len)

        # =========================
        # CONGELAR ENCABEZADO
        # =========================
        worksheet.freeze_panes(1, 0)

        # =========================
        # FILTRO AUTOMÁTICO
        # =========================
        worksheet.autofilter(
            0, 0,
            df_filtrado.shape[0],
            df_filtrado.shape[1] - 1
        )

        # =========================
        # COLORES CELDA POR CELDA
        # =========================
        col_t = df_filtrado.columns.get_loc("T-Test")
        col_f = df_filtrado.columns.get_loc("F-Test")
        col_corr = df_filtrado.columns.get_loc("Corr. Coef.")
        col_offset = df_filtrado.columns.get_loc("Offset")

        for row_num, row in df_filtrado.iterrows():
            excel_row = row_num + 1

            # Bordes generales
            for col_num in range(len(df_filtrado.columns)):
                valor = row[col_num]

                if pd.isna(valor) or valor in [float("inf"), float("-inf")]:
                    worksheet.write(excel_row, col_num, "", format_border)
                else:
                    worksheet.write(excel_row, col_num, valor, format_border)


            # T-Test
            try:
                if float(row["T-Test"]) < 0.005:
                    worksheet.write(excel_row, col_t, row["T-Test"], format_green)
                else:
                    worksheet.write(excel_row, col_t, row["T-Test"], format_red)
            except:
                pass

            # F-Test
            try:
                if float(row["F-Test"]) < 0.005:
                    worksheet.write(excel_row, col_f, row["F-Test"], format_red)
                else:
                    worksheet.write(excel_row, col_f, row["F-Test"], format_green)
            except:
                pass

            # Corr. Coef.
            try:
                val = float(row["Corr. Coef."])
                if val >= 0.95:
                    worksheet.write(excel_row, col_corr, val, format_green)
                elif 0.90 <= val <= 0.94:
                    worksheet.write(excel_row, col_corr, val, format_yellow)
                else:
                    worksheet.write(excel_row, col_corr, val, format_red)
            except:
                pass

            # Offset
            try:
                val = float(row["Offset"])
                if abs(val) > 0.5:
                    worksheet.write(excel_row, col_offset, val, format_purple)
            except:
                pass

        # =========================
        # HOJA RESUMEN KPI
        # =========================
        resumen = workbook.add_worksheet("Resumen")

        total = len(df_filtrado)
        fallas_t = (pd.to_numeric(df_filtrado["T-Test"], errors="coerce") >= 0.005).sum()
        fallas_corr = (pd.to_numeric(df_filtrado["Corr. Coef."], errors="coerce") < 0.95).sum()
        offsets_altos = (pd.to_numeric(df_filtrado["Offset"], errors="coerce").abs() > 0.5).sum()

        resumen.write("A1", "Total mediciones")
        resumen.write("B1", total)

        resumen.write("A2", "T-Test fuera de rango")
        resumen.write("B2", fallas_t)

        resumen.write("A3", "Correlación baja (<0.95)")
        resumen.write("B3", fallas_corr)

        resumen.write("A4", "Offset alto (>0.5)")
        resumen.write("B4", offsets_altos)

    excel_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Excel PRO",
        data=excel_data,
        file_name="Comparacion_CMM_PRO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Carga ambos archivos TXT para continuar")

#streamlit run "C:\Users\maripes3\Documents\Comparaciones\MIS DATOS\MIS DATOS\ComparacionVsCMM.py"