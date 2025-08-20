import streamlit as st
import pandas as pd
import io
import os
import gc

# ---- psutil opcional ----
try:
    import psutil
except Exception:
    psutil = None

st.set_page_config(page_title="Gestor Inscripciones y Comparador Nube", layout="wide")
st.title("📊 Gestor de Inscripciones: Consulta, Vista Previa, Mapeo Inteligente y Registros Nuevos en Excel Nube/Local")

def mostrar_memoria(prefix=""):
    if psutil is None:
        st.sidebar.markdown("💾 **Uso de memoria:** N/D (instala psutil)")
        return
    process = psutil.Process(os.getpid())
    memoria_mb = process.memory_info().rss / 1024 / 1024
    st.sidebar.markdown(f"💾 **{prefix}Uso de memoria:** {memoria_mb:.2f} MB")

mostrar_memoria("Inicio • ")

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def obtener_enlace_descarga(url: str):
    if not url:
        return url
    if "download=1" not in url:
        if "?" in url:
            url = url + "&download=1"
        else:
            url = url + "?download=1"
    return url

@st.cache_data(show_spinner=False)
def url_to_bytes(url: str) -> bytes:
    import requests
    url_descarga = obtener_enlace_descarga(url)
    r = requests.get(url_descarga, timeout=60)
    r.raise_for_status()
    return r.content

@st.cache_data(show_spinner=False)
def read_excel_from_bytes(file_bytes: bytes, sheet: str, header_row: int, usecols=None) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=header_row, usecols=usecols, engine="openpyxl")

@st.cache_data(show_spinner=False)
def get_columns_from_bytes(file_bytes: bytes, sheet: str, header_row: int):
    df_head = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=header_row, nrows=0, engine="openpyxl")
    return list(df_head.columns)

st.sidebar.header("Fuente SW11 (BDUnidad)")
url_sw11 = st.sidebar.text_input("Enlace compartido SW11 (OneDrive/SharePoint)")
file_sw11 = st.sidebar.file_uploader("O cargar archivo SW11 local (.xlsx)", type="xlsx", key="sw11_file")
hoja_sw11 = st.sidebar.text_input("Nombre hoja SW11", value="bduNIDAD")
header_sw11 = st.sidebar.number_input("Fila de encabezado SW11 (0-indexada)", value=0, min_value=0)

st.sidebar.header("Fuente Promoción")
url_promo = st.sidebar.text_input("Enlace compartido Promoción (OneDrive/SharePoint)")
file_promo = st.sidebar.file_uploader("O cargar archivo Promoción local (.xlsx)", type="xlsx", key="promo_file")
hoja_promo = st.sidebar.text_input("Nombre hoja Promoción", value="Tecnico")
header_promo = st.sidebar.number_input("Fila de encabezado Promoción (0-indexada)", value=1, min_value=0)

st.sidebar.markdown("---")
opt_cols = st.sidebar.checkbox("⚡ Seleccionar columnas antes de cargar (ahorra memoria)", value=True)
st.sidebar.markdown("---")

with st.expander("ℹ️ Instrucciones"):
    st.write("Activa **Seleccionar columnas** para cargar solo lo necesario y ahorrar memoria.")

data_sw11, data_promo = None, None
sw11_bytes, promo_bytes = None, None
msg = None

try:
    if url_sw11:
        sw11_bytes = url_to_bytes(url_sw11)
    elif file_sw11:
        sw11_bytes = file_sw11.getvalue()

    if url_promo:
        promo_bytes = url_to_bytes(url_promo)
    elif file_promo:
        promo_bytes = file_promo.getvalue()
except Exception as e:
    msg = f"Error al descargar/leer archivo: {e}"

usecols_sw11 = None
usecols_promo = None
try:
    if opt_cols and sw11_bytes:
        cols_sw11_all = get_columns_from_bytes(sw11_bytes, hoja_sw11, int(header_sw11))
        sugeridos_bd = ["Cédula", "Primer nombre", "Mail", "Teléfono", "Nombre programa", "Estado", "Cohorte"]
        default_bd = [c for c in sugeridos_bd if c in cols_sw11_all] or cols_sw11_all
        usecols_sw11 = st.sidebar.multiselect("Columnas a cargar de SW11:", cols_sw11_all, default=default_bd)
    if opt_cols and promo_bytes:
        cols_promo_all = get_columns_from_bytes(promo_bytes, hoja_promo, int(header_promo))
        sugeridos_pr = ["Número de Documento de Identidad", "Nombre", "Correo", "Número de teléfono", "Programa", "Estados", "Periodo Académico"]
        default_pr = [c for c in sugeridos_pr if c in cols_promo_all] or cols_promo_all
        usecols_promo = st.sidebar.multiselect("Columnas a cargar de Promoción:", cols_promo_all, default=default_pr)
except Exception as e:
    msg = f"No se pudieron leer columnas para la selección previa: {e}"

try:
    if sw11_bytes:
        data_sw11 = read_excel_from_bytes(sw11_bytes, hoja_sw11, int(header_sw11), usecols=usecols_sw11)
    if promo_bytes:
        data_promo = read_excel_from_bytes(promo_bytes, hoja_promo, int(header_promo), usecols=usecols_promo)
except Exception as e:
    msg = f"Error al abrir Excel: {e}"

if msg:
    st.error(msg)

if (data_sw11 is not None) and (data_promo is not None):
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Vista previa SW11")
        st.dataframe(data_sw11.head(20), use_container_width=True)
    with col2:
        st.subheader("Vista previa Promoción")
        st.dataframe(data_promo.head(20), use_container_width=True)
    st.markdown("---")

    mostrar_memoria("Post-carga • ")

    st.header("🔄 Mapeo de columnas SW11 ➡️ Promoción")
    sugeridos = {col: col for col in data_sw11.columns if col in data_promo.columns}
    if st.button("Descargar plantilla de mapeo"):
        df_plantilla = pd.DataFrame({
            "Columna SW11": list(data_sw11.columns),
            "Columna Promoción Sugerida": [sugeridos.get(col, "") for col in data_sw11.columns]
        })
        st.download_button("Descargar plantilla", data=to_excel_bytes(df_plantilla), file_name="plantilla_mapeo_sw11.xlsx")

    mapeo = {}
    cols = st.columns(3)
    for i, col_bd in enumerate(data_sw11.columns):
        sugerida = sugeridos.get(col_bd, "")
        col_promo = cols[i%3].selectbox(
            f"SW11: `{col_bd}` ➡️ Promoción:",
            ["(Sin mapeo)"] + list(data_promo.columns),
            index=(list(data_promo.columns).index(sugerida)+1) if sugerida in data_promo.columns else 0,
            key=f"map_{col_bd}"
        )
        if col_promo != "(Sin mapeo)":
            mapeo[col_bd] = col_promo

    st.header("🆕 Registros nuevos en Promoción (no están en SW11)")
    if len(mapeo) == 0:
        st.warning("Realiza al menos un mapeo para comparar.")
    else:
        col_bd = list(mapeo.keys())[0]
        col_promo = mapeo[col_bd]
        st.info(f"Comparando por: **SW11:** `{col_bd}` ➡️ **Promoción:** `{col_promo}`")

        bd_ids = data_sw11[col_bd].dropna().astype(str).str.strip().str.lower().unique()
        promo_ids = data_promo[col_promo].dropna().astype(str).str.strip().str.lower()
        nuevos_mask = ~promo_ids.isin(bd_ids)

        nuevos = data_promo.loc[nuevos_mask]
        st.success(f"Total registros nuevos: {nuevos.shape[0]}")
        st.dataframe(nuevos.head(500), use_container_width=True)

        if nuevos.shape[0] > 0:
            st.download_button("Descargar nuevos registros (.xlsx)", data=to_excel_bytes(nuevos),
                               file_name="nuevos_registros.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            nuevos_para_sw11 = pd.DataFrame()
            for col_bd_, col_promo_ in mapeo.items():
                if col_promo_ in nuevos.columns:
                    nuevos_para_sw11[col_bd_] = nuevos[col_promo_].values
                else:
                    nuevos_para_sw11[col_bd_] = pd.NA
            nuevos_para_sw11 = nuevos_para_sw11.reindex(columns=data_sw11.columns)
            sw11_actualizado = pd.concat([data_sw11, nuevos_para_sw11], ignore_index=True)

            st.header("🟢 Vista previa: Cómo quedaría SW11 actualizado")
            st.dataframe(sw11_actualizado.tail(30), use_container_width=True)
            st.download_button("Descargar SW11 actualizado (.xlsx)", data=to_excel_bytes(sw11_actualizado),
                               file_name="sw11_actualizado.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            mostrar_memoria("Post-mapeo • ")
            del nuevos_para_sw11, sw11_actualizado
            gc.collect()
else:
    st.info("Carga ambos archivos y ajusta los parámetros para continuar.")
