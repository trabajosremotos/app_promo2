import streamlit as st
import pandas as pd
import requests
import io

st.set_page_config(page_title="Gestor Inscripciones y Comparador Nube", layout="wide")
st.title("📊 Gestor de Inscripciones: Consulta, Vista Previa, Mapeo Inteligente y Registros Nuevos en Excel Nube/Local")

# --- Utilidad para exportar a Excel como bytes ---
def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def obtener_enlace_descarga(url):
    if not url:
        return url
    if "download=1" not in url:
        if "?" in url:
            url = url + "&download=1"
        else:
            url = url + "?download=1"
    return url

def cargar_excel_url(url, hoja, header_row):
    url_descarga = obtener_enlace_descarga(url)
    r = requests.get(url_descarga)
    r.raise_for_status()
    return pd.read_excel(io.BytesIO(r.content), sheet_name=hoja, header=header_row, engine="openpyxl")

def cargar_excel_upload(uploaded_file, hoja, header_row):
    return pd.read_excel(uploaded_file, sheet_name=hoja, header=header_row, engine="openpyxl")

def sugerencia_mapeo(cols_bd, cols_promo):
    sugeridos = {}
    for col in cols_bd:
        col_str = str(col).lower()
        for cp in cols_promo:
            if col_str == str(cp).lower():
                sugeridos[col] = cp
                break
            elif col_str.replace(" ", "") in str(cp).lower().replace(" ", ""):
                sugeridos[col] = cp
    return sugeridos

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

with st.expander("ℹ️ Instrucciones"):
    st.write("""
    - Puedes usar archivos de la nube (OneDrive/SharePoint) o cargar archivos locales.
    - Ajusta los nombres de las hojas y las filas de encabezado según tu archivo.
    - Se mostrarán vistas previas, sugerencia de mapeo y comparación de registros nuevos.
    - Puedes descargar la plantilla de mapeo para automatizar futuras cargas.
    """)

# --- CARGA DE DATOS ---
data_sw11, data_promo = None, None
msg = None

if url_sw11:
    try:
        data_sw11 = cargar_excel_url(url_sw11, hoja_sw11, int(header_sw11))
    except Exception as e:
        msg = f"Error al cargar SW11 desde nube: {e}\n\n➡️ Si estás en una red empresarial o proxy, puede que esté bloqueado el acceso a OneDrive desde Python. Descarga manualmente el archivo e intenta cargarlo como archivo local."
elif file_sw11:
    try:
        data_sw11 = cargar_excel_upload(file_sw11, hoja_sw11, int(header_sw11))
    except Exception as e:
        msg = f"Error al cargar SW11 local: {e}"

if url_promo:
    try:
        data_promo = cargar_excel_url(url_promo, hoja_promo, int(header_promo))
    except Exception as e:
        msg = f"Error al cargar Promoción desde nube: {e}\n\n➡️ Si estás en una red empresarial o proxy, puede que esté bloqueado el acceso a OneDrive desde Python. Descarga manualmente el archivo e intenta cargarlo como archivo local."
elif file_promo:
    try:
        data_promo = cargar_excel_upload(file_promo, hoja_promo, int(header_promo))
    except Exception as e:
        msg = f"Error al cargar Promoción local: {e}"

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

    # --- MAPEADOR ---
    st.header("🔄 Mapeo de columnas SW11 ➡️ Promoción")
    sugeridos = sugerencia_mapeo(data_sw11.columns, data_promo.columns)
    if st.button("Descargar plantilla de mapeo"):
        df_plantilla = pd.DataFrame({
            "Columna SW11": list(data_sw11.columns),
            "Columna Promoción Sugerida": [sugeridos.get(col, "") for col in data_sw11.columns]
        })
        st.download_button("Descargar plantilla", data=to_excel_bytes(df_plantilla), file_name="plantilla_mapeo_sw11.xlsx")

    st.markdown("**Mapea manualmente las columnas o deja las sugeridas:**")
    mapeo = {}
    cols = st.columns(3)
    for i, col_bd in enumerate(data_sw11.columns):
        sugerida = sugeridos.get(col_bd, "")
        col_promo = cols[i%3].selectbox(f"SW11: `{col_bd}` ➡️ Promoción:", ["(Sin mapeo)"] + list(data_promo.columns), index=(list(data_promo.columns).index(sugerida)+1) if sugerida in data_promo.columns else 0, key=f"map_{col_bd}")
        if col_promo != "(Sin mapeo)":
            mapeo[col_bd] = col_promo

    # --- COMPARADOR ---
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
        st.dataframe(nuevos, use_container_width=True)
        if nuevos.shape[0] > 0:
            st.download_button(
                "Descargar nuevos registros",
                data=to_excel_bytes(nuevos),
                file_name="nuevos_registros.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Vista previa: cómo quedaría SW11 actualizado con los nuevos registros integrados (en el formato de SW11)
            nuevos_para_sw11 = pd.DataFrame()
            for col_bd, col_promo in mapeo.items():
                if col_promo in nuevos.columns:
                    nuevos_para_sw11[col_bd] = nuevos[col_promo].values
                else:
                    nuevos_para_sw11[col_bd] = pd.NA
            nuevos_para_sw11 = nuevos_para_sw11.reindex(columns=data_sw11.columns)
            sw11_actualizado = pd.concat([data_sw11, nuevos_para_sw11], ignore_index=True)

            st.header("🟢 Vista previa: Cómo quedaría SW11 actualizado")
            st.dataframe(sw11_actualizado.tail(30), use_container_width=True)
            st.download_button(
                "Descargar SW11 actualizado (.xlsx)",
                data=to_excel_bytes(sw11_actualizado),
                file_name="sw11_actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Carga ambos archivos y ajusta los parámetros para continuar.")
