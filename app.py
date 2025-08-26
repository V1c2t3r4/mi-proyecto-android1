import io
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Capacidad por Subestación", layout="wide")
st.title("Capacidad por Subestación – Demo sin Tkinter")

# Subida de archivos
col1, col2 = st.columns(2)
with col1:
    f_trans = st.file_uploader("Sube Capacidad_Transformadores.xlsx", type=["xlsx"])
with col2:
    f_info = st.file_uploader("Sube informacion_importante.xlsx", type=["xlsx"])

if f_trans and f_info:
    try:
        df_trans = pd.read_excel(f_trans)
        df_info  = pd.read_excel(f_info)
        st.success("Archivos leídos correctamente.")
    except Exception as e:
        st.error(f"Error leyendo archivos: {e}")
        st.stop()

    # Normalizaciones básicas de ejemplo
    df_trans['Nombre Subestación'] = df_trans['Nombre Subestación'].astype(str).str.upper().str.strip()
    df_trans['Capacidad'] = pd.to_numeric(df_trans['Capacidad'], errors='coerce').fillna(0)

    df_info['SUBESTACION'] = df_info['SUBESTACION'].astype(str).str.upper().str.strip()
    df_info['SUBESTACION_EQ'] = df_info['SUBESTACION'].apply(lambda x: x if x.startswith("S/E ") else "S/E " + x)
    df_info['POTENCIA_MW'] = pd.to_numeric(df_info['POTENCIA_MW'], errors='coerce').fillna(0)

    # Resumen simple: capacidad por subestación vs potencia conectada (CONECTADO/ICC/SCR)
    estados_descuentan = {"CONECTADO", "ICC", "SCR"}
    df_cap = df_trans.groupby('Nombre Subestación', as_index=False)['Capacidad'].sum()

    df_desc = (df_info[df_info['ESTADO_PMGD'].astype(str).str.upper().isin(estados_descuentan)]
               .groupby('SUBESTACION_EQ', as_index=False)['POTENCIA_MW'].sum()
               .rename(columns={'SUBESTACION_EQ':'Nombre Subestación',
                                'POTENCIA_MW':'Potencia Descontada (MW)'}))

    resumen = df_cap.merge(df_desc, on='Nombre Subestación', how='left').fillna({'Potencia Descontada (MW)':0})
    resumen['Capacidad Disponible (MW)'] = resumen['Capacidad'] - resumen['Potencia Descontada (MW)']

    st.subheader("Resumen")
    st.dataframe(resumen.sort_values('Nombre Subestación').reset_index(drop=True))

    # Descargar Excel con timestamp (opcional)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        resumen.to_excel(w, index=False, sheet_name="RESUMEN")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    st.download_button(
        "Descargar RESUMEN en Excel",
        data=out.getvalue(),
        file_name=f"Resumen_Subestaciones_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Sube ambos archivos para continuar.")
