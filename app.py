"""
app.py
======
Interfaz web para el análisis de asistencia mensual — Austral Pack S.A.

Ejecutar localmente:
    streamlit run app.py
"""

import io
import streamlit as st

from analisis_asistencia import (
    construir_df,
    analizar_permisos,
    analizar_atrasos,
    analizar_salidas,
    analizar_horas,
    calcular_criticidad,
    generar_excel,
    UMBRAL_ATRASO_MAX,
)

# ===========================================================
# CONFIGURACIÓN DE PÁGINA
# ===========================================================
st.set_page_config(
    page_title="Análisis de Asistencia · Austral Pack",
    page_icon="📊",
    layout="centered",
)

st.markdown(
    """
    <style>
        .block-container { padding-top: 2rem; }
        .stDownloadButton > button { background-color: #9e050d; color: white; }
        .stDownloadButton > button:hover { background-color: #e0151e; color: white; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ===========================================================
# HEADER
# ===========================================================
st.title("📊 Análisis de Asistencia")
st.caption("Austral Pack S.A. — carga los tres archivos Excel para generar el reporte.")
st.divider()

# ===========================================================
# CARGA DE ARCHIVOS
# ===========================================================
col1, col2, col3 = st.columns(3)
with col1:
    file_general = st.file_uploader(
        "Informe General",
        type=["xlsx"],
        help="informe-asistencia-general.xlsx",
    )
with col2:
    file_semanal = st.file_uploader(
        "Informe Semanal",
        type=["xlsx"],
        help="informe-asistencia-semanal.xlsx",
    )
with col3:
    file_nomina = st.file_uploader(
        "Nómina para turnos",
        type=["xlsx"],
        help="Nomina para turnos.xlsx",
    )

st.divider()

# ===========================================================
# BOTÓN GENERAR
# ===========================================================
if file_general and file_semanal and file_nomina:
    if st.button("Generar reporte", type="primary", use_container_width=True):
        # Limpiar resultado anterior si se vuelve a generar
        for key in ("excel_bytes", "crit", "n_areas", "n_personas", "min_atraso", "min_perdidos"):
            st.session_state.pop(key, None)

        try:
            with st.status("Procesando archivos...", expanded=True) as status:

                st.write("Cargando datos...")
                df, df_gen, df_nom = construir_df(file_general, file_semanal, file_nomina)

                st.write("Analizando permisos...")
                perms = analizar_permisos(df_gen, df_nom)

                st.write("Analizando atrasos...")
                df_atr, personas_atr, por_area_atr, max_atr = analizar_atrasos(df)

                mask_raw = (
                    ~df["con_permiso"]
                    & df["et_min"].notna()
                    & df["er_min"].notna()
                )
                n_anomalias = int(
                    (
                        (df[mask_raw]["er_min"] >= df[mask_raw]["st_min"].fillna(9999))
                        | (
                            (df[mask_raw]["er_min"] - df[mask_raw]["et_min"])
                            > UMBRAL_ATRASO_MAX
                        )
                    ).sum()
                )

                st.write("Analizando salidas...")
                por_area_sal = analizar_salidas(df)

                st.write("Analizando horas...")
                por_persona_hrs, por_area_hrs, perdida = analizar_horas(df)

                st.write("Calculando índice de criticidad...")
                crit = calcular_criticidad(df_atr)

                st.write("Generando Excel...")
                resultados = (
                    df_gen, perms, df_atr, personas_atr, por_area_atr, max_atr,
                    por_area_sal, por_persona_hrs, por_area_hrs, perdida,
                    crit, n_anomalias,
                )
                output = io.BytesIO()
                generar_excel(resultados, output)
                output.seek(0)

                # Guardar en session_state para que persista tras el re-render
                st.session_state["excel_bytes"]  = output.getvalue()
                st.session_state["crit"]         = crit
                st.session_state["n_areas"]      = len(crit)
                st.session_state["n_personas"]   = int(crit["n_personas"].sum())
                st.session_state["min_atraso"]   = int(crit["minutos_atraso_total"].sum())
                st.session_state["min_perdidos"] = int(perdida["minutos_perdidos"].sum())

                status.update(label="Reporte generado con éxito ✓", state="complete")

        except Exception as e:
            st.error(f"Error al procesar los archivos: {e}")
            st.exception(e)

else:
    st.info("Carga los tres archivos Excel para continuar.")

# ===========================================================
# RESULTADO
# ===========================================================
if "excel_bytes" in st.session_state:
    st.success("El reporte está listo para descargar.")

    # KPIs rápidos
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Áreas", st.session_state["n_areas"])
    c2.metric("Personas c/ atraso", st.session_state["n_personas"])
    c3.metric("Min atraso total", f"{st.session_state['min_atraso']:,}")
    c4.metric("Min perdidos", f"{st.session_state['min_perdidos']:,}")

    crit = st.session_state["crit"]
    st.markdown(f"**Área más crítica:** {crit.iloc[0]['area']} — índice {crit.iloc[0]['indice_0_10']} / 10")

    st.download_button(
        label="⬇️  Descargar Excel",
        data=st.session_state["excel_bytes"],
        file_name="analisis_asistencia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
