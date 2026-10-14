import sys
from pathlib import Path

# Pfad absichern, damit Imports aus src/ reibungslos funktionieren
root_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

import os
import streamlit as st
import plotly.express as px
import pandas as pd

from src.etl.importer import load_config, import_csv_data, validate_schema
from src.etl.anonymizer import anonymize_dataframe
from src.analytics.kpi_calculator import calculate_kpis
from src.reporting.excel_exporter import export_to_excel

# --- Page Configuration ---
st.set_page_config(
    page_title="AutoData Pipeline Control Center",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AutoData Pipeline – Interactive Control Center")
st.markdown("Willkommen im Cockpit! Hier kannst du die Pipeline ausführen, Daten analysieren und Berichte herunterladen.")

# --- Sidebar Configuration ---
st.sidebar.header("Pipeline-Steuerung")

config = load_config()
raw_dir = config["paths"]["raw_data_dir"]
sample_file = os.path.join(raw_dir, "sample_operations.csv")

st.sidebar.info(f"**Datenquelle:** `{sample_file}`")

run_pipeline = st.sidebar.button("🚀 Pipeline ausführen", type="primary")

# --- Pipeline Pipeline Execution Logic ---
if run_pipeline or "df_processed" in st.session_state:
    if run_pipeline:
        with st.spinner("Pipeline läuft... (Import -> Validierung -> Anonymisierung -> KPIs)"):
            # ETL & Analytics Durchlauf
            df_raw = import_csv_data(sample_file)
            df_clean = validate_schema(df_raw)
            df_anon = anonymize_dataframe(df_clean)
            kpis = calculate_kpis(df_anon)
            
            # Excel-Report generieren
            export_to_excel(df_anon, kpis)
            
            # In Session State speichern, damit Streamlit bei Interaktionen den Zustand behält
            st.session_state["df_processed"] = df_anon
            st.session_state["kpis"] = kpis
            
        st.success("Pipeline erfolgreich ausgeführt!")

    df = st.session_state["df_processed"]
    kpis = st.session_state["kpis"]

    # --- KPI Kacheln (Metrics) ---
    st.subheader(" Key Performance Indicators")
    col1, col2, col3, col4 = st.columns(4)
    
    col1.metric("Gesamtdatensätze", kpis.get("total_records", 0))
    col2.metric("Ø Durchlaufzeit", f"{kpis.get('overall_avg_duration', 0.0)} Tage")
    col3.metric("Max. Durchlaufzeit", f"{kpis.get('max_duration', 0.0)} Tage")
    col4.metric("Min. Durchlaufzeit", f"{kpis.get('min_duration', 0.0)} Tage")

    st.divider()

    # --- Charts & Visualisierungen ---
    st.subheader("📈 Interaktive Analysen")
    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.markdown("#### Status-Verteilung")
        status_counts = kpis.get("status_counts", {})
        if status_counts:
            df_status = pd.DataFrame(list(status_counts.items()), columns=["Status", "Anzahl"])
            fig_status = px.pie(
                df_status, 
                names="Status", 
                values="Anzahl", 
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            st.plotly_chart(fig_status, use_container_width=True)
        else:
            st.info("Keine Status-Daten vorhanden.")

    with chart_col2:
        st.markdown("#### Verteilung der Durchlaufzeiten")
        duration_col = "durchlaufzeit_tage" if "durchlaufzeit_tage" in df.columns else ("duration" if "duration" in df.columns else None)
        if duration_col:
            fig_duration = px.histogram(
                df, 
                x=duration_col, 
                nbins=10, 
                title="Durchlaufzeiten in Tagen",
                labels={duration_col: "Tage"},
                color_discrete_sequence=["#2b5c8f"]
            )
            st.plotly_chart(fig_duration, use_container_width=True)
        else:
            st.info("Keine Spalte für Durchlaufzeiten im Datensatz.")

    st.divider()

    # --- Datentabelle & Excel-Download ---
    st.subheader("📋 Verarbeitete Daten")
    st.dataframe(df, use_container_width=True)

    # Excel Download Button
    excel_path = os.path.join(config["paths"]["processed_data_dir"], "pipeline_report.xlsx")
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📥 Excel-Bericht herunterladen",
                data=f,
                file_name="pipeline_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("👈 Klicke in der Seitenleiste auf **'Pipeline ausführen'**, um das Dashboard zu starten.")