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
from src.reporting.exporter import export_data

# --- Page Konfigurationen ---
st.set_page_config(
    page_title="AutoData Pipeline Control Center",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AutoData Pipeline – Interaktives Kontrollzentrum")
st.markdown("Willkommen im Cockpit! Hier kannst du die Pipeline ausführen, Daten analysieren und Berichte in verschiedenen Formaten herunterladen.")

# --- Sidebar Konfigurationen ---
st.sidebar.header("Pipeline-Steuerung")

config = load_config()
raw_dir = config["paths"]["raw_data_dir"]
sample_file = os.path.join(raw_dir, "sample_operations.csv")

st.sidebar.info(f"**Datenquelle:** `{sample_file}`")

# Format-Auswahl für den Export direkt in der Sidebar
export_format = st.sidebar.selectbox(
    "📤 Gewünschtes Export-Format",
    ["Excel", "CSV", "JSON", "PDF (Visueller Report)"]
)

run_pipeline = st.sidebar.button("🚀 Pipeline ausführen", type="primary")

# Mapping für Dateinamen, Endungen und Mime-Types
format_config_map = {
    "Excel": ("pipeline_report.xlsx", "excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    "CSV": ("pipeline_report.csv", "csv", "text/csv"),
    "JSON": ("pipeline_report.json", "json", "application/json"),
    "PDF (Visueller Report)": ("pipeline_report.pdf", "pdf", "application/pdf")
}

file_name, fmt_key, mime_type = format_config_map[export_format]

# --- Pipeline Logik ---
if run_pipeline or "df_processed" in st.session_state:
    if run_pipeline:
        with st.spinner(f"Pipeline läuft... (Generiere {export_format}-Bericht)"):
            # ETL & Analytics Durchlauf
            df_raw = import_csv_data(sample_file)
            df_clean = validate_schema(df_raw)
            df_anon = anonymize_dataframe(df_clean)
            kpis = calculate_kpis(df_anon)
            
            # Report im gewählten Format generieren
            export_data(df_anon, kpis, fmt=fmt_key, filename=file_name)
            
            # In Session State speichern
            st.session_state["df_processed"] = df_anon
            st.session_state["kpis"] = kpis
            
        st.success(f"Pipeline erfolgreich ausgeführt & {export_format}-Bericht erstellt!")

    # Session-state Daten abrufen
    df = st.session_state["df_processed"]
    kpis = st.session_state["kpis"]

    # --- KPI Kacheln (Metrics) ---
    st.subheader("📌 Key Performance Indicators")
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
                df_status, names="Status", values="Anzahl", hole=0.4,
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
                df, x=duration_col, nbins=10, title="Durchlaufzeiten in Tagen",
                labels={duration_col: "Tage"}, color_discrete_sequence=["#2b5c8f"]
            )
            st.plotly_chart(fig_duration, use_container_width=True)
        else:
            st.info("Keine Spalte für Durchlaufzeiten im Datensatz.")

    st.divider()

    # --- Datentabelle & Dynamischer Download-Button ---
    st.subheader("📋 Verarbeitete Daten")
    st.dataframe(df, use_container_width=True)

    report_path = os.path.join(config["paths"]["processed_data_dir"], file_name)
    if os.path.exists(report_path):
        with open(report_path, "rb") as f:
            st.download_button(
                label=f"📥 {export_format} herunterladen",
                data=f,
                file_name=file_name,
                mime=mime_type
            )

else:
    st.info("👈 Wähle links in der Seitenleiste dein Format aus und klicke auf **'Pipeline ausführen'**.")