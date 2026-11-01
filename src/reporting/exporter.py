import sys
from pathlib import Path

# Pfad absichern, damit imports aus src/ greifen
root_dir = Path(__file__).resolve().parent.parent.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

import os
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from src.etl.importer import load_config, import_csv_data, validate_schema
from src.etl.anonymizer import anonymize_dataframe
from src.analytics.kpi_calculator import calculate_kpis


def _get_output_path(output_filename: str) -> str:
    """Hilfsfunktion: Erstellt das Zielverzeichnis und liefert den vollständigen Pfad zurück."""
    config = load_config()
    processed_dir = config["paths"]["processed_data_dir"]
    os.makedirs(processed_dir, exist_ok=True)
    return os.path.join(processed_dir, output_filename)


# 1. EXCEL EXPORT
def export_to_excel(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.xlsx"):
    """Exportiert Daten und KPIs in eine Excel-Datei mit mehreren Tabellenblättern."""
    output_path = _get_output_path(output_filename)
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Verarbeitete_Daten", index=False)
        
        summary_data = {
            "Metrik": ["Gesamtdatensätze", "Ø Durchlaufzeit gesamt (Tage)"],
            "Wert": [kpis["total_records"], kpis["overall_avg_duration"]]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="KPI_Zusammenfassung", index=False)
        
    print(f"[INFO] Excel-Bericht gespeichert unter: '{output_path}'")


# 2. CSV EXPORT
def export_to_csv(df: pd.DataFrame, output_filename: str = "pipeline_report.csv"):
    """Exportiert die verarbeiteten Daten als CSV-Datei."""
    output_path = _get_output_path(output_filename)
    df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"[INFO] CSV-Datei gespeichert unter: '{output_path}'")


# 3. JSON EXPORT
def export_to_json(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.json"):
    """Exportiert Daten und KPIs gemeinsam in eine strukturierte JSON-Datei."""
    output_path = _get_output_path(output_filename)
    
    payload = {
        "kpis": kpis,
        "data": df.to_dict(orient="records")
    }
    
    import json
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=4, ensure_ascii=False)
        
    print(f"[INFO] JSON-Datei gespeichert unter: '{output_path}'")



# 4. PDF EXPORT (Visueller Report laut Flipchart-Skizze)
def export_to_pdf(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.pdf"):
    """
    Generiert einen visuellen PDF-Report exakt nach Flipchart-Layout:
    - Kreisdiagramm oben
    - Balkendiagramm darunter
    - Daten-Tabelle
    - Durchlaufzeit Max & Min als Text unten
    """
    output_path = _get_output_path(output_filename)
    temp_dir = os.path.join(os.path.dirname(output_path), "temp_assets")
    os.makedirs(temp_dir, exist_ok=True)

    # 1. Automatische Spaltenerkennung (Status/Kategorie & Durchlaufzeiten)
    cat_col = "status" if "status" in df.columns else ("Kategorie" if "Kategorie" in df.columns else df.columns[0])
    
    # Suche nach Durchlaufzeit-Spalten aus dem Pipeline-Datensatz
    time_col = None
    for col in ["durchlaufzeit_tage", "duration", "Durchlaufzeit"]:
        if col in df.columns:
            time_col = col
            break
    
    if not time_col:
        numeric_cols = df.select_dtypes(include=["number"]).columns
        time_col = numeric_cols[0] if len(numeric_cols) > 0 else df.columns[1]

    # Daten kopieren & Werte explizit als Numerik erzwingen (NaNs mit 0 auffüllen)
    df_chart = df.copy()
    df_chart[time_col] = pd.to_numeric(df_chart[time_col], errors="coerce").fillna(0)

    # --- 1. Kreisdiagramm (Pie Chart) ---
    pie_path = os.path.join(temp_dir, "pie_chart.png")
    plt.figure(figsize=(4, 4))
    pie_data = df_chart.groupby(cat_col)[time_col].sum()
    plt.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', startangle=140, 
            colors=['#3498db', '#2ecc71', '#e74c3c', '#f1c40f', '#9b59b6'])
    plt.title('Verteilung der Durchlaufzeiten', fontsize=12, fontweight='bold')
    plt.tight_layout()
    plt.savefig(pie_path, dpi=200)
    plt.close()

    # --- 2. Balkendiagramm (Bar Chart) ---
    bar_path = os.path.join(temp_dir, "bar_chart.png")
    plt.figure(figsize=(6, 3))
    plt.bar(df_chart[cat_col].astype(str), df_chart[time_col], color='#3498db')
    plt.title('Durchlaufzeit nach Status', fontsize=12, fontweight='bold')
    plt.xlabel(cat_col)
    plt.ylabel('Zeit (Tage)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(bar_path, dpi=200)
    plt.close()

    # --- 3. PDF Layout zusammenbauen ---
    doc = SimpleDocTemplate(output_path, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=15,
        textColor=colors.HexColor('#2c3e50')
    )

    story.append(Paragraph("AutoData-Pipeline: Analyse-Report", title_style))
    story.append(Spacer(1, 10))

    # Oben: Kreisdiagramm
    story.append(Image(pie_path, width=180, height=180))
    story.append(Spacer(1, 15))

    # Mitte: Balkendiagramm
    story.append(Image(bar_path, width=360, height=180))
    story.append(Spacer(1, 15))

    # Datentabelle (A, B, C, D)
    table_data = [list(df.columns)] + df.values.tolist()
    formatted_table_data = [[str(cell) for cell in row] for row in table_data]
    
    t = Table(formatted_table_data, hAlign='LEFT')
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#bdc3c7'))
    ]))
    story.append(t)
    story.append(Spacer(1, 15))

    # Unten: Max & Min Werte als Textform
    max_idx = df_chart[time_col].idxmax()
    min_idx = df_chart[time_col].idxmin()
    
    max_val = df_chart.loc[max_idx, time_col]
    max_cat = df_chart.loc[max_idx, cat_col]
    
    min_val = df_chart.loc[min_idx, time_col]
    min_cat = df_chart.loc[min_idx, cat_col]

    text_style = ParagraphStyle('SummaryText', parent=styles['Normal'], fontSize=11, leading=16)
    
    summary_html = f"""
    <b>Zahlen in Textform:</b><br/>
    • <b>Längste Durchlaufzeit (Max):</b> Status '{max_cat}' mit {max_val} Tagen<br/>
    • <b>Kürzeste Durchlaufzeit (Min):</b> Status '{min_cat}' mit {min_val} Tagen
    """
    story.append(Paragraph(summary_html, text_style))

    # Dokument erzeugen
    doc.build(story)
    
    # Aufräumen der Temp-Bilder
    if os.path.exists(pie_path): os.remove(pie_path)
    if os.path.exists(bar_path): os.remove(bar_path)
    
    print(f"[INFO] PDF-Bericht (Layout laut Skizze) gespeichert unter: '{output_path}'")


# ALL-IN-ONE EXPORTER ROUTER
def export_data(df: pd.DataFrame, kpis: dict, fmt: str = "excel", filename: str = None):
    """
    Zentraler Einstiegspunkt für den Datenexport.
    Unterstützt die Formate: 'excel', 'csv', 'json', 'pdf'.
    """
    fmt = fmt.lower()
    if fmt == "excel":
        export_to_excel(df, kpis, filename or "pipeline_report.xlsx")
    elif fmt == "csv":
        export_to_csv(df, filename or "pipeline_report.csv")
    elif fmt == "json":
        export_to_json(df, kpis, filename or "pipeline_report.json")
    elif fmt == "pdf":
        export_to_pdf(df, kpis, filename or "pipeline_report.pdf")
    else:
        raise ValueError(f"Nicht unterstütztes Export-Format: '{fmt}'. Erlaubt: excel, csv, json, pdf")


if __name__ == "__main__":
    # Testlauf für alle 4 Formate
    config = load_config()
    sample_file = os.path.join(config["paths"]["raw_data_dir"], "sample_operations.csv")
    
    df = import_csv_data(sample_file)
    df = validate_schema(df)
    df_anon = anonymize_dataframe(df)
    kpis = calculate_kpis(df_anon)
    
    print("\n--- Starte Multi-Format Export Test ---")
    export_data(df_anon, kpis, fmt="excel")
    export_data(df_anon, kpis, fmt="csv")
    export_data(df_anon, kpis, fmt="json")
    export_data(df_anon, kpis, fmt="pdf")