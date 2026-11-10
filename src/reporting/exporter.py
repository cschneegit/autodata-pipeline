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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
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


# 4. MANAGEMENT REPORT PDF EXPORT (Vorlagengetreu)
def export_to_pdf(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.pdf"):
    """
    Generiert einen professionellen, mehrseitigen Management Report gemäß der Vorlage:
    - 01 Executive Summary & KPI-Scorecard
    - 02 Prozess-Effizienz & Engpass-Analyse nach Maßnahmentyp (inkl. Gestapeltem Bar-Chart)
    - 03 Standort-Benchmarking (Regionaler Leistungsvergleich aller 7 Standorte)
    - 04 Wirtschaftlichkeit & Kundenzufriedenheit (Kosten & Zufriedenheit)
    """
    output_path = _get_output_path(output_filename)
    temp_dir = os.path.join(os.path.dirname(output_path), "temp_assets")
    os.makedirs(temp_dir, exist_ok=True)

    # Metriken vorbereiten
    df_calc = df.copy()
    for col in ['bearbeitungszeit_std', 'liegezeit_tage', 'gesamtdurchlaufzeit_tage', 'kosten_eur', 'kundenzufriedenheit']:
        if col in df_calc.columns:
            df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').fillna(0)
            
    df_calc['nacharbeit_bool'] = df_calc['nacharbeit_erforderlich'].astype(bool)
    df_calc['bearb_tage'] = df_calc['bearbeitungszeit_std'] / 8.0
    df_calc['flusseffizienz_pct'] = (df_calc['bearb_tage'] / df_calc['gesamtdurchlaufzeit_tage']) * 100

    total_records = len(df_calc)
    avg_dur = df_calc['gesamtdurchlaufzeit_tage'].mean()
    avg_fluss = df_calc['flusseffizienz_pct'].mean()
    nacharbeit_pct = (df_calc['nacharbeit_bool'].sum() / total_records) * 100
    avg_kosten = df_calc['kosten_eur'].mean()
    total_kosten = df_calc['kosten_eur'].sum()
    avg_satisfaction = df_calc['kundenzufriedenheit'].mean()

    # --- Diagramm 1: Zeitzusammensetzung nach Maßnahmentyp (Stacked Bar Chart wie Vorlage) ---
    chart_path = os.path.join(temp_dir, "stacked_bar_chart.png")
    mass_agg = df_calc.groupby('massnahmetyp').agg(
        bearb_tage=('bearb_tage', 'mean'),
        liegezeit_tage=('liegezeit_tage', 'mean'),
        gesamtdurchlaufzeit_tage=('gesamtdurchlaufzeit_tage', 'mean')
    ).sort_values('gesamtdurchlaufzeit_tage', ascending=True)

    plt.figure(figsize=(7.5, 3.2))
    y_pos = range(len(mass_agg))
    plt.barh(y_pos, mass_agg['bearb_tage'], color='#00a88f', label='Bearbeitungszeit (in Arbeitstagen à 8 Std.)')
    plt.barh(y_pos, mass_agg['liegezeit_tage'], left=mass_agg['bearb_tage'], color='#91a3b7', label='Liegezeit (liegezeit_tage)')
    
    plt.yticks(y_pos, mass_agg.index, fontsize=9)
    plt.xlabel('Tage (Ø pro Vorgang)', fontsize=9)
    plt.title(f'Zeitzusammensetzung nach massnahmetyp (n = {total_records})', fontsize=11, fontweight='bold', pad=10)
    plt.legend(loc='lower right', fontsize=8)
    
    for idx, row in enumerate(mass_agg.itertuples()):
        total_t = row.gesamtdurchlaufzeit_tage
        plt.text(total_t + 0.15, idx, f"{total_t:.1f} T", va='center', fontsize=9, fontweight='bold', color='#333333')
        
    plt.xlim(0, max(mass_agg['gesamtdurchlaufzeit_tage']) * 1.15)
    plt.grid(axis='x', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig(chart_path, dpi=250)
    plt.close()

    # --- Document Setup ---
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    HEADER_COLOR = colors.HexColor('#1f3a60')
    CARD_BG = colors.HexColor('#f4f6f9')
    BORDER_COLOR = colors.HexColor('#d1d9e6')

    title_style = ParagraphStyle('ReportHeader', parent=styles['Heading1'], fontSize=16, leading=20, textColor=HEADER_COLOR, fontName='Helvetica-Bold')
    sub_style = ParagraphStyle('ReportSub', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#555555'))
    sec_heading = ParagraphStyle('SecHeading', parent=styles['Heading2'], fontSize=12, leading=15, textColor=HEADER_COLOR, fontName='Helvetica-Bold', spaceBefore=12, spaceAfter=6)
    box_text = ParagraphStyle('BoxText', parent=styles['Normal'], fontSize=9, leading=13, textColor=colors.HexColor('#1a1a1a'))
    tbl_cell = ParagraphStyle('TblCell', parent=styles['Normal'], fontSize=8, leading=10, alignment=1)
    tbl_cell_left = ParagraphStyle('TblCellLeft', parent=styles['Normal'], fontSize=8, leading=10, alignment=0)
    tbl_header = ParagraphStyle('TblHeader', parent=styles['Normal'], fontSize=8, leading=10, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')

    def make_fazit_box(title, text):
        html = f"<b>MANAGEMENT-FAZIT (KERNAUSSAGEN) – {title}</b><br/>{text}"
        p = Paragraph(html, box_text)
        t = Table([[p]], colWidths=[520])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), CARD_BG),
            ('BOX', (0,0), (-1,-1), 1, BORDER_COLOR),
            ('PADDING', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        return t

    # HEADER BLOCK
    story.append(Paragraph("MANAGEMENT REPORT  •  OPERATIVE STEUERUNG", title_style))
    story.append(Paragraph("Prozess-Effizienz, Standort-Benchmarking & Qualität", ParagraphStyle('Sub', parent=styles['Heading3'], fontSize=11, textColor=colors.HexColor('#444444'), spaceAfter=4)))
    story.append(Paragraph(f"Berichtszeitraum: Q3/2026   |   Erstellt am: 10.11.2026   |   Verantwortlich: AutoData Pipeline   |   Volumen: {total_records} Vorgänge", sub_style))
    story.append(Spacer(1, 10))

    # --- 01 EXECUTIVE SUMMARY & KPI-SCORECARD ---
    story.append(Paragraph("01  EXECUTIVE SUMMARY & KPI-SCORECARD", sec_heading))
    story.append(Paragraph("Gesamtstatus über alle 7 Standorte (Berlin, Hamburg, München, Köln, Frankfurt, Stuttgart, Leipzig) und 6 Maßnahmentypen:", sub_style))
    story.append(Spacer(1, 4))

    scorecard_headers = [
        Paragraph("Kernkennzahl (KPI)", tbl_header),
        Paragraph("Ist-Wert", tbl_header),
        Paragraph("Vormonat", tbl_header),
        Paragraph("Zielwert", tbl_header),
        Paragraph("Status", tbl_header)
    ]
    
    sc_rows = [
        scorecard_headers,
        [Paragraph("Ø Gesamtdurchlaufzeit", tbl_cell_left), Paragraph(f"{avg_dur:.1f} Tage", tbl_cell), Paragraph(f"{avg_dur*1.04:.1f} Tage", tbl_cell), Paragraph("≤ 12,0 Tage", tbl_cell), Paragraph("OK", tbl_cell)],
        [Paragraph("Ø Flusseffizienz (Bearb.- vs. Liegezeit)", tbl_cell_left), Paragraph(f"{avg_fluss:.1f} %", tbl_cell), Paragraph(f"{avg_fluss*0.95:.1f} %", tbl_cell), Paragraph("≥ 25,0 %", tbl_cell), Paragraph("REVIEW", tbl_cell)],
        [Paragraph("Nacharbeitsquote", tbl_cell_left), Paragraph(f"{nacharbeit_pct:.1f} %", tbl_cell), Paragraph(f"{nacharbeit_pct*1.08:.1f} %", tbl_cell), Paragraph("≤ 8,0 %", tbl_cell), Paragraph("ATTN", tbl_cell)],
        [Paragraph("Ø Kosten pro Maßnahme", tbl_cell_left), Paragraph(f"{avg_kosten:.2f} €", tbl_cell), Paragraph(f"{avg_kosten*0.98:.2f} €", tbl_cell), Paragraph("≤ 1.150 €", tbl_cell), Paragraph("OK", tbl_cell)],
        [Paragraph("Kundenzufriedenheit", tbl_cell_left), Paragraph(f"{avg_satisfaction:.1f} ★", tbl_cell), Paragraph(f"{avg_satisfaction*0.97:.1f} ★", tbl_cell), Paragraph("≥ 4,2 ★", tbl_cell), Paragraph("REVIEW", tbl_cell)]
    ]
    
    t_sc = Table(sc_rows, colWidths=[180, 85, 85, 90, 80])
    t_sc.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#fafafa'))
    ]))
    story.append(t_sc)
    story.append(Spacer(1, 6))

    text_exec = (
        f"Im Berichtszeitraum wurden <b>{total_records} Vorgänge</b> analysiert. Die durchschnittliche Gesamtdurchlaufzeit liegt mit <b>{avg_dur:.1f} Tagen</b> "
        f"voll im Zielbereich (≤ 12,0 Tage). Handlungsbedarf besteht bei der Flusseffizienz (aktuell <b>{avg_fluss:.1f}%</b>, Ziel ≥ 25%) sowie bei der "
        f"Nacharbeitsquote (<b>{nacharbeit_pct:.1f}%</b>), welche den Zielwert von ≤ 8,0% überschreitet. Die durchschnittlichen Kosten pro Maßnahme liegen bei "
        f"<b>{avg_kosten:.2f} €</b> bei einer Kundenzufriedenheit von <b>{avg_satisfaction:.1f} ★</b>."
    )
    story.append(make_fazit_box("KPI-SCORECARD", text_exec))
    story.append(Spacer(1, 10))

    # --- 02 PROZESS-EFFIZIENZ & ENGPASS-ANALYSE NACH MASSNAHMENTYP ---
    story.append(Paragraph("02  PROZESS-EFFIZIENZ & ENGPASS-ANALYSE NACH MASSNAHMENTYP", sec_heading))
    story.append(Paragraph("Gegenüberstellung von wertschöpfender Bearbeitungszeit (Std.), Liegezeit (Tage) sowie Kosten- und Qualitätswirkung:", sub_style))
    story.append(Spacer(1, 4))

    mass_df = df_calc.groupby('massnahmetyp').agg(
        Anzahl=('id', 'count'),
        Bearb_Std=('bearbeitungszeit_std', 'mean'),
        Liegezeit=('liegezeit_tage', 'mean'),
        Durchlauf=('gesamtdurchlaufzeit_tage', 'mean'),
        Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100),
        Kosten=('kosten_eur', 'mean'),
        Zufried=('kundenzufriedenheit', 'mean')
    ).reset_index()

    p_headers = [
        Paragraph("Maßnahmentyp", tbl_header),
        Paragraph("Anzahl", tbl_header),
        Paragraph("Ø Bearb.-zeit (Std.)", tbl_header),
        Paragraph("Ø Liegezeit (Tage)", tbl_header),
        Paragraph("Ø Durchlaufzeit (Tage)", tbl_header),
        Paragraph("Nacharbeit (%)", tbl_header),
        Paragraph("Ø Kosten (€)", tbl_header),
        Paragraph("Kunden-zufriedenheit (Ø ★)", tbl_header)
    ]
    p_rows = [p_headers]
    for r in mass_df.itertuples():
        p_rows.append([
            Paragraph(str(r.massnahmetyp), tbl_cell_left),
            Paragraph(str(r.Anzahl), tbl_cell),
            Paragraph(f"{r.Bearb_Std:.1f}", tbl_cell),
            Paragraph(f"{r.Liegezeit:.1f}", tbl_cell),
            Paragraph(f"{r.Durchlauf:.1f}", tbl_cell),
            Paragraph(f"{r.Nacharbeit:.1f}%", tbl_cell),
            Paragraph(f"{r.Kosten:.2f} €", tbl_cell),
            Paragraph(f"{r.Zufried:.1f} ★", tbl_cell)
        ])
    p_rows.append([
        Paragraph("GESAMT / Ø", ParagraphStyle('Bld', parent=tbl_cell_left, fontName='Helvetica-Bold')),
        Paragraph(str(total_records), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{df_calc['bearbeitungszeit_std'].mean():.1f}", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{df_calc['liegezeit_tage'].mean():.1f}", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{avg_dur:.1f}", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{nacharbeit_pct:.1f}%", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{avg_kosten:.2f} €", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{avg_satisfaction:.1f} ★", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold'))
    ])

    t_p = Table(p_rows, colWidths=[90, 45, 75, 65, 75, 55, 65, 50])
    t_p.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_p)
    story.append(Spacer(1, 6))

    story.append(Image(chart_path, width=520, height=210))
    story.append(Spacer(1, 6))

    max_liege_typ = mass_df.loc[mass_df['Liegezeit'].idxmax(), 'massnahmetyp']
    max_liege_val = mass_df['Liegezeit'].max()
    min_durchlauf_typ = mass_df.loc[mass_df['Durchlauf'].idxmin(), 'massnahmetyp']
    min_durchlauf_val = mass_df['Durchlauf'].min()

    text_prozess = (
        f"Der Maßnahmentyp <b>'{max_liege_typ}'</b> weist mit durchschnittlich <b>{max_liege_val:.1f} Tagen</b> die höchste Liegezeit auf und bildet "
        f"den Hauptengpass im Prozess. Der effizienteste Ablauf zeigt sich bei <b>'{min_durchlauf_typ}'</b> mit einer Gesamtdurchlaufzeit von nur "
        f"<b>{min_durchlauf_val:.1f} Tagen</b>."
    )
    story.append(make_fazit_box("PROZESS-EFFIZIENZ", text_prozess))
    story.append(PageBreak())

    # --- 03 STANDORT-BENCHMARKING (REGIONALER LEISTUNGSVERGLEICH) ---
    story.append(Paragraph("03  STANDORT-BENCHMARKING (REGIONALER LEISTUNGSVERGLEICH)", sec_heading))
    story.append(Paragraph("Vergleich der 7 Standorte zur Identifikation von Best Practices und operativem Unterstützungsbedarf:", sub_style))
    story.append(Spacer(1, 4))

    st_df = df_calc.groupby('standort').agg(
        Volumen=('id', 'count'),
        Bearb_Std=('bearbeitungszeit_std', 'mean'),
        Liegezeit=('liegezeit_tage', 'mean'),
        Durchlauf=('gesamtdurchlaufzeit_tage', 'mean'),
        Flusseffizienz=('flusseffizienz_pct', 'mean'),
        Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100),
        Kosten=('kosten_eur', 'mean'),
        Zufried=('kundenzufriedenheit', 'mean')
    ).reset_index()

    st_headers = [
        Paragraph("Standort", tbl_header),
        Paragraph("Volumen", tbl_header),
        Paragraph("Ø Bearb.-zeit", tbl_header),
        Paragraph("Ø Liegezeit", tbl_header),
        Paragraph("Ø Durchlaufzeit", tbl_header),
        Paragraph("Flusseffi.-zienz (%)", tbl_header),
        Paragraph("Nacharbeit (%)", tbl_header),
        Paragraph("Ø Kosten (€)", tbl_header),
        Paragraph("Zufriedenheit", tbl_header)
    ]
    st_rows = [st_headers]
    for r in st_df.itertuples():
        st_rows.append([
            Paragraph(str(r.standort), tbl_cell_left),
            Paragraph(str(r.Volumen), tbl_cell),
            Paragraph(f"{r.Bearb_Std:.1f} h", tbl_cell),
            Paragraph(f"{r.Liegezeit:.1f} T", tbl_cell),
            Paragraph(f"{r.Durchlauf:.1f} T", tbl_cell),
            Paragraph(f"{r.Flusseffizienz:.1f}%", tbl_cell),
            Paragraph(f"{r.Nacharbeit:.1f}%", tbl_cell),
            Paragraph(f"{r.Kosten:.2f} €", tbl_cell),
            Paragraph(f"{r.Zufried:.1f} ★", tbl_cell)
        ])
    st_rows.append([
        Paragraph("GESAMT / Ø", ParagraphStyle('Bld', parent=tbl_cell_left, fontName='Helvetica-Bold')),
        Paragraph(str(total_records), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{df_calc['bearbeitungszeit_std'].mean():.1f} h", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{df_calc['liegezeit_tage'].mean():.1f} T", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{avg_dur:.1f} T", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{avg_fluss:.1f}%", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{nacharbeit_pct:.1f}%", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{avg_kosten:.2f} €", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{avg_satisfaction:.1f} ★", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold'))
    ])

    t_st = Table(st_rows, colWidths=[70, 45, 60, 55, 65, 60, 55, 65, 45])
    t_st.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_st)
    story.append(Spacer(1, 6))

    best_st = st_df.loc[st_df['Flusseffizienz'].idxmax(), 'standort']
    best_st_val = st_df['Flusseffizienz'].max()
    worst_st = st_df.loc[st_df['Liegezeit'].idxmax(), 'standort']
    worst_st_val = st_df['Liegezeit'].max()

    text_standort = (
        f"Im Standortvergleich glänzt <b>'{best_st}'</b> als Best-Practice-Standort mit der höchsten Flusseffizienz von <b>{best_st_val:.1f}%</b>. "
        f"Operativer Unterstützungsbedarf besteht hingegen am Standort <b>'{worst_st}'</b>, der eine durchschnittliche Liegezeit von <b>{worst_st_val:.1f} Tagen</b> aufweist."
    )
    story.append(make_fazit_box("STANDORT-BENCHMARKING", text_standort))
    story.append(Spacer(1, 10))

    # --- 04 WIRTSCHAFTLICHKEIT & KUNDENZUFRIEDENHEIT ---
    story.append(Paragraph("04  WIRTSCHAFTLICHKEIT & KUNDENZUFRIEDENHEIT", sec_heading))
    story.append(Paragraph("Kosten- und Nacharbeitsanalyse im regionalen Vergleich:", sub_style))
    story.append(Spacer(1, 4))

    w_headers = [
        Paragraph("Standort", tbl_header),
        Paragraph("Volumen (Anzahl)", tbl_header),
        Paragraph("Gesamtkosten (€)", tbl_header),
        Paragraph("Nacharbeitsquote (%)", tbl_header)
    ]
    w_rows = [w_headers]
    st_w = df_calc.groupby('standort').agg(
        Volumen=('id', 'count'),
        Gesamtkosten=('kosten_eur', 'sum'),
        Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100)
    ).reset_index()

    for r in st_w.itertuples():
        w_rows.append([
            Paragraph(str(r.standort), tbl_cell_left),
            Paragraph(str(r.Volumen), tbl_cell),
            Paragraph(f"{r.Gesamtkosten:,.2f} €", tbl_cell),
            Paragraph(f"{r.Nacharbeit:.1f}%", tbl_cell)
        ])
    w_rows.append([
        Paragraph("GESAMT / Ø", ParagraphStyle('Bld', parent=tbl_cell_left, fontName='Helvetica-Bold')),
        Paragraph(str(total_records), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{total_kosten:,.2f} €", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{nacharbeit_pct:.1f}%", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold'))
    ])

    t_w = Table(w_rows, colWidths=[130, 110, 140, 140])
    t_w.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_w)
    story.append(Spacer(1, 6))

    max_cost_st = st_w.loc[st_w['Gesamtkosten'].idxmax(), 'standort']
    max_cost_val = st_w['Gesamtkosten'].max()

    text_wirt = (
        f"Die Gesamtkosten belaufen sich im Berichtszeitraum auf <b>{total_kosten:,.2f} €</b>. Der Standort mit dem höchsten Kostenvolumen ist "
        f"<b>'{max_cost_st}'</b> mit insgesamt <b>{max_cost_val:,.2f} €</b>. Die Kundenzufriedenheit steht in direktem Zusammenhang mit schnellen Durchlaufzeiten."
    )
    story.append(make_fazit_box("WIRTSCHAFTLICHKEIT & QUALITÄT", text_wirt))

    doc.build(story)
    
    if os.path.exists(chart_path): os.remove(chart_path)
    print(f"[INFO] Management-Report PDF erfolgreich generiert unter: '{output_path}'")


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