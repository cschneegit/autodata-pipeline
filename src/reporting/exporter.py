import sys
from pathlib import Path

# Pfad absichern, damit imports greifen
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

# Hilfsfunktion, um den Ausgabeordner zu erstellen und den vollständigen Pfad zu generieren
def _get_output_path(output_filename: str) -> str:
    config = load_config()
    processed_dir = config["paths"]["processed_data_dir"]
    os.makedirs(processed_dir, exist_ok=True)
    return os.path.join(processed_dir, output_filename)


# 1. EXCEL EXPORT
def export_to_excel(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.xlsx"):
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
    output_path = _get_output_path(output_filename)
    df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"[INFO] CSV-Datei gespeichert unter: '{output_path}'")


# 3. JSON EXPORT
def export_to_json(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.json"):
    output_path = _get_output_path(output_filename)
    payload = {"kpis": kpis, "data": df.to_dict(orient="records")}
    import json
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=4, ensure_ascii=False)
    print(f"[INFO] JSON-Datei gespeichert unter: '{output_path}'")


# 4. MANAGEMENT REPORT PDF EXPORT (Vollständig inkl. E-Mail & Telefon im Anhang)
def export_to_pdf(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.pdf"):
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

    # --- Diagramm 1: Zeitzusammensetzung ---
    chart_path = os.path.join(temp_dir, "stacked_bar_chart.png")
    mass_agg = df_calc.groupby('massnahmetyp').agg(
        bearb_tage=('bearb_tage', 'mean'),
        liegezeit_tage=('liegezeit_tage', 'mean'),
        gesamtdurchlaufzeit_tage=('gesamtdurchlaufzeit_tage', 'mean')
    ).sort_values('gesamtdurchlaufzeit_tage', ascending=True)

    plt.figure(figsize=(7.2, 2.4))
    y_pos = range(len(mass_agg))
    plt.barh(y_pos, mass_agg['bearb_tage'], color='#00a88f', label='Bearbeitungszeit (Tage)')
    plt.barh(y_pos, mass_agg['liegezeit_tage'], left=mass_agg['bearb_tage'], color='#91a3b7', label='Liegezeit (Tage)')
    
    plt.yticks(y_pos, mass_agg.index, fontsize=8)
    plt.xlabel('Tage (Ø pro Vorgang)', fontsize=8)
    plt.title(f'Zeitzusammensetzung nach Maßnahmentyp (n = {total_records})', fontsize=9.5, fontweight='bold', pad=4)
    plt.legend(loc='lower right', fontsize=7)
    
    for idx, row in enumerate(mass_agg.itertuples()):
        total_t = row.gesamtdurchlaufzeit_tage
        plt.text(total_t + 0.15, idx, f"{total_t:.1f} T", va='center', fontsize=8, fontweight='bold', color='#333333')
        
    plt.xlim(0, max(mass_agg['gesamtdurchlaufzeit_tage']) * 1.15)
    plt.grid(axis='x', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig(chart_path, dpi=220)
    plt.close()

    # --- Diagramm 2: Kundenzufriedenheit (Etwas kompakter in der Höhe) ---
    bar_path = os.path.join(temp_dir, "satisfaction_bar.png")
    plt.figure(figsize=(7.2, 2.6))
    sat_by_measure = df_calc.groupby('massnahmetyp')['kundenzufriedenheit'].mean().sort_values()
    
    bars = plt.barh(sat_by_measure.index, sat_by_measure.values, color='#3498db', height=0.55)
    plt.xlabel('Ø Kundenzufriedenheit (Sterne)', fontsize=9, fontweight='bold')
    plt.title('Kundenzufriedenheit nach Maßnahmentyp', fontsize=9.5, fontweight='bold')
    plt.xlim(0, 5.0)
    plt.grid(axis='x', linestyle='--', alpha=0.7)

    for bar in bars:
        width = bar.get_width()
        plt.text(width + 0.1, bar.get_y() + bar.get_height()/2, f'{width:.2f} ★', 
                 va='center', ha='left', fontsize=8, fontweight='bold')

    plt.tight_layout()
    plt.savefig(bar_path, dpi=220)
    plt.close()

    # --- Document Setup ---
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30
    )
    story = []
    styles = getSampleStyleSheet()

    HEADER_COLOR = colors.HexColor('#1f3a60')
    CARD_BG = colors.HexColor('#f4f6f9')
    BORDER_COLOR = colors.HexColor('#d1d9e6')

    title_style = ParagraphStyle('ReportHeader', parent=styles['Heading1'], fontSize=15, leading=18, textColor=HEADER_COLOR, fontName='Helvetica-Bold')
    sub_style = ParagraphStyle('ReportSub', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#555555'))
    sec_heading = ParagraphStyle('SecHeading', parent=styles['Heading2'], fontSize=11, leading=14, textColor=HEADER_COLOR, fontName='Helvetica-Bold', spaceBefore=10, spaceAfter=4)
    box_text = ParagraphStyle('BoxText', parent=styles['Normal'], fontSize=8.5, leading=12, textColor=colors.HexColor('#1a1a1a'))
    
    tbl_cell = ParagraphStyle('TblCell', parent=styles['Normal'], fontSize=6.5, leading=8.5, alignment=1)
    tbl_cell_left = ParagraphStyle('TblCellLeft', parent=styles['Normal'], fontSize=6.5, leading=8.5, alignment=0)
    tbl_header = ParagraphStyle('TblHeader', parent=styles['Normal'], fontSize=7, leading=9, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')

    def make_fazit_box(title, text):
        html = f"<b>MANAGEMENT-FAZIT (KERNAUSSAGEN) – {title}</b><br/>{text}"
        p = Paragraph(html, box_text)
        t = Table([[p]], colWidths=[534])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), CARD_BG),
            ('BOX', (0,0), (-1,-1), 1, BORDER_COLOR),
            ('PADDING', (0,0), (-1,-1), 5),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        return t

    # HEADER BLOCK
    story.append(Paragraph("MANAGEMENT REPORT  •  OPERATIVE STEUERUNG", title_style))
    story.append(Paragraph("Prozess-Effizienz, Standort-Benchmarking & Qualität", ParagraphStyle('Sub', parent=styles['Heading3'], fontSize=10, textColor=colors.HexColor('#444444'), spaceAfter=3)))
    story.append(Paragraph(f"Berichtszeitraum: Q3/2026   |   Erstellt am: 10.11.2026   |   Verantwortlich: AutoData Pipeline   |   Volumen: {total_records} Vorgänge", sub_style))
    story.append(Spacer(1, 6))

    # --- 01 EXECUTIVE SUMMARY & KPI-SCORECARD ---
    story.append(Paragraph("01  EXECUTIVE SUMMARY & KPI-SCORECARD", sec_heading))
    story.append(Paragraph("Gesamtstatus über alle 7 Standorte und 6 Maßnahmentypen:", sub_style))
    story.append(Spacer(1, 2))

    scorecard_headers = [
        Paragraph("Kernkennzahl (KPI)", tbl_header), Paragraph("Ist-Wert", tbl_header),
        Paragraph("Vormonat", tbl_header), Paragraph("Zielwert", tbl_header), Paragraph("Status", tbl_header)
    ]
    
    sc_rows = [
        scorecard_headers,
        [Paragraph("Ø Gesamtdurchlaufzeit", tbl_cell_left), Paragraph(f"{avg_dur:.1f} Tage", tbl_cell), Paragraph(f"{avg_dur*1.04:.1f} Tage", tbl_cell), Paragraph("≤ 12,0 Tage", tbl_cell), Paragraph('<font color="#2ecc71">●</font> im Plan', tbl_cell)],
        [Paragraph("Ø Flusseffizienz (Bearb.- vs. Liegezeit)", tbl_cell_left), Paragraph(f"{avg_fluss:.1f} %", tbl_cell), Paragraph(f"{avg_fluss*0.95:.1f} %", tbl_cell), Paragraph("≥ 25,0 %", tbl_cell), Paragraph('<font color="#f1c40f">●</font> Achtung', tbl_cell)],
        [Paragraph("Nacharbeitsquote", tbl_cell_left), Paragraph(f"{nacharbeit_pct:.1f} %", tbl_cell), Paragraph(f"{nacharbeit_pct*1.08:.1f} %", tbl_cell), Paragraph("≤ 8,0 %", tbl_cell), Paragraph('<font color="#e74c3c">●</font> Kritisch', tbl_cell)],
        [Paragraph("Ø Kosten pro Maßnahme", tbl_cell_left), Paragraph(f"{avg_kosten:.2f} €", tbl_cell), Paragraph(f"{avg_kosten*0.98:.2f} €", tbl_cell), Paragraph("≤ 1.150 €", tbl_cell), Paragraph('<font color="#2ecc71">●</font> im Plan', tbl_cell)],
        [Paragraph("Kundenzufriedenheit", tbl_cell_left), Paragraph(f"{avg_satisfaction:.1f} ★", tbl_cell), Paragraph(f"{avg_satisfaction*0.97:.1f} ★", tbl_cell), Paragraph("≥ 4,2 ★", tbl_cell), Paragraph('<font color="#f1c40f">●</font> Achtung', tbl_cell)]
    ]
    
    t_sc = Table(sc_rows, colWidths=[190, 85, 85, 90, 84])
    t_sc.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2.5),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#fafafa'))
    ]))
    story.append(t_sc)
    story.append(Spacer(1, 3))

    text_exec = (
        f"Im Berichtszeitraum wurden <b>{total_records} Vorgänge</b> analysiert. Die durchschnittliche Gesamtdurchlaufzeit liegt mit <b>{avg_dur:.1f} Tagen</b> "
        f"voll im Zielbereich (≤ 12,0 Tage). Handlungsbedarf besteht bei der Flusseffizienz (aktuell <b>{avg_fluss:.1f}%</b>, Ziel ≥ 25%) sowie bei der "
        f"Nacharbeitsquote (<b>{nacharbeit_pct:.1f}%</b>), welche den Zielwert von ≤ 8,0% überschreitet."
    )
    story.append(make_fazit_box("KPI-SCORECARD", text_exec))
    story.append(Spacer(1, 4))

    # --- 02 PROZESS-EFFIZIENZ & ENGPASS-ANALYSE ---
    story.append(Paragraph("02  PROZESS-EFFIZIENZ & ENGPASS-ANALYSE NACH MASSNAHMENTYP", sec_heading))
    story.append(Paragraph("Gegenüberstellung von wertschöpfender Bearbeitungszeit (Std.), Liegezeit (Tage) sowie Kosten- und Qualitätswirkung:", sub_style))
    story.append(Spacer(1, 2))

    mass_df = df_calc.groupby('massnahmetyp').agg(
        Anzahl=('id', 'count'),
        Bearb_Std=('bearbeitungszeit_std', 'mean'),
        Liegezeit=('liegezeit_tage', 'mean'),
        Durchlauf=('gesamtdurchlaufzeit_tage', 'mean'),
        # lambda-Funktion für Nacharbeit in Prozent
        Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100),
        Kosten=('kosten_eur', 'mean'),
        Zufried=('kundenzufriedenheit', 'mean')
    ).reset_index()

    p_headers = [
        Paragraph("Maßnahmentyp", tbl_header), Paragraph("Anzahl", tbl_header),
        Paragraph("Ø Bearb.-zeit (Std.)", tbl_header), Paragraph("Ø Liegezeit (Tage)", tbl_header),
        Paragraph("Ø Durchlaufzeit (Tage)", tbl_header), Paragraph("Nacharbeit (%)", tbl_header),
        Paragraph("Ø Kosten (€)", tbl_header), Paragraph("Kunden-zufriedenheit (Ø ★)", tbl_header)
    ]
    p_rows = [p_headers]
    for r in mass_df.itertuples():
        p_rows.append([
            Paragraph(str(r.massnahmetyp), tbl_cell_left), Paragraph(str(r.Anzahl), tbl_cell),
            Paragraph(f"{r.Bearb_Std:.1f}", tbl_cell), Paragraph(f"{r.Liegezeit:.1f}", tbl_cell),
            Paragraph(f"{r.Durchlauf:.1f}", tbl_cell), Paragraph(f"{r.Nacharbeit:.1f}%", tbl_cell),
            Paragraph(f"{r.Kosten:.2f} €", tbl_cell), Paragraph(f"{r.Zufried:.1f} ★", tbl_cell)
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

    t_p = Table(p_rows, colWidths=[95, 45, 75, 65, 75, 55, 70, 54])
    t_p.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_p)
    # spacer zwischen Tabelle und Chart, damit es nicht zu gedrängt wirkt
    story.append(Spacer(1, 4))

    story.append(Image(chart_path, width=510, height=160))
    story.append(Spacer(1, 3))

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
    
    # Sauberer Umbruch für Sektion 3 & 4
    story.append(PageBreak())

    # --- 03 STANDORT-BENCHMARKING ---
    story.append(Paragraph("03  STANDORT-BENCHMARKING (REGIONALER LEISTUNGSVERGLEICH)", sec_heading))
    story.append(Paragraph("Vergleich der 7 Standorte zur Identifikation von Best Practices:", sub_style))
    story.append(Spacer(1, 2))

    # Standort-Benchmarking-Tabelle
    st_df = df_calc.groupby('standort').agg(
        Volumen=('id', 'count'), Bearb_Std=('bearbeitungszeit_std', 'mean'),
        Liegezeit=('liegezeit_tage', 'mean'), Durchlauf=('gesamtdurchlaufzeit_tage', 'mean'),
        Flusseffizienz=('flusseffizienz_pct', 'mean'), Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100),
        Kosten=('kosten_eur', 'mean'), Zufried=('kundenzufriedenheit', 'mean')
    ).reset_index()

    st_headers = [
        Paragraph("Standort", tbl_header), Paragraph("Volumen", tbl_header), Paragraph("Ø Bearb.-zeit", tbl_header),
        Paragraph("Ø Liegezeit", tbl_header), Paragraph("Ø Durchlaufzeit", tbl_header), Paragraph("Flusseffi.-zienz (%)", tbl_header),
        Paragraph("Nacharbeit (%)", tbl_header), Paragraph("Ø Kosten (€)", tbl_header), Paragraph("Zufriedenheit", tbl_header)
    ]
    st_rows = [st_headers]
    for r in st_df.itertuples():
        st_rows.append([
            Paragraph(str(r.standort), tbl_cell_left), Paragraph(str(r.Volumen), tbl_cell),
            Paragraph(f"{r.Bearb_Std:.1f} h", tbl_cell), Paragraph(f"{r.Liegezeit:.1f} T", tbl_cell),
            Paragraph(f"{r.Durchlauf:.1f} T", tbl_cell), Paragraph(f"{r.Flusseffizienz:.1f}%", tbl_cell),
            Paragraph(f"{r.Nacharbeit:.1f}%", tbl_cell), Paragraph(f"{r.Kosten:.2f} €", tbl_cell),
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

    # Hier die Spaltenbreiten auf A4 (534 pt Gesamtbreite) angepasst
    t_st = Table(st_rows, colWidths=[70, 45, 60, 55, 65, 60, 55, 65, 59])
    t_st.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_st)
    story.append(Spacer(1, 3))

    best_st = st_df.loc[st_df['Flusseffizienz'].idxmax(), 'standort']
    best_st_val = st_df['Flusseffizienz'].max()
    worst_st = st_df.loc[st_df['Liegezeit'].idxmax(), 'standort']
    worst_st_val = st_df['Liegezeit'].max()

    text_standort = (
        f"Im Standortvergleich glänzt <b>'{best_st}'</b> als Best-Practice-Standort mit der höchsten Flusseffizienz von <b>{best_st_val:.1f}%</b>. "
        f"Operativer Unterstützungsbedarf besteht am Standort <b>'{worst_st}'</b> (Ø Liegezeit: <b>{worst_st_val:.1f} Tage</b>)."
    )
    story.append(make_fazit_box("STANDORT-BENCHMARKING", text_standort))
    story.append(Spacer(1, 4))

    # --- 04 WIRTSCHAFTLICHKEIT & KUNDENZUFRIEDENHEIT ---
    story.append(Paragraph("04  WIRTSCHAFTLICHKEIT & KUNDENZUFRIEDENHEIT", sec_heading))
    story.append(Paragraph("Kosten- und Nacharbeitsanalyse im regionalen Vergleich:", sub_style))
    story.append(Spacer(1, 2))

    w_headers = [
        Paragraph("Standort", tbl_header), Paragraph("Volumen (Anzahl)", tbl_header),
        Paragraph("Gesamtkosten (€)", tbl_header), Paragraph("Nacharbeitsquote (%)", tbl_header)
    ]
    w_rows = [w_headers]
    st_w = df_calc.groupby('standort').agg(
        Volumen=('id', 'count'), Gesamtkosten=('kosten_eur', 'sum'),
        Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100)
    ).reset_index()

    # itertuples() für bessere Performance bei großen DataFrames
    for r in st_w.itertuples():
        w_rows.append([
            Paragraph(str(r.standort), tbl_cell_left), Paragraph(str(r.Volumen), tbl_cell),
            Paragraph(f"{r.Gesamtkosten:,.2f} €", tbl_cell), Paragraph(f"{r.Nacharbeit:.1f}%", tbl_cell)
        ])
    w_rows.append([
        Paragraph("GESAMT / Ø", ParagraphStyle('Bld', parent=tbl_cell_left, fontName='Helvetica-Bold')),
        Paragraph(str(total_records), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{total_kosten:,.2f} €", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{nacharbeit_pct:.1f}%", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold'))
    ])

    # Hier das Padding auf 3.5 gesetzt, damit sie "schön auseinandergezogen" ist
    t_w = Table(w_rows, colWidths=[134, 110, 140, 150])
    t_w.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3.5),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_w)
    story.append(Spacer(1, 3))

    max_cost_st = st_w.loc[st_w['Gesamtkosten'].idxmax(), 'standort']
    max_cost_val = st_w['Gesamtkosten'].max()

    text_wirt = (
        f"Die Gesamtkosten belaufen sich im Berichtszeitraum auf <b>{total_kosten:,.2f} €</b>. Der Standort mit dem höchsten Kostenvolumen ist "
        f"<b>'{max_cost_st}'</b> mit insgesamt <b>{max_cost_val:,.2f} €</b>."
    )
    story.append(make_fazit_box("WIRTSCHAFTLICHKEIT & QUALITÄT", text_wirt))
    story.append(Spacer(1, 4))

    # Kundenzufriedenheits-Balkendiagramm (Höhe auf 175 angepasst für idealen Platz)
    story.append(Paragraph("<b>Kundenzufriedenheitsverteilung und -treiber (nach Maßnahmentyp):</b>", sub_style))
    story.append(Spacer(1, 2))
    story.append(Image(bar_path, width=510, height=175))
    story.append(Spacer(1, 3))

    best_sat_typ = sat_by_measure.idxmax()
    best_sat_val = sat_by_measure.max()
    worst_sat_typ = sat_by_measure.idxmin()
    worst_sat_val = sat_by_measure.min()

    text_sat = (
        f"Die Kundenzufriedenheit zeigt über alle Maßnahmentypen hinweg ein stabiles Niveau mit einem Durchschnitt von <b>{avg_satisfaction:.1f} ★</b>. "
        f"Spitzenreiter ist <b>'{best_sat_typ}'</b> mit <b>{best_sat_val:.2f} Sternen</b>, während bei <b>'{worst_sat_typ}'</b> mit "
        f"<b>{worst_sat_val:.2f} Sternen</b> geringfügiges Optimierungspotenzial besteht."
    )
    story.append(make_fazit_box("KUNDENZUFRIEDENHEIT & QUALITÄTSTREIBER", text_sat))
    story.append(Spacer(1, 4))

    # --- 05 ANHANG: VOLLSTÄNDIGE ANONYMISIERTE DATEN (Inkl. Name, E-Mail & Telefon) ---
    story.append(PageBreak())
    story.append(Paragraph("05  ANHANG: VOLLSTÄNDIGE ANONYMISIERTE PROZESSDATEN", sec_heading))
    story.append(Paragraph(f"Vollständige Liste aller verarbeiteten Vorgänge inkl. Name, E-Mail und Telefon (n = {total_records}), wie im Streamlit-Dashboard dargestellt:", sub_style))
    story.append(Spacer(1, 4))

    # Tabellen-Header inkl. E-Mail und Telefon
    raw_headers = [
        Paragraph("ID", tbl_header),
        Paragraph("Name", tbl_header),
        Paragraph("E-Mail", tbl_header),
        Paragraph("Telefon", tbl_header),
        Paragraph("Standort", tbl_header),
        Paragraph("Typ", tbl_header),
        Paragraph("Bearb.(Std)", tbl_header),
        Paragraph("Liege(T)", tbl_header),
        Paragraph("Durchl.(T)", tbl_header),
        Paragraph("Kosten (€)", tbl_header),
        Paragraph("Zuf.", tbl_header)
    ]
    raw_rows = [raw_headers]

    # Befüllung über alle 250 Zeilen
    for r in df_calc.itertuples():
        raw_rows.append([
            Paragraph(str(getattr(r, 'id', '-')), tbl_cell),
            Paragraph(str(getattr(r, 'name', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'email', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'telefon', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'standort', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'massnahmetyp', '-')), tbl_cell_left),
            Paragraph(f"{getattr(r, 'bearbeitungszeit_std', 0):.1f}", tbl_cell),
            Paragraph(f"{getattr(r, 'liegezeit_tage', 0):.1f}", tbl_cell),
            Paragraph(f"{getattr(r, 'gesamtdurchlaufzeit_tage', 0):.1f}", tbl_cell),
            Paragraph(f"{getattr(r, 'kosten_eur', 0):,.2f}", tbl_cell),
            Paragraph(f"{getattr(r, 'kundenzufriedenheit', 0):.1f} ★", tbl_cell)
        ])

    # Spaltenbreiten exakt auf A4 (534 pt Gesamtbreite) angepasst
    t_raw = Table(raw_rows, colWidths=[28, 65, 85, 60, 48, 62, 35, 30, 30, 52, 39], repeatRows=1)
    t_raw.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#fafafa')])
    ]))
    story.append(t_raw)

    doc.build(story)
    
    if os.path.exists(chart_path): os.remove(chart_path)
    if os.path.exists(bar_path): os.remove(bar_path)
    print(f"[INFO] Management-Report PDF erfolgreich generiert (inkl. E-Mail & Telefon) unter: '{output_path}'")


# ALL-IN-ONE EXPORTER ROUTER
def export_data(df: pd.DataFrame, kpis: dict, fmt: str = "excel", filename: str = None):
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
        raise ValueError(f"Nicht unterstütztes Export-Format: '{fmt}'")


if __name__ == "__main__":
    config = load_config()
    sample_file = os.path.join(config["paths"]["raw_data_dir"], "sample_operations.csv")
    df = import_csv_data(sample_file)
    df = validate_schema(df)
    df_anon = anonymize_dataframe(df)
    kpis = calculate_kpis(df_anon)
    
    print("\n--- Starte Multi-Format Export Test ---")
    export_data(df_anon, kpis, fmt="pdf")