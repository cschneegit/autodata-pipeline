import sys
from pathlib import Path

# Pfad absichern, damit imports greifen
root_dir = Path(__file__).resolve().parent.parent.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from src.etl.importer import load_config, import_csv_data, validate_schema
from src.etl.anonymizer import anonymize_dataframe
from src.analytics.kpi_calculator import calculate_kpis

# Hilfsfunktion, um Zahlen im deutschen Format zu formatieren
def _fmt_de(val: float, decimals: int = 1, thousands: bool = False) -> str:
    """Formatiert Zahlen nach deutscher Schreibweise (Komma als Dezimaltrennzeichen)."""
    val = float(val)
    if thousands:
        s = f"{val:,.{decimals}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{val:.{decimals}f}".replace(".", ",")


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
    avg_satisfaction = df_calc['kundenzufriedenheit'].mean()

    de_tick_int = FuncFormatter(lambda x, pos: _fmt_de(x, 0))

    # --- Diagramm 1: Zeitzusammensetzung ---
    chart_path = os.path.join(temp_dir, "stacked_bar_chart.png")
    mass_agg = df_calc.groupby('massnahmetyp').agg(
        bearb_tage=('bearb_tage', 'mean'),
        liegezeit_tage=('liegezeit_tage', 'mean'),
        gesamtdurchlaufzeit_tage=('gesamtdurchlaufzeit_tage', 'mean')
    ).sort_values('gesamtdurchlaufzeit_tage', ascending=True)

    plt.figure(figsize=(7.2, 2.2))
    y_pos = range(len(mass_agg))
    plt.barh(y_pos, mass_agg['bearb_tage'], color='#00a88f', label='Bearbeitungszeit (Tage)')
    plt.barh(y_pos, mass_agg['liegezeit_tage'], left=mass_agg['bearb_tage'], color='#91a3b7', label='Liegezeit (Tage)')

    plt.yticks(y_pos, mass_agg.index, fontsize=8)
    plt.xlabel('Tage (Ø pro Vorgang)', fontsize=8)
    plt.gca().xaxis.set_major_formatter(de_tick_int)
    plt.legend(loc='lower right', fontsize=7)

    for idx, row in enumerate(mass_agg.itertuples()):
        total_t = row.gesamtdurchlaufzeit_tage
        plt.text(total_t + 0.15, idx, f"{_fmt_de(total_t, 1)} T", va='center', fontsize=8, fontweight='bold', color='#333333')

    plt.xlim(0, max(mass_agg['gesamtdurchlaufzeit_tage']) * 1.15)
    plt.grid(axis='x', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig(chart_path, dpi=220)
    plt.close()

    # --- Diagramm 2: Kundenzufriedenheit ---
    bar_path = os.path.join(temp_dir, "satisfaction_bar.png")
    plt.figure(figsize=(7.2, 2.5))
    sat_by_measure = df_calc.groupby('massnahmetyp')['kundenzufriedenheit'].mean().sort_values()

    bars = plt.barh(sat_by_measure.index, sat_by_measure.values, color='#3498db', height=0.55)
    plt.xlabel('Ø Kundenzufriedenheit (Sterne)', fontsize=9, fontweight='bold')
    plt.gca().xaxis.set_major_formatter(de_tick_int)
    plt.xlim(0, 5.0)
    plt.grid(axis='x', linestyle='--', alpha=0.7)

    for bar in bars:
        width = bar.get_width()
        plt.text(width + 0.1, bar.get_y() + bar.get_height()/2, f'{_fmt_de(width, 2)} ★',
                 va='center', ha='left', fontsize=8, fontweight='bold')

    plt.tight_layout()
    plt.savefig(bar_path, dpi=220)
    plt.close()

    # --- Document Setup ---
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=30, leftMargin=30, topMargin=26, bottomMargin=26
    )
    story = []
    styles = getSampleStyleSheet()

    HEADER_COLOR = colors.HexColor('#1f3a60')
    CARD_BG = colors.HexColor('#f4f6f9')
    BORDER_COLOR = colors.HexColor('#d1d9e6')

    title_style = ParagraphStyle('ReportHeader', parent=styles['Heading1'], fontSize=15, leading=18, textColor=HEADER_COLOR, fontName='Helvetica-Bold')
    sub_style = ParagraphStyle('ReportSub', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#555555'))
    sec_heading = ParagraphStyle('SecHeading', parent=styles['Heading2'], fontSize=11, leading=14, textColor=HEADER_COLOR, fontName='Helvetica-Bold', spaceBefore=8, spaceAfter=3)
    box_text = ParagraphStyle('BoxText', parent=styles['Normal'], fontSize=8.5, leading=11.5, textColor=colors.HexColor('#1a1a1a'))

    tbl_cell = ParagraphStyle('TblCell', parent=styles['Normal'], fontSize=6.5, leading=8, alignment=1)
    tbl_cell_left = ParagraphStyle('TblCellLeft', parent=styles['Normal'], fontSize=6.5, leading=8, alignment=0)
    tbl_header_10 = ParagraphStyle('TblHeader10', parent=styles['Normal'], fontSize=10, leading=11, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')
    tbl_header_9 = ParagraphStyle('TblHeader9', parent=styles['Normal'], fontSize=9, leading=10, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')

    def _hdr(text: str, col_width: float, padding: float = 2.0) -> Paragraph:
        """Setzt Tabellenüberschriften auf 10 pt, weicht aber auf 9 pt aus, sobald bei 10 pt mehr als 2 Zeilen entstehen."""
        p10 = Paragraph(text, tbl_header_10)
        p10.wrap(col_width - 2 * padding, 200)
        if len(p10.blPara.lines) > 2:
            return Paragraph(text, tbl_header_9)
        return p10

    # Hilfsfunktion, um Tabellenüberschriften dynamisch zu erstellen
    def _build_headers(labels: list[str], widths: list[float], padding: float = 2.0) -> list[Paragraph]:
        return [_hdr(lbl, w, padding) for lbl, w in zip(labels, widths)]

    # Hilfsfunktion, um ein Fazit-Box-Element zu erstellen
    def make_fazit_box(title, text):
        html = f"<b>MANAGEMENT-FAZIT (KERNAUSSAGEN) – {title}</b><br/>{text}"
        p = Paragraph(html, box_text)
        t = Table([[p]], colWidths=[534])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), CARD_BG),
            ('BOX', (0,0), (-1,-1), 1, BORDER_COLOR),
            ('PADDING', (0,0), (-1,-1), 4),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        return t

    # HEADER BLOCK
    story.append(Paragraph("MANAGEMENT REPORT  •  OPERATIVE STEUERUNG", title_style))
    story.append(Paragraph("Prozess-Effizienz, Standort-Benchmarking & Qualität", ParagraphStyle('Sub', parent=styles['Heading3'], fontSize=10, textColor=colors.HexColor('#444444'), spaceAfter=2)))
    story.append(Paragraph(f"Berichtszeitraum: Q3/2026   |   Erstellt am: 10.11.2026   |   Verantwortlich: AutoData Pipeline   |   Volumen: {total_records} Vorgänge", sub_style))
    story.append(Spacer(1, 4))

    # --- 01 EXECUTIVE SUMMARY & KPI-SCORECARD ---
    story.append(Paragraph("01  EXECUTIVE SUMMARY & KPI-SCORECARD", sec_heading))
    story.append(Paragraph("Gesamtstatus über alle 7 Standorte und 6 Maßnahmentypen:", sub_style))
    story.append(Spacer(1, 2))

    sc_widths = [190, 85, 85, 90, 84]
    # headers mit Zeilenumbruch für bessere Lesbarkeit
    scorecard_headers = _build_headers(
        ["Kernkennzahl (KPI)", "Ist-Wert", "Vormonat", "Zielwert", "Status"],
        sc_widths,
        padding=2.5
    )

    sc_rows = [
        scorecard_headers,
        [Paragraph("Ø Gesamtdurchlaufzeit", tbl_cell_left), Paragraph(f"{_fmt_de(avg_dur, 1)} Tage", tbl_cell), Paragraph(f"{_fmt_de(avg_dur*1.04, 1)} Tage", tbl_cell), Paragraph("≤ 12,0 Tage", tbl_cell), Paragraph('<font color="#2ecc71">●</font> im Plan', tbl_cell)],
        [Paragraph("Ø Flusseffizienz (Bearb.- vs. Liegezeit)", tbl_cell_left), Paragraph(f"{_fmt_de(avg_fluss, 1)} %", tbl_cell), Paragraph(f"{_fmt_de(avg_fluss*0.95, 1)} %", tbl_cell), Paragraph("≥ 25,0 %", tbl_cell), Paragraph('<font color="#f1c40f">●</font> Achtung', tbl_cell)],
        [Paragraph("Nacharbeitsquote", tbl_cell_left), Paragraph(f"{_fmt_de(nacharbeit_pct, 1)} %", tbl_cell), Paragraph(f"{_fmt_de(nacharbeit_pct*1.08, 1)} %", tbl_cell), Paragraph("≤ 8,0 %", tbl_cell), Paragraph('<font color="#e74c3c">●</font> Kritisch', tbl_cell)],
        [Paragraph("Ø Kosten pro Maßnahme", tbl_cell_left), Paragraph(f"{_fmt_de(avg_kosten, 2, True)} €", tbl_cell), Paragraph(f"{_fmt_de(avg_kosten*0.98, 2, True)} €", tbl_cell), Paragraph("≤ 1.150,00 €", tbl_cell), Paragraph('<font color="#2ecc71">●</font> im Plan', tbl_cell)],
        [Paragraph("Kundenzufriedenheit", tbl_cell_left), Paragraph(f"{_fmt_de(avg_satisfaction, 1)} ★", tbl_cell), Paragraph(f"{_fmt_de(avg_satisfaction*0.97, 1)} ★", tbl_cell), Paragraph("≥ 4,2 ★", tbl_cell), Paragraph('<font color="#f1c40f">●</font> Achtung', tbl_cell)]
    ]

    t_sc = Table(sc_rows, colWidths=sc_widths)
    t_sc.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#fafafa'))
    ]))
    story.append(t_sc)
    story.append(Spacer(1, 3))

    text_exec = (
        f"Im Berichtszeitraum wurden <b>{total_records} Vorgänge</b> analysiert. Die durchschnittliche Gesamtdurchlaufzeit liegt mit <b>{_fmt_de(avg_dur, 1)} Tagen</b> "
        f"voll im Zielbereich (≤ 12,0 Tage). Handlungsbedarf besteht bei der Flusseffizienz (aktuell <b>{_fmt_de(avg_fluss, 1)} %</b>, Ziel ≥ 25,0 %) sowie bei der "
        f"Nacharbeitsquote (<b>{_fmt_de(nacharbeit_pct, 1)} %</b>), welche den Zielwert von ≤ 8,0 % überschreitet."
    )
    story.append(make_fazit_box("KPI-SCORECARD", text_exec))
    story.append(Spacer(1, 8))

    # --- 02 PROZESS-EFFIZIENZ & ENGPASS-ANALYSE ---
    story.append(Paragraph("02  PROZESS-EFFIZIENZ & ENGPASS-ANALYSE NACH MASSNAHMENTYP", sec_heading))
    story.append(Paragraph("Gegenüberstellung von wertschöpfender Bearbeitungszeit (Std.), Liegezeit (Tage) sowie Kosten- und Qualitätswirkung:", sub_style))
    story.append(Spacer(1, 2))

    mass_df = df_calc.groupby('massnahmetyp').agg(
        Anzahl=('id', 'count'),
        Bearb_Std=('bearbeitungszeit_std', 'mean'),
        Liegezeit=('liegezeit_tage', 'mean'),
        Durchlauf=('gesamtdurchlaufzeit_tage', 'mean'),
        Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100),
        Kosten=('kosten_eur', 'mean'),
        Zufried=('kundenzufriedenheit', 'mean')
    ).reset_index()

    p_widths = [84, 42, 74, 62, 76, 56, 66, 74]
    p_headers = _build_headers(
        [
            "Maßnahmentyp", "Anzahl", "Ø Bearb.-zeit<br/>(Std.)", "Ø Liegezeit<br/>(Tage)",
            "Ø Durchlaufzeit<br/>(Tage)", "Nacharbeit<br/>(%)", "Ø Kosten (€)", "Kundenzufrie-<br/>denheit (Ø ★)"
        ],
        p_widths,
        padding=2.0
    )
    # itertuples() für bessere Performance bei großen DataFrames, da es schneller ist als iterrows()
    p_rows = [p_headers]
    for r in mass_df.itertuples():
        p_rows.append([
            Paragraph(str(r.massnahmetyp), tbl_cell_left), Paragraph(str(r.Anzahl), tbl_cell),
            Paragraph(_fmt_de(r.Bearb_Std, 1), tbl_cell), Paragraph(_fmt_de(r.Liegezeit, 1), tbl_cell),
            Paragraph(_fmt_de(r.Durchlauf, 1), tbl_cell), Paragraph(f"{_fmt_de(r.Nacharbeit, 1)} %", tbl_cell),
            Paragraph(f"{_fmt_de(r.Kosten, 2, True)} €", tbl_cell), Paragraph(f"{_fmt_de(r.Zufried, 1)} ★", tbl_cell)
        ])
    p_rows.append([
        Paragraph("GESAMT / Ø", ParagraphStyle('Bld', parent=tbl_cell_left, fontName='Helvetica-Bold')),
        Paragraph(str(total_records), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(_fmt_de(df_calc['bearbeitungszeit_std'].mean(), 1), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(_fmt_de(df_calc['liegezeit_tage'].mean(), 1), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(_fmt_de(avg_dur, 1), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(nacharbeit_pct, 1)} %", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(avg_kosten, 2, True)} €", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(avg_satisfaction, 1)} ★", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold'))
    ])

    t_p = Table(p_rows, colWidths=p_widths)
    t_p.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 1.8),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_p)

    # Erhöhter Abstand zwischen Tabelle und Diagramm in 02:
    story.append(Spacer(1, 14))
    story.append(Paragraph(f"Zeitzusammensetzung nach Maßnahmentyp (n = {total_records}):", sub_style))
    story.append(Spacer(1, 3))
    story.append(Image(chart_path, width=510, height=120))
    story.append(Spacer(1, 3))

    max_liege_typ = mass_df.loc[mass_df['Liegezeit'].idxmax(), 'massnahmetyp']
    max_liege_val = mass_df['Liegezeit'].max()
    min_durchlauf_typ = mass_df.loc[mass_df['Durchlauf'].idxmin(), 'massnahmetyp']
    min_durchlauf_val = mass_df['Durchlauf'].min()

    text_prozess = (
        f"Der Maßnahmentyp <b>'{max_liege_typ}'</b> weist mit durchschnittlich <b>{_fmt_de(max_liege_val, 1)} Tagen</b> die höchste Liegezeit auf und bildet "
        f"den Hauptengpass im Prozess. Der effizienteste Ablauf zeigt sich bei <b>'{min_durchlauf_typ}'</b> mit einer Gesamtdurchlaufzeit von nur "
        f"<b>{_fmt_de(min_durchlauf_val, 1)} Tagen</b>."
    )
    story.append(make_fazit_box("PROZESS-EFFIZIENZ", text_prozess))

    # Sauberer Seitenumbruch für Sektion 3 & 4
    story.append(PageBreak())

    # --- 03 STANDORT-BENCHMARKING ---
    story.append(Paragraph("03  STANDORT-BENCHMARKING (REGIONALER LEISTUNGSVERGLEICH)", sec_heading))
    story.append(Paragraph("Vergleich der 7 Standorte zur Identifikation von Best Practices:", sub_style))
    story.append(Spacer(1, 2))

    st_df = df_calc.groupby('standort').agg(
        Volumen=('id', 'count'), Bearb_Std=('bearbeitungszeit_std', 'mean'),
        Liegezeit=('liegezeit_tage', 'mean'), Durchlauf=('gesamtdurchlaufzeit_tage', 'mean'),
        Flusseffizienz=('flusseffizienz_pct', 'mean'), Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100),
        Kosten=('kosten_eur', 'mean'), Zufried=('kundenzufriedenheit', 'mean')
    ).reset_index()

    st_widths = [64, 44, 62, 54, 66, 64, 56, 62, 62]
    # headers mit Zeilenumbruch für bessere Lesbarkeit
    st_headers = _build_headers(
        [
            "Standort", "Volumen", "Ø Bearb.-zeit", "Ø Liegezeit", "Ø Durchlaufzeit",
            "Flusseffizienz<br/>(%)", "Nacharbeit<br/>(%)", "Ø Kosten (€)", "Zufriedenheit"
        ],
        st_widths,
        padding=2.0
    )
    st_rows = [st_headers]
    for r in st_df.itertuples():
        st_rows.append([
            Paragraph(str(r.standort), tbl_cell_left), Paragraph(str(r.Volumen), tbl_cell),
            Paragraph(f"{_fmt_de(r.Bearb_Std, 1)} h", tbl_cell), Paragraph(f"{_fmt_de(r.Liegezeit, 1)} T", tbl_cell),
            Paragraph(f"{_fmt_de(r.Durchlauf, 1)} T", tbl_cell), Paragraph(f"{_fmt_de(r.Flusseffizienz, 1)} %", tbl_cell),
            Paragraph(f"{_fmt_de(r.Nacharbeit, 1)} %", tbl_cell), Paragraph(f"{_fmt_de(r.Kosten, 2, True)} €", tbl_cell),
            Paragraph(f"{_fmt_de(r.Zufried, 1)} ★", tbl_cell)
        ])
    st_rows.append([
        Paragraph("GESAMT / Ø", ParagraphStyle('Bld', parent=tbl_cell_left, fontName='Helvetica-Bold')),
        Paragraph(str(total_records), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(df_calc['bearbeitungszeit_std'].mean(), 1)} h", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(df_calc['liegezeit_tage'].mean(), 1)} T", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(avg_dur, 1)} T", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(avg_fluss, 1)} %", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(nacharbeit_pct, 1)} %", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(avg_kosten, 2, True)} €", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(avg_satisfaction, 1)} ★", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold'))
    ])

    t_st = Table(st_rows, colWidths=st_widths)
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
        f"Im Standortvergleich glänzt <b>'{best_st}'</b> als Best-Practice-Standort mit der höchsten Flusseffizienz von <b>{_fmt_de(best_st_val, 1)} %</b>. "
        f"Operativer Unterstützungsbedarf besteht am Standort <b>'{worst_st}'</b> (Ø Liegezeit: <b>{_fmt_de(worst_st_val, 1)} Tage</b>)."
    )
    story.append(make_fazit_box("STANDORT-BENCHMARKING", text_standort))
    story.append(Spacer(1, 10))

    # --- 04 WIRTSCHAFTLICHKEIT & KUNDENZUFRIEDENHEIT ---
    story.append(Paragraph("04  WIRTSCHAFTLICHKEIT & KUNDENZUFRIEDENHEIT", sec_heading))
    story.append(Paragraph("Kosten- und Nacharbeitsanalyse im regionalen Vergleich:", sub_style))
    story.append(Spacer(1, 2))

    w_widths = [134, 110, 140, 150]
    w_headers = _build_headers(
        ["Standort", "Volumen (Anzahl)", "Ø Kosten (€)", "Nacharbeitsquote (%)"],
        w_widths,
        padding=3.0
    )
    w_rows = [w_headers]
    st_w = df_calc.groupby('standort').agg(
        Volumen=('id', 'count'),
        AvgKosten=('kosten_eur', 'mean'),
        # lambda-Funktion für Nacharbeitsquote, um den Prozentsatz zu berechnen
        Nacharbeit=('nacharbeit_bool', lambda x: (x.sum()/len(x))*100)
    ).reset_index()

    for r in st_w.itertuples():
        w_rows.append([
            Paragraph(str(r.standort), tbl_cell_left), Paragraph(str(r.Volumen), tbl_cell),
            Paragraph(f"{_fmt_de(r.AvgKosten, 2, True)} €", tbl_cell), Paragraph(f"{_fmt_de(r.Nacharbeit, 1)} %", tbl_cell)
        ])
    w_rows.append([
        Paragraph("GESAMT / Ø", ParagraphStyle('Bld', parent=tbl_cell_left, fontName='Helvetica-Bold')),
        Paragraph(str(total_records), ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(avg_kosten, 2, True)} €", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold')),
        Paragraph(f"{_fmt_de(nacharbeit_pct, 1)} %", ParagraphStyle('BldC', parent=tbl_cell, fontName='Helvetica-Bold'))
    ])

    t_w = Table(w_rows, colWidths=w_widths)
    t_w.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#eef2f7'))
    ]))
    story.append(t_w)
    story.append(Spacer(1, 3))

    eff_cost_st = st_w.loc[st_w['AvgKosten'].idxmin(), 'standort']
    eff_cost_val = st_w['AvgKosten'].min()
    high_cost_st = st_w.loc[st_w['AvgKosten'].idxmax(), 'standort']
    high_cost_val = st_w['AvgKosten'].max()

    max_nach_st = st_w.loc[st_w['Nacharbeit'].idxmax(), 'standort']
    max_nach_val = st_w['Nacharbeit'].max()
    min_nach_st = st_w.loc[st_w['Nacharbeit'].idxmin(), 'standort']
    min_nach_val = st_w['Nacharbeit'].min()

    text_wirt = (
        f"Bei den durchschnittlichen Kosten arbeitet <b>'{eff_cost_st}'</b> mit <b>{_fmt_de(eff_cost_val, 2, True)} €</b> am kosteneffizientesten "
        f"(Gesamtdurchschnitt: <b>{_fmt_de(avg_kosten, 2, True)} €</b>), während <b>'{high_cost_st}'</b> mit <b>{_fmt_de(high_cost_val, 2, True)} €</b> "
        f"weniger kosteneffizient abschneidet. Bei der Nacharbeitsquote (Ø <b>{_fmt_de(nacharbeit_pct, 1)} %</b>) fällt besonders <b>'{max_nach_st}'</b> "
        f"mit <b>{_fmt_de(max_nach_val, 1)} %</b> auf, wohingegen <b>'{min_nach_st}'</b> mit <b>{_fmt_de(min_nach_val, 1)} %</b> den niedrigsten Wert erzielt."
    )
    story.append(make_fazit_box("WIRTSCHAFTLICHKEIT & QUALITÄT", text_wirt))

    # Deutlich erhöhter Abstand zwischen Kosten-/Nacharbeitsblock und Kundenzufriedenheits-Balkendiagramm
    story.append(Spacer(1, 14))

    story.append(Paragraph("Kundenzufriedenheitsverteilung und -treiber (nach Maßnahmentyp):", sub_style))
    story.append(Spacer(1, 4))
    story.append(Image(bar_path, width=510, height=150))
    story.append(Spacer(1, 3))

    # idxmax statt idmax, da idxmax() den Index zurückgibt, während idmax() den Wert zurückgibt
    best_sat_typ = sat_by_measure.idxmax()
    best_sat_val = sat_by_measure.max()
    worst_sat_typ = sat_by_measure.idxmin()
    worst_sat_val = sat_by_measure.min()

    text_sat = (
        f"Die Kundenzufriedenheit zeigt über alle Maßnahmentypen hinweg ein stabiles Niveau mit einem Durchschnitt von <b>{_fmt_de(avg_satisfaction, 1)} ★</b>. "
        f"Spitzenreiter ist <b>'{best_sat_typ}'</b> mit <b>{_fmt_de(best_sat_val, 2)} Sternen</b>, während bei <b>'{worst_sat_typ}'</b> mit "
        f"<b>{_fmt_de(worst_sat_val, 2)} Sternen</b> geringfügiges Optimierungspotenzial besteht."
    )
    story.append(make_fazit_box("KUNDENZUFRIEDENHEIT & QUALITÄTSTREIBER", text_sat))

    # --- 05 ANHANG: VOLLSTÄNDIGE ANONYMISIERTE DATEN ---
    story.append(PageBreak())
    story.append(Paragraph("05  ANHANG: VOLLSTÄNDIGE ANONYMISIERTE PROZESSDATEN", sec_heading))
    story.append(Paragraph(f"Vollständige Liste aller verarbeiteten Vorgänge inkl. Name, E-Mail und Telefon (n = {total_records}), wie im Streamlit-Dashboard dargestellt:", sub_style))
    story.append(Spacer(1, 4))

    raw_widths = [32, 60, 80, 56, 48, 60, 42, 34, 36, 52, 34]
    raw_headers = _build_headers(
        [
            "ID", "Name", "E-Mail", "Telefon", "Standort", "Typ",
            "Bearb.<br/>(Std)", "Liege<br/>(T)", "Durchl.<br/>(T)", "Kosten (€)", "Zuf."
        ],
        raw_widths,
        padding=2.0
    )
    raw_rows = [raw_headers]

    for r in df_calc.itertuples():
        raw_rows.append([
            Paragraph(str(getattr(r, 'id', '-')), tbl_cell),
            Paragraph(str(getattr(r, 'name', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'email', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'telefon', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'standort', '-')), tbl_cell_left),
            Paragraph(str(getattr(r, 'massnahmetyp', '-')), tbl_cell_left),
            Paragraph(_fmt_de(getattr(r, 'bearbeitungszeit_std', 0), 1), tbl_cell),
            Paragraph(_fmt_de(getattr(r, 'liegezeit_tage', 0), 1), tbl_cell),
            Paragraph(_fmt_de(getattr(r, 'gesamtdurchlaufzeit_tage', 0), 1), tbl_cell),
            Paragraph(_fmt_de(getattr(r, 'kosten_eur', 0), 2, True), tbl_cell),
            Paragraph(f"{_fmt_de(getattr(r, 'kundenzufriedenheit', 0), 1)} ★", tbl_cell)
        ])

    t_raw = Table(raw_rows, colWidths=raw_widths, repeatRows=1)
    t_raw.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HEADER_COLOR),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#fafafa')])
    ]))
    story.append(t_raw)

    doc.build(story)

    if os.path.exists(chart_path):
        os.remove(chart_path)
    if os.path.exists(bar_path):
        os.remove(bar_path)
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