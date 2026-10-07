import sys
from pathlib import Path

# Pfad absichern, damit imports aus src/ greifen
root_dir = Path(__file__).resolve().parent.parent.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

import os
import pandas as pd
from src.etl.importer import load_config, import_csv_data, validate_schema
from src.etl.anonymizer import anonymize_dataframe
from src.analytics.kpi_calculator import calculate_kpis

def export_to_excel(df: pd.DataFrame, kpis: dict, output_filename: str = "pipeline_report.xlsx"):
    """Exportiert die verarbeiteten Daten und KPIs in einen strukturierten Excel-Bericht."""
    config = load_config()
    processed_dir = config["paths"]["processed_data_dir"]
    
    # Sicherstellen, dass das Verzeichnis existiert
    os.makedirs(processed_dir, exist_ok=True)
    output_path = os.path.join(processed_dir, output_filename)
    
    # Excel-Writer mit Pandas initialisieren (nutzt openpyxl im Hintergrund)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # 1. Hauptdaten (anonymisiert) auf ein Tabellenblatt schreiben
        df.to_excel(writer, sheet_name="Verarbeitete_Daten", index=False)
        
        # 2. KPI-Übersicht in ein separates Tabellenblatt schreiben
        summary_data = {
            "Metrik": [
                "Gesamtdatensätze",
                "Ø Durchlaufzeit gesamt (Tage)"
            ],
            "Wert": [
                kpis["total_records"],
                kpis["overall_avg_duration"]
            ]
        }
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name="KPI_Zusammenfassung", index=False)
        
    print(f"[INFO] Excel-Bericht erfolgreich unter '{output_path}' gespeichert.")

if __name__ == "__main__":
    # Testlauf für den Excel-Export
    config = load_config()
    sample_file = os.path.join(config["paths"]["raw_data_dir"], "sample_operations.csv")
    
    # Pipeline-Schritte durchlaufen
    df = import_csv_data(sample_file)
    df = validate_schema(df)
    df_anon = anonymize_dataframe(df)
    kpis = calculate_kpis(df_anon)
    
    # Export ausführen
    export_to_excel(df_anon, kpis)