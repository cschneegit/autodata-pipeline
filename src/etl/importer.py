import os
import yaml
import pandas as pd
from pathlib import Path

def load_config(config_path: str = "config/pipeline_config.yaml") -> dict:
    """Lädt die zentrale Pipeline-Konfiguration."""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def import_csv_data(file_path: str) -> pd.DataFrame:
    """Liest eine Rohdaten-CSV ein und gibt ein Pandas DataFrame zurück."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Die Datei {file_path} wurde nicht gefunden.")
    
    df = pd.read_csv(file_path)
    print(f"[INFO] Erfolgreich eingelesen: {file_path} ({len(df)} Zeilen)")
    return df

def validate_schema(df: pd.DataFrame) -> pd.DataFrame:
    """Validiert das DataFrame auf erforderliche Basisschemas und bereinigt fehlende Werte."""
    config = load_config()
    strategy = config["processing"].get("missing_value_strategy", "drop")
    
    # Beispielhafte Pflichtspaltenprüfung
    required_columns = ["id", "datum", "standort", "status"]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Pflichtspalte '{col}' fehlt im Datensatz!")
            
    # Umgang mit fehlenden Werten
    initial_count = len(df)
    if strategy == "drop":
        df = df.dropna(subset=["id", "datum"])
    elif strategy == "fill_zero":
        df = df.fillna(0)
        
    print(f"[INFO] Validierung abgeschlossen. Zeilen vorab: {initial_count}, nach Bereinigung: {len(df)}")
    return df

if __name__ == "__main__":
    # Kleiner Testlauf direkt im Skript
    config = load_config()
    raw_dir = config["paths"]["raw_data_dir"]
    sample_file = os.path.join(raw_dir, "sample_operations.csv")
    
    df_raw = import_csv_data(sample_file)
    df_clean = validate_schema(df_raw)
    print(df_clean.head())