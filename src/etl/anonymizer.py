import hashlib
import pandas as pd
from src.etl.importer import load_config

def hash_value(val: str) -> str:
    """Erstellt einen anonymisierten Hash-Wert für einen String."""
    if pd.isna(val) or val == "":
        return val
    # SHA256 Hash erzeugen und auf 12 Zeichen kürzen (für bessere Lesbarkeit)
    return hashlib.sha256(str(val).encode("utf-8")).hexdigest()[:12]

def anonymize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Anonymisiert sensible Spalten basierend auf der Konfigurationsdatei."""
    config = load_config()
    cols_to_anonymize = config["processing"].get("anonymize_columns", [])
    
    df_encoded = df.copy()
    for col in cols_to_anonymize:
        if col in df_encoded.columns:
            df_encoded[col] = df_encoded[col].apply(hash_value)
            print(f"[INFO] Spalte '{col}' erfolgreich anonymisiert.")
        else:
            print(f"[WARNUNG] Spalte '{col}' aus Config nicht im DataFrame gefunden.")
            
    return df_encoded

if __name__ == "__main__":
    # Testlauf für den Anonymizer
    from src.etl.importer import import_csv_data, validate_schema
    import os
    
    config = load_config()
    sample_file = os.path.join(config["paths"]["raw_data_dir"], "sample_operations.csv")
    
    df = import_csv_data(sample_file)
    df = validate_schema(df)
    df_anon = anonymize_dataframe(df)
    
    print("\n--- Daten nach Anonymisierung ---")
    print(df_anon[["id", "name", "email", "standort"]])