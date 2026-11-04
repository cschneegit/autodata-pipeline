import pandas as pd

def calculate_kpis(df: pd.DataFrame) -> dict:
    """Berechnet zentrale und erweiterte KPIs aus dem DataFrame für Reporting und Dashboard."""
    if df.empty:
        return {"total_records": 0, "overall_avg_duration": 0.0}

    total_records = len(df)
    
    # Durchlaufzeiten analysieren (unterstützt neue und alte Spaltennamen)
    duration_col = None
    for col in ["gesamtdurchlaufzeit_tage", "durchlaufzeit_tage", "duration"]:
        if col in df.columns:
            duration_col = col
            break

    if duration_col:
        overall_avg_duration = float(df[duration_col].mean())
        max_duration = float(df[duration_col].max())
        min_duration = float(df[duration_col].min())
        
        # Ausführliche Detail-Infos für die Textform ermitteln
        max_idx = df[duration_col].idxmax()
        min_idx = df[duration_col].idxmin()
        
        cat_col = "status" if "status" in df.columns else df.columns[0]
        id_col = "id" if "id" in df.columns else df.columns[0]
        
        max_detail = f"Datensatz ID {df.loc[max_idx, id_col]} (Status: {df.loc[max_idx, cat_col]})"
        min_detail = f"Datensatz ID {df.loc[min_idx, id_col]} (Status: {df.loc[min_idx, cat_col]})"
    else:
        overall_avg_duration = 0.0
        max_duration = 0.0
        min_duration = 0.0
        max_detail = "N/A"
        min_detail = "N/A"
        
    status_counts = df["status"].value_counts().to_dict() if "status" in df.columns else {}

    kpis = {
        "total_records": total_records,
        "overall_avg_duration": round(overall_avg_duration, 2),
        "max_duration": round(max_duration, 2),
        "min_duration": round(min_duration, 2),
        "max_detail": max_detail,
        "min_detail": min_detail,
        "status_counts": status_counts
    }
    
    return kpis