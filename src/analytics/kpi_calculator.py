import pandas as pd

def calculate_kpis(df: pd.DataFrame) -> dict:
    """Berechnet zentrale und erweiterte KPIs aus dem DataFrame für Reporting und Dashboard."""
    if df.empty:
        return {"total_records": 0, "overall_avg_duration": 0.0}

    total_records = len(df)
    
    # Durchlaufzeiten analysieren (unterstützt 'durchlaufzeit_tage' oder 'duration')
    duration_col = None
    for col in ["durchlaufzeit_tage", "duration"]:
        if col in df.columns:
            duration_col = col
            break

    if duration_col:
        overall_avg_duration = float(df[duration_col].mean())
        max_duration = float(df[duration_col].max())
        min_duration = float(df[duration_col].min())
    else:
        overall_avg_duration = 0.0
        max_duration = 0.0
        min_duration = 0.0
        
    # Status-Verteilung berechnen (falls 'status' vorhanden ist)
    status_counts = df["status"].value_counts().to_dict() if "status" in df.columns else {}

    kpis = {
        "total_records": total_records,
        "overall_avg_duration": round(overall_avg_duration, 2),
        "max_duration": round(max_duration, 2),
        "min_duration": round(min_duration, 2),
        "status_counts": status_counts
    }
    
    return kpis