# Offizielles, schlankes Python-Image als Basis
FROM python:3.11-slim

# Arbeitsverzeichnis im Container festlegen
WORKDIR /app

# Abhängigkeiten zuerst kopieren (nutzt Docker-Caching optimal aus)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Den restlichen Projektcode in den Container kopieren
COPY . .

# Den Port für Streamlit freigeben
EXPOSE 8501

# Befehl zum Starten des Streamlit-Dashboards im Container
CMD ["streamlit", "run", "src/dashboard.py", "--server.address=0.0.0.0", "--server.port=8501"]