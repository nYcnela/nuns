# Używamy lekkiego obrazu Pythona
FROM python:3.9-slim

# Ustawiamy katalog roboczy w kontenerze
WORKDIR /app

# Kopiujemy plik z zależnościami
COPY requirements.txt .

# Instalujemy zależności
RUN pip install --no-cache-dir -r requirements.txt

# Kopiujemy resztę plików aplikacji
COPY . .

# Otwieramy port 8501 (domyślny dla Streamlit)
EXPOSE 8501

# Uruchamiamy aplikację, nasłuchując na wszystkich interfejsach (0.0.0.0)
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0"]
