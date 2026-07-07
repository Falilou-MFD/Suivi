FROM python:3.11-slim

# Reflex a besoin de Node.js (qui inclut npm) et unzip
RUN apt-get update && apt-get install -y curl unzip && \
    curl -fsSL https://deb.nodesource.com/setup_18.x | bash - && \
    apt-get install -y nodejs && \
    rm -rf /var/lib/apt/lists/*

RUN apt-get update && apt-get install -y \
    wget \
    unzip \
    && rm -rf /var/lib/apt/lists/*


RUN wget https://github.com/duckdb/duckdb/releases/download/v1.1.3/duckdb_cli-linux-amd64.zip \
    && unzip duckdb_cli-linux-amd64.zip -d /usr/local/bin \
    && rm duckdb_cli-linux-amd64.zip \
    && chmod +x /usr/local/bin/duckdb

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Initialisation du projet Reflex
RUN reflex init --name suivi_app --template blank

# 💡 SOLUTION : On force Reflex à utiliser NPM au lieu de Bun pour éviter le bug d'extraction sous Docker
RUN echo 'import reflex as rx\n\nconfig = rx.Config(\n    app_name="suivi_app",\n    use_bun=False,\n)' > rxconfig.py

# On remplace le fichier de base par le nôtre
COPY suivi_app.py /app/suivi_app/suivi_app.py

CMD ["reflex", "run", "--env", "prod"]