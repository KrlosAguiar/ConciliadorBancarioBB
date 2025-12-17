FROM python:3.9-slim

# Define a pasta de trabalho
WORKDIR /app

# Instala apenas o essencial para o Python e limpeza de cache
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copia os arquivos do GitHub para dentro do servidor
COPY . .

# Instala as bibliotecas do seu requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Expõe a porta que o Streamlit usa
EXPOSE 8080

# Comando para rodar o aplicativo
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0"]
