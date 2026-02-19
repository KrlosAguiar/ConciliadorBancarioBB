FROM python:3.9-slim

# Define a pasta de trabalho
WORKDIR /app

# Instala apenas o essencial e o unrar-free
# O próprio pacote já cria o atalho 'unrar', não precisamos fazer manualmente

RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    unrar-free \
    pkg-config \
    libcairo2-dev \
    && rm -rf /var/lib/apt/lists/*

# Copia os arquivos do GitHub para dentro do servidor
COPY . .

# Instala as bibliotecas do seu requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Expõe a porta que o Streamlit usa
EXPOSE 8080

# Comando para rodar o aplicativo
ENTRYPOINT ["streamlit", "run", "Página Inicial.py", "--server.port=8080", "--server.address=0.0.0.0"]
