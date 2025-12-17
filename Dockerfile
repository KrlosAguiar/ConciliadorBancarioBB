# Usa uma imagem leve do Python
FROM python:3.9-slim

# Define pasta de trabalho
WORKDIR /app

# Instala dependências do sistema operacional (necessário para PDF/Excel)
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    software-properties-common \
    && rm -rf /var/lib/apt/lists/*

# Copia os arquivos da sua pasta para o servidor
COPY . .

# Instala as bibliotecas Python
RUN pip3 install -r requirements.txt

# Expõe a porta 8080 (Padrão do Google Cloud Run)
EXPOSE 8080

# Comando para iniciar o site na porta correta
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0"]
