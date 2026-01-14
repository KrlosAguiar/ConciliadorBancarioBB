FROM python:3.9-slim

# Define a pasta de trabalho
WORKDIR /app

# Instala o essencial E O UNRAR (Necessário para arquivos .rar)
# Adicionei 'unrar-free' na lista
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    unrar-free \
    && rm -rf /var/lib/apt/lists/*

# --- TRUQUE IMPORTANTE ---
# O pacote linux instala como "unrar-free", mas a biblioteca Python busca por "unrar".
# Criamos um link simbólico para "enganar" o sistema e fazer funcionar direto.
RUN ln -s /usr/bin/unrar-free /usr/bin/unrar

# Copia os arquivos do GitHub para dentro do servidor
COPY . .

# Instala as bibliotecas do seu requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Expõe a porta que o Streamlit usa
EXPOSE 8080

# Comando para rodar o aplicativo (Mantido o seu original)
ENTRYPOINT ["streamlit", "run", "Página Inicial.py", "--server.port=8080", "--server.address=0.0.0.0"]
