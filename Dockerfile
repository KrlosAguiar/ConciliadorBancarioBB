FROM python:3.9-slim

# Define a pasta de trabalho
WORKDIR /app

# --- AQUI ESTÁ A MUDANÇA NECESSÁRIA ---
# Adicionei 'unrar-free' na lista de instalações.
# Sem isso, o Python vai dar erro ao tentar abrir arquivos .rar
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    unrar-free \
    && rm -rf /var/lib/apt/lists/*

# Cria um "apelido" para o sistema entender que 'unrar' é o mesmo que 'unrar-free'
RUN ln -s /usr/bin/unrar-free /usr/bin/unrar

# Copia os arquivos do GitHub para dentro do servidor
COPY . .

# Instala as bibliotecas do seu requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Expõe a porta que o Streamlit usa
EXPOSE 8080

# --- MANTIVE SEU COMANDO ORIGINAL ---
# O erro acontecia aqui porque eu tinha mudado o nome do arquivo.
# Mantendo "Página Inicial.py", o Cloud Run vai achar seu código.
ENTRYPOINT ["streamlit", "run", "Página Inicial.py", "--server.port=8080", "--server.address=0.0.0.0"]
