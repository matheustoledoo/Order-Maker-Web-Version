# Usa uma imagem base com Python 3.9 slim
FROM python:3.9-slim

# Instala dependências básicas e o LibreOffice
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Cria um diretório de trabalho
WORKDIR /app

# Copia os arquivos do projeto para dentro do container
COPY . .

# Instala as dependências do Python
RUN pip install --no-cache-dir -r requirements.txt

# Define a variável de ambiente PORT para o Railway
ENV PORT=5000

# Expõe a porta do Flask
EXPOSE 5000

# Comando para iniciar o aplicativo
CMD ["python", "app.py"]
