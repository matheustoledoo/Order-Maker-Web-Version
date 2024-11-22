FROM python:3.9-slim

# Atualiza o sistema e instala dependências necessárias
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-dejavu \
    && apt-get clean

# Configura o diretório de trabalho
WORKDIR /app

# Copia os arquivos para o contêiner
COPY . .

# Instala as dependências do Python
RUN pip install --no-cache-dir -r requirements.txt

# Expõe a porta para o servidor Flask
EXPOSE 8080

# Comando para rodar a aplicação
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "app:app"]
