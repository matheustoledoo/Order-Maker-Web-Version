# Usar uma imagem Python oficial
FROM python:3.9-slim

# Diretório de trabalho
WORKDIR /app

# Copiar arquivos para o container
COPY . /app

# Instalar dependências do sistema e do Python
RUN pip install --no-cache-dir -r requirements.txt

# Expor a porta para o Railway
EXPOSE 5000

# Comando para iniciar a aplicação
CMD ["python", "app.py"]
