# Usa uma imagem base com Python 3.x instalada
FROM python:3.9-slim

# Atualiza o sistema e instala o LibreOffice
RUN apt-get update && apt-get install -y libreoffice

# Cria um diretório de trabalho dentro do container
WORKDIR /app

# Copia os arquivos do projeto para o container
COPY . .

# Instala as dependências do projeto
RUN pip install --no-cache-dir -r requirements.txt

# Define a variável de ambiente PORT (opcional, depende do serviço de deploy)
ENV PORT=5000

# Expõe a porta da aplicação Flask
EXPOSE 5000

# Comando para iniciar o servidor
CMD ["python", "app.py"]
