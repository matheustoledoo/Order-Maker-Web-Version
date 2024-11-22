import os
import uuid
import time
import tempfile
from flask import Flask, render_template, request, send_file
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import platform  # Para detectar o sistema operacional

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.path.abspath("files")  # Caminho absoluto
app.config["ALLOWED_EXTENSIONS"] = {"pptx"}

# Funções auxiliares
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in app.config["ALLOWED_EXTENSIONS"]

def aplicar_formatacao(paragraph, fonte="Codec Pro", tamanho=24, cor=(0, 0, 0)):
    for run in paragraph.runs:
        run.font.name = fonte
        run.font.size = Pt(tamanho)
        run.font.color.rgb = RGBColor(*cor)

def substituir_valores_marcadores(slide, marcador, valor):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if marcador in paragraph.text:
                    paragraph.text = paragraph.text.replace(marcador, valor)
                    aplicar_formatacao(paragraph)

def adicionar_lista_incremental(slide, marcador, lista):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if marcador in paragraph.text:
                    texto_atual = paragraph.text.strip()
                    if texto_atual != marcador:
                        paragraph.text = texto_atual
                    else:
                        paragraph.text = marcador
                    aplicar_formatacao(paragraph)

                    for item in lista:
                        novo_paragraph = shape.text_frame.add_paragraph()
                        novo_paragraph.text = item
                        aplicar_formatacao(novo_paragraph)

def adicionar_objetos_dinamicos(slide, lista_objetos):
    left = Inches(6)
    top = Inches(3.2)
    width = Inches(1.5)
    height = Inches(0.5)
    espacamento_vertical = Inches(0.9)
    limite_caracteres = 32

    for obj in lista_objetos:
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = False
        text_frame.auto_size = False

        linhas = [obj[i:i+limite_caracteres] for i in range(0, len(obj), limite_caracteres)]
        for linha in linhas:
            paragraph = text_frame.add_paragraph()
            paragraph.text = linha
            aplicar_formatacao(paragraph)
        top += espacamento_vertical

def convert_to_pdf(pptx_path):
    """
    Converte o arquivo PPTX para PDF usando o PowerPoint via COM no Windows
    ou LibreOffice no Linux (Railway).
    """
    if platform.system() == "Windows":
        # Conversão no Windows com PowerPoint
        import pythoncom
        import comtypes.client
        pythoncom.CoInitialize()
        ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
        ppt_app.Visible = 1
        presentation = None

        try:
            if not os.path.exists(pptx_path):
                raise FileNotFoundError(f"O arquivo {pptx_path} não foi encontrado!")

            pdf_path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
            presentation = ppt_app.Presentations.Open(pptx_path, WithWindow=False)
            presentation.SaveAs(pdf_path, 32)
            return pdf_path
        finally:
            if presentation:
                presentation.Close()
            ppt_app.Quit()
    else:
        # Conversão no Railway com LibreOffice
        pdf_path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
        command = f"libreoffice --headless --convert-to pdf --outdir {os.path.dirname(pdf_path)} {pptx_path}"
        if os.system(command) != 0:
            raise Exception("Erro ao converter para PDF com LibreOffice.")
        return pdf_path

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            arquivo = request.form["arquivo"]
            caminho_arquivo = os.path.join(app.config["UPLOAD_FOLDER"], arquivo)
            prs = Presentation(caminho_arquivo)

            # Dados do formulário
            nome_cliente = request.form.get("nome_cliente", "")
            valor_servico = request.form.get("valor_servico", "")
            valor_mobilizacao = request.form.get("valor_mobilizacao", "")
            objetos = request.form.get("objetos", "").splitlines()
            action = request.form.get("action")

            # Substituir valores nos slides
            substituir_valores_marcadores(prs.slides[1], "{", nome_cliente)

            # Criar um arquivo PPTX temporário
            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as temp_pptx:
                prs.save(temp_pptx.name)
                output_path = temp_pptx.name

            # Converter para PDF se necessário
            if action == "pdf":
                output_path = convert_to_pdf(output_path)

            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"Erro no processamento: {e}"

    arquivos = [f for f in os.listdir(app.config["UPLOAD_FOLDER"]) if allowed_file(f)]
    return render_template("index.html", arquivos=arquivos)

if __name__ == "__main__":
    if not os.path.exists(app.config["UPLOAD_FOLDER"]):
        os.makedirs(app.config["UPLOAD_FOLDER"])
    app.run(debug=True)
