# TESTE COM A FUNÇÃO QUE FOI SOLICITADA EM 23/10/24
import re
import docx
import pdfkit
from io import BytesIO

# Função para extrair conteúdo e encontrar palavras-chave
def extrair_conteudo_e_palavras_chave(docx_path):
    doc = docx.Document(docx_path)
    conteudo = ""
    palavras_chave = []

    # Varre todos os parágrafos do documento
    for paragrafo in doc.paragraphs:
        texto = paragrafo.text
        conteudo += texto + "\n"
        
        # Encontra palavras dentro de {{chave}} usando regex
        chaves_encontradas = re.findall(r'\{\{(.*?)\}\}', texto)
        palavras_chave.extend(chaves_encontradas)
    
    return conteudo, palavras_chave

# Função para converter o conteúdo para PDF, destacando as palavras-chave
def converter_para_pdf_in_memory(conteudo, palavras_chave):
    # Formata o conteúdo, destacando as palavras-chave
    for palavra in palavras_chave:
        conteudo = conteudo.replace(f"{{{{{palavra}}}}}", f"<strong>{palavra}</strong>")
    
    # Cria um HTML básico para converter em PDF
    html_content = f"<html><body><pre>{conteudo}</pre></body></html>"

    # Configura para gerar o PDF em memória
    pdf_output = BytesIO()
    
    # Se o wkhtmltopdf não estiver no PATH, você pode configurar o caminho
    caminho_wkhtmltopdf = r'"C:\\Arquivos de Programas\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"'
    pdfkit_config = pdfkit.configuration(wkhtmltopdf=caminho_wkhtmltopdf)

    try:
        # Usa pdfkit para gerar o PDF a partir do HTML e armazena no buffer de memória
        pdfkit.from_string(html_content, pdf_output, configuration=pdfkit_config)
        
        # Retorna o conteúdo do PDF em formato binário
        pdf_output.seek(0)
        return pdf_output.read()

    except Exception as e:
        print(f"Erro ao gerar o PDF: {e}")
        return None

def processar_e_converter(docx_path):
    # Extrai o conteúdo do arquivo DOCX e as palavras-chave
    conteudo, palavras_chave = extrair_conteudo_e_palavras_chave(docx_path)
    
    # Converte o conteúdo para PDF, destacando as palavras-chave
    pdf_data = converter_para_pdf_in_memory(conteudo, palavras_chave)
    
    if pdf_data:
        print("PDF gerado em memória com sucesso.")
        return pdf_data  # Retorna o PDF em binário para ser salvo ou manipulado
    else:
        print("Falha ao gerar o PDF.")
        return None

# Exemplo de uso
docx_file = "exemplo.docx"  # Substitua pelo caminho do seu arquivo
pdf_bytes = processar_e_converter(docx_file)

# Opcionalmente, você pode salvar o PDF gerado em memória em um arquivo para verificar o resultado
with open("saida_palavras_chave.pdf", "wb") as f:
    f.write(pdf_bytes)
