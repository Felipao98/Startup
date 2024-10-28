# codigo para a conversão do arquivo sem a necessidade de salvar
import docx
import pdfkit
from io import BytesIO

def extrair_conteudo_doc(input_path):
    # Abrindo o arquivo .docx e extraindo o conteúdo
    doc = docx.Document(input_path)
    conteudo = ""

    for paragrafo in doc.paragraphs:
        conteudo += paragrafo.text + "\n"
    
    return conteudo

def converter_para_pdf_in_memory(texto):
    # Criando um HTML básico para converter em PDF
    html_content = f"<html><body><pre>{texto}</pre></body></html>"

    # Configurando para gerar o PDF em memória
    pdf_output = BytesIO()
    
    # Se o wkhtmltopdf não estiver no PATH, você pode configurar o caminho
    caminho_wkhtmltopdf = r'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
    pdfkit_config = pdfkit.configuration(wkhtmltopdf=caminho_wkhtmltopdf)

    try:
        # Usando pdfkit para gerar o PDF a partir do HTML e armazenando no buffer de memória
        pdfkit.from_string(html_content, pdf_output, configuration=pdfkit_config)
        
        # Retornando o conteúdo do PDF em formato binário
        pdf_output.seek(0)
        return pdf_output.read()

    except Exception as e:
        print(f"Erro ao gerar o PDF: {e}")
        return None

def converter_doc_para_pdf_em_memoria(doc_path):
    # Extrai o conteúdo do arquivo DOCX
    conteudo = extrair_conteudo_doc(doc_path)
    
    # Converte o conteúdo para PDF em memória
    pdf_data = converter_para_pdf_in_memory(conteudo)
    
    if pdf_data:
        print("PDF gerado em memória com sucesso.")
        return pdf_data  # Retorna o PDF em binário para ser usado conforme necessário
    else:
        print("Falha ao gerar o PDF.")
        return None

# Exemplo de uso
docx_file = "Texto AClabDidatico.docx"  # Substitua pelo seu caminho de arquivo
pdf_bytes = converter_doc_para_pdf_em_memoria(docx_file)

# Se quiser salvar o PDF em disco para testar (apenas opcional para ver o resultado):
with open("saida_em_memoria.pdf", "wb") as f:
    f.write(pdf_bytes)
