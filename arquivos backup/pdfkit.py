import docx
import re
import pdfkit

def extrair_marcacoes(texto):
    # Usando expressão regular para capturar as marcações entre {{ }}
    padrao = r'\{\{(.*?)\}\}'
    marcacoes = re.findall(padrao, texto)
    return marcacoes

def processar_documento(doc_path):
    # Abrindo o documento docx
    doc = docx.Document(doc_path)
    conteudo_extraido = ""
    
    for paragrafo in doc.paragraphs:
        # Adiciona o parágrafo ao conteúdo extraído
        conteudo_extraido += paragrafo.text + "\n"
    
    # Extraindo marcações
    marcacoes = extrair_marcacoes(conteudo_extraido)
    
    # Retorna o conteúdo completo e as marcações encontradas
    return conteudo_extraido, marcacoes

def salvar_como_pdf(texto, pdf_path):
    # Salvando o conteúdo como HTML temporário
    html_content = f"<html><body><pre>{texto}</pre></body></html>"
    
    # Converte o HTML em PDF usando pdfkit
    pdfkit.from_string(html_content, pdf_path)

def main(docx_file, pdf_file):
    # Processa o documento .docx para extrair o texto e marcações
    conteudo, marcacoes = processar_documento(docx_file)
    
    print("Conteúdo do documento:")
    print(conteudo)
    
    print("\nMarcações encontradas:")
    for m in marcacoes:
        print(f" - {m}")
    
    # Salva o conteúdo do documento em PDF
    salvar_como_pdf(conteudo, pdf_file)
    print(f"\nDocumento salvo como PDF: {pdf_file}")

if __name__ == "__main__":
    docx_file = "exemplo.docx"  # Substitua pelo seu arquivo .docx
    pdf_file = "saida.pdf"      # Nome do arquivo de saída PDF
    
    main(docx_file, pdf_file)
