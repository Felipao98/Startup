import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from docx2pdf import convert

def convert_doc_to_pdf(input_path, output_path=None):
    # Verifica se o arquivo é .doc ou .docx
    if not (input_path.endswith(".docx") or input_path.endswith(".doc")):
        raise ValueError("O arquivo deve ser .docx ou .doc")

    # Se o caminho de saída não for fornecido, salva na mesma pasta com extensão .pdf
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"

    try:
        # Converte o arquivo .docx para PDF
        convert(input_path, output_path)
        print(f"Arquivo convertido com sucesso: {output_path}")
    except Exception as e:
        print(f"Ocorreu um erro na conversão: {e}")

def select_and_convert():
    # Cria uma janela para selecionar o arquivo
    Tk().withdraw()  # Oculta a janela principal do Tkinter
    input_file = askopenfilename(title="Selecione o arquivo DOCX ou DOC", 
                                 filetypes=[("Documentos do Word", "*.docx;*.doc")])
    
    if input_file:
        # Converte o arquivo selecionado para PDF
        convert_doc_to_pdf(input_file)
    else:
        print("Nenhum arquivo foi selecionado.")

select_and_convert()
