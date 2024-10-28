import pdfkit

path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

file = 'TextoAClabDidatico.html'

# url = 'https://www.google.com/'

config = pdfkit.configuration(wkhtmltopdf = path)

pdfkit.from_file(file, output_path='TextoAClabDidatico.pdf', configuration=config)

# pdfkit.from_url(url, output_path='pagina.pdf', configuration=config)