import os
import mimetypes
from pdf2docx import Converter
import img2pdf
from docx2pdf import convert as docx2pdf_convert
import openpyxl
import pdfkit
import fitz  
import tabula



def png_para_pdf(png_path, pdf_path):

    with open(pdf_path, "wb") as f:
        f.write(img2pdf.convert(png_path))
        


def pdf_para_png(pdf_path, png_path):

    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        output_path = f"{png_path.rstrip('.png')}_{page_num + 1}.png"
        pix.save(output_path)



def docx_para_pdf(docx_path, pdf_path):
    docx2pdf_convert(docx_path, pdf_path)



def pdf_para_docx(pdf_path, docx_path):

    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()



def xlsx_para_pdf(xlsx_path, pdf_path):
    # Carregar a planilha
    wb = openpyxl.load_workbook(xlsx_path)
    sheet = wb.active

    # Salvar como HTML temporário
    html_path = "temp.html"
    with open(html_path, "w") as f:
        f.write('<html><body><table border="1">')
        for row in sheet.iter_rows(values_only=True):
            f.write('<tr>')
            for cell in row:
                f.write(f'<td>{cell}</td>')
            f.write('</tr>')
        f.write('</table></body></html>')


    # Converter HTML para PDF
    pdfkit.from_file(html_path, pdf_path)
    os.remove(html_path)



def pdf_para_xlsx(pdf_path, xlsx_path):
    tabula.convert_into(pdf_path, xlsx_path, output_format="xlsx")



def converter_arquivo(input_path, output_path):
    mime_type, _ = mimetypes.guess_type(input_path)
    
    if mime_type == 'application/pdf':

        if output_path.endswith('.docx'):
            pdf_para_docx(input_path, output_path)

        elif output_path.endswith('.png'):
            pdf_para_png(input_path, output_path)

        elif output_path.endswith('.xlsx'):
            pdf_para_xlsx(input_path, output_path)

        else:
            print("Formato de saída não suportado para PDF.")

    elif mime_type == 'image/png':
        if output_path.endswith('.pdf'):
            png_para_pdf(input_path, output_path)

        else:
            print("Formato de saída não suportado para PNG.")
    elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':

        if output_path.endswith('.pdf'):
            docx_para_pdf(input_path, output_path)

        else:

            print("Formato de saída não suportado para DOCX.")
    elif mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':

        if output_path.endswith('.pdf'):
            xlsx_para_pdf(input_path, output_path)
        else:
            print("Formato de saída não suportado para XLSX.")

    else:
        print("Tipo de arquivo não suportado.")

# Exemplos de uso
converter_arquivo("Teste.docx", "teste5.pdf")
