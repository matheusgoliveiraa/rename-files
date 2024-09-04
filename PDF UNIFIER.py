import os
import PyPDF2
from tkinter import Tk, filedialog

def select_folder():
    """Abre uma caixa de diálogo para selecionar uma pasta."""
    root = Tk()
    root.withdraw()  # Oculta a janela principal
    folder_selected = filedialog.askdirectory()
    return folder_selected

def merge_pdfs(pdf1_path, pdf2_path, output_path):
    """Une dois arquivos PDF em um único arquivo PDF."""
    try:
        pdf_writer = PyPDF2.PdfWriter()

        # Adiciona páginas de ambos os PDFs ao PdfWriter
        for pdf_path in [pdf1_path, pdf2_path]:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    pdf_writer.add_page(page)

        # Salva o PDF unificado
        with open(output_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)
        print(f"Arquivos '{pdf1_path}' e '{pdf2_path}' foram unificados em '{output_path}'")
    except Exception as e:
        print(f"Erro ao unificar PDFs: {e}")

def normalize_filename(filename):
    """Remove espaços e zeros à esquerda dos primeiros 15 caracteres do nome do arquivo."""
    return filename[:15].lstrip('0').replace(" ", "")

def main():
    """Função principal que executa o processo de seleção de pastas e unificação de PDFs."""
    try:
        print("Selecione a primeira pasta:")
        folder1 = select_folder()
        print("Selecione a segunda pasta:")
        folder2 = select_folder()
        print("Selecione a pasta onde deseja criar a nova pasta 'Contribuição Associativa Unificada':")
        base_output_folder = select_folder()
        output_folder = os.path.join(base_output_folder, "Contribuição Associativa Unificada")
        os.makedirs(output_folder, exist_ok=True)

        # Mapeia arquivos PDF nas duas pastas com base nos primeiros 15 caracteres normalizados do nome
        pdfs1 = {normalize_filename(f): os.path.join(folder1, f) for f in os.listdir(folder1) if f.endswith('.pdf')}
        pdfs2 = {normalize_filename(f): os.path.join(folder2, f) for f in os.listdir(folder2) if f.endswith('.pdf')}

        # Unifica PDFs que estão presentes em ambas as pastas
        for key in pdfs1:
            if key in pdfs2:
                output_filename = os.path.basename(pdfs1[key])
                output_path = os.path.join(output_folder, output_filename)
                merge_pdfs(pdfs1[key], pdfs2[key], output_path)
            else:
                print(f"PDF correspondente para '{key}' não encontrado na segunda pasta.")
    except Exception as e:
        print(f"Erro no processo: {e}")

if __name__ == "__main__":
    main()
