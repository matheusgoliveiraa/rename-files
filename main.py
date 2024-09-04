import os
import pandas as pd
import fuzzywuzzy

from tkinter import Tk, filedialog
from fuzzywuzzy import process
 
def select_folder():
    root = Tk()
    root.withdraw()  
    folder_selected = filedialog.askdirectory()
    return folder_selected
 
def rename_pdfs(folder_path, excel_path):
   
    if not os.path.exists(excel_path):
        print(f"Erro: O arquivo Excel não foi encontrado no caminho especificado: {excel_path}")
        return
    
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
    except Exception as e:
        print(f"Erro ao carregar o arquivo Excel: {e}")
        return
 
    names_column = df.iloc[:, 2].astype(str).tolist()
   
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            try:
              
                if len(filename) > 61:
                    
                    remaining_name = filename[61:].replace('_', ' ').strip()
                   
                   
                    name_without_extension = remaining_name.replace('.pdf', '')
                   
                    
                    closest_match, score = process.extractOne(name_without_extension, names_column)
                   
                    if score >= 80:  
                        new_name_prefix = df[df.iloc[:, 2] == closest_match].iloc[0, 0]
                        new_name = f"{new_name_prefix} {name_without_extension}.pdf"
                       
                        
                        old_file = os.path.join(folder_path, filename)
                        new_file = os.path.join(folder_path, new_name)
                       
                        old_file = os.path.normpath(old_file)
                        new_file = os.path.normpath(new_file)
 
                        # Log dos caminhos para diagnóstico
                        print(f"Tentando renomear:")
                        print(f"Antigo: '{old_file}'")
                        print(f"Novo: '{new_file}'")
 
                        # Verifica se o arquivo antigo realmente existe
                        if not os.path.exists(old_file):
                            print(f"Erro: O arquivo '{old_file}' não foi encontrado para renomear.")
                            continue  # Pula para o próximo arquivo
 
                        # Renomeia o arquivo
                        try:
                            os.rename(old_file, new_file)
                            print(f"Arquivo '{filename}' renomeado para '{new_name}'")
                        except FileNotFoundError:
                            print(f"Erro: O arquivo '{old_file}' não foi encontrado para renomear.")
                        except PermissionError:
                            print(f"Erro: Permissão negada ao tentar renomear o arquivo '{old_file}'.")
                        except OSError as e:
                            print(f"Erro: Problema com o caminho do arquivo '{old_file}' ou '{new_file}': {e}")
                        except Exception as e:
                            print(f"Erro inesperado ao tentar renomear o arquivo '{old_file}': {e}")
                    else:
                        print(f"Nome restante '{remaining_name}' não encontrou uma correspondência suficiente na coluna C do Excel.")
                else:
                    print(f"Nome do arquivo '{filename}' é muito curto para renomear.")
           
            except Exception as e:
                print(f"Erro ao processar o arquivo '{filename}': {e}")
 
def main():
    """Função principal que executa o processo de seleção de pasta e renomeação de arquivos."""
    print("Selecione a pasta onde os arquivos PDF estão localizados:")
    folder = select_folder()
   
    excel_path = r'seu_arquivo.xlsx'
   
    if folder:
        rename_pdfs(folder, excel_path)
    else:
        print("Nenhuma pasta selecionada.")
 
if __name__ == "__main__":
    main()
 