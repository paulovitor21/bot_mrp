import time
import os
import win32com.client
import xlwings as xw
from botcity.core import DesktopBot

class Bot(DesktopBot):
    def action(self, execution=None):
        pass

    def open_excel(self, excel_file_path):
        os.startfile(excel_file_path)
        time.sleep(15)  # Aguardar o Excel carregar completamente

        try:
            wb = xw.Book(excel_file_path)  # Abrir ou pegar o arquivo ativo
            return wb
        except Exception as e:
            print(f"Erro ao abrir o arquivo no Excel: {e}")
            return None

    def process_supplier(self, file_componel, master_all_path):
        print(f"Processing plan assy: {file_componel}")

        wb = self.open_excel(master_all_path)
        if not wb:
            return

        ws = wb.sheets["Supplier"]  # Selecionar a aba desejada
        ws.activate()
        print(f"Aba ativa: {ws.name}")

        # Salvar e fechar o arquivo
        try:
            wb.save()
            wb.close()
            print("Supplier data has been updated successfully.")
        except Exception as e:
            print(f"Erro ao salvar/fechar o arquivo: {e}")

if __name__ == '__main__':
    bot = Bot()
    file_bom = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\BOM Status 03062025 0935.xlsx"
    bot.process_bom(file_bom)