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

    def process_bom(self, file_bom, master_all_path):
        print(f"Processing bom file: {file_bom}")

        wb = self.open_excel(master_all_path)
        if not wb:
            return

        ws = wb.sheets["BOM"]  # Selecionar a aba desejada
        ws.activate()
        print(f"Aba ativa: {ws.name}")

        # Chamar a função BOM_Assy_Master_Click da macro usando win32com.client
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True  # Exibe a interface do Excel
            # excel.DisplayAlerts = False  # Desabilita as mensagens de alerta do Excel
            excel.Application.Run(f"'{wb.name}'!Module1.BOM_Assy_Master_Click", file_bom)
            print("Macro BOM_Assy_Master_Click executada com sucesso.")
        except Exception as e:
            print(f"Erro ao executar a macro: {e}")

        # Salvar e fechar o arquivo
        try:
            wb.save()
            wb.close()
            print("BOM data has been updated successfully.")
        except Exception as e:
            print(f"Erro ao salvar/fechar o arquivo: {e}")

if __name__ == '__main__':
    bot = Bot()
    file_bom = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\BOM Status 03062025 0935.xlsx"
    bot.process_bom(file_bom)