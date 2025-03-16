import pandas as pd
import time
import os
import win32com.client
import xlwings as xw
from botcity.core import DesktopBot

class Bot(DesktopBot):
    def action(self, execution=None):
        pass

    def process_bom(self, file_bom):
        print(f"Processing bom file: ")

        # Abrir o Excel e carregar o workbook existente
        excel_file_path = r'C:\Users\Paulo\Desktop\bot_mrp\bot_mrp\05.03_Master_All_Sourcing_.xlsb'
        os.startfile(excel_file_path)
        time.sleep(15)  # Aguardar o Excel carregar completamente

        # Tentar conectar ao arquivo aberto no Excel
        try:
            wb = xw.Book(excel_file_path)  # Abrir ou pegar o arquivo ativo
        except Exception as e:
            print(f"Erro ao abrir o arquivo no Excel: {e}")
            return

        ws = wb.sheets["BOM"]  # Selecionar a aba desejada
        ws.activate()
        print(f"Aba ativa: {ws.name}")


        
        # Chamar a função Onhand_Chave_Click da macro usando win32com.client
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True  # Exibe a interface do Excel
            workbook = excel.Workbooks.Open(excel_file_path)
            formulario_nome = "Module1"
            excel.Application.Run(f"{formulario_nome}.BOM_Assy_Master_Click", file_bom)
            print("Macro BOM_Assy_Master_Click executada com sucesso.")
            
        except Exception as e:
            print(f"Erro ao executar a macro: {e}")

        # Salvar e fechar o arquivo
        try:
            wb.save()
            wb.close()
            print("bom updated successfully.")
        except Exception as e:
            print(f"Erro ao salvar/fechar o arquivo: {e}")

if __name__ == '__main__':
    bot = Bot()
    file_onhand = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Onhand Status 03062025 0935.xlsx"
    bot.process_onhand(file_onhand)