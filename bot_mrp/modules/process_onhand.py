import pandas as pd
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

    def process_onhand(self, file_onhand, master_all_path):
        print(f"Processing onhand file: {file_onhand}")

        # Carregar os dados do onhand
        onhand_df = pd.read_excel(file_onhand)

        # Selecionar apenas as colunas específicas usando índices
        selected_columns = onhand_df.iloc[:, [0, 1, 2, 4, 6, 7, 8, 15, 16]]

        # Substituir valores nulos por string vazia para evitar erros ao inserir no Excel
        selected_columns = selected_columns.fillna("")

        # Converter os dados para lista de listas
        data_to_insert = selected_columns.values.tolist()

        # Abrir o Excel e carregar o workbook existente
        wb = self.open_excel(master_all_path)
        if not wb:
            return

        ws = wb.sheets["Onhand"]  # Selecionar a aba desejada
        ws.activate()
        print(f"Aba ativa: {ws.name}")

        # Limpar os dados antigos, mantendo o cabeçalho
        ws.range("B2:J2").expand().clear_contents()

        # Inserir os dados no Excel de uma vez só
        try:
            ws.range("B2").value = data_to_insert
        except Exception as e:
            print(f"Erro ao inserir dados: {e}")

        # Chamar a função Onhand_Chave_Click da macro usando win32com.client
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True  # Exibe a interface do Excel
            excel.DisplayAlerts = False  # Desabilita as mensagens de alerta do Excel
            excel.Application.Run(f"'{wb.name}'!Module1.Onhand_Chave_Click")
            print("Macro Onhand_Chave_Click executada com sucesso.")
        except Exception as e:
            print(f"Erro ao executar a macro: {e}")

        # Salvar e fechar o arquivo
        try:
            wb.save()
            wb.close()
            print("Onhand data has been updated successfully.")
        except Exception as e:
            print(f"Erro ao salvar/fechar o arquivo: {e}")

if __name__ == '__main__':
    bot = Bot()
    file_onhand = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Onhand Status 03062025 0935.xlsx"
    bot.process_onhand(file_onhand)