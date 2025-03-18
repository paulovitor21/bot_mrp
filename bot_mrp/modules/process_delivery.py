import pandas as pd
import time
import os
import xlwings as xw
from botcity.core import DesktopBot

class Bot(DesktopBot):
    def action(self, execution=None):
        pass

    def open_excel(self, excel_file_path):
        os.startfile(excel_file_path)
        time.sleep(10)  # Aguardar o Excel carregar completamente

        try:
            wb = xw.Book(excel_file_path)  # Abrir ou pegar o arquivo ativo
            return wb
        except Exception as e:
            print(f"Erro ao abrir o arquivo no Excel: {e}")
            return None

    def process_delivery(self, file_delivery, master_all_path):
        print(f"Processing delivery file: {file_delivery}")

        # Carregar os dados do delivery
        delivery_df = pd.read_excel(file_delivery)

        # Abrir o Excel e carregar o workbook existente
        wb = self.open_excel(master_all_path)
        if not wb:
            return

        ws = wb.sheets["Delivery"]  # Selecionar a aba desejada
        ws.activate()
        print(f"Aba ativa: {ws.name}")

        # Selecionar e limpar os dados antigos, mantendo o cabeçalho
        ws.range("C2").expand().clear_contents()

        # Inserir os novos dados
        ws.range("C2").value = delivery_df.values.tolist()

        # Salvar e fechar o arquivo
        try:
            wb.save()
            wb.close()
            print("Delivery data has been updated in master_all.xlsx")
        except Exception as e:
            print(f"Erro ao salvar/fechar o arquivo: {e}")

if __name__ == '__main__':
    bot = Bot()
    file_delivery = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Delivery Status 03062025 0935.xlsx"
    bot.process_delivery(file_delivery)