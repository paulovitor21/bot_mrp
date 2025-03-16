import pandas as pd
import time
import os
import xlwings as xw
from botcity.core import DesktopBot

class Bot(DesktopBot):
    def action(self, execution=None):
        pass

    def process_delivery(self, file_delivery):
        print(f"Processing delivery file: {file_delivery}")
        # Carregar os dados do delivery
        delivery_df = pd.read_excel(file_delivery)

        # Abrir o Excel e carregar o workbook existente
        excel_file_path = r'C:\Users\Paulo\Desktop\bot_mrp\bot_mrp\05.03_Master_All_Sourcing_.xlsb'
        os.startfile(excel_file_path)
        time.sleep(10)  # Aguarde o Excel abrir

        # Usar xlwings para navegar para a aba 'delivery'
        wb = xw.Book(excel_file_path)  # Abre o arquivo ou pega o ativo
        ws = wb.sheets["Delivery"]  # Nome da aba desejada
        ws.activate()
        print(f"Aba ativa: {ws.name}")

        # Selecionar e limpar os dados antigos, mantendo o cabeçalho
        ws.range("C2").expand().clear_contents()

        # Inserir os novos dados
        ws.range("C2").value = delivery_df.values.tolist()

        # Salvar o workbook
        wb.save()
        wb.close()

        print("Delivery data has been updated in master_all.xlsx")

if __name__ == '__main__':
    bot = Bot()
    file_delivery = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Delivery Status 03062025 0935.xlsx"
    bot.process_delivery(file_delivery)