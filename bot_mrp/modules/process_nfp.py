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

    def process_nfp(self, file_nfp, master_all_path):
        print(f"Processing NFP file: {file_nfp}")

        # Abrir o Excel e carregar o workbook existente
        wb = self.open_excel(master_all_path)
        if not wb:
            return

        ws = wb.sheets["PlanPPH"]  # Selecionar a aba desejada
        ws.activate()
        print(f"Aba ativa: {ws.name}")

        # Selecionar e limpar os dados antigos, mantendo o cabeçalho
        ws.range("A1").expand().clear_contents()

        # Carregar os dados do NFP a partir da aba especificada
        nfp_df = pd.read_excel(file_nfp, sheet_name='Summy Daily_ByLine', header=0)

        # Selecionar as colunas B, C e K em diante
        colunas_selecionadas = [1, 2] + list(range(10, nfp_df.shape[1]))
        dados_para_copiar = nfp_df.iloc[:, colunas_selecionadas]

        # Copiar os dados para a área de transferência
        dados_para_copiar.to_clipboard(index=False, header=True)

        # Colar os dados no Excel
        ws.range("A1").api.PasteSpecial(Paste=-4104)  # xlPasteAll

        # Salvar e fechar o arquivo
        try:
            wb.save()
            wb.close()
            print("NFP data has been updated in master_all.xlsx")
        except Exception as e:
            print(f"Erro ao salvar/fechar o arquivo: {e}")

if __name__ == '__main__':
    bot = Bot()
    file_nfp = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\NFP 03062025 0935.xlsx"
    master_all_path = r"C:\Users\Paulo\Desktop\bot_mrp\bot_mrp\05.03_Master_All_Sourcing_.xlsb"
    bot.process_nfp(file_nfp, master_all_path)