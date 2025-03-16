# import win32com.client

# # Abrir o Excel
# excel = win32com.client.Dispatch("Excel.Application")
# excel.Visible = True  # Exibe a interface do Excel

# # Abrir a planilha
# workbook = excel.Workbooks.Open(r"C:\Users\Paulo\Desktop\bot_mrp\bot_mrp\05.03_Master_All_Sourcing_.xlsb")

# # Nome do UserForm onde está a função
# formulario_nome = "Module1"

# # Carregar o formulário manualmente e chamar a função
# # userform = excel.Application.VBE.ActiveVBProject.VBComponents(formulario_nome)
# excel.Application.Run(f"{formulario_nome}.OnhandChave_Click")

# # Opcional: Fechar e salvar
# workbook.Close(SaveChanges=True)
# excel.Quit()

import win32com.client

# Caminho do arquivo Excel com a macro
caminho_arquivo = r"C:\Users\Paulo\Desktop\bot_mrp\bot_mrp\05.03_Master_All_Sourcing_.xlsb"
# Caminho do arquivo BOM que será passado para a macro
caminho_bom = r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\0305_Bom_Master.xlsb"

# Inicia o Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  # Defina como True se quiser ver a execução

# Abre o arquivo Excel com a macro
workbook = excel.Workbooks.Open(caminho_arquivo)

# Chama a macro e passa o caminho do arquivo BOM como argumento
excel.Application.Run("BOM_Assy_Master_Click", caminho_bom)

# Salva e fecha
workbook.Close(SaveChanges=True)
excel.Quit()
