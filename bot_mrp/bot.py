# Import for integration with BotCity Maestro SDK
from botcity.maestro import *
from modules.process_delivery import Bot  # Import the Bot class
from modules.process_onhand import Bot as OnhandBot  # Import the Bot class
from modules.process_bom import Bot as BomBot  # Import the Bot class
from modules.process_up_plan_assy import Bot as UpPlanAssyBot  # Import the Bot class
from modules.process_supplier import Bot as SupplierBot  # Import the Bot class



# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    master_all_path = r'C:\Users\Paulo\Desktop\bot_mrp\bot_mrp\05.03_Master_All_Sourcing_.xlsb'

    # 1 - etapa bom
    file_bom = (r"C:\Users\Paulo\Documents\Trabalho\base_data_09_12_24\base_qlik\1213_Bom_Master.xlsb")
    bom_bot = BomBot()
    bom_bot.process_bom(file_bom, master_all_path)

    # 2 - etapa delivery
    file_delivery =(r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Delivery Status 03062025 0935.xlsx")
    bot = Bot()
    bot.process_delivery(file_delivery, master_all_path)
    
    # 3 - etapa onhand
    file_onhand = (r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Integration Onhand Inquiry20250306.xlsx")
    onhand_bot = OnhandBot()
    onhand_bot.process_onhand(file_onhand, master_all_path)

    # 4 - Up_PlanAssy_Click
    plan_assy_bot = UpPlanAssyBot()
    plan_assy_bot.process_plan_assy(master_all_path)

    # 5 - supplier componel
    file_componel = (r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\07.03.25-LG-COMPONEL_-ESTOQUE_\07.03.25 LG COMPONEL_ ESTOQUE_.xlsx")
    supplier_bot = SupplierBot()
    supplier_bot.process_supplier(file_componel, master_all_path)



    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


if __name__ == '__main__':
    main()
