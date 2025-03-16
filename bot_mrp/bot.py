# Import for integration with BotCity Maestro SDK
from botcity.maestro import *
from modules.process_delivery import Bot  # Import the Bot class
from modules.process_onhand import Bot as OnhandBot  # Import the Bot class
from modules.process_bom import Bot as BomBot  # Import the Bot class



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

    # 1 - etapa bom
    file_bom = (r"C:\Users\Paulo\Documents\Trabalho\base_data_09_12_24\base_qlik\1213_Bom_Master.xlsb")
    bom_bot = BomBot()
    bom_bot.process_bom(file_bom)

    # 2 - etapa delivery
    file_delivery =(r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Delivery Status 03062025 0935.xlsx")
    bot = Bot()
    bot.process_delivery(file_delivery)
    
    # 3 - etapa onhand
    file_onhand = (r"C:\Users\Paulo\Desktop\tarefa\planilhas_base\Integration Onhand Inquiry20250306.xlsx")
    onhand_bot = OnhandBot()
    onhand_bot.process_onhand(file_onhand)



    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


if __name__ == '__main__':
    main()
