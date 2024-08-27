from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin
from botcity.maestro import *

BotMaestroSDK.RAISE_NOT_CONNECTED = False

lista = [
    ["João", 28],
    ["Maria", 35],
    ["Pedro", 22],
    ["Ana", 29],
    ["Luís", 31],
    ["Sofia", 27],
    ["Carlos", 24],
    ["Marta", 33],
    ["Ricardo", 30],
    ["Laura", 26],
    ["André", 29],
    ["Isabela", 32],
    ["Tiago", 25]]

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot_excel = BotExcelPlugin()

    var_strCaminhoArquivo = './Excel-BotCity.xlsx'
    bot_excel.read(var_strCaminhoArquivo)

    for linha in bot_excel.as_list():
      print(linha)

    # Adicionando uma nova linha na instância
    bot_excel.add_row(['Felipe', '32'])

    # Reescrever os dados no arquivo xlsx.
    bot_excel.write(var_strCaminhoArquivo)

    for linha in bot_excel.as_list():
      print(linha)

    # Capturando o total de linhas da nossa instância
    indice = len(bot_excel.as_list())

    #Removendo a última linha
    bot_excel.remove_row(indice)

    # Escrevendo os novos dados na planilha
    bot_excel.write(var_strCaminhoArquivo)

    for linha in bot_excel.as_list():
      print(linha)

    # Adicionando várias linhas ao mesmo tempo
    bot_excel.add_rows(lista)

    # Escrevendo o arquivo Excel com os dados
    bot_excel.write(var_strCaminhoArquivo)

    for linha in bot_excel.as_list():
      print(linha)

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()