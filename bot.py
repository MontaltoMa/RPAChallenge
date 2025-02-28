# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

#Import for Clickium 
from clicknium import clicknium as cc, locator, ui

#Import for Pandas
import pandas as pd

#import for Selenium Driver
from selenium.webdriver.common.by import By

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

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    # bot.browser = Browser.FIREFOX

    # Uncomment to set the WebDriver path
    bot.driver_path = r"C:\BotCity\chromedriver.exe"

    # bot.driver_path = "<path to your WebDriver binary>"

    #Baixa o arquivo em uma pasta especifica 
    bot.download_folder_path = r"C:\Users\marcos.martins\Downloads"

    # Opens the BotCity website.
    bot.browse("https://rpachallenge.com/")

    bot.maximize_window()

    # Implement here your logic...
    
    #Localiza o botão de Download
    bot.find_element("//a[text()=' Download Excel ']", By.XPATH).click()
    print("Feito o download do Excel")

    #Caminho final do Arquivo      
    var_strCaminhoArquivo = r"C:\Users\marcos.martins\Downloads\challenge.xlsx"
    print("Transferido para a pasta de downloads")

    var_intTempoEspera = 60

    #Espera o donwload ser finalizado
    bot.wait_for_downloads(timeout= var_intTempoEspera)

    #Volta para a página do RPAChallenge
    bot.back()
    
    #Localiza o botão de Start
    bot.find_element("//button[text()='Start']", By.XPATH).click()
    print("Iniciado RPAChallenge")
    
    #Lê o arquivo excel
    var_dfClientes2 = pd.read_excel(var_strCaminhoArquivo)

    """
    For para percorrer todas as linhas preenchidas do excel!
    Localiza o XPath de cada elemento a ser preenchido
    E escreve ele em cada espaço especifico
    """
    print("Iniciado Loop de preenchimento")
    for index, linha in var_dfClientes2.iterrows():
        var_strFirstName = linha['First Name']
        var_strLastName = linha['Last Name ']
        var_strCompanyName = linha['Company Name']
        var_strRoleInCompany = linha['Role in Company']
        var_strAddress = linha['Address']
        var_strEmail = linha['Email']
        var_strPhoneNumber = str(linha['Phone Number'])

        
        bot.find_element("//input[@ng-reflect-name='labelFirstName']", By.XPATH).send_keys(var_strFirstName)

        bot.find_element("//input[@ng-reflect-name='labelLastName']", By.XPATH).send_keys(var_strLastName)

        bot.find_element("//input[@ng-reflect-name='labelCompanyName']", By.XPATH).send_keys(var_strCompanyName)

        bot.find_element("//input[@ng-reflect-name='labelRole']", By.XPATH).send_keys(var_strRoleInCompany)

        bot.find_element("//input[@ng-reflect-name='labelAddress']", By.XPATH).send_keys(var_strAddress)

        bot.find_element("//input[@ng-reflect-name='labelEmail']", By.XPATH).send_keys(var_strEmail)

        bot.find_element("//input[@ng-reflect-name='labelPhone']", By.XPATH).send_keys(var_strPhoneNumber)
            
        bot.find_element("//input[@type='submit']", By.XPATH).click()

    print("Finalizado loop de preenchimento!")
    # Wait 3 seconds before closing
    bot.wait(5000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK.",
    #     total_items=0,
    #     processed_items=0,
    #     failed_items=0
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()



