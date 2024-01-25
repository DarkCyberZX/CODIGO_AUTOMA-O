from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

def Login(driver):
    try:
            # Vá para a página do Google
        driver.get("https://gool.cittati.com.br/Login.aspx?ReturnUrl=%2fHome%2fInicio.aspx")

        # Espere até que o elemento esteja presente
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable ((By.ID, "ucTrocarModulo_moduloUrbano"))
        )
        elemento.click()

        elementologin = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable ((By.ID, "ucLogarUsuario_txtLogin"))
        )
        elementologin.send_keys("*************")

        elementosenha = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ucLogarUsuario_txtSenha"))
        )
        elementosenha.send_keys("*******")

        elemento_ok = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ucLogarUsuario_btnLogar"))
        )
        elemento_ok.click()
        Menu(driver)
    except:
        Login(driver)
        
    
def Menu(driver): 
    try:
        elementomenu = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "item_menu_2"))
        )
        elementomenu.click()
        elementodetalhamento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "50505"))
        )
        elementodetalhamento.click()
        Detalhamento(driver)
        
    except:
        Login(driver)
        
def Detalhamento(driver):
    global dia, mes, ano, tempo, empresa, dia_final
    try:
        if dia > dia_final:
            driver.quit()
        elementodata = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_txtData"))
        )
        driver.execute_script("arguments[0].removeAttribute('onkeypress');", elementodata)
        driver.execute_script(f"arguments[0].value = '{dia}/{mes}/{ano}';", elementodata)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", elementodata)

        # Localize o elemento select
        select_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_ddlEmpresa"))
        )

        select = Select(select_element)
        select.select_by_index(empresa)
        time.sleep(3)
        select_element_sentido = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_ddlSentido"))
        )
        time.sleep(2)
        select_sentido = Select(select_element_sentido)
        select_sentido.select_by_index(2)

        elemento_linhas = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_chkSelevionarTodosLinhas"))
        )
        time.sleep(1)
        elemento_linhas.click()
        elemento_ordenar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_chkOrdenarHorarioPartida"))
        )
        time.sleep(1)
        elemento_ordenar.click()

        elemento_excel = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_btnExportarExcel"))
        )
        elemento_excel.click()
        time.sleep(tempo)
        dia += 1
        Detalhamento2(driver)
    except:
        Login(driver)
        
def Detalhamento2(driver):
    global dia, mes, ano, tempo, empresa, dia_final
    try:
        if dia > dia_final:
            driver.quit()
        elementodata = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_txtData"))
        )
        driver.execute_script("arguments[0].removeAttribute('onkeypress');", elementodata)
        driver.execute_script(f"arguments[0].value = '{dia}/{mes}/{ano}';", elementodata)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", elementodata)
        elemento_excel = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_contentFiltroPesquisa_btnExportarExcel"))
        )
        elemento_excel.click()
        time.sleep(tempo)
        dia += 1
        Detalhamento2(driver)
    
    except:
        Login(driver)

input("Pressione Enter para continuar...")
dia = int(input("Qual o dia inicial? "))
dia_final = int(input("Qual o dia final? "))
mes = int(input("Qual o mês? "))
ano = int(input("Qual o ano? "))
tempo = int(input("Qual o tempo de espera? "))
empresa = int(input("Qual a empresa?\n" 
                "1- São Francisco\n"
                "2- Real Alagoas\n"
                "3- Cidade de Maceió\n "))

print("Iniciando...\n"
      "NÃO FECHE ESSE PROGRAMA.")

driver = webdriver.Edge()
Login(driver)
