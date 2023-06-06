from openpyxl import load_workbook
import os

from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera

from selenium.webdriver.common.by import By

'''
navegadorFormulario = opcoesSelenium.Chrome()
navegadorFormulario.get("https://pt.surveymonkey.com/r/RD8SMQ3")

tempoEspera.sleep(3)

#Preenche nome
navegadorFormulario.find_element(By.NAME, "135842344").send_keys("Nicole Ferreira")

tempoEspera.sleep(3)

#Preenche email
navegadorFormulario.find_element(By.NAME, "135842392").send_keys("nicole.ferreira@gmail.com")

tempoEspera.sleep(3)

#Preenche telefone
navegadorFormulario.find_element(By.NAME, "135842405").send_keys("(11) 11111 - 1111")

tempoEspera.sleep(3)

#Preenche Sobre
navegadorFormulario.find_element(By.NAME, "135842659").send_keys("Sei automatizar processos e planilhas com Python")

tempoEspera.sleep(3)


#Preenche Radio Button Feminino
navegadorFormulario.find_element(By.ID, "135842623_1011908190_label").click()

tempoEspera.sleep(3)

#Clica para enviar as informações
navegadorFormulario.find_element(By.XPATH, '//*[@id="patas"]/main/article/section/form/div[2]/button').click()
'''
#--------------------------------------------------------------------------------------------------------------------
# SEGUNDA PARTE
#--------------------------------------------------------------------------------------------------------------------



nomeCaminhoArquivo = "C:\\Users\\Pedro W\Desktop\\PYTHON_GERAL\\PythonRPA\\Preenchendo Formulários na Web com Python e Excel\\DadosFormulario.xlsx"
planilha_aberta = load_workbook(nomeCaminhoArquivo)

#Seleciona a sheet que tem as informacoes a serem passadas para o formulario on-line
sheet_selecionada = planilha_aberta["Dados"]

for linha in range(2, len(sheet_selecionada['A']) + 1):

    nome = sheet_selecionada['A%s' % linha].value
    email = sheet_selecionada['B%s' % linha].value
    telefone = sheet_selecionada['C%s' % linha].value
    sexo= sheet_selecionada['D%s' % linha].value
    sobre = sheet_selecionada['E%s' % linha].value

    tempoEspera.sleep(2)

    navegadorFormulario = opcoesSelenium.Chrome()
    navegadorFormulario.get("https://pt.surveymonkey.com/r/RD8SMQ3")

    tempoEspera.sleep(6)

    #Preenche nome
    navegadorFormulario.find_element(By.NAME, "135842344").send_keys(nome)

    tempoEspera.sleep(3)

    #Preenche email
    navegadorFormulario.find_element(By.NAME, "135842392").send_keys(email)

    tempoEspera.sleep(3)

    #Preenche telefone
    navegadorFormulario.find_element(By.NAME, "135842405").send_keys(telefone)

    tempoEspera.sleep(3)

    #Preenche Sobre
    navegadorFormulario.find_element(By.NAME, "135842659").send_keys(sobre)

    tempoEspera.sleep(3)

    if sexo == "Masculino":

        #Preenche Radio Button Masculino
        navegadorFormulario.find_element(By.ID, "135842623_1011908189_label").click()

    else: 

        #Preenche Radio Button Feminino
        navegadorFormulario.find_element(By.ID, "135842623_1011908190_label").click()

    tempoEspera.sleep(3)

    #Clica para enviar as informações
    navegadorFormulario.find_element(By.XPATH, '//*[@id="patas"]/main/article/section/form/div[2]/button').click()



