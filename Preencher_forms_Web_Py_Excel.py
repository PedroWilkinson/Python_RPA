from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera

from selenium.webdriver.common.by import By

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

#--------------------------------------------------------------------------------------------------------------------
# SEGUNDA PARTE
#--------------------------------------------------------------------------------------------------------------------

from openpyxl import load_workbook
import os

nomeCaminhoArquivo = ""


