from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys 
import pyautogui as tempoEspera

#IMportando o elemento By
from selenium.webdriver.common.by import By

#Abre o navegador do Google Chrome e abre o site do bussca Cep
navegador = opcoesSelenium.Chrome()
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

#Aguarda 3 segundos para dar tempo do computador processar as informacoes:
tempoEspera.sleep(3)

#Imprime o CEP no campo Digite um CEP ou um Endereco:
navegador.find_element(By.NAME, "endereco").send_keys("05892387")

#Aguarda 3 segundos para dar tempo do computador processar as informacoes:
tempoEspera.sleep(3)

#Clicando no botao btn_pesquisar para pesquisar
navegador.find_element(By.NAME, "btn_pesquisar").click()

#Aguarda 3 segundos para dar tempo do computador processar as informacoes:
tempoEspera.sleep(3)