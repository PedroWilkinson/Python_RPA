#Importamos o selenium para trabalhar com as páginas da web
from selenium import webdriver as opcoes_selenium_aula
from selenium.webdriver.common.keys import Keys

#Importando a biblioteca do pyautogui para trabalhar com o tempo e teclas do teclado
import pyautogui as tempoPausaComputador

#Usuando o pyautogui para controlar o teclado
import pyautogui as teclasAtalhoTeclado

#Usando o By para trabalhar com as atualizações mais recentes
from selenium.webdriver.common.by import By

#Passamos autorização ao acesso as configurações do Chrome
meuNavegador = opcoes_selenium_aula.Chrome()
meuNavegador.get('https://www.google.com.br/')

tempoPausaComputador.sleep(4)

#Procurando pelo elemento NAME e quando encontrar vou escrever Dolar hoje
meuNavegador.find_element(By.NAME, "q").send_keys("Dolar hoje")

tempoPausaComputador.sleep(4)

#Retorna para o campo name q
#Faz a busca do valor que está digitado no campo NAME q
meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(4)


valorDolarPeloGoogle = meuNavegador.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

tempoPausaComputador.sleep(4)

print(valorDolarPeloGoogle)

#-----------------------------------------------------------------------------------------------------

tempoPausaComputador.sleep(2)

#Retorna para o campo name q
#Faz a busca do valor que está digitado no campo NAME q
meuNavegador.find_element(By.NAME, "q").send_keys("")

tempoPausaComputador.sleep(4)

#Estamos usando o pyautogui para apertar a tecla TAB
teclasAtalhoTeclado.press('tab')

tempoPausaComputador.sleep(4)

#Estamos usando o pyautogui para apertar a tecla enter
#Enter para limpar o campo de pesquisa
teclasAtalhoTeclado.press('enter')

tempoPausaComputador.sleep(4)

#Procurando pelo elemento NAME e quando encontrar vou escrever Dolar hoje
meuNavegador.find_element(By.NAME, "q").send_keys("Euro hoje")

tempoPausaComputador.sleep(4)

#Retorna para o campo name q
#Faz a busca do valor que está digitado no campo NAME q
meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(4)


valorEuroPeloGoogle = meuNavegador.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

tempoPausaComputador.sleep(4)

print(valorEuroPeloGoogle)

#----------------------------------------------------------------------------------------------------------------------------------

import xlsxwriter 
import os

nomeCaminhoArquivo = "C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Extraindo Valor do Dolar e Euro e Salvando no Excel\\Imprime Dolar e Euro Google.xlsx"
planilhaCriada = xlsxwriter.Workbook(nomeCaminhoArquivo)
sheet1 = planilhaCriada.add_worksheet()

#Escervendo nas células
sheet1.write("A1", "Dolar")
sheet1.write("B1", "Euro")
sheet1.write("A2", valorDolarPeloGoogle)
sheet1.write("B2", valorEuroPeloGoogle)

#Substituir a vírgula por ponto 
valorDolarPeloGoogle = valorDolarPeloGoogle.replace(',','.')
valorEuroPeloGoogle = valorEuroPeloGoogle.replace(',','.')

#Convertendo o valor do Dolar e do Euro de Sting para Float
valor_Dolar_Tipo_Float = float(valorDolarPeloGoogle)
valor_Euro_Tipo_Float = float(valorEuroPeloGoogle)

sheet1.write("A3", valor_Dolar_Tipo_Float)
sheet1.write("B3", valor_Euro_Tipo_Float)

#Fechando o arquivo do Excel que está em segundo plano
planilhaCriada.close()

#Abro o arquivo
os.startfile(nomeCaminhoArquivo)

