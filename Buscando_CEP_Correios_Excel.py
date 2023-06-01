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

#Pega os dados da rua no site do busca CEP pelo XPATH
rua = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]')[0].text
print(rua)

#Pega os dados da bairro no site do busca CEP pelo XPATH
bairro = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[2]')[0].text
print(bairro)

#Pega os dados da cidade no site do busca CEP pelo XPATH
cidade = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[3]')[0].text
print(cidade)

#Pega os dados da CEP no site do busca CEP pelo XPATH
cep = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[4]')[0].text
print(cep)

#------------------------------------------------------------------------------------------------------------------------------------------
# SEGUNDA ETAPA DO CODIGO/PROJETO
#------------------------------------------------------------------------------------------------------------------------------------------

from openpyxl import load_workbook
import os

#Pegamos o caminho + nome do arquivo no computador
nome_arquivo_cep = "C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Extrair Endereco Site Busca CEP\\PesquisaEndereco_1.xlsx"
planilhaDadosEndereco = load_workbook(nome_arquivo_cep)

#Selecionando a sheet / planilha de Dados
sheet_selecionada = planilhaDadosEndereco["Dados"]

#Pegamos a ultima linha preenchida na coluna A e acrescentamos + 1
linha = len(sheet_selecionada['A']) + 1
colunaA = "A" + str(linha)  #Criando a variavel para juntar A + a ultima linha ex: A2
colunaB = "B" + str(linha)  #Criando a variavel para juntar B + a ultima linha ex: B2
colunaC = "C" + str(linha)  #Criando a variavel para juntar C + a ultima linha ex: C2
colunaD = "D" + str(linha)  #Criando a variavel para juntar D + a ultima linha ex: D2

#Imprimindo as informacoes do site na planilha
sheet_selecionada[colunaA] = rua #A2
sheet_selecionada[colunaB] = bairro #B2
sheet_selecionada[colunaC] = cidade #C2
sheet_selecionada[colunaD] = cep #D2

#Salvando o arquivo do Excel com as novas informacoes
planilhaDadosEndereco.save(filename=nome_arquivo_cep)

#Abrindo e apresentando na tela o arquivo com os dados
os.startfile(nome_arquivo_cep)







