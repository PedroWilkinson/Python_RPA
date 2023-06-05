from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera

navegadorFormulario = opcoesSelenium.Chrome()
navegadorFormulario.get("https://pt.surveymonkey.com/r/RD8SMQ3")

tempoEspera.sleep(6)

#Preenche nome
navegadorFormulario.find_element_by_name("135842344").send_keys("Nicole Ferreira")

#Preenche email
navegadorFormulario.find_element_by_name("135842392").send_keys("nicole.ferreira@gmail.com")

#Preenche telefone
navegadorFormulario.find_element_by_name("135842405").send_keys("(11) 11111 - 1111")

#Preenche Sobre
navegadorFormulario.find_element_by_name("135842659").send_keys("Sei automatizar processos e planilhas com Python")


#Preenche Radio Button Feminino
navegadorFormulario.find_element_by_id("135842623_1011908190_label").click()

tempoEspera.sleep(6)

#Preenche Radio Button Feminino
navegadorFormulario.find_element_by_xpath('//*[@id="patas"]/main/article/section/form/div[2]/button').click()





