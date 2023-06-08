import xlsxwriter as opcoesDoXlsxWriter
import os

nomeCaminhoArquivo = 'C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Xlsxwriter\\PintaFundoEFonte.xlsx'
minhaPlanilha = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)

sheetDados = minhaPlanilha.add_worksheet("Dados")

#Altera a cor do fundo da celula
corFundo = minhaPlanilha.add_format({'fg_color':'yellow'})

#Altera a cor da fonte
corFonte = minhaPlanilha.add_format()
corFonte.set_font_color('blue')

#Adicionando dados na Sheet
sheetDados.write("A1", "Nome", corFundo)
sheetDados.write("B1", "Idade", corFundo)
sheetDados.write("A2", "Amanda", corFonte)
sheetDados.write("B2", 21, corFonte)
sheetDados.write("A3", "Pedro", corFonte)
sheetDados.write("B3", 28, corFonte)


#Fechando o arquivo
minhaPlanilha.close()


#Abrindo o arquivo
os.startfile(nomeCaminhoArquivo)

