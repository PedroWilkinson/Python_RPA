import xlsxwriter as opcoesDoXlsxWriter
import os

nomeCaminhoArquivo = 'C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Xlsxwriter\\PintaFundoEFonteExemplo.xlsx'
minhaPlanilha = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)

sheetDados = minhaPlanilha.add_worksheet("Dados")

#Altera a cor do fundo da celula
#corFundo = minhaPlanilha.add_format({'fg_color':'yellow'})

#Altera a cor da fonte
corFonte = minhaPlanilha.add_format()
corFonte.set_font_color('blue')

corFonteFundo = minhaPlanilha.add_format({'align':'center',
                                          'font_color': 'white',
                                          'bold': True,
                                          'bg_color':'gray'})

#Adicionando dados na Sheet
sheetDados.write("A1", "Nome", corFonteFundo)
sheetDados.write("B1", "Idade", corFonteFundo)
sheetDados.write("A2", "Amanda", corFonte)
sheetDados.write("B2", 21, corFonte)
sheetDados.write("A3", "Pedro", corFonte)
sheetDados.write("B3", 28, corFonte)


#Fechando o arquivo
minhaPlanilha.close()


#Abrindo o arquivo
os.startfile(nomeCaminhoArquivo)

