import xlsxwriter as opcoesDoXlsxWriter
import os

nomeCaminhoArquivo = 'C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Xlsxwriter\\Formulas.xlsx'
minhaPlanilha = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)

sheetDados = minhaPlanilha.add_worksheet("Dados")

#Altera a cor do fundo da celula
#corFundo = minhaPlanilha.add_format({'fg_color':'yellow'})

#Altera a cor da fonte
corFonte = minhaPlanilha.add_format()
corFonte.set_font_color('blue')

#Mesclando cores na celula
corFonteFundo = minhaPlanilha.add_format({'align':'center',
                                          'font_color': 'white',
                                          'bold': True,
                                          'bg_color':'gray'})

#Adicionando TITULOS
sheetDados.write("A1", "Número 1")
sheetDados.write("B1", "Número 2")
sheetDados.write("C1", "Fórmula")

#Adicionando valores na coluna A
sheetDados.write("A2", 10)
sheetDados.write("A3", 6)
sheetDados.write("A4", 8)
sheetDados.write("A5", 6)
sheetDados.write("A8", "Ana")

#Adicionando valores na coluna B
sheetDados.write("B2", 7)
sheetDados.write("B3", 5)
sheetDados.write("B4", 3)
sheetDados.write("B5", 1)
sheetDados.write("B8", "Paula")

#Adicionando formulas na coluna C
sheetDados.write_formula("C2", "=A2+B2")
sheetDados.write_formula("C3", "=A3-B3")
sheetDados.write_formula("C4", "=A4*B4")
sheetDados.write_formula("C5", "=A5/B5")
sheetDados.write_formula("C8", '=CONCATENATE(A8," ",B8)')

#Coluna Tamanho 15
sheetDados.set_column('A:C', 15)

#Fechando o arquivo
minhaPlanilha.close()


#Abrindo o arquivo
os.startfile(nomeCaminhoArquivo)

