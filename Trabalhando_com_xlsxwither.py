import xlsxwriter as opcoesDoXlsxWriter
import os

nomeCaminhoArquivo = 'C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Xlsxwriter\\FormatacaoCondicional.xlsx'
workbook = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)

sheetPadrao = workbook.add_worksheet("Dados")


formatoMaior = workbook.add_format({'bg_color' : 'green',
                                    'font_color' : 'white'})

formatoMenor = workbook.add_format({'bg_color' : 'red',
                                    'font_color' : 'white'})


inserirDados = [
    ["Coluna 1", "Coluna 2", "Coluna 3", "Coluna 4"],
    [34, 50, 12, 34],
    [59, 58, 76, 51],
    [43, 80, 34, 12],
    [91, 58, 73, 19],
]

sheetPadrao.write('A1', "Celulas com valores >= 50 estao em verde e < 50 estao em vermelho")

for linha, range in enumerate(inserirDados):
    sheetPadrao.write_row(linha + 2, 1, range) #Linha 3, Coluna 2

sheetPadrao.conditional_format('B4:E7', {'type' : 'cell',
                                         'criteria' : '>=',
                                         'value' : 50,
                                         'format' : formatoMaior})    

sheetPadrao.conditional_format('B4:E7', {'type' : 'cell',
                                         'criteria' : '<',
                                         'value' : 50,
                                         'format' : formatoMenor})  


'''
sheetPadrao.write("B3", "Imagem logo Udemy")

sheetPadrao.insert_image('B5', 'C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Xlsxwriter\\logo.png')



add_merge_celulas = workbook.add_format({
    'bold': True,
    'border': 6,
    'align' : 'center',
    'valign' : 'vcenter',
    'size' : 46,
    'fg_color' : 'red',
    'font_color' : 'white',

})


sheetPadrao.merge_range('B3:I5', 'Aula de Merge Celulas', add_merge_celulas)


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
'''


#Fechando o arquivo
workbook.close()


#Abrindo o arquivo
os.startfile(nomeCaminhoArquivo)

