from openpyxl import load_workbook
import os

from openpyxl.styles import Color, PatternFill, Font, Border, Side
from openpyxl.styles import colors
from openpyxl.cell import Cell

caminho_nome_arquivo = "C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Openpyxl\\Formulas.xlsx"
planilha_aberta = load_workbook(filename=caminho_nome_arquivo)

#Seleciona a sheet de Aluno
sheet_selecionada = planilha_aberta['Aluno']

sheet_selecionada



'''
#Popula as informacoes que vao para a planilha
dadosTabela = [
    ['Nome', 'Idade'],
    ['Berenice', 28],
    ['Caio', 32],
    ['Nicole', 34]
]

#O Append pega toda a lista e passa parar a sheet
for linhaPlanilha in dadosTabela:
    sheet_selecionada.append(linhaPlanilha)


corTitulo = PatternFill(start_color= '00FFFF00',
                       end_color= '00FFFF00',
                       fill_type='solid')

corCelulas = PatternFill(start_color= '00CCFFCC',
                       end_color= '00CCFFCC',
                       fill_type='solid')

sheet_selecionada["A1"].fill = corTitulo
sheet_selecionada["B1"].fill = corTitulo

for linha in range(2, len(sheet_selecionada['A']) + 1):

    celulaColunaA = "A" + str(linha)
    celulaColunaB = "B" + str(linha)

    sheet_selecionada[celulaColunaA].fill = corCelulas
    sheet_selecionada[celulaColunaB].fill = corCelulas


#Deleta as Linhas
sheet_selecionada.delete_rows(3)
sheet_selecionada.delete_rows(3)
sheet_selecionada.delete_rows(5)

#Deleta a Coluna
sheet_selecionada.delete_cols(2)
'''

#Salva a planilha com as alteracoes
planilha_aberta.save(filename=caminho_nome_arquivo)

#Abre a planilha
os.startfile(caminho_nome_arquivo)
