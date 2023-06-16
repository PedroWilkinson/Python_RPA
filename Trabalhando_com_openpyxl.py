from openpyxl import load_workbook
import os

caminho_nome_arquivo = "C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Openpyxl\\InserirDados.xlsx"
planilha_aberta = load_workbook(filename=caminho_nome_arquivo)

#Seleciona a sheet de Aluno
sheet_selecionada = planilha_aberta['Aluno']

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



'''
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
