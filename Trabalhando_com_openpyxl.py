from openpyxl import load_workbook
import os

caminho_nome_arquivo = "C:\\Users\\Pedro W\\Desktop\\PYTHON_GERAL\\PythonRPA\\Openpyxl\\DeletarLinhaColuna.xlsx"
planilha_aberta = load_workbook(filename=caminho_nome_arquivo)

#Seleciona a sheet de Aluno
sheet_selecionada = planilha_aberta['Aluno']

#Deleta as Linhas
sheet_selecionada.delete_rows(3)
sheet_selecionada.delete_rows(3)
sheet_selecionada.delete_rows(5)

#Deleta a Coluna
sheet_selecionada.delete_cols(2)

#Salva a planilha com as alteracoes
planilha_aberta.save(filename=caminho_nome_arquivo)

#Abre a planilha
os.startfile(caminho_nome_arquivo)
