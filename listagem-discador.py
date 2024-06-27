import glob
import pandas as pd
import xlrd
from openpyxl import Workbook

lista_arquivo = []
for arquivo in glob.glob(r"M:\Atendimento\Oferta Ativa\MK Bairros\Bases Bairros\2024\*xls"):
    lista_arquivo.append(arquivo)

# Salvando Blacklist

df = pd.read_excel(r'M:\Atendimento\Oferta Ativa\MK Bairros\Importante\01 - Contatos excluídos.xlsx', sheet_name='Blacklist')

def lista_blacklist(nome_coluna, data_frame):
    valores_coluna = data_frame[nome_coluna]
    valores_coluna = valores_coluna.dropna()
    valores_coluna = valores_coluna.apply(str)

    blacklist = []
    for i in valores_coluna:
        l = i.replace('.0', '')
        if l != '0':
            blacklist.append(l)
    
    return blacklist

coluna1 = lista_blacklist('Tel 1', df)
coluna2 = lista_blacklist('Tel 2', df)
coluna3 = lista_blacklist('Tel 3', df)
coluna4 = lista_blacklist('Tel 4', df)
coluna5 = lista_blacklist('Cel1', df)

headhunter = coluna1 + coluna2 + coluna3 + coluna4 + coluna5

def excluir_linhas_excel(arquivo_entrada, arquivo_saida, lista_exclusao):
    # Abrir o arquivo Excel com xlrd
    workbook = xlrd.open_workbook(arquivo_entrada)
    sheet = workbook.sheet_by_index(0)

    # Criar uma nova planilha para salvar as modificações com openpyxl
    wb = Workbook()
    ws = wb.active

    # Lista para manter o índice das linhas a serem excluídas
    linhas_excluir = set()

    # Identificar as linhas a serem excluídas
    for row_index in range(sheet.nrows):
        row = sheet.row_values(row_index)
        if any(item in row for item in lista_exclusao):
            linhas_excluir.add(row_index)

    # Escrever as linhas não excluídas na nova planilha
    for row_index in range(sheet.nrows):
        if row_index not in linhas_excluir:
            row_values = sheet.row_values(row_index)
            ws.append(row_values)

    # Salvar as modificações
    wb.save(arquivo_saida)

for i in lista_arquivo:
    arquivo_entrada = i
    nome_arquivo = i.replace('M:\\Atendimento\\Oferta Ativa\\MK Bairros\\Bases Bairros\\2024\\', '')
    arquivo_saida = f'M:\\Atendimento\\Oferta Ativa\\MK Bairros\\Bases Bairros\\2024\\Base Discador Convertida\\{nome_arquivo}.xlsx'
    lista_exclusao = headhunter

    excluir_linhas_excel(arquivo_entrada, arquivo_saida, lista_exclusao)
    print(f"Arquivo {nome_arquivo} Final criado.")

#Lista de arquivos
lista_arquivoxlsx = glob.glob(r"M:\Atendimento\Oferta Ativa\MK Bairros\Bases Bairros\2024\Base Discador Convertida\*.xlsx")

# Iterate over each Excel file
for i in lista_arquivoxlsx:
   
    excel = pd.read_excel(i)
    telefones_df_new = pd.DataFrame(excel, columns=['Tel 1', 'Tel 2', 'Tel 3', 'Tel 4', 'Cel1', 'Cel2', 'Nome', 'Bairro', 'Email'])
    telefones_organizados_new = telefones_df_new.melt(id_vars=['Nome', 'Bairro', 'Email'], 
                                           value_vars=['Tel 1', 'Tel 2', 'Tel 3', 'Tel 4', 'Cel1', 'Cel2'], 
                                           var_name='Phone_Type', 
                                           value_name='Phone')
    telefones_organizados_new = telefones_organizados_new.dropna(subset=['Phone'])
    telefones_organizados_new = telefones_organizados_new[telefones_organizados_new['Phone'] != 0]
    telefones_organizados_new = telefones_organizados_new.drop(columns=['Phone_Type'])
    output_file_path_new = f'M:\\Atendimento\\Oferta Ativa\\MK Bairros\\Bases Bairros\\2024\\Base Discador Convertida\\Base Discador Final\\{i.split("\\")[-1]}'
    telefones_organizados_new.to_excel(output_file_path_new, index=False)

    print(f"{i} -> Completo")
