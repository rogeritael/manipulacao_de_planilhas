from openpyxl import Workbook
from datetime import date

# acao = input('Qual codigo da açao voce quer processar?').upper()
acao = 'bidi4'

with open(f'./dados/{acao}.txt', 'r') as arquivo_cotacao:
    # Pega todas as linhas do arquivo
    linhas = arquivo_cotacao.readlines()
    
    # formata linha a linha da forma que queremos
    linhas = [linha.replace('\n', '').split(';') for linha in linhas]

    #cria um planilha em memoria
    workbook = Workbook()

    # pega a atual ativa e muda o nome
    aba_atual = workbook.active
    aba_atual.title = "Dados"

    # criaçao do cabeçalho
    aba_atual.append(['DATA', 'COTAÇAO', 'BANDA INFERIOR', 'BANDA SUPERIOR'])

    indice = 2
    for linha in linhas:
        # DATA
        ano_mes_dia = linha[0].split(' ')[0].split('-')
        data = date(
            int(ano_mes_dia[0]),
            int(ano_mes_dia[1]),
            int(ano_mes_dia[2])
        )

        #COTACAO
        cotacao = float(linha[1])


        #Preenchendo as colunas dinamicamente
        aba_atual[f'A{indice}'] = data
        aba_atual[f'B{indice}'] = cotacao

        #BANDA INFERIOR
        aba_atual[f'C{indice}'] = f'=AVERAGE(B{indice}:B{indice+19}) - 2*STDEV(B{indice}:B{indice+19})'
        #BANDA SUPERIOR
        aba_atual[f'D{indice}'] = f'=AVERAGE(B{indice}:B{indice+19}) + 2*STDEV(B{indice}:B{indice+19}'

        indice += 1








    workbook.save('./planilha/dados.xlsx')