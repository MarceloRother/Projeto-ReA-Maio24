import glob
import os
import pandas as pd
import csv
import copy

# Declaração da classe Dados para armazenar os dados de cada linha do CSV
class Dados:
    id = ''
    nome = ''
    saldoInicio = ''
    entradas = ''
    baixas = ''
    saldoFinal = ''

# Padrão para encontrar arquivos .xls no diretório atual
directory_path = 'MOVIMENTAÇÃO DOS ESTOQUES UNIDADES 12 MESES\\TOP CENTER'
file_pattern = os.path.join(directory_path, '*.xls')

for xls_file in glob.glob(file_pattern):
    # Ler arquivo .xls
    df = pd.read_excel(xls_file)

    csv_file = xls_file.replace('.xls', '.csv')

    # Salvar o DataFrame em um arquivo Excel no formato .csv
    df.to_csv(csv_file, index=False)

    ## REMOVENDO ESPAÇOS

    # Abre o arquivo .csv para leitura com especificação de codificação UTF-8
    with open(csv_file, 'r', encoding='utf8') as arqEntrada:
        # Cria um leitor CSV para iterar sobre as linhas do arquivo
        leitor = csv.reader(arqEntrada)

        # Declara uma lista para armazenar o conjunto final dos dados selecionados
        dadosFinal = list()

        cont = 0

        # Itera sobre cada linha do arquivo
        for linha in leitor:
            if cont > 7:
                if 'Página' not in linha[16]:
                    # Declara uma lista para armazenar os dados da linha atual que não são vazios ou "-"
                    dadosLinha = list()

                    # Itera sobre cada célula da linha
                    for celula in linha:
                        # Verifica se a célula não está vazia e não é "-"
                        if celula != "" and celula != "-":
                            # Adiciona a célula à lista dadosLinha
                            dadosLinha.append(celula)
            
                    # Se dadosLinha não estiver vazia, adiciona à lista dadosFinal
                    if dadosLinha:
                        dadosFinal.append(dadosLinha)
                else:
                    cont = 1
            else:
                cont += 1 

    # Criação de diretório de saída, se não existir
    directory_path_ws = os.path.join(directory_path, 'ARQUIVOS EXTRAS')
    if not os.path.exists(directory_path_ws):
        os.makedirs(directory_path_ws)
    
    csv_file_ws = os.path.join(directory_path_ws, os.path.basename(csv_file).replace('.csv', ' - SE.csv'))

    with open(csv_file_ws, 'w', encoding='utf8') as arqSaida:
        # Cria um escritor CSV para escrever as linhas no novo arquivo
        escritor = csv.writer(arqSaida)
        
        # Itera sobre cada conjunto de dados armazenado em dadosFinal
        for dados in dadosFinal:
            # Escreve a linha no arquivo de saída
            escritor.writerow(dados)

        # Exibe uma mensagem de confirmação
        print("Espaços removidos!")

    # Inicializa uma lista para armazenar objetos da classe Dados
    listaDeDados = []

    ## SELEÇÃO DE DADOS

    # Abre o arquivo CSV para leitura com especificação de codificação UTF-8
    with open(csv_file_ws, 'r', encoding='utf8') as arqEntrada:
        # Cria um leitor CSV para iterar sobre as linhas do arquivo
        leitor = csv.reader(arqEntrada)
        
        # Cria uma instância da classe Dados
        dados = Dados()
        # Itera sobre cada linha do arquivo
        for linha in leitor:
            # Verifica se a linha não está vazia
            if len(linha) > 0:
                # Verifica se a linha começa com 'Material:' para capturar os dados iniciais
                if linha[0] == 'Material:':
                    dados.id = linha[1]
                    dados.nome = linha[2]
                    dados.saldoInicio = linha[4]
                # Verifica se a linha começa com 'Saldo atual:' para capturar o saldo final
                elif linha[0] == 'Saldo atual:':
                    dados.saldoFinal = linha[1]
                    listaDeDados.append(copy.copy(dados))
                elif '2024-' not in linha[0] and 'Data' not in linha[0] and 'Local' not in linha[0] and 'Período' not in linha[0] and '2023-' not in linha[0]:
                    dados.entradas = linha[0]
                    dados.baixas = linha[1]
        print('Dados selecionados!')   

    csv_file_final = os.path.join(directory_path_ws, os.path.basename(csv_file).replace('.csv', ' - FINAL.csv'))

    with open(csv_file_final, 'w', encoding='utf8') as arqFinal:
        # Cria um escritor CSV para escrever as linhas no novo arquivo
        escritor = csv.writer(arqFinal)

        # Escreve o cabeçalho no arquivo de saída
        escritor.writerow(('ID', 'NOME', 'Saldo Inicial', 'Entradas', 'Baixas', 'Saldo Final'))
        # Itera sobre cada objeto Dados na lista listaDeDados
        for dado in listaDeDados:
            # Escreve os dados do objeto Dados no arquivo de saída
            escritor.writerow((dado.id, dado.nome, dado.saldoInicio, dado.entradas, dado.baixas, dado.saldoFinal))

    # Ler o arquivo CSV
    df = pd.read_csv(csv_file_final)

    # Salvar o DataFrame em um arquivo Excel no formato .xlsx

    # Criação de diretório final, se não existir
    directory_path_final = os.path.join(directory_path, 'FINAL')
    if not os.path.exists(directory_path_final):
        os.makedirs(directory_path_final)

    xlsx_file_final = os.path.join(directory_path_final, os.path.basename(csv_file_final).replace(' - FINAL.csv', ' - FINAL.xlsx'))
    df.to_excel(xlsx_file_final, index=False)

print("Processamento concluído!")
