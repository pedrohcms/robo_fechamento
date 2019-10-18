# -*- coding: utf8 -*-
#Author: Pedro Henrique Correa Mota da Silva

import pandas as pd
import xlrd
import os
from datetime import date
import time

init_time = time.time()
op = 1

while op == 1:
    print('Digite 1 para processar arquivo excel')
    print('Digite 0 para sair')
    op = int(input())
     
    if op == 1:
        path = input('Digite o caminho do arquivo excel, com a extensão: \n')

        if os.path.isfile(path) == False:
            print('Arquivo não encontrado!')
            continue
        else:
            print('Arquivo recebido com sucesso!')

            sheet_name = input('Digite o nome da planilha: ')

            resposta = input('A panilha foi ordenada de maneira crescente nos campos de interface e empresa, respectivamente Pressione S - SIM |N - NÃO?: ')

            if resposta == 'N':
                print('Ordene as colunas primeiro e depois utilize o algoritmo')
                exit()

            try:
                document = pd.read_excel(path, sheet_name=sheet_name)    
            except:
                print('A panilha desejada não foi encontrada!')
                continue

            if os.path.isdir('./resultados') == False:
                os.mkdir('./resultados')

            print('Iniciando o Processamento!')

            #Criando o arquivo txt caso ele não exista
            txt = open('./resultados/'+sheet_name+'_'+str(date.today().year)+'.txt', 'w')

            init_time = time.time()

            #Iniciando os DataFrames
            interface = pd.DataFrame(document, columns= ['Interface'])
            empr = pd.DataFrame(document, columns= ['Empr'])
            cl = pd.DataFrame(document, columns= ['CL'])
            conta = pd.DataFrame(document, columns= ['Conta'])
            montante = pd.DataFrame(document, columns= ['Valor do Montante'])
            pep = pd.DataFrame(document, columns= ['Elemento PEP'])
            chave_ref = pd.DataFrame(document, columns= ['Chv.ref.1'])
            data_doc = pd.DataFrame(document, columns= ['Data do Doc'])
            contrato = pd.DataFrame(document, columns= ['Contrato'])
            data_lancamento = pd.DataFrame(document, columns= ['Data Lançamento'])
            historico = pd.DataFrame(document, columns= ['Denominação'])

            for linha in range(document.shape[0]):
                string_linha = '&SdtTexto.Add(\''
                
                #empresa
                string = str(empr.loc[linha][0])

                if len(string) != 5:
                    string = string.zfill(5)

                string_linha += string
                
                #Débito ou crédito
                string = str(cl.loc[linha][0])

                string_linha += string

                #Conta contabil
                string = str(conta.loc[linha][0])

                string_linha += string

                #Montante
                string = str(montante.loc[linha][0])

                if '.' in string:
                    string = string.split('.')
                    
                    if len(string[1]) == 1:
                        string[1] += '0'
                    
                    string = string[0] + string[1]
                    string = string.zfill(15)

                string_linha += string
                
                #PEP
                string = str(pep.loc[linha][0])
                
                string = string + ' ' * (23 - len(string)) 

                string_linha += string
                
                #Chav. Ref. 1
                string = str(chave_ref.loc[linha][0])
                
                if string == 'nan':
                    string = ' ' * (12) 
                else:
                    string = string + ' ' * (12 - len(string))

                string_linha += string

                #Data Documento
                string = data_doc.loc[linha][0]

                string = format(string, "%Y%m%d")

                string_linha += string

                #Contrato
                string = str(contrato.loc[linha][0])

                string = string.zfill(6)

                string_linha += string

                #Data do lançamento
                string = data_lancamento.loc[linha][0]

                string = format(string, "%d/%m/%Y")

                string_linha += string
                
                #Histórico(Denominação)
                string = str(historico.loc[linha][0])

                if len(string) >= 50:
                    string = string[0:50]
                else:
                    string = string + ' ' * (50 - len(string))

                string_linha += string

                #Colocando o ')
                string_linha += '\')'

                #Hora de salvar no arquivo
                txt.write(string_linha)
                txt.write('\n')

                if linha < document.shape[0]-1: 
                    interface_atual = str(interface.loc[linha][0])
                    interface_prox = str(interface.loc[linha+1][0])

                    if interface_atual != interface_prox:
                        txt.write('Do \'Processar\'')
                        txt.write('\n')
                
                print('Linha: '+str(linha))

            txt.write('Do \'Processar\'')

            txt.close()

            end_time = time.time()

            print('O tempo de execução foi '+ str(end_time - init_time) + ' segundos')

    elif op == 0:
        break
    else:
        print('Opção inválida')