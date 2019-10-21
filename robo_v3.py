# -*- coding: utf8 -*-
#Author: Pedro Henrique Correa Mota da Silva

import pandas as pd
import xlrd
import os
from datetime import date
from dateutil.parser import parse
import time

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

            resposta = input('A panilha foi ordenada de maneira crescente nos campos de interface e empresa, respectivamente Pressione S - SIM |N - NÃO?: ')

            if resposta == 'N':
                print('Ordene as colunas primeiro e depois utilize o algoritmo')
                exit()
            
            sheet_name = input('Digite o nome da planilha: ')

            try:
                document = pd.read_excel(path, sheet_name=sheet_name).astype(str)
            except:
                print('A panilha desejada não foi encontrada!')
                continue

            if os.path.isdir('./resultados') == False:
                os.mkdir('./resultados')

            print('Iniciando o Processamento!')

            #Criando o arquivo txt caso ele não exista
            txt = open('./resultados/'+sheet_name+'_'+str(date.today().year)+'.txt', 'w')

            #Iniciando os DataFrames
            interface = document['Interface'].values.tolist()
            empr = document['Empr'].values.tolist()
            cl = document['CL'].values.tolist()
            conta = document['Conta'].values.tolist()
            montante = document['Valor do Montante'].values.tolist()
            pep = document['Elemento PEP'].values.tolist()
            chave_ref = document['Chv.ref.1'].values.tolist()
            data_doc = document['Data do Doc'].values.tolist()
            contrato = document['Contrato'].values.tolist()
            data_lancamento = document['Data Lançamento'].values.tolist()
            historico = document['Denominação'].values.tolist()

            init_time = time.perf_counter()

            for linha in range(document.shape[0]):
                string_linha = '&SdtTexto.Add(\''
                
                #empresa
                string = empr[linha]
                
                if len(string) > 5 or len(string) < 3:
                    print('O valor do campo Empresa é inválido')
                    break
                else:
                    string = string.zfill(5)

                string_linha += string
                
                #Débito ou crédito(CL)
                string = cl[linha]

                if (string == 'C' or 'D') and len(string) == 1:
                    string_linha += string
                else:
                    print('O valor do campo CL é invalido')
                    break

                #Conta contabil
                string = conta[linha]

                if len(string) != 10:
                    print('O valor do campo Conta contábil é inválido')
                    break

                string_linha += string

                #Montante
                string = montante[linha]

                if len(string) > 15:
                    print('O valor do campo montante é inválido')
                    break

                if '.' in string:
                    string = string.split('.')
                    
                    if len(string[1]) == 1:
                        string[1] += '0'
                    
                    string = string[0] + string[1]
                    string = string.zfill(15)

                string_linha += string
                
                #PEP
                string = pep[linha]
                
                if len(string) == 15:
                    string = string + ' ' * (23 - len(string)) 

                    string_linha += string
                else:
                    print('O valor do campo Elemento PEP é inválido')
                    break
                
                #Chav. Ref. 1
                string = chave_ref[linha]
                
                if string == 'nan':
                    string = ' ' * (12) 
                elif len(string) == 12:
                    pass
                else:
                    print('O valor do campo chave referência é inválido')
                    break

                string_linha += string

                #Data Documento
                string = data_doc[linha]

                if len(string) != 10:
                    print('O valor do campo data é inválido')
                    break
            
                string = parse(string)
                string = format(string, "%Y%m%d")

                string_linha += string

                #Contrato
                string = contrato[linha]
                
                if len(string) < 5 or len(string) > 6:
                    print('O valor do campo contrato é inválido')
                    break
                elif len(string) == 5:
                    print('AVISO! O valor do campo contrato na linha: '+ linha + ' tem 5 números')
                
                string = string.zfill(6)

                string_linha += string

                #Data do lançamento
                string = data_lancamento[linha]

                if len(string) != 10:
                    print('O valor do campo data é inválido')
                    break

                string = parse(string)
                string = format(string, "%d/%m/%Y")

                string_linha += string
                
                #Histórico(Denominação)
                string = historico[linha]

                if len(string) >= 50:
                    string = string[0:50]
                else:
                    string = string + ' ' * (50 - len(string))

                string_linha += string

                #Coilocando o ')
                string_linha += '\')'

                #Hora de salvar no arquivo
                txt.write(string_linha)
                txt.write('\n')

                if linha < document.shape[0]-1: 
                    interface_atual = interface[linha][0]
                    interface_prox = interface[linha+1][0]

                    if interface_atual != interface_prox:
                        txt.write('Do \'Processar\'')
                        txt.write('\n')
                
                print('Linha: '+str(linha))

            end_time = time.perf_counter()

            txt.write('Do \'Processar\'')

            txt.close()

            print('O tempo de execução foi '+ str(end_time - init_time) + ' segundos')

    elif op == 0:
        print('Saindo')
        break
    else:
        print('Opção inválida')