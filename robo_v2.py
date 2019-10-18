# -*- coding: utf8 -*-
#Author: Pedro Henrique Correa Mota da Silva

import pandas as pd
import xlrd
import os
from datetime import date
from dateutil.parser import parse
import time

#Criando o arquivo txt caso ele não exista
txt = open('setembro_2019.txt', 'w')

document = pd.read_excel('Ajuste Contabil FEchamento Toledo 2.xlsx', sheet_name='Setembro').astype(str)

interface = pd.DataFrame(document, columns=['Interface'])
empr = pd.DataFrame(document, columns=['Empr'])
cl = pd.DataFrame(document, columns=['CL'])
conta = pd.DataFrame(document, columns=['Conta'])
montante = pd.DataFrame(document, columns=['Valor do Montante'])
pep = pd.DataFrame(document, columns=['Elemento PEP'])
chave_ref = pd.DataFrame(document, columns=['Chv.ref.1'])
data_doc = pd.DataFrame(document, columns=['Data do Doc'])
contrato = pd.DataFrame(document, columns=['Contrato'])
data_lancamento = pd.DataFrame(document, columns=['Data Lançamento'])
historico = pd.DataFrame(document, columns=['Denominação'])

init_time = time.perf_counter()

for linha in range(document.shape[0]):
    string_linha = '&SdtTexto.Add(\''
                
    #empresa
    string = empr.iloc[linha][0]

    if len(string) != 5:
        string = string.zfill(5)

        string_linha += string
                
    #Débito ou crédito
    string = cl.iloc[linha][0]

    string_linha += string

    #Conta contabil
    string = conta.iloc[linha][0]

    string_linha += string

    #Montante
    string = montante.iloc[linha][0]

    if '.' in string:
        string = string.split('.')
                    
    if len(string[1]) == 1:
        string[1] += '0'
                    
    string = string[0] + string[1]
    string = string.zfill(15)

    string_linha += string
                
    #PEP
    string = pep.iloc[linha][0]
                
    string = string + ' ' * (23 - len(string)) 

    string_linha += string
                
    #Chav. Ref. 1
    string = chave_ref.iloc[linha][0]
                
    if string == 'nan':
        string = ' ' * (12) 
    else:
        string = string + ' ' * (12 - len(string))

    string_linha += string

    #Data Documento
    string = data_doc.iloc[linha][0]

    string = parse(string)
    string = format(string, "%Y%m%d")

    string_linha += string

    #Contrato
    string = contrato.iloc[linha][0]

    string = string.zfill(6)

    string_linha += string

    #Data do lançamento
    string = data_lancamento.iloc[linha][0]

    string = parse(string)
    string = format(string, "%d/%m/%Y")

    string_linha += string
                
    #Histórico(Denominação)
    string = historico.iloc[linha][0]

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
        interface_atual = str(interface.iloc[linha][0])
        interface_prox = str(interface.iloc[linha+1][0])

    if interface_atual != interface_prox:
        txt.write('Do \'Processar\'')
        txt.write('\n')
                
    print('Linha: '+str(linha))

txt.write('Do \'Processar\'')

txt.close()

end_time = time.perf_counter()

print('O tempo de execução foi '+ str(end_time - init_time) + 'segundos')