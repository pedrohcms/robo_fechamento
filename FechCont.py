import pandas as pd
import datetime as dt
import os
import time
from xlrd import XLRDError

print('FAVOR ORDERNAR A SEGUNDA E PRIMEIRA COLUNA DA PLANILHA ANTES!')

ColunaOrdenadas = False
ColunaOrdenadasInput = ''

while ColunaOrdenadas == False:
    ColunaOrdenadasInput = input('As Colunas interface e Empresa foram ordenadas de MENOR para MAIOR separadamente nessa mesma ordem? (S - SIM | N - Não)')
    
    if ColunaOrdenadasInput == 'S':
        ColunaOrdenadas = True

DirArq = False
while DirArq == False:
    ArqDir = input('Digite o diretório de onde o arquivo .txt será gravado: \n')
    
    if os.path.isfile(ArqDir):
        DirArq = True
        print('Diretório do arquivo encontrado!\n')
    else:
        print('Diretório do arquivo não existente!\n')
        
#Abrindo o arquivo txt para ser gravado
f = open(ArqDir, "w")

DirExs = False
while DirExs == False:
    #É informado o diretório da planilha
    Diretorio = input('Digite o diretório da planilha (Junto com o nome do arquivo e extensão): \n')

    if os.path.isfile(Diretorio):
        DirExs = True
        print('Diretório do arquivo encontrado!\n')
    else:
        print('Diretório do arquivo não existente!\n')

#Abrindo o arquivo Excel
file1 = pd.ExcelFile(Diretorio)

PlnDir = False
while PlnDir == False:
    #Colocar o nome da planilha
    SheetName = input('Digite o nome da aba da planilha: \n')

    #Lendo o excel da planilha responsável
    try:
        df1 = pd.read_excel(file1, sheet_name = SheetName)
        PlnDir = True
    except XLRDError:
        print('Nome da aba não existente!\n')

QtdLin = len(df1.index) #Quantidade total de Linhas
QtdCol = len(df1.columns) #Quantidade total de colunas

#Ordenação de colunas (Nesse caso não foi usado pois essa ordenação é diferente do excel)
#df1 = df1.nsmallest(QtdLin, ['Interface'])
#df1 = df1.nsmallest(QtdLin, ['Empr'])
#df1 = df1.sort_values('Interface', ascending=True) 
#df1 = df1.sort_values('Empr', ascending=True)

i = 1 #Linha
j = 0 #Coluna
BreakLoop = False

Colunas = [0, 2, 3, 5, 7, 8, 10, 11, 9, 4] #As colunas referentes à montagem

#Horario em que começou
inicio = time.perf_counter()
#TempoPrevisto = QtdLin*0.022
#TempoPrevistoFormat = '{0:.2f}'.format(TempoPrevisto)

#print("\n Tempo de execução previsto: " + str(TempoPrevistoFormat) + " segundos. \n")

print('Processando...')

for i in range(QtdLin):
    if BreakLoop == False:
        for j in range(10):

            Txt = str(df1.iloc[i,Colunas[j]]) #Pegando o valor da célula

            TxtInt = int(df1.iloc[i,[1]]) #Pega o valor da 'Interface'
            df1['Interface'].astype(str).astype(int) #convertendo a coluna  para int

            if j == 0: #Coluna Empresa
                
                if Txt == 'nan' or len(Txt) == 0:
                    print("\nERRO! Coluna 'Empr'(Empresa) com o valor vazio!")
                    BreakLoop = True
                else:
                    TxtEmp = Txt
                    TxtEmpAnt = str(df1.iloc[i-1,Colunas[j]]) #Pega a empresa da linha anterior

                    IntAnt = int(df1.iloc[i-1,[1]]) #Pega a interface anterior

                    if i > 1: #Se não for a primeira linha
                        if TxtInt != IntAnt: #A cada interface diferente é colocado uma linha 'Do 'Processar''
                            f.write("Do 'Processar'\n")
                        elif TxtEmpAnt != TxtEmp: #A cada empresa diferente é colocado uma linha 'Do 'Processar''
                            f.write("Do 'Processar'\n")

                    f.write("&SdtTexto.Add('")

                    #Preencher 0 a esquerda
                    iTxt = 0
                    TxtFlt = 5 - len(TxtEmp)
                    for iTxt in range(TxtFlt):
                        f.write('0')
                        iTxt += 1

                    f.write(Txt)

            if j == 1: #Coluna Debito ou Credito
                if Txt == 'nan' or len(Txt) == 0:
                    print("\nERRO! Coluna 'CL'(Débito/Crédito) com o valor vazio!")
                    BreakLoop = True
                else:
                    f.write(Txt)

            if j == 2: #Coluna Conta
                if Txt == 'nan' or len(Txt) == 0 or len(Txt) < 10:
                    print("\nERRO! Coluna 'Conta' com o valor errado!")
                    BreakLoop = True
                else:
                    f.write(Txt)

            if j == 3: #Coluna Valor do Montante
                if Txt == 'nan' or len(Txt) == 0:
                    print("\nERRO! Coluna 'Valor do Montante' com o valor vazio!")
                    BreakLoop = True
                else:
                    iTxt = 0 #Contador de 0 a esquerda
                    DotCount = 0 #Contador depois da virgula

                    for letter in Txt: #For no texto do valor
                        if DotCount > 0:
                            DotCount += 1 

                        if letter == '.': #caso ache a virgula, começa a contar depois da virgula
                            DotCount += 1
                    
                    if DotCount == 2: #Caso seja 1 casa depois da virgula, adicionar um 0
                        Txt = Txt + '0'
                    
                    TxtRpc = Txt.replace(".", "") #TxtRpc é o valor sem a virgula
                    
                    TxtFlt = 15 - len(TxtRpc) #TxtFlt é a diferença do tamanho do valor com a quantidade de casas que deve ocupar
                        
                    for iTxt in range(TxtFlt):
                        f.write('0') #Preenche com zero à esquerda de acordo com a diferença
                        iTxt += 1
                         
                    f.write(TxtRpc)#Qual valor será colocado

            if j == 4: #Coluna do PEP
                if Txt == 'nan' or len(Txt) == 0 or len(Txt) < 15:
                    print("\nERRO! Coluna 'Elemento PEP' com o valor errado!")
                    BreakLoop = True
                elif len(Txt) < 23:
                    iTxt = 0
                    f.write(Txt)
                    TxtFlt = 23 - len(Txt)
                    for iTxt in range(TxtFlt):
                        f.write(' ')
                        iTxt += 1

            if j == 5: #Coluna Chave Ref
                if len(Txt) < 12 and Txt != 'nan':
                    TxtFlt = 12 - len(Txt)
                    f.write(Txt)
                    tTxt = 0
                    for iTxt in range(TxtFlt):
                        f.write(' ')
                        iTxt += 1
                elif Txt == 'nan': #Se o valor é vazio
                    iTxt = 0
                    for iTxt in range(12):
                        f.write(' ')
                        iTxt += 1
                else:
                    f.write(Txt)

            if j == 6: #Coluna Data do Documento
                if Txt == 'nan' or len(Txt) == 0 or len(Txt) < 19:
                    print("\nERRO! Coluna 'Data do Doc'(Data do Documento) com o valor errado!")
                    BreakLoop = True
                else:
                    TxtDt = dt.datetime.strptime(Txt, '%Y-%m-%d %H:%M:%S').strftime('%Y%m%d') #Convertendo o formato da data
                    f.write(TxtDt)

            if j == 7: #Coluna do Contrato
                if len(Txt) == 0:
                    print("\nERRO! Coluna 'Contrato' com o valor errado!")
                    BreakLoop = True
                else:
                    if len(Txt) < 6:
                        #Caso o valor seja menor que 6 e possa ser adicionado 0 a esquerda 
                        iTxt = 0
                        TxtFlt = 6 - len(Txt)
                        for iTxt in range(TxtFlt):
                            f.write('0')
                        f.write(Txt)
                    else:
                        f.write(Txt)

            if j == 8: #Coluna de Data do Lançamento
                if Txt == 'nan' or len(Txt) == 0 or len(Txt) < 19:
                    print("\nERRO! Coluna 'Data Lançamento'(Data do Lançamento) com o valor errado!")
                    BreakLoop = True
                else:
                    TxtDt = dt.datetime.strptime(Txt, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
                    f.write(TxtDt)

            if j == 9: #Coluna do Histórico
                if Txt == 'nan' or len(Txt) == 0:
                    BreakLoop = True
                else:
                    f.write(Txt)
                    iTxt = 0
                    TxtFlt = 50 - len(Txt)
                    for iTxt in range(TxtFlt):
                        f.write(' ')

                f.write("')")

            j += 1

            if j >= 10: #Quando terminar de verificar todas as colunas
                j = 0
                i += 1
                f.write('\n')

                #Mostrar mensagem enquanto processa
                print('Linha: ' + str(i))
    else:
        break

f.write("Do 'Processar'")   
            
f.close()

#Horario que terminou a execução
fim = time.perf_counter()
tempo = fim - inicio
tempo = '{0:.2f}'.format(tempo)

print("Arquivo gerado em: " + ArqDir + "\n")
print("Tempo de execução: " + str(tempo) + " segundos.")

print("Favor verificar a posição de caracter de cada informação!")