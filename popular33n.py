
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import pyodbc
from datetime import datetime, time
import openpyxl as xl
from win32com.client import Dispatch
import win32com.client as wincl
import tkinter
from datetime import date
from tkinter import messagebox
import attCliente
import attTransportadora
import popularLips

def carregar33n():
    #Definir quantas vezes o loop vai rodar baseado no dia do mes
    i = 0

    data_atual = date.today()
    data_em_texto = data_atual.strftime('%d/%m/%Y')
    agora = date.today()

    #Até o dia 10, roda 1 vez
    if(agora.day <= 10):
        i = 2

    #Entre os dias 10 e 21, roda 2 vezes
    if(agora.day > 10 and agora.day < 21):
        i = 3

    #Apartir do dia 21, roda 3 vezes
    if(agora.day >= 21):
        i = 4

    #Variaveis auxiliares
    j = 1
    k = 1
    bases = 0
    aux = 1

    while(aux < i):
        #Verifica quantos arquvivos existem para rodar o FOR
        if(os.path.isfile("C:\Alfred\Bases\\baseSAP" + str(aux) + ".xls")):
            bases = bases + 1
        aux = aux + 1

    for j in range(bases):
        #Tratar a base do SAP
        xl = Dispatch("Excel.Application")
        xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

        #Abre arquivos
        wb1 = xl.Workbooks.Open(Filename="C:\Alfred\Bases\\baseSAP" + str(j + 1) + ".xls")
        wb2 = xl.Workbooks.Open(Filename="C:\Alfred\Macros\TratarSAP.xlsm")

        #Roda a macro
        wb2.Application.Run("TratarSAP.xlsm!Módulo1.Limpar")

        #Copia a planilha do arquivo A no arquivo B
        ws1 = wb1.Worksheets(1)
        ws1.Copy(Before=wb2.Worksheets(1))
        wb1.Close(SaveChanges=False)

        #Roda a macro
        wb2.Application.Run("TratarSAP.xlsm!Módulo1.Executar")

        #Salva e fecha arquivo
        wb2.Close(SaveChanges=True)
        xl.Quit()

    ##########################################################################################################################3
        #Trata a base para ser salva no banco de dados
        xl = Dispatch("Excel.Application")
        xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

        #Abre arquivos
        wb1 = xl.Workbooks.Open(Filename="C:\Alfred\Bases\\baseSAP.xls")
        wb2 = xl.Workbooks.Open(Filename="C:\Alfred\Macros\TratarBase.xlsm")

        #Roda a macro
        wb2.Application.Run("TratarBase.xlsm!Módulo1.Limpar_planilha")

        #Copia a planilha do arquivo A no arquivo B
        ws1 = wb1.Worksheets(1)
        ws1.Copy(Before=wb2.Worksheets(1))
        wb1.Close(SaveChanges=False)

        #Roda a macro
        wb2.Application.Run("TratarBase.xlsm!Módulo1.TratarBase")

        #Salva e fecha arquivo
        wb2.Close(SaveChanges=True)
        xl.Quit()

    ##########################################################################################################################
        #Conecta com o banco de dados
        # conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\henrique.malone\Unilever\Alfred - Documentos\Fretes.mdb;')
        conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\henrique.malone\OneDrive - Unilever\Alfred\Fretes.mdb;')
        cursor = conn.cursor()

        vet = [] #vetor que vai receber o conteudo das linhas

        #Abre arquivo com as informações a serem salvas no banco de dados
        xlsx = pd.ExcelFile('C:/Alfred/BasesTratadas/33NTratada.xlsx')

        #Seleciona a planilha correta
        df = pd.read_excel(xlsx, '33N')

        linhas = df.shape[0] #conta quantas linhas tem a planilha

        for i in range(0, linhas): #percorre linhas
            for j in range(0, 52): #percorre colunas
                vet.append(df.iloc[i, j]) #preencher o vetor com o conteudo da linha 

            cursor.execute("""insert into 33N(dt,
                                            entrega,
                                            empresa,
                                            tipo_transporte,
                                            descricao_transporte,
                                            centro,
                                            descricao_centro,
                                            meso_regiao,
                                            cod_cliente, 
                                            estado,
                                            cidade,
                                            cod_transportadora,
                                            tipo_veiculo,
                                            descricao_veiculo,
                                            identificacao,
                                            distancia,
                                            processo_entrega,
                                            descricao_processo_expedicao,
                                            criacao_dt,
                                            data_embarque,
                                            criacao_doc_custo,
                                            data_fatura,
                                            data_pod,
                                            descricao_tipo_carga,
                                            num_etapa,
                                            status_transporte,
                                            peso_remessa,
                                            unidade_peso,
                                            unidade_distribuicao,
                                            valor_liquido_total_doc_custo,
                                            valor_frete,
                                            pedagio,
                                            valor_cofins,
                                            valor_icms,
                                            valor_iss,
                                            valor_iva,
                                            valor_pis,
                                            valor_etapa1,
                                            valor_retencao,
                                            moeda,
                                            cat_item_custo,
                                            fatura,
                                            status_qtd_faturada,
                                            liberacao_frs,
                                            num_nf,
                                            num_ocorrencia,
                                            num_doc_custo,
                                            valor_fatura,
                                            status_class_contabil,
                                            status_transf,
                                            tipo_agrupamento_etapa,
                                            agrupamento_etapa                                
                                            ) 
                                values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                                int(vet[0]), str(vet[1]), vet[2], vet[3], vet[4], int(vet[5]), vet[6], vet[7], vet[8], vet[9], vet[10],
                                int(vet[11]), vet[12], vet[13], int(vet[14]), int(vet[15]), vet[16], vet[17], vet[18], vet[19], vet[20], 
                                vet[21], vet[22], vet[23], int(vet[24]), int(vet[25]), int(vet[26]), vet[27], int(vet[28]), int(vet[29]), 
                                int(vet[30]), int(vet[31]), int(vet[32]), int(vet[33]), int(vet[34]), int(vet[35]), int(vet[36]), int(vet[37]), 
                                int(vet[38]), vet[39], vet[40], str(vet[41]), vet[42], vet[43], vet[44], int(vet[45]), int(vet[46]), int(vet[47]),
                                vet[48], vet[49], vet[50], vet[51])   
            conn.commit()
            vet.clear()

        #Atualiza lista de clientes
        attCliente.atualizar()

        # #Atualiza lista de transportadoras
        attTransportadora.atualizar()
        
        # #Popular lips
        popularLips.lips(k)
        k = k + 1

        #Remove o arquivo criado baseSAP
        os.remove("C:\\Alfred\\Bases\\baseSAP.xls")




