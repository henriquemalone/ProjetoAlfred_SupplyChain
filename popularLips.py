import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pyodbc
import os
from datetime import datetime, time
import openpyxl as xl
from win32com.client import Dispatch
import win32com.client as wincl
import tkinter
from tkinter import messagebox

def lips(k):
    xl = Dispatch("Excel.Application")
    xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

    #Abre arquivos
    wb1 = xl.Workbooks.Open(Filename="C:\Alfred\BasesTratadas\\33NTratada.xlsx")
    wb2 = xl.Workbooks.Open(Filename="C:\Alfred\Bases\lips" + str(k) + ".XLSX")
    wb3 = xl.Workbooks.Open(Filename="C:\Alfred\Macros\TratarNF.xlsm")

    #Roda macros
    wb3.Application.Run("TratarNF.xlsm!Módulo4.limpar_abas")
    wb3.Application.Run("TratarNF.xlsm!Módulo2.Limpar_Planilha")

    #Copia a planilha do arquivo A no arquivo B
    ws1 = wb1.Worksheets(2)
    ws1.Copy(Before=wb3.Worksheets(1))

    #Copia a planilha do arquivo A no arquivo B
    ws1 = wb2.Worksheets(1)
    ws1.Copy(Before=wb3.Worksheets(1))

    #Roda macro
    wb2.Application.Run("TratarNF.xlsm!Módulo2.Tratar_NF")

    #Salva e fecha arquivo
    wb2.Close(SaveChanges=True)
    xl.Quit()

    ##########################################################################################################################3
    #Conecta com o banco de dados
    # conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\henrique.malone\Unilever\Alfred - Documentos\LipsNf.mdb;')
    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\henrique.malone\OneDrive - Unilever\Alfred\LipsNf2.mdb;')
    cursor = conn.cursor()

    vet = [] #vetor que vai receber o conteudo das linhas

    #Abre arquivo
    xlsx = pd.ExcelFile('C:/Alfred/BasesTratadas/LipsTratada.xlsx')

    #Seleciona a planilha
    df = pd.read_excel(xlsx, 'Sheet1')

    linhas = df.shape[0] #conta quantas linhas tem a planilha

    for i in range(0, linhas): #percorre linhas
        for j in range(0, 14): #percorre colunas
                vet.append(df.iloc[i, j]) #preencher o vetor com o conteudo da linha       

        cursor.execute("""insert into LipsNF(Fornecimento,
                                            Peso_Bruto,
                                            Peso_liquido,
                                            Material,
                                            Produto,
                                            Valor,
                                            Origem,
                                            Cliente,
                                            Cidade,
                                            Estado,
                                            Data_criacao,
                                            Transportadora,
                                            dt,
                                            NF) 
                                        values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", str(vet[0]), float(vet[1]), float(vet[2]), 
                                        str(vet[3]), vet[4], int(vet[5]), vet[6], vet[7], vet[8], vet[9], vet[10], int(vet[11]), int(vet[12]), 
                                        vet[13])
        conn.commit()
        vet.clear()

    #Deleta os arquivos
    os.remove("C:\\Alfred\\BasesTratadas\\33NTratada.xlsx")
    os.remove("C:\\Alfred\\BasesTratadas\\LipsTratada.xlsx")
    os.remove("C:\\Alfred\\Bases\\lips" + str(k) + ".XLSX")
    os.remove("C:\\Alfred\\Bases\\baseSAP" + str(k) + ".xls")