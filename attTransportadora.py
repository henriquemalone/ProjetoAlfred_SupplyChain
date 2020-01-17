import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pyodbc
from datetime import datetime, time
import tkinter
from tkinter import messagebox

def atualizar():
    try:
        conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\henrique.malone\OneDrive - Unilever\Alfred\Fretes.mdb;')
        cursor = conn.cursor()

        vet = [] #vetor que vai receber o conteudo das linhas

        xlsx = pd.ExcelFile('C:/Alfred/BasesTratadas/33NTratada.xlsx')

        df = pd.read_excel(xlsx, 'Transportadora')

        linhas = df.shape[0] #conta quantas linhas tem a planilha

        for i in range(0, linhas): #percorre linhas
            for j in range(0, 2): #percorre colunas
                vet.append(df.iloc[i, j]) #preencher o vetor com o conteudo da linha       

            cursor.execute("""Select transportadora_etapa from Transportadora where transportadora_etapa = ?""", int(vet[0]))

            record = cursor.fetchone()

            if not record:
                cursor.execute("""INSERT INTO Transportadora VALUES(?, ?)""", int(vet[0]), vet[1])
            else:
                cursor.execute("""UPDATE Transportadora SET nome = ? WHERE transportadora_etapa = ?""", vet[1], int(vet[0]))

            conn.commit()
            vet.clear()
    except:
        messagebox.showerror("Error", "Erro ao atualizar a base de transportadoras!")