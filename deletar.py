from datetime import datetime, time
import openpyxl as xl
from win32com.client import Dispatch
import win32com.client as wincl
import pyodbc
import win32com.client  
from datetime import date
import pyautogui
import time
import openpyxl as xl
import pynput
import os.path

def limparAntigos():
    periodo = [1, 10, 11, 20, 21, 28, 29, 30, 31]
    datas = []
    aux = 0
    i = 0

    data_atual = date.today()
    data_em_texto = data_atual.strftime('%d/%m/%Y')
    agora = date.today()


    #No mes de janeiro MES vai ser igual a 12 (dezembro)
    if(agora.month-1 == 0):
        mes = 12
    else:
        mes = agora.month-1

    if(agora.day == 1):
        #Preenche macro SAP de acordo com os meses janeiro, março, maio, julho, agosto e outubro
        if(mes == 1 or mes == 3 or mes == 5  or mes == 7 or mes == 8 or mes == 10 
            or mes == 12):
            while(aux < 3):
                datas.append(str(periodo[i]) + "/" + str(mes) + "/" + str(agora.year-2))
                if(aux == 2):
                    datas.append(str(periodo[i+4]) + "/" + str(mes) + "/" + str(agora.year-2))
                else:
                    datas.append(str(periodo[i+1]) + "/" + str(mes) + "/" + str(agora.year-2))
                aux = aux + 1
                i = i + 2

        if(mes == 2):
            while(aux < 3):
                datas.append(str(periodo[i]) + "/" + str(mes) + "/" + str(agora.year-2))
                if(aux == 2):
                    datas.append(str(periodo[i+2]) + "/" + str(mes) + "/" + str(agora.year-2))
                else:
                    datas.append(str(periodo[i+1]) + "/" + str(mes) + "/" + str(agora.year-2))
                aux = aux + 1
                i = i + 2

        if(mes == 4 or mes == 6 or mes == 9  or mes == 11):
            while(aux < 3):
                datas.append(str(periodo[i]) + "/" + str(mes) + "/" + str(agora.year-2))
                if(aux == 2):
                    datas.append(str(periodo[i+3]) + "/" + str(mes) + "/" + str(agora.year-2))
                else:
                    datas.append(str(periodo[i+1]) + "/" + str(mes) + "/" + str(agora.year-2))
                aux = aux + 1
                i = i + 2

    aux = 2
    i = 0

    if(agora.day > 1):
        if (mes == 12):
            mes = 1
        else:
            mes = mes + 1

        datas.append("1" + "/" + str(mes) + "/" + str(agora.year-2))
        while(aux <= agora.day):
            datas.append(str(aux) + "/" + str(mes) + "/" + str(agora.year-2))
            aux = aux + 1
            i = i + 1

    #Pega ultima posição do vetor
    indice = len(datas) - 1
    ultimo = datas[indice]

    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\henrique.malone\OneDrive - Unilever\Alfred\Fretes.mdb;')
    cursor = conn.cursor()    

    cursor.execute("""delete * from 33N where criacao_dt >= ? and criacao_dt <= ?""", datas[0], ultimo)
    conn.commit()


def limparAntigosLips():
    periodo = [1, 10, 11, 20, 21, 28, 29, 30, 31]
    datas = []
    aux = 0
    i = 0

    data_atual = date.today()
    data_em_texto = data_atual.strftime('%d/%m/%Y')
    agora = date.today()


    #No mes de janeiro MES vai ser igual a 12 (dezembro)
    if(agora.month-1 == 0):
        mes = 12
    else:
        mes = agora.month-1

    if(agora.day == 1):
        #Preenche macro SAP de acordo com os meses janeiro, março, maio, julho, agosto e outubro
        if(mes == 1 or mes == 3 or mes == 5  or mes == 7 or mes == 8 or mes == 10 
            or mes == 12):
            while(aux < 3):
                datas.append(str(periodo[i]) + "/" + str(mes) + "/" + str(agora.year-2))
                if(aux == 2):
                    datas.append(str(periodo[i+4]) + "/" + str(mes) + "/" + str(agora.year-2))
                else:
                    datas.append(str(periodo[i+1]) + "/" + str(mes) + "/" + str(agora.year-2))
                aux = aux + 1
                i = i + 2

        if(mes == 2):
            while(aux < 3):
                datas.append(str(periodo[i]) + "/" + str(mes) + "/" + str(agora.year-2))
                if(aux == 2):
                    datas.append(str(periodo[i+2]) + "/" + str(mes) + "/" + str(agora.year-2))
                else:
                    datas.append(str(periodo[i+1]) + "/" + str(mes) + "/" + str(agora.year-2))
                aux = aux + 1
                i = i + 2

        if(mes == 4 or mes == 6 or mes == 9  or mes == 11):
            while(aux < 3):
                datas.append(str(periodo[i]) + "/" + str(mes) + "/" + str(agora.year-2))
                if(aux == 2):
                    datas.append(str(periodo[i+3]) + "/" + str(mes) + "/" + str(agora.year-2))
                else:
                    datas.append(str(periodo[i+1]) + "/" + str(mes) + "/" + str(agora.year-2))
                aux = aux + 1
                i = i + 2

    aux = 2
    i = 0

    if(agora.day > 1):
        if (mes == 12):
            mes = 1
        else:
            mes = mes + 1

        datas.append("1" + "/" + str(mes) + "/" + str(agora.year-2))
        while(aux <= agora.day):
            datas.append(str(aux) + "/" + str(mes) + "/" + str(agora.year-2))
            aux = aux + 1
            i = i + 1

    #Pega ultima posição do vetor
    indice = len(datas) - 1
    ultimo = datas[indice]

    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\henrique.malone\OneDrive - Unilever\Alfred\\LipsNf2.mdb;')
    cursor = conn.cursor()    

    cursor.execute("""delete * from LipsNF where Data_criacao >= ? and Data_criacao <= ?""", datas[0], ultimo)
    conn.commit()