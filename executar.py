import subprocess
import os
import pyautogui
from pynput.keyboard import Key, Controller as KeyboardController
from pynput.mouse import Button, Controller as MouseController
import time
from pynput.mouse import Listener
import logging
import macrosSAP
import popular33n
import popularLips
import deletar

aux = 0
bases = 0

#Abre o SAP
print("Abrindo SAP...")
os.startfile("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")

time.sleep(5) #aguarda 5 segs

pyautogui.press("return") #pree enter

time.sleep(8) #aguarda 8 segs

print("Baixando bases 33N do SAP...")
macrosSAP.rodarMacro33n() #Chama funções de macrosSAP

#Verifica quantos reports foram gerados pelo sap e armazena na variavel bases
aux = 1
while(aux < 4):
    if(os.path.isfile("C:\Alfred\Bases\\baseSAP" + str(aux) + ".xls")):
        bases = bases + 1
    aux = aux + 1

x = 1

print("Baixando bases lips do SAP...")
while(x < bases+1):
    os.startfile("C:\Alfred\Bases\\baseSAP" + str(x) + ".xls") #abre arquivo excel

    time.sleep(5) #aguarda 5 segs

    #Serie de comandos para simuar teclas do computador
    pyautogui.press("left") #press tecla esquerda
    pyautogui.press("return") #press enter

    #press para baixo 3x
    aux = 0
    while(aux < 3):
        pyautogui.press("down")
        aux = aux + 1

    #press para baixo 22x
    aux = 0
    while(aux < 23):
        pyautogui.press("right")
        aux = aux + 1

    pyautogui.keyDown('shift') #press shift
    pyautogui.hotkey("ctrl", "down") #press ctrl + para baixo
    pyautogui.hotkey("ctrl", "down") #press ctrl + para baixo
    pyautogui.keyUp('shift') #press shift
    pyautogui.hotkey("ctrl", "c") #press ctrl + C

    macrosSAP.rodarMacroLips(x)    #Chama funções de macrosSAP

    x = x + 1

#Fecha arquvos excel aberto
os.system("taskkill /im EXCEL.EXE")
#Confirmação o fechamento
pyautogui.press("right")
pyautogui.press("return")

#Fecha SAP aberto
os.system("taskkill /im saplogon.exe")
os.system("taskkill /im saplogon.exe")

print("Limpando dados antigos 33N...")
deletar.limparAntigos() #Chama funções de deletar

print("Limpando dados antigos Lips...")
deletar.limparAntigosLips() #Chama funções de deletar

print("Populando banco de dados...")
popular33n.carregar33n() #Chama funções de popular33n
 













