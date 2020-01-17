import win32com.client  
from datetime import date
import pyautogui
import time
import openpyxl as xl
import pynput
import os.path

def rodarMacro33n():
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
                datas.append(str(periodo[i]) + "." + str(mes) + "." + str(agora.year))
                if(aux == 2):
                    datas.append(str(periodo[i+4]) + "." + str(mes) + "." + str(agora.year))
                else:
                    datas.append(str(periodo[i+1]) + "." + str(mes) + "." + str(agora.year))
                aux = aux + 1
                i = i + 2

        if(mes == 2):
            while(aux < 3):
                datas.append(str(periodo[i]) + "." + str(mes) + "." + str(agora.year))
                if(aux == 2):
                    datas.append(str(periodo[i+2]) + "." + str(mes) + "." + str(agora.year))
                else:
                    datas.append(str(periodo[i+1]) + "." + str(mes) + "." + str(agora.year))
                aux = aux + 1
                i = i + 2

        if(mes == 4 or mes == 6 or mes == 9  or mes == 11):
            while(aux < 3):
                datas.append(str(periodo[i]) + "." + str(mes) + "." + str(agora.year))
                if(aux == 2):
                    datas.append(str(periodo[i+3]) + "." + str(mes) + "." + str(agora.year))
                else:
                    datas.append(str(periodo[i+1]) + "." + str(mes) + "." + str(agora.year))
                aux = aux + 1
                i = i + 2

    aux = 2
    i = 0

    if(agora.day > 1):
        if (mes == 12):
            mes = 1
        else:
            mes = mes + 1

        datas.append("1" + "." + str(mes) + "." + str(agora.year))
        while(aux <= agora.day):
            datas.append(str(aux) + "." + str(mes) + "." + str(agora.year))
            aux = aux + 1
            i = i + 1

    #Pega ultima posição do vetor
    indice = len(datas) - 1
    ultimo = datas[indice]
    ontem = datas[indice-1]

    aux = 1
    i = 0

    # print(datas)

    if (agora.day > 1 and agora.day <= 10):
        SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
        session = SapGui.FindById("ses[0]")  

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "Z0SWt0033TRN"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "2236"
        session.findById("wnd[0]/usr/ctxtS_TPLST-LOW").text = "BR01"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = datas[0]
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = ontem
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").caretPosition = 9
        session.findById("wnd[0]/usr/btn%_S_SHTYP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "ZA04"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "ZA08"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "ZA09"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "ZA15"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "ZA16"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_BUKRS_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "9311"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_TPLST_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BR02"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Alfred\Bases"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "baseSAP1.xls"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    elif(agora.day > 1):
        SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine 
        session = SapGui.FindById("ses[0]")  

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "Z0SWt0033TRN"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "2236"
        session.findById("wnd[0]/usr/ctxtS_TPLST-LOW").text = "BR01"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = "1.1.2020"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = "10.1.2020"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").caretPosition = 9
        session.findById("wnd[0]/usr/btn%_S_SHTYP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "ZA04"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "ZA08"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "ZA09"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "ZA15"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "ZA16"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_BUKRS_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "9311"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_TPLST_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BR02"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Alfred\Bases"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "baseSAP1.xls"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()

    if(agora.day > 10 and agora.day <= 20):
        SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
        session = SapGui.FindById("ses[0]")  

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "Z0SWt0033TRN"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "2236"
        session.findById("wnd[0]/usr/ctxtS_TPLST-LOW").text = "BR01"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = datas[10]
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = ontem
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").caretPosition = 9
        session.findById("wnd[0]/usr/btn%_S_SHTYP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "ZA04"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "ZA08"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "ZA09"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "ZA15"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "ZA16"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_BUKRS_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "9311"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_TPLST_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BR02"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Alfred\Bases"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "baseSAP2.xls"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    elif(agora.day >= 21):
        SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
        session = SapGui.FindById("ses[0]")  

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "Z0SWt0033TRN"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "2236"
        session.findById("wnd[0]/usr/ctxtS_TPLST-LOW").text = "BR01"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = "11.1.2020"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = "20.1.2020"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").caretPosition = 9
        session.findById("wnd[0]/usr/btn%_S_SHTYP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "ZA04"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "ZA08"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "ZA09"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "ZA15"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "ZA16"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_BUKRS_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "9311"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_TPLST_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BR02"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Alfred\Bases"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "baseSAP2.xls"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()

    if(agora.day > 20):
        SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
        session = SapGui.FindById("ses[0]")  

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "Z0SWt0033TRN"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "2236"
        session.findById("wnd[0]/usr/ctxtS_TPLST-LOW").text = "BR01"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = "21.1.2019"
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = ultimo
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").caretPosition = 9
        session.findById("wnd[0]/usr/btn%_S_SHTYP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "ZA04"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "ZA08"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "ZA09"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "ZA15"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "ZA16"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_BUKRS_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "9311"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/btn%_S_TPLST_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BR02"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Alfred\Bases"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "baseSAP3.xls"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()


def rodarMacroLips(x):
    SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
    session = SapGui.FindById("ses[0]")  

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "z0swt0039"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "LIPS"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
    session.findById("wnd[0]/usr/tblSAPLZ0SWT0033_SE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLZ0SWT0033_SE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&LOAD")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Alfred\Bases"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "lips" + str(x) + ".XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
