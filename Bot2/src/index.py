#import os
import win32com.client
#import sys
import subprocess
import time
from openpyxl import Workbook
from openpyxl import load_workbook
from usefulFunctions import *

#poo desde video 26

class sapInterfaceJob():
    def __init__(self):
        pass

    def startSAP(self, environment):
        global sapGuiAuto, application, connection, session, session1, paths, login, user, psw
        wb = load_workbook(r'C:\Users\crist\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\Venado\Cris\Traslado tesorería Bot\config.xlsx')
        ws = wb['config']
        ws1 = wb['sapLogin']

        paths = {'accountPath': ws['B1'].value,
                'SAPPath': ws['B2'].value,
                'migraPath': ws['B3'].value}
        
        # accountPath = str(paths['accountPath'])
        # SAPPath = str(paths['SAPPath'])
        # migraPath = str(paths['migraPath'])

        

        login = {'user': ws1['B1'].value,
                 'psw': ws1['B2'].value}

        # user = str(login['user'])
        # psw = str(login['psw'])
        
       
        #command2 =r"D:\Program Files (x86)\ERPSAP\SAPgui\saplogon.exe"
        proc = subprocess.Popen([paths['SAPPath'], '-new-tab'])
        time.sleep(2)

        sapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(sapGuiAuto) == win32com.client.CDispatch:
            pass

        application = sapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sapGuiAuto = None
            pass

        #connection = application.OpenConnection("SAP QAS", True)
        connection = application.OpenConnection(environment, True)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            sapGuiAuto = None
            pass

        session = connection.Children(0)
        #session1 = connection.Children(1)
       
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sapGuiAuto = None
            pass

        # if not type(session1) == win32com.client.CDispatch:
        #     connection = None
        #     application = None
        #     sapGuiAuto = None
        #     pass

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = login['user']
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = login['psw']
        session.findById("wnd[0]").sendVKey(0)

    def loadBankAccounts(self):

        # session1.findById("wnd[0]/tbar[0]/okcd").text= "f-02"
        # session1.findById("wnd[0]").sendVKey(0)
        z = today()
        x = f"C:\\Users\\crist\\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\\Venado\\Cris\\Traslado tesorería Bot\\Cuentas recaudadoras\\{z}\\CUENTAS DE CAJA IVSA.xlsx" 
        y = xlsxFormatting(x)
        wb2 = load_workbook(y)
        ws2 = wb2['CAJAS RECAUDADORAS']
        i = 5
        j = 0
        k = 7

        while True:
            session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
            session.findById("wnd[0]").sendVKey(0)
            accountNumber1 = ws2[f'C{i}'].value
            accountNumberStr1 = str(accountNumber1).replace(' ', '')
            accountNumber2 = ws2[f'D{i}'].value
            accountNumberStr2 = str(accountNumber2).replace(' ', '')
            if len(accountNumberStr1)==9 and len(accountNumberStr2)==9 and type(accountNumber1)== int and type(accountNumber2)== int:
                session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = accountNumberStr1
                session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "GV01"
                session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").setFocus
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                while True:
                    try:
                        a = f'wnd[0]/usr/lbl[9,{k}]'
                        f = f'wnd[0]/usr/lbl[53,{k}]'
                        c = f'wnd[0]/usr/lbl[64,{k}]'
                        im = f'wnd[0]/usr/lbl[67,{k}]'
                        asignacion = session.findById(a).text
                        fecha = session.findById(f).text
                        ct = session.findById(c).text
                        importe = session.findById(im).text
                    
                        asignacion = str(asignacion).replace(' ', '')
                        fecha = str(fecha).replace(' ', '')
                        ct = str(ct).replace(' ', '')
                        importe = str(importe).replace(' ', '')

                        k+=1
                    except Exception as e:
                        print(e)
                        break

                    match ct:
                        case '40':
                            c2 = f'wnd[0]/usr/lbl[64,{k+1}]'
                            ct2 = session.findById(c2).text
                            ct2 = str(ct2).replace(' ', '')
                            match ct2:
                                case '40':
                                    pass
                                
                                case '50':
                                    pass

                                case _:
                                    pass

                        case '50':
                            pass

                        case _:
                            pass
            else:
                j+=1
                if j >= 2:
                    break
                continue
                


            
    
    


if __name__=='__main__':
    environment= "QAS - EHP8 on HANA"
    bot5SapInterface = sapInterfaceJob()
    bot5SapInterface.startSAP(environment)
    bot5SapInterface.loadBankAccounts()
