import win32com.client
import subprocess
import time
from openpyxl import Workbook
from openpyxl import load_workbook
import re
from usefulFunctions import *

#poo desde video 26

class sapInterfaceJob():
    def __init__(self):
        pass

    def startSAP(self, environment):
        global sapGuiAuto, application, connection, session, session1, paths, login, user, psw
        wb = load_workbook(r"C:\Users\crist\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\Venado\Cris\Traslado de tesoreria B5\config.xlsx")
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
        x = f"C:\\Users\\crist\\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\\Venado\\Cris\\Traslado de tesoreria B5\\Cuentas recaudadoras\\{z}\\CUENTAS DE CAJA IVSA.xlsx" 
        y = xlsxFormatting(x)
        wb2 = load_workbook(y)
        ws2 = wb2['CAJAS RECAUDADORAS']
        i = 5
        j = 0

        while True:
            session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
            session.findById("wnd[0]").sendVKey(0)
            accountNumber1 = ws2[f'C{i}'].value
            accountNumberStr1 = str(accountNumber1).replace(' ', '')
            accountNumber2 = ws2[f'D{i}'].value
            accountNumberStr2 = str(accountNumber2).replace(' ', '')
            bank =  ws2[f'E{i}'].value
            bank = str(bank).strip()
            k = 11
            if len(accountNumberStr1)==9 and len(accountNumberStr2)==9 and type(accountNumber1)== int and type(accountNumber2)== int:
                session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = accountNumberStr1
                session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "GV01"
                session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").setFocus
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                a = f'wnd[0]/usr/lbl[9,{k}]'
                nd = f'wnd[0]/usr/lbl[28,{k}]'
                f = f'wnd[0]/usr/lbl[53,{k}]'
                c = f'wnd[0]/usr/lbl[64,{k}]'
                im = f'wnd[0]/usr/lbl[67,{k}]'
                asignacion = session.findById(a).text
                ndoc = session.findById(nd).text
                fecha = session.findById(f).text
                ct = session.findById(c).text
                importe = session.findById(im).text
                per = 7         
                rec = session.findById('wnd[0]/usr/lbl[37,1]').text
                rec = str(rec)
                txtCabDoc = 'TRASLADO A ' + bank
                print(txtCabDoc)
                r2 = re.search('RECAUDADORA', rec).span()
                r2 = r2[1]
                r2+=1
                rec = rec[r2:]
                rec = rec.strip()
                rec = rec.replace(' ', '.')
            
                asignacion = str(asignacion).replace(' ', '')
                ndoc = str(ndoc).replace(' ', '')
                fecha = str(fecha).replace(' ', '')
                l = fecha.index('.')
                fecha = fecha[:l+3]
                ct = str(ct).replace(' ', '')
                importe = str(importe).replace(' ', '')

                texto = 'LP.TRASPASO ' + rec + ' A ' + bank + ' ' + fecha
# PROCESO -------------------------------------------------------------
                session.EndTransaction()

                session.findById("wnd[0]/tbar[0]/okcd").text = "f-02"
                session.findById("wnd[0]").sendVKey(0)

                session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = '30.10.2022'
                session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = '30.10.2022'
                session.findById("wnd[0]/usr/txtBKPF-XBLNR").text = rec
                session.findById("wnd[0]/usr/txtBKPF-BKTXT").text = txtCabDoc
                session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = accountNumberStr2
                session.findById("wnd[0]/usr/txtBKPF-MONAT").text = per
                session.findById("wnd[0]/tbar[0]/btn[0]").press()

                session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importe 
                session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = texto
                session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = '50'
                session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = accountNumberStr1
                session.findById("wnd[0]/tbar[0]/btn[0]").press()

                session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importe 
                session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = texto
                session.findById("wnd[0]/mbar/menu[0]/menu[3]").select()

                validacion = session.findById("wnd[0]/usr/txtRF05A-AZSAL").text
                validacion = str(validacion)
                validacion = validacion.replace(' ', '')
                validacion = validacion.replace('.', '')
                validacion = validacion.replace(',', '.')
                validacion = float(validacion)
                if validacion == 0:
                    print('Validación de saldo 0 correcto')
                else:
                    print(f'ERROR DE VALIDACIÓN DE SALDO 0 EN ASIGNACIÓN: {asignacion}')
                
                try:
                    session.findById("wnd[0]/tbar[0]/btn[11]").press()
                
                except Exception as e:
                    print(f'No se pudo guardar: {e}')
                
                doc = session.findById("wnd[0]/sbar/pane[0]").text
                doc = doc.replace(' ', '')
                doc = doc[4:13]
                print(doc)

                #session.EndTransaction()

                break
                
                


            
    
    


if __name__=='__main__':
    environment= "QAS - EHP8 on HANA"
    bot5SapInterface = sapInterfaceJob()
    bot5SapInterface.startSAP(environment)
    bot5SapInterface.loadBankAccounts()
