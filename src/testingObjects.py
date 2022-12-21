from usefulFunctions import *
from probando import *
import os

class testingObjects():
    def __init__(self):
        self.currentPathParentFolder = getCurrentPath()
        self.currentPathGrandpaFolder = Path(currentPathParentFolder).parent
        self.xlsxForTest = os.path.join(currentPathGrandpaFolder,"TestTemplates")
        self.logPath = os.path.join(self.currentPathParentFolder,"log.txt")

        # self.k = 7
        self.i = 3
        self.j = 0
        self.jMax = 3

    # def testXlsxCreator(self):
    #     y = testingObjects().getFilterBankList()
    #     try:        
    #         os.mkdir(self.xlsxForTest)
    #     except Exception as e:
    #         print('El archivo ya ha sido creado')


    def testXlsxCharger(self):
        x = sapInterfaceJob()
        x.startSAP()
        try:        
            os.mkdir(self.xlsxForTest)
        except Exception as e:
            print('El archivo ya ha sido creado')

        accountsDictionary = self.getAccountsDictionaryOfLists()
        self.wb2.close()
        for bank in accountsDictionary:
            p = 0
            xlsxName = '%s - %s.xlsx' %(bank[-4:], today())
            wb3Path = os.path.join(self.xlsxForTest, xlsxName)
            wb3 = Workbook()
            wb3.save(wb3Path)
            wb3 = load_workbook(wb3Path)
            ws3 = wb3['Sheet']
            ws3.title = bank
            ws3 = wb3[bank]
            for account in accountsDictionary[bank][0]:
                x.getFbl3nMenu()
                x.bank = bank
                try:
                    x.getAccountTableChildren(account)
                    n = accountsDictionary[bank][0].index(account)                
                except Exception as e:
                    writeLog('\n', e, x.logPath)
                    continue
                parametersList = x.getWholeParametersList()
                approvedParametersList = x.wichMigraVerification(parametersList)
                for o in range(len(approvedParametersList[0])):
                    fecha = approvedParametersList[2][o]
                    importe = approvedParametersList[4][o]
                    ws3[f'A{p+o+17}'] = fecha
                    ws3[f'E{p+o+17}'] = importe
                p+=len(approvedParametersList[0]) 

                parametersList = []
                approvedParametersList = []   
                
                wb3.save(wb3Path)
                wb3.close()
        x.proc.kill()
                
    def getBankList(self):

        y =  os.path.join(currentPathGrandpaFolder,"CUENTAS FORMATEADAS.xlsx")
        self.wb2 = load_workbook(y)
        self.ws2 = self.wb2['CAJAS RECAUDADORAS']

        # xlsxCellsRange = []
        xlsxBankList = []

        while True:
            self.accountNumber1 = self.ws2[f'C{self.i}'].value
            self.accountNumberStr1 = str(self.accountNumber1).replace(' ', '')
            self.accountNumber2 = self.ws2[f'D{self.i}'].value
            self.accountNumberStr2 = str(self.accountNumber2).replace(' ', '')
            self.bank =  self.ws2[f'E{self.i}'].value
            self.bank = str(self.bank).strip()

            if len(self.accountNumberStr1)==9 and len(self.accountNumberStr2)==9 and type(self.accountNumber1)== int and type(self.accountNumber2)== int:
                # xlsxCellsRange.append(self.i)
                xlsxBankList.append(self.bank)
                self.i+=1
            else:
                self.i+=1
                self.j+=1
                if self.j > self.jMax:
                    break
                else:
                    continue
        return xlsxBankList

    def getFilterBankList(self):
        bankList = self.getBankList()
        filterBankList = []
        for bank in bankList:
            if bank not in filterBankList:
                filterBankList.append(bank)
        filterBankList.pop(0)
        filterBankList.pop(12)
        # filterBankList = filterBankList.pop(13)
        return filterBankList
    
    def getAccountsDictionaryOfLists(self):
        x = sapInterfaceJob()  
        x.chargeXlsxSheet()
        xlsxRange = x.getExcelRange()
        print('Este es el rango del xlsx: ', xlsxRange)
        bankFilterList = self.getFilterBankList()
        accountDictionary = {}

        for bank in bankFilterList:
            accountDictionary[f'{bank}'] = [[], []]
            for r in xlsxRange:                      
                x.accountNumber1 = x.ws2[f'C{r}'].value
                x.accountNumberStr1 = str(x.accountNumber1).replace(' ', '')
                # x.accountNumber2 = x.ws2[f'D{r}'].value
                # x.accountNumberStr2 = str(x.accountNumber2).replace(' ', '')
                x.bank =  x.ws2[f'E{r}'].value
                x.bank = str(x.bank).strip()
                self.rec =  self.ws2[f'B{r}'].value
                self.rec = str(self.rec)
                r2 = re.search('RECAUDADORA', self.rec).span()
                r2 = r2[1]
                r2+=1
                self.rec = self.rec[r2:]
                self.rec = self.rec.strip()
                self.rec = self.rec.replace(' ', '.')
                
                self.rec = self.rec.replace('AGENCIA', 'AG')
                # self.rec = self.rec.replace('CENTRAL', 'CTL')
                # self.txtCabDoc = 'TRASLADO A ' + self.bank
                
                if x.bank == bank:                   
                    accountDictionary[f'{bank}'][0].append(x.accountNumberStr1)
                    accountDictionary[f'{bank}'][1].append(x.rec)
        return accountDictionary

    # def getAccountTableChildren(self, account):
    #     self.session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = account
    #     self.session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "GV01"
    #     self.session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").setFocus
    #     self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        # self.rec = self.session.findById('wnd[0]/usr/lbl[37,1]').text
        # self.rec = str(self.rec)
        # r2 = re.search('RECAUDADORA', self.rec).span()
        # r2 = r2[1]
        # r2+=1
        # self.rec = self.rec[r2:]
        # self.rec = self.rec.strip()
        # self.rec = self.rec.replace(' ', '.')
        
        # self.rec = self.rec.replace('AGENCIA', 'AG')
        # self.rec = self.rec.replace('CENTRAL', 'CTL')
        # self.txtCabDoc = 'TRASLADO A ' + self.bank



if __name__=='__main__':
    m = testingObjects().testXlsxCharger()