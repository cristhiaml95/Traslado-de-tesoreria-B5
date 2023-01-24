from usefulObjets import sapInterfaceJob as sij
import win32com.client
from usefulFunctions import writeLog


class assignmentPaste:
    def __init__(self):
        self.SAP_job = None
        self.xlApp = None
        self.dailyMigrationAccountsPath = None

    def assignmentPaste(self, ETVflow):
        self.SAP_job = sij()
        self.SAP_job.r = None
        self.SAP_job.chargeListOfNames()
        self.SAP_job.startSAP()
        self.SAP_job.chargeXlsxSheet()
        self.SAP_job.ETVflow = ETVflow
        self.SAP_job.tMigracion = 2
        self.SAP_job.ws2 = self.SAP_job.wsDist
                
        self.SAP_job.xlsxRange = self.SAP_job.getExcelRange()
        
        # print('Este es el rango del xls: ', self.SAP_job.xlsxRange)
        for self.SAP_job.r in self.SAP_job.xlsxRange:
            x = self.SAP_job.subProcess_1()
            if x == -1:
                continue
            # mensaje = f'Extrayendo asignaciones de: {self.SAP_job.accountNumberStr1}'
            # writeLog('\n', mensaje, self.SAP_job.logPath)
            assignmentsList = self.SAP_job.approvedParametersList[7]
        
            for i, assignment in enumerate(assignmentsList):
                self.SAP_job.ws2.cell(row = self.SAP_job.r, column = 8+i).value = assignment

        self.SAP_job.wb2.save(self.SAP_job.dailyMigrationAccountsPath)
        self.dailyMigrationAccountsPath = self.SAP_job.dailyMigrationAccountsPath
        self.SAP_job.proc.kill()
        mensaje2 = f'Asignaciones extraidas correctamente: {self.SAP_job.bank} {self.SAP_job.accountNumberStr1}'
        writeLog('\n', mensaje2, self.SAP_job.logPath)

    def openExcel(self):
        path = self.dailyMigrationAccountsPath
        mensaje3 = 'Abriendo excel de cuentas, elija las asignaciones a migrar, luego guarde y cierre el archivo. Y presione "Migrar" en el GUI para continuar el proceso.'
        writeLog('\n', mensaje3, self.SAP_job.logPath)
        self.xlApp = win32com.client.Dispatch("Excel.Application")
        self.xlApp.Visible = True
        self.xlApp.Workbooks.Open(path)
        # self.xlApp.Workbooks(path).Activate()
        


if __name__ == '__main__':
    assignmentPaste().assignmentPaste(2)
    





     

        