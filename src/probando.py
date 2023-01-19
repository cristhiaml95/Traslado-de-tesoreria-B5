import win32com.client

xlApp = win32com.client.Dispatch("Excel.Application")
xlApp.Visible = True
xlApp.Workbooks.Open(r'C:\Users\crist\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\Venado\Cris\Bot5\Cuentas recaudadoras\CUENTAS DE CAJA IVSA.xlsx')
