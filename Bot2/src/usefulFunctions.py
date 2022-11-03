from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
import os

def today():
    fullTime = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=-4)
    currenteDateStr = fullTime.strftime("%d.%m.%Y")
    return currenteDateStr

def xlsxFormatting(x):
    z1 = ''.join([today(), '-F'])
    #z = today()
    #x = f"C:\\Users\\crist\\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\\Venado\\Cris\\Traslado tesorería Bot\\Cuentas recaudadoras\\{z}\\CUENTAS DE CAJA IVSA.xlsx"
    wb1 = load_workbook(x)
    try:
        a = f'C:\\Users\\crist\\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\\Venado\\Cris\\Traslado tesorería Bot\\Cuentas recaudadoras 2\\{z1}'
        os.mkdir(a)
    except Exception as e:
        print(e)
    y = f"C:\\Users\\crist\\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\\Venado\\Cris\\Traslado tesorería Bot\\Cuentas recaudadoras 2\\{z1}\\CUENTAS FORMATEADAS.xlsx"
    wb1.save(y)
    wb2 = load_workbook(y)
    ws2 = wb2['CAJAS RECAUDADORAS']
    mergeRangesList = []
    for i in ws2.merged_cell_ranges:
        i = str(i)
        j = i.index(':')
        #print(i, ' ', j)
        if i[0]==i[j+1] and i[0]=='D' or i[0]==i[j+1]=='E' and i[0]=='E':
            if int(i[1:j])>=3:
                mergeRangesList.append(i)
    #print(mergeRangesList[:])

    for k in mergeRangesList:
        #print(k)
        ws2.unmerge_cells(k)
        l = k.index(':')
        
        a = k[1:l]
        b = k[l+2:]
        #print(a)
        #print(b)
        a = int(a)
        b = int(b)
        a+=1
        b+=1
        
        for m in range(a, b):
            n = ''.join([k[0], str(m)])
            #print( ws2[f'{n}'])
            #print(ws2[k[:l]])
            ws2[f'{n}'] = ws2[k[:l]].value



    wb2.save(y)
    return y

if __name__=='__main__': 
    z = today()
    x = f"C:\\Users\\crist\\OneDrive - UNIVERSIDAD NACIONAL DE INGENIERIA\\Venado\\Cris\\Traslado tesorería Bot\\Cuentas recaudadoras\\{z}\\CUENTAS DE CAJA IVSA.xlsx" 
    xlsxFormatting(x)