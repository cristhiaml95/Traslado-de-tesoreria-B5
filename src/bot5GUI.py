import tkinter as tk
import pyautogui as pg
from usefulObjets import sapInterfaceJob
import os
from usefulFunctions import currentPathParentFolder, copyANDeraseFile, copyFile, copyANDeraseFile2, asig_ndocToMigra, copyFile2, writeLog
from PIL import ImageTk, Image
import time
from botAssignment import assignmentPaste as ap

class bot5GUI:
    def __init__(self):
        self.w, self.h = pg.size()

        self.wd = tk.Tk()
        self.header = None
        self.options1Square = None
        self.options2Square = None
        self.options3Square = None
        self.pasteMigraSquare = None
        self.selOp1 = None
        self.selOp2 = None
        self.selOp3 = None
        self.op1_1 = None
        self.op1_2 = None
        self.op1_3 = None
        self.op2_1 = None
        self.op2_2 = None

        self.cbOp = None
        self.cbOp_1 = None
        self.pmOp = None
        self.pmOp_1 = None

        self.selMigraChoise = None
        self.dFechaChoise = None
        self.photoPath = None

        self.PUNoChoise = None

    def windowDesign(self):
        self.wd.title("BOT 5 - MENU")
        size = f'{int(self.w/5)}x{int(self.h/1.8)}+{int(self.w*2/5)}+{int(self.h*2/9)}'
        self.wd.geometry(size)
        self.photoPath = os.path.join(currentPathParentFolder, 'Acceso', 'Bot5.ico')
        photo = ImageTk.PhotoImage(Image.open(self.photoPath))
        self.wd.iconbitmap(self.photoPath)
        self.wd.wm_iconphoto(True, photo)
        #self.wd.resizable(0,0)
        self.wd.configure(background='light sky blue')

    def window1content(self):

        self.header = tk.Label(self.wd, text='BIENVENIDO AL BOT 5 \nMIGRACIONES', bg='light sky blue', font=('consolas 24 bold', 15))
        self.header.pack()

        self.options1Square = tk.LabelFrame(self.wd, bg='light sky blue', text='Elija correctamente lo desea migrar: ')
        self.options1Square.pack()

        self.selOp1 = tk.IntVar()
        #selOp1.set(0)

        self.op1_1 = tk.Radiobutton(self.options1Square, text = 'Solo agencias', bg = 'light sky blue', variable = self.selOp1, value = 1, width=20, anchor='w', command= self.op1_1Command)
        self.op1_1.pack()     
       
        self.op1_2 = tk.Radiobutton(self.options1Square, text = 'Solo distribuidoras', bg = 'light sky blue', variable = self.selOp1, value = 2, width=20, anchor='w', command= self.op1_23Command)
        self.op1_2.pack()

        self.op1_3 = tk.Radiobutton(self.options1Square, text = 'Ambas', bg = 'light sky blue', variable = self.selOp1, value = 3, width=20, anchor='w', command= self.op1_23Command)
        self.op1_3.pack()

        self.cbOp = tk.BooleanVar()
        self.cbOp.set(False)

        self.cbOp_1 = tk.Checkbutton(self.options1Square, text = 'Diferencia de fechas: ', bg = 'light sky blue', variable = self.cbOp, width=20, anchor='w', command= self.cbOp_1Command)
        self.cbOp_1['state'] = 'disabled'
        self.cbOp_1.pack()

        self.dFecha = tk.StringVar()
        self.dFecha.set('1')

        self.dFecha_1 = tk.Entry(self.options1Square, width=20)
        self.dFecha_1.insert(0, '1')
        self.dFecha_1['state'] = 'disabled'
        # self.dFecha_1.delete(0, 'end')
        self.dFecha_1.pack()

        self.anyLabel2 = tk.Label(self.options1Square, bg='light sky blue', text='                          ')
        self.anyLabel2.pack()

        self.options2Square = tk.LabelFrame(self.wd, bg='light sky blue', text='              Flujo de migracion:                ')
        self.options2Square.pack()

        self.selOp2 = tk.IntVar()

        self.op2_1 = tk.Radiobutton(self.options2Square, text = 'Distribuidora - ETV', bg = 'light sky blue', variable = self.selOp2, value = 1, width=20, anchor='w', command= self.op2_123Command)
        self.op2_1['state'] = 'disabled'
        self.op2_1.pack()

        self.op2_2 = tk.Radiobutton(self.options2Square, text = 'ETV - banco', bg = 'light sky blue', variable = self.selOp2, value = 2, width=20, anchor='w', command= self.op2_123Command)
        self.op2_2['state'] = 'disabled'
        self.op2_2.pack()

        self.op2_3 = tk.Radiobutton(self.options2Square, text = 'DIRECTO', bg = 'light sky blue', variable = self.selOp2, value = 3, width=20, anchor='w', command= self.op2_123Command)
        self.op2_3['state'] = 'disabled'
        self.op2_3.pack()

        self.getAssignmentsButton = tk.Button(self.wd, text = 'EXTRAER ASIGNACIONES', command = self.getAssignmentsCommand)
        self.getAssignmentsButton['state'] = 'disabled'
        self.getAssignmentsButton.pack()

        self.options3Square = tk.LabelFrame(self.wd, bg='light sky blue', text='     ¿Realizó la validación manual?      ')
        self.options3Square.pack()

        self.selOp3 = tk.IntVar()

        self.op3_1 = tk.Radiobutton(self.options3Square, text = 'Sí', bg = 'light sky blue', variable = self.selOp3, value = 1, width=20, anchor='w', command= self.op3_12Command)
        self.op3_1['state'] = 'disabled'
        self.op3_1.pack()

        self.op3_2 = tk.Radiobutton(self.options3Square, text = 'No', bg = 'light sky blue', variable = self.selOp3, value = 2, width=20, anchor='w', command= self.op3_12Command)
        self.op3_2['state'] = 'disabled'
        self.op3_2.pack()

        # self.nextButton = tk.Button(self.wd, text = 'SIEUIENTE', command = self.nextButtonCommand)
        # self.nextButton.pack()

        self.pasteMigraSquare = tk.LabelFrame(self.wd, bg='light sky blue', text='     Pegar números de documento     ')
        self.pasteMigraSquare.pack()

        self.pmOp = tk.BooleanVar()
        self.pmOp.set(False)

        self.pmOp_1 = tk.Checkbutton(self.pasteMigraSquare, text = 'Habilitar pegado.', bg = 'light sky blue', variable = self.pmOp, width=20, anchor='w', command= self.pmOp_1Command)
        self.pmOp_1.pack()

        self.pmButton = tk.Button(self.pasteMigraSquare, text = 'PEGAR', command = self.pmButtonCommand)
        self.pmButton['state'] = 'disabled'
        self.pmButton.pack()

        self.anyLabel1 = tk.Label(self.pasteMigraSquare, bg='light sky blue', text='                          ')
        self.anyLabel1.pack()

        self.anyLabel3 = tk.Label(self.wd, bg='light sky blue', text='                          ')
        self.anyLabel3.pack()

        self.goButton = tk.Button(self.wd, text = 'MIGRAR', command = self.goButtonCommand)
        self.goButton['state'] = 'disabled'
        self.goButton.pack()

        self.wd.mainloop()


    def noChoise1PU(self):
        self.PUNoChoise = tk.Toplevel(self.wd)
        self.PUNoChoise.title('INFO')
        size = f'{int(self.w/6)}x{int(self.h/8)}+{int(self.w*5/12)}+{int(self.h*9/16)}'
        self.PUNoChoise.geometry(size)
        alert = tk.Label(self.PUNoChoise, text='Debe elegir una opcion.', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUNoChoise, text='   Ok   ', command=self.PUNoChoise.destroy)
        alert.pack()
        okButton.pack()
        self.PUNoChoise.mainloop()

    
    def valConfirmation(self):
        self.PUAgConfirmation = tk.Toplevel(self.wd)
        self.PUAgConfirmation.title('ALERTA')
        size = f'{int(self.w/6)}x{int(self.h/8)}+{int(self.w*5/12)}+{int(self.h*9/16)}'
        self.PUAgConfirmation.geometry(size)
        alert = tk.Label(self.PUAgConfirmation, text='¿Desea continuar?', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUAgConfirmation, text='     Ok      ', command= self.agORdistMigrationProcess)
        cancelButton = tk.Button(self.PUAgConfirmation, text='Cancelar', command=self.PUAgConfirmation.destroy)
        alert.pack()
        okButton.pack()
        cancelButton.pack()
        self.PUAgConfirmation.mainloop()

    def bothConfirmation(self):
        self.PUValConfirmation = tk.Toplevel(self.wd)
        self.PUValConfirmation.title('ALERTA')
        size = f'{int(self.w/6)}x{int(self.h/8)}+{int(self.w*5/12)}+{int(self.h*9/16)}'
        self.PUValConfirmation.geometry(size)
        alert = tk.Label(self.PUValConfirmation, text='¿Desea continuar?', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUValConfirmation, text='     Ok      ', command= self.bothMigrationProcess)
        cancelButton = tk.Button(self.PUValConfirmation, text='Cancelar', command=self.PUValConfirmation.destroy)
        alert.pack()
        okButton.pack()
        cancelButton.pack()
        self.PUValConfirmation.mainloop()

    def goButtonCommand(self):
        if self.dFecha_1.get() == '' or self.cbOp.get() == False:
            self.dFecha.set('1')
        else:
            self.dFecha.set(self.dFecha_1.get())

        self.dFechaChoise = self.dFecha.get()
        self.selMigraChoise = self.selOp1.get()
        match self.selMigraChoise:
            case 0:
                self.noChoise1PU()                
            case 1:
                try:
                    self.dFechaChoise = int(self.dFechaChoise)
                except:
                    self.noChoise1PU()
                self.valConfirmation()
            case 2:
                if self.selOp2.get() == 0 or self.selOp3.get() == 0:
                    self.noChoise1PU()
                else:
                    if self.selOp3.get() == 2:
                        self.wd.destroy()
                    else:     
                        self.valConfirmation()
            case 3:
                try:
                    self.dFechaChoise = int(self.dFechaChoise)
                except:
                    self.noChoise1PU()
                if self.selOp2.get() == 0 or self.selOp3.get() == 0:
                    self.noChoise1PU()
                else:
                    if self.selOp3.get() == 2:
                        self.wd.destroy()
                    else:
                        self.bothConfirmation()            

    def op1_1Command(self):
        self.dFecha.set('1')
        self.dFecha_1.delete(0, tk.END)
        self.dFecha_1.insert(0, '1')
        self.dFecha_1['state'] = 'disabled'
        self.cbOp.set(False)
        
        self.cbOp_1['state'] = 'normal'

        self.selOp2.set(0)
        self.op2_1['state'] = 'disabled'
        self.op2_2['state'] = 'disabled'
        self.op2_3['state'] = 'disabled'
        self.getAssignmentsButton['state'] = 'disabled'
        self.selOp3.set(0)
        self.op3_1['state'] = 'disabled'
        self.op3_2['state'] = 'disabled'
        self.goButton['state'] = 'normal'

    def cbOp_1Command(self):
        if self.cbOp.get():
            self.dFecha_1['state'] = 'normal'
        else:
            self.dFecha_1['state'] = 'disabled'

    def pmOp_1Command(self):
        if self.pmOp.get():
            self.pmButton['state'] = 'normal'
        else:
            self.pmButton['state'] = 'disabled'

    def op1_23Command(self):
        self.dFecha.set('1')
        self.dFecha_1.delete(0, tk.END)
        self.dFecha_1.insert(0, '1')
        self.dFecha_1['state'] = 'disabled'
        self.cbOp.set(False)
        if self.selOp1.get() == 2:
            self.cbOp_1['state'] = 'disabled'
        else:
            self.cbOp_1['state'] = 'normal'
            self.dFecha.set('1')
            self.dFecha_1.delete(0, tk.END)
            self.dFecha_1.insert(0, '1')
            self.dFecha_1['state'] = 'disabled'
            self.cbOp.set(False)

        self.dFecha_1['state'] = 'disabled'
        self.selOp2.set(0)
        self.op2_1['state'] = 'normal'
        self.op2_2['state'] = 'normal'
        self.op2_3['state'] = 'normal'
        self.getAssignmentsButton['state'] = 'disabled'
        self.selOp3.set(0)
        self.op3_1['state'] = 'disabled'
        self.op3_2['state'] = 'disabled'
        self.goButton['state'] = 'disabled'

    def op2_123Command(self):
        self.selOp3.set(0)
        self.op3_1['state'] = 'normal'
        self.op3_2['state'] = 'normal'
        self.goButton['state'] = 'disabled'
        self.getAssignmentsButton['state'] = 'normal'

    def op3_12Command(self):
        # self.goButton['state'] = 'normal'
        pass

    def pasteDone(self):
        self.PUNoChoise = tk.Toplevel(self.wd)
        self.PUNoChoise.title('INFO')
        size = f'{int(self.w/6)}x{int(self.h/8)}+{int(self.w*5/12)}+{int(self.h*9/16)}'
        self.PUNoChoise.geometry(size)
        alert = tk.Label(self.PUNoChoise, text='Pegado terminado.', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUNoChoise, text='   Ok   ', command=self.PUNoChoise.destroy)
        alert.pack()
        okButton.pack()
        self.PUNoChoise.mainloop()

    def getAssignmentsCommand(self):
        self.goButton['state'] = 'disabled'
        assignments = ap()
        assignments.assignmentPaste(self.selOp2.get())
        assignments.openExcel()
        self.goButton['state'] = 'normal'


    def agORdistMigrationProcess(self):
        self.PUAgConfirmation.destroy()
        time.sleep(1)
        self.wd.destroy()
        migra = sapInterfaceJob()
        migra.dFecha = self.dFechaChoise
        migra.process(self.selOp1.get(), self.selOp2.get())
        copyANDeraseFile('logs.txt')
        copyANDeraseFile('CUENTAS DE CAJA IVSA.xlsx')
        copyFile('CUENTAS DE CAJA IVSA.xlsx')

    def bothMigrationProcess(self):
        self.PUValConfirmation.destroy()
        time.sleep(1)
        self.wd.destroy()
        migraAG = sapInterfaceJob()
        migraAG.dFecha = self.dFechaChoise
        migraAG.process(1, 0)
        migraDIS = sapInterfaceJob()
        migraDIS.process(2, self.selOp2.get())
        copyANDeraseFile('logs.txt')
        copyANDeraseFile('CUENTAS DE CAJA IVSA.xlsx')
        copyFile('CUENTAS DE CAJA IVSA.xlsx')

    def pmButtonCommand(self):
        try:
            meanwhileXlsx = os.path.join(currentPathParentFolder, 'Migraciones', 'ASIG-NDOC')
            asig_ndocToMigra(meanwhileXlsx, sapInterfaceJob().logPathMig)
            try:
                copyANDeraseFile2('logs.txt')
            except:
                writeLog('\n', 'No existe archivos logs.txt en carpeta Migraciones', sapInterfaceJob().logPathMig)
            copyANDeraseFile2('ASIG-NDOC.xlsx')
            copyFile2('ASIG-NDOC.xlsx')
            self.pasteDone()        
        except:
            writeLog('\n', 'Error en el pegado, llamar al programador CLLM +51 932446031', sapInterfaceJob().logPathMig)

    def fullGUI(self):
        self.windowDesign()
        self.window1content()

if __name__ == '__main__':
    bot5 = bot5GUI()
    bot5.fullGUI()