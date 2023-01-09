import tkinter as tk
import pyautogui as pg

class bot5GUI:
    def __init__(self):
        self.w, self.h = pg.size()

        self.wd = tk.Tk()
        self.header = None
        self.options1Square = None
        self.options2Square = None
        self.options3Square = None
        self.selOp1 = None
        self.selOp2 = None
        self.selOp3 = None
        self.op1_1 = None
        self.op1_2 = None
        self.op1_3 = None
        self.op2_1 = None
        self.op2_2 = None

        self.nextButton = None
        self.selMigraChoise = None

        self.PUNoChoise = None

    def windowDesign(self):
        self.wd.title("BOT 5 - MENU")
        size = f'{int(self.w/5)}x{int(self.h/2.5)}+{int(self.w*2/5)}+{int(self.h*3/10)}'
        self.wd.geometry(size)
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

        self.op1_2 = tk.Radiobutton(self.options1Square, text = 'Solo distribuidoras', bg = 'light sky blue', variable = self.selOp1, value = 2, width=20, anchor='w', command= self.op1_2Command)
        self.op1_2.pack()

        self.op1_3 = tk.Radiobutton(self.options1Square, text = 'Ambas', bg = 'light sky blue', variable = self.selOp1, value = 3, width=20, anchor='w', command= self.op1_3Command)
        self.op1_3.pack()

        self.options2Square = tk.LabelFrame(self.wd, bg='light sky blue', text='Flujo de migracion: ')
        self.options2Square.pack()

        self.selOp2 = tk.IntVar()

        self.op2_1 = tk.Radiobutton(self.options2Square, text = 'Caja - ETV', bg = 'light sky blue', variable = self.selOp2, value = 1, width=20, anchor='w', command= self.op2_1Command)
        self.op2_1['state'] = 'disabled'
        self.op2_1.pack()

        self.op2_2 = tk.Radiobutton(self.options2Square, text = 'ETV - banco', bg = 'light sky blue', variable = self.selOp2, value = 2, width=20, anchor='w', command= self.op2_2Command)
        self.op2_2['state'] = 'disabled'
        self.op2_2.pack()

        self.options3Square = tk.LabelFrame(self.wd, bg='light sky blue', text='¿Realizó la validación manual?')
        self.options3Square.pack()

        self.selOp3 = tk.IntVar()

        self.op3_1 = tk.Radiobutton(self.options3Square, text = 'Sí', bg = 'light sky blue', variable = self.selOp3, value = 1, width=20, anchor='w', command= self.op3_1Command)
        self.op3_1['state'] = 'disabled'
        self.op3_1.pack()

        self.op3_2 = tk.Radiobutton(self.options3Square, text = 'No', bg = 'light sky blue', variable = self.selOp3, value = 2, width=20, anchor='w', command= self.op3_2Command)
        self.op3_2['state'] = 'disabled'
        self.op3_2.pack()

        # self.nextButton = tk.Button(self.wd, text = 'SIEUIENTE', command = self.nextButtonCommand)
        # self.nextButton.pack()

        self.goButton = tk.Button(self.wd, text = 'MIGRAR', command = self.goButtonCommand)
        #self.goButton['state'] = 'disabled'
        self.goButton.pack()

        self.wd.mainloop()


    def noChoise1PU(self):
        self.PUNoChoise = tk.Toplevel(self.wd)
        self.PUNoChoise.title('INFO')
        size = f'{int(self.w/10)}x{int(self.h/10)}+{int(self.w*9/20)}+{int(self.h*9/20)}'
        self.PUNoChoise.geometry(size)
        alert = tk.Label(self.PUNoChoise, text='Debe elegir una opcion.', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUNoChoise, text='Ok', command=self.PUNoChoise.destroy)
        alert.pack()
        okButton.pack()
        self.PUNoChoise.mainloop()

    
    def agConfirmation(self):
        self.PUAgConfirmation = tk.Toplevel(self.wd)
        self.PUAgConfirmation.title('ALERTA')
        size = f'{int(self.w/10)}x{int(self.h/10)}+{int(self.w*9/20)}+{int(self.h*9/20)}'
        self.PUAgConfirmation.geometry(size)
        alert = tk.Label(self.PUAgConfirmation, text='¿Desea continuar?', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUAgConfirmation, text='Ok', command= self.agMigrationProcess)
        cancelButton = tk.Button(self.PUAgConfirmation, text='Cancelar', command=self.PUAgConfirmation.destroy)
        alert.pack()
        okButton.pack()
        cancelButton.pack()
        self.PUAgConfirmation.mainloop()

    def valConfirmation(self):
        self.PUValConfirmation = tk.Toplevel(self.wd)
        self.PUValConfirmation.title('ALERTA')
        size = f'{int(self.w/10)}x{int(self.h/10)}+{int(self.w*9/20)}+{int(self.h*9/20)}'
        self.PUValConfirmation.geometry(size)
        alert = tk.Label(self.PUValConfirmation, text='¿Desea continuar?', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUValConfirmation, text='Ok', command= self.distMigrationProcess)
        cancelButton = tk.Button(self.PUValConfirmation, text='Cancelar', command=self.PUValConfirmation.destroy)
        alert.pack()
        okButton.pack()
        cancelButton.pack()
        self.PUValConfirmation.mainloop()

    def bothConfirmation(self):
        self.PUValConfirmation = tk.Toplevel(self.wd)
        self.PUValConfirmation.title('ALERTA')
        size = f'{int(self.w/10)}x{int(self.h/10)}+{int(self.w*9/20)}+{int(self.h*9/20)}'
        self.PUValConfirmation.geometry(size)
        alert = tk.Label(self.PUValConfirmation, text='¿Desea continuar?', font=('consolas 24 bold', 12))
        okButton = tk.Button(self.PUValConfirmation, text='Ok', command= self.bothMigrationProcess)
        cancelButton = tk.Button(self.PUValConfirmation, text='Cancelar', command=self.PUValConfirmation.destroy)
        alert.pack()
        okButton.pack()
        cancelButton.pack()
        self.PUValConfirmation.mainloop()

   

    def goButtonCommand(self):
        self.selMigraChoise = self.selOp1.get()
        match self.selMigraChoise:
            case 0:
                self.noChoise1PU()                
            case 1:
                self.agConfirmation()
            case 2:     
                self.valConfirmation()
            case 3:
                self.bothConfirmation()            

    def op1_1Command(self):
        self.selOp2.set(0)
        self.op2_1['state'] = 'disabled'
        self.op2_2['state'] = 'disabled'
        self.selOp3.set(0)
        self.op3_1['state'] = 'disabled'
        self.op3_2['state'] = 'disabled'

    def op1_2Command(self):
        self.op2_1['state'] = 'normal'
        self.op2_2['state'] = 'normal'
        self.selOp2.set(0)
        self.op3_1['state'] = 'disabled'
        self.op3_2['state'] = 'disabled'

    def op1_3Command(self):
        self.op2_1['state'] = 'normal'
        self.op2_2['state'] = 'normal'
        self.selOp2.set(0)
        self.op3_1['state'] = 'disabled'
        self.op3_2['state'] = 'disabled'

    def op2_1Command(self):
        self.op3_1['state'] = 'normal'
        self.op3_2['state'] = 'normal'
        self.selOp3.set(0)

    def op2_2Command(self):
        self.op3_1['state'] = 'normal'
        self.op3_2['state'] = 'normal'
        self.selOp3.set(0)

    def op3_1Command(self):
        pass

    def op3_2Command(self):
        pass

    def agMigrationProcess(self):
        self.PUAgConfirmation.destroy()
        pass

    def distMigrationProcess(self):
        self.PUValConfirmation.destroy()
        pass

    def bothMigrationProcess(self):
        pass

    


    def fullGUI(self):
        self.windowDesign()
        self.window1content()

if __name__ == '__main__':
    bot5 = bot5GUI()
    bot5.fullGUI()