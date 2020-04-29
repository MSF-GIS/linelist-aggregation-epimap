from tkinter import *
from tkinter import filedialog

class MainDialog(Tk):
    def createWidgets(self):
        self.MainLabel = Label(self, text="XLSX File Path")
        self.MainLabel.grid(row = 0, column=0)#padx=5, pady=10
        self.ChooseFileButton = Button(self,text=' ... ', command= self.chooseFileMethod)
        self.ChooseFileButton.grid(row = 0, column=3, padx=10, pady=10, ipadx = 10)
        self.FPSv = StringVar()
        self.FPSv.trace("w",self.alterButton)
        self.FilePathDisplayer = Entry(self,width=40,textvariable=self.FPSv)
        self.FilePathDisplayer.grid(row = 0, column = 1,columnspan = 2)
        self.CancelDialogButton = Button(self,text=' Cancel ', command = self.quitDialog)
        self.CancelDialogButton.grid(row = 1, column = 1,ipadx = 20)
        self.RunDialogButton = Button(self,text=' Run ',state=DISABLED, command = self.runAggregator)
        self.RunDialogButton.grid(row = 1, column = 2, ipadx = 20)
        self.LogBox = Text(self, height = 10, width = 60)
        self.LogBox.grid(row = 2, column = 0, columnspan = 4, padx=10, pady=10)
    def chooseFileMethod(self):
        XLSXPath = filedialog.askopenfilename(filetypes= [('Excel worksheet','*.xlsx')])
        self.FilePathDisplayer.delete(0,END)
        self.FilePathDisplayer.insert(0,XLSXPath)
    def alterButton(self,*args):
        if len(self.FilePathDisplayer.get())> 0:
            self.RunDialogButton.config(state = NORMAL)
        else:
            self.RunDialogButton.config(state = DISABLED)
    def quitDialog(self):
        self.destroy()
    def runAggregator(self):
        pass
    def __init__(self):
        Tk.__init__(self)
        self.title('Aggregator Tool Dialog')
        self.iconbitmap('D:\GIS\Projets\D2020_MSF_Python\msf.ico')
        self.createWidgets()
        # Run the script

DialogObj = MainDialog()
DialogObj.mainloop()