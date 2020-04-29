import logging
import os
import traceback
from tkinter import *
from tkinter import filedialog

from aggregator import aggregate, export_data_frame_to_excel, get_data_dir


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
        self.LogBox = Text(self, height = 10, width = 60, state=DISABLED)
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
        try:
            df = aggregate(os.path.join(get_data_dir(), self.FilePathDisplayer.get()))
            export_data_frame_to_excel(df, 'AggregatedLinelist.xlsx')
            self.LogBox.config(state=NORMAL)
            self.LogBox.delete("1.0", END)
            self.LogBox.insert(END, 'SUCCESS : File AggregatedLinelist.xlsx created')
            self.LogBox.config(state=DISABLED)
        except Exception as e:
            logging.exception('Error: ' + str(e))
            self.LogBox.config(state=NORMAL)
            self.LogBox.delete("1.0", END)
            self.LogBox.insert(END, str(e) + '\n')
            self.LogBox.insert(END, ''.join(traceback.format_tb(e.__traceback__)))
            self.LogBox.config(state=DISABLED)
    def __init__(self):
        Tk.__init__(self)
        self.title('Aggregator Tool Dialog')
        self.resizable(width=False, height=False)
        self.iconbitmap(os.path.join(get_data_dir(), '..', 'res', 'msf.ico'))
        self.createWidgets()


if __name__ == '__main__':
    logging.basicConfig(filename='logs.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    DialogObj = MainDialog()
    DialogObj.mainloop()
