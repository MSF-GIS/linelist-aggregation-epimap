import datetime
import logging
import os
import traceback
from tkinter import *
from tkinter import filedialog

from aggregator import aggregate, export_data_frame_to_excel, get_data_dir


class MainDialog(Tk):
    def createWidgets(self):
        self.InputMainLabel = Label(self, text="Input XLSX File Path:")
        self.InputMainLabel.grid(row = 0, column=0)#padx=5, pady=10
        self.ChooseInputFileButton = Button(self,text=' ... ', command= self.chooseInputFileMethod)
        self.ChooseInputFileButton.grid(row = 0, column=3, padx=10, pady=10, ipadx = 10)
        self.iFPSv = StringVar()
        self.iFPSv.trace("w",self.alterButton)
        self.InputFilePathDisplayer = Entry(self, width=75, textvariable=self.iFPSv)
        self.InputFilePathDisplayer.grid(row = 0, column = 1, columnspan = 2)
        self.OutputMainLabel = Label(self, text="Output XLSX File Path:")
        self.OutputMainLabel.grid(row = 1, column=0)
        self.ChooseOutputFileButton = Button(self,text=' ... ', command= self.chooseOutputFileMethod)
        self.ChooseOutputFileButton.grid(row = 1, column=3, padx=10, pady=10, ipadx = 10)
        default_output_filename = 'cartong_epimap_' + datetime.datetime.now().strftime('%Y_%m_%d_%H_%M_%S') +'.xlsx'
        self.oFPSv = StringVar(value=os.path.join(os.getcwd(), default_output_filename))
        self.oFPSv.trace("w",self.alterButton)
        self.OutputFilePathDisplayer = Entry(self, width=75, textvariable=self.oFPSv)
        self.OutputFilePathDisplayer.grid(row = 1, column = 1, columnspan = 2)
        self.CancelDialogButton = Button(self,text=' Cancel ', command = self.quitDialog)
        self.CancelDialogButton.grid(row = 2, column = 1,ipadx = 20)
        self.RunDialogButton = Button(self,text=' Run ',state=DISABLED, command = self.runAggregator)
        self.RunDialogButton.grid(row = 2, column = 2, ipadx = 20)
        self.LogBox = Text(self, height = 10, width = 80, state=DISABLED)
        self.LogBox.grid(row = 3, column = 0, columnspan = 4, padx=10, pady=10)
    def chooseInputFileMethod(self):
        XLSXPath = filedialog.askopenfilename(filetypes= [('Excel worksheet','*.xlsx')])
        self.InputFilePathDisplayer.delete(0, END)
        self.InputFilePathDisplayer.insert(0, XLSXPath)
    def chooseOutputFileMethod(self):
        XLSXPath = filedialog.asksaveasfilename(filetypes= [('Excel worksheet','*.xlsx')])
        self.OutputFilePathDisplayer.delete(0, END)
        self.OutputFilePathDisplayer.insert(0, XLSXPath)
    def alterButton(self,*args):
        if len(self.InputFilePathDisplayer.get())> 0:
            self.RunDialogButton.config(state = NORMAL)
        else:
            self.RunDialogButton.config(state = DISABLED)
    def quitDialog(self):
        self.destroy()
    def runAggregator(self):
        try:
            start_datetime = datetime.datetime.now()
            df = aggregate(os.path.join(get_data_dir(), self.InputFilePathDisplayer.get()))
            export_data_frame_to_excel(df, self.OutputFilePathDisplayer.get())
            self.LogBox.config(state=NORMAL)
            self.LogBox.delete("1.0", END)
            self.LogBox.insert(END, return_session_logs(self.logs_filename, start_datetime))
            self.LogBox.insert(END, f'SUCCESS : File {self.OutputFilePathDisplayer.get()} created')
            self.LogBox.config(state=DISABLED)
            if 'cartong_epimap_' in self.oFPSv.get():
                default_output_filename = 'cartong_epimap_' + datetime.datetime.now().strftime(
                    '%Y_%m_%d_%H_%M_%S') + '.xlsx'
                self.oFPSv.set(os.path.join(os.getcwd(), default_output_filename))
        except Exception as e:
            logging.exception('Error: ' + str(e))
            self.LogBox.config(state=NORMAL)
            self.LogBox.delete("1.0", END)
            self.LogBox.insert(END, return_session_logs(self.logs_filename, start_datetime))
            self.LogBox.config(state=DISABLED)
    def __init__(self, logs_filename):
        Tk.__init__(self)
        self.logs_filename = logs_filename
        self.title('Aggregator Tool Dialog')
        self.resizable(width=False, height=False)
        self.iconbitmap(os.path.join(get_data_dir(), '..', 'res', 'msf.ico'))
        self.createWidgets()


def return_session_logs(logs_filename, start_datetime):
    session_logs = ''
    is_session_log = False
    with open(logs_filename, 'r') as f:
        for line in f.readlines():
            if line.startswith('20'):
                log_datetime = datetime.datetime.strptime(line[:23], '%Y-%m-%d %H:%M:%S,%f')
                is_session_log = (log_datetime >= start_datetime)
                if is_session_log:
                    session_logs += line[34:]
            elif is_session_log:
                session_logs += line
    return session_logs


if __name__ == '__main__':
    logs_filename = 'logs.log'
    logging.basicConfig(filename=logs_filename, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    DialogObj = MainDialog(logs_filename)
    DialogObj.mainloop()
