from tkiinter import *
from tkinter import ttk
from tkinter import filedialog

class Root(Tk):
  def __init__(self):
    super(Root, self).__init__()
    self.title("Spreadsheet Widget")
    self.minsize(600, 400)
    self.wm_iconbitmap("icon.ico")
      
    self.labelFrame ttk.LabelFrame(self, text = "Open A File")
    self.labelFrame.grid (column = 0, row = 1, padx = 20, pady = 20)
      
    self.button()
      
  def button(self):
    self.button = tk.button(self.labelFrame, text = "Browse A File", command = self.fileDialog)
    self.button.grid(column = 1, row = 1)
    
  def button(self):
    self.button = tk.button(self.labelFrame, text = "Browse A File", command = self.fileDialog)
    self.button.grid(column = 1, row = 2)
    
  def fileDialog(self):
    self.filename = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetype = (("xls", "*.xlsx"), ("ALL FILES", "*.*"))
    self.label = ttk.Label(self.labelFrame, text = "")
    self.label.configure(text = self.filename)
