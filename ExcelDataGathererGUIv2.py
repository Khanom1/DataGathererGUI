# -*- coding: utf-8 -*-
"""
Created on Thu Sep 01 10:59:10 2016

@author: rkhanom
"""

from Tkinter import *
from PIL import ImageTk, Image
import tkFileDialog
import os
import getpass
import pandas as pd
import tkMessageBox
from functools import partial
import xlrd 
from xlrd import open_workbook

user = getpass.getuser()

class ImageAnalyzer(Frame):
    def __init__(self, master):
        self.create_widgets()        
  
#-------------------------------------------------------------------------------------------------------------------------      
    def create_widgets(self):
        """Create all widgets!""" 

        self.buttonID1 = 1
        self.button = Button(text = "Select Destination File", command = partial(self.OpenFile, 1))
        self.button.grid(row = 10, column = 100)
        
        self.buttonID2 = 2
        self.button2 = Button(text = "Select Directory of Batch Files", command = partial(self.OpenFile, 2))
        self.button2.pack()
        self.button2.grid(row = 80, column = 100)
        
        self.label1 = Label(text="Worksheet name:")     
        self.label1.pack()
        self.label1.grid(row=160, column=100)
 
        self.label2 = Label(text="Cell(s):")     
        self.label2.pack()
        self.label2.grid(row=240, column=100)
                      
        self.buttonID3 = 4
        self.button3 = Button(text = "OK")
        self.button3["command"] = self.GoGather
        self.button3.pack()
        self.button3.grid(row = 320, column = 100)
        
        self.dest = StringVar()
        self.dest.set(" ")        
        self.entryWidget1 = Entry(textvariable = self.dest)    
        self.entryWidget1["width"] = 50
        self.entryWidget1.pack(side=RIGHT)
        self.entryWidget1.grid(row = 40, column = 100)
        
        self.batch = StringVar()
        self.batch.set(" ")        
        self.entryWidget2 = Entry(textvariable = self.batch)    
        self.entryWidget2["width"] = 50
        self.entryWidget2.pack(side=RIGHT)
        self.entryWidget2.grid(row = 120, column = 100)
        
        #self.batch = StringVar()
        #self.batch.set("")
        self.entryWidgetSheet = Entry()
        self.entryWidgetSheet["width"] = 50
        self.entryWidgetSheet.pack(side=RIGHT)
        self.entryWidgetSheet.grid(row = 180, column = 100)
        
        self.v = StringVar()
        self.v.set("A1")        
        self.entryWidgetCell = Entry(textvariable = self.v)    
        self.entryWidgetCell["width"] = 50
        self.entryWidgetCell.pack(side=RIGHT)
        self.entryWidgetCell.grid(row = 280, column = 100)         
        
#-------------------------------------------------------------------------------------------------------------------------
               
    def OpenFile(self, idofbutton):
        self._buttonID = idofbutton
        
        if self._buttonID == 1:   
            self.filepath = tkFileDialog.askopenfilename(initialdir='C:/Users/%s' % user)
            self.dest.set(self.filepath)
        elif self._buttonID == 2:
            self.dirpath = tkFileDialog.askdirectory(initialdir='C:/Users/%s' % user)
            print(self.dirpath)
            self.batch.set(self.dirpath)
                          
    def GetEntries(self, widget):
        self._entryWidget = widget        
        if self._entryWidget.get().strip() == "":
            tkMessageBox.showerror("TKinter Entry Widget", "Enter a Cell Value or range, e.g. A1:A4")
            print("here")
        else:
            self.CopyCell = self._entryWidget.get().strip()
            print(self.CopyCell)            
            return self.CopyCell              
        
    def GoGather(self):
        #self._filename = self.filepath
        
        self.targetCell = self.GetEntries(self.entryWidgetCell)
        self.targetSheet = self.GetEntries(self.entryWidgetSheet)
        self.destFile = self.GetEntries(self.entryWidget1)
        self.dirpath = self.GetEntries(self.entryWidget2)
        
        self.listOfFiles = next(os.walk(self.dirpath)) [2]
        print(self.listOfFiles)        
        self.list1 = [x.encode('UTF8') for x in self.listOfFiles]
        print(self.list1)
        
        
        for i in [x.encode('UTF8') for x in self.listOfFiles]:  
            self.openTargetBook = open_workbook(self.dirpath + '/' + i)
            self.openTargetSheet = self.openTargetbook.sheet_by_name(self.targetSheet)
            self.value = self.ws.cell(1,2)
            self.openDest
            print(self.value.value)           
            #self.xlDest = pd.ExcelFile(self._filepath)
            #print(self.xlDest.sheet_names)

            
"""place following in a __main__ to create a test case and module out of this file"""        

if __name__ == "__main__":
    root = Tk()
    root.geometry("300x200")
    root.grid()
    
    app = ImageAnalyzer(root)
    root.mainloop()




    