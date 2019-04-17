# -*- coding: utf-8 -*-
"""
Created on Thu Apr 11 14:23:53 2019

@author: Qiao Zhang
"""
 
from tkinter import *
#from tkinter import Label
from tkinter import ttk,BOTH
from tkinter import filedialog

class TKDC:

    def __init__(self, master):
        self.master = master
        self.factor = 1.3
        self.w = 800*self.factor
        self.h = 450*self.factor
        self.ws = self.master.winfo_screenwidth()
        self.hs = self.master.winfo_screenheight()
        self.x = (self.ws/2) - (self.w/2)    
        self.y = (self.hs/2) - (self.h/2)
        self.master.geometry('%dx%d+%d+%d' % (self.w, self.h, self.x, self.y))
        self.frame = Frame(master)
        self.frame.place(relx=.5,rely=.5,anchor="center")
        self.master.title("Data Check Program - CFT")
        self.first_label = Label(master, text ="Hello",font = ("Times New Roman",15))
        self.second_label = Label(master, text ="Please click the button to select your Data Submission Folder",font = ("Times New Roman",15))
        self.third_label = Label(master,text = "...",font =("arial",10))
        self.forth_label = Label(master,text = "...",font =("arial",10))
        self.button = Button(master,text="Select Folder",command=self.clicked,width = 10,height =2,cursor='heart')
        style = ttk.Style()
        self.bar = ttk.Progressbar(master,style='black.Horizontal.TProgressbar',length=300)
        self.bar.config(mode = 'determinate', maximum = 10, value = 0)
        self.run_button = Button(master,text="Run", command = self.runfile,width = 10,height =2)
        self.fifth_label = Label(master,text = "...",font =("arial",10))

        

        # LAYOUT
        self.first_label.pack(padx=10, pady=10) #.grid(row=0, column=4)
        self.second_label.pack(padx=10, pady=10) #grid(row=1, column=4, columnspan=2)
        self.button.pack(padx=10, pady=10) #grid(row=2, column=4, columnspan=2)
        self.bar.pack(padx=10, pady=10)
        self.third_label.pack(padx=10, pady=10) #grid(row=3, column=4, columnspan=2)
        self.forth_label.pack(padx=10, pady=10)
        self.run_button.pack(padx=10, pady=10)
        self.fifth_label.pack(padx=10, pady=10)
        
    def clicked(self):
        global path
        path = ""
        self.third_label.configure(text ="directory path: ...")
        self.forth_label.configure(text ="...")
        self.fifth_label.configure(text ="...")
        self.run_button.configure(text = "Run")
        self.bar.config(mode = 'indeterminate')
        self.bar.start()
        path = filedialog.askdirectory()

        #path = filedialog.askopenfile()
        #print(path.name) 
        self.third_label.configure(text = "Directory Path: " + path)
        self.forth_label.configure(text = "If directory path correct, please click the Run button")
    
    def runfile(self):
        try:
            import data_check
            self.run_button.configure(text = "Running...")
            tem = data_check.data_check(path)
            self.fifth_label.configure(text = "Result Folder: " + tem.result_folder)
            self.bar.stop()
            self.bar.config(mode = 'determinate', maximum = 10, value = 10)
            self.run_button.configure(text = "Done!")
        except:
            self.run_button.configure(text="Please select Data Folder!")
    
root = Tk()
TKDC(root)

root.mainloop()

