from tkinter import*
from PIL import Image,ImageTk
from tkinter import ttk
import openpyxl,xlrd
from openpyxl import Workbook
import pathlib
import datetime
import pandas as pd
class SCANClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1900x400+0+480")
        self.root.title("ENTER DATA")
        self.root.config(bg='white')
        self.root.focus_force()
        self.no=StringVar()
        LIST=[]
        def submit():
            EXCEL_FILE = r'REPORT.xls'
            df=pd.read_excel(EXCEL_FILE)
            A={}
            a=self.no.get()
            A["TROLLEY NUMBER"]=a
            A["DATETIME"]=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            LIST.append(a)
            if(len(a)==7 and (a.startswith("DST") or a.startswith("CDT") or a.startswith("BIT") or a.startswith("BST")) and LIST.count(a)<=1 ):
                new_record=pd.DataFrame(A,index=[0])
                df=pd.concat([df,new_record],ignore_index=True)
                df.to_excel(EXCEL_FILE, index=False)
                Label(self.root, text= "    DATA ENTERED!        ", font=("goudy old style",25,"bold"),bg='white',fg="black").place(x=790,y=200)
                txt_search.delete(0, END)
            elif(LIST.count(a)>1):
                Label(self.root, text= "TRY AFTER SOME TIME!        ", font=("goudy old style",25,"bold"),bg='white',fg="black").place(x=790,y=200)
                txt_search.delete(0, END)
            else:
                Label(self.root, text= "    INVALID BARCODE!        ", font=("goudy old style",25,"bold"),bg='white',fg="black").place(x=790,y=200)
                txt_search.delete(0, END)

        
        title=Label(self.root,text="SCAN THE BARCODE IN THE TROLLEY",font=("goudy old style",20,"bold"),bg='#0f4d7d',fg="white",cursor="hand2").place(x=50,y=20,width=1800)
        
        SearchFrame=LabelFrame(self.root,text="SCAN TROLLEY",font=("goudy old style",20,"bold"),bd=2,relief=RIDGE,bg='white',fg="black")
        SearchFrame.place(x=600,y=70,width=700,height=100) 
        
        txt_search=Entry(SearchFrame,textvariable=self.no,font=("goudy old style",20),bg="lightyellow")
        txt_search.place(x=190,y=2)
        
        btn_search=Button(SearchFrame,text="SUBMIT",command=submit,font=("goudy old style",20,"bold"),bg="cyan",fg="black",cursor="hand2").place(x=510,y=1,width=150,height=40)
        root.bind("<Return>",lambda event:submit())

if __name__=="__main__":
    root=Tk()
    obj=SCANClass(root)
    root.mainloop()
