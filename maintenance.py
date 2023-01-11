from tkinter import*
from PIL import Image,ImageTk
from tkinter import ttk
import openpyxl,xlrd
from openpyxl import Workbook
import pathlib
import datetime
import pandas as pd

class MAINTENANCEClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1900x480+0+450")
        self.root.title("MAINTENANCE ENTRY")
        self.root.config(bg='white')
        self.root.focus_force()
        self.no=StringVar()
        LIST=[]
        def submit():
            Trollies=[]
            ARUN={}
            DateTime=[]
            POP={}
            OPO={}
            OUTT=[]
            IN=[]
            loc=(r'REPORT.xls')
            wb = xlrd.open_workbook(loc)
            wb1=xlrd.open_workbook(loc)
            sheet = wb.sheet_by_name('Sheet1')
            sheet.cell_value(0,0)
            sh1 = wb.sheet_by_index(0)
            col = sh1.col_values(0)
            for i in range(sheet.nrows):
                Trollies.append(str(sheet.row_values(i)[0]))
                DateTime.append(str(sheet.row_values(i)[1]))
            Trollies.remove('TROLLEY NUMBER')
            DateTime.remove('DATETIME')
            NU=Trollies[::-1]
            TimeDate=DateTime[::-1]
            for i in range(1, len(NU)):
                ARUN[NU[i]] = TimeDate[i]
            Trollyset=set(Trollies)
            K=len(Trollyset)
            for trolly in Trollies:
                if(Trollies.count(trolly)%2==0):
                    OUTT.append(trolly)
                else:
                    IN.append(trolly)
            for trolly in OUTT:
                if(trolly in ARUN):
                    POP[trolly]=ARUN[trolly]
            for trolly in IN:
                if(trolly in ARUN):
                    OPO[trolly]=ARUN[trolly]

            EXCEL_FILE = r'MAINTENANCE.xls'
            df=pd.read_excel(EXCEL_FILE)
            A={}
            a=self.no.get()
            A["MAINTENANCE TROLLEY NUMBER"]=a
            A["DATE AND TIME"]=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            LIST.append(a)               
            if (len(a)==7 and (a.startswith("DST") or a.startswith("CDT") or a.startswith("BIT") or a.startswith("BST")) and LIST.count(a)<=1 and a not in OPO):
                new_record=pd.DataFrame(A,index=[0])
                df=pd.concat([df,new_record],ignore_index=True)
                df.to_excel(EXCEL_FILE,index=False)
                Label(self.root, text= "        DATA ENTERED!                   \n                                         ", font=("goudy old style",25,"bold"),bg='white',fg="black").place(x=760,y=400)
                txt_search.delete(0, END)
            elif(a in OPO):
                Label(self.root, text= "  TROLLEY IS IN! \n ACCOUNT IT FOR OUT            ", font=("goudy old style",25,"bold"),bg='white',fg="black").place(x=760,y=400)
                txt_search.delete(0, END)
            elif(LIST.count(a)>1 ):
                Label(self.root, text= "TRY AFTER SOME TIME!                   \n                                                 ", font=("goudy old style",25,"bold"),bg='white',fg="black").place(x=760,y=400)
                txt_search.delete(0, END)
            else:
                Label(self.root, text= "     INVALID BARCODE!                   \n                                                                          ", font=("goudy old style",25,"bold"),bg='white',fg="black").place(x=760,y=400)
                txt_search.delete(0, END)

        title=Label(self.root,text="SCAN THE BARCODE IN THE TROLLEY FOR MAINTENANCE",font=("goudy old style",20,"bold"),bg='#0f4d7d',fg="white",cursor="hand2").place(x=50,y=200,width=1800)
        
        SearchFrame=LabelFrame(self.root,text="SCAN TROLLEY",font=("goudy old style",20,"bold"),bd=2,relief=RIDGE,bg='white',fg="black")
        SearchFrame.place(x=600,y=270,width=700,height=100)

        TROLLIESSS=Label(self.root,text=f"TOTAL TROLLIES\n {0}",bd=5,relief=RIDGE,bg="#607d8b",fg="white",font=("goudy old style",25,"bold"))
        TROLLIESSS.place(x=700,y=20,height=150,width=550)

        txt_search=Entry(SearchFrame,textvariable=self.no,font=("goudy old style",20),bg="lightyellow")
        txt_search.place(x=190,y=2)
        
        btn_search=Button(SearchFrame,text="SUBMIT",command=submit,font=("goudy old style",20,"bold"),bg="cyan",fg="black",cursor="hand2").place(x=510,y=1,width=150,height=40)
        root.bind("<Return>",lambda event:submit())

        def info():
            def chec():
                Trollies=[]
                Date=[]
                ARUN={}
                Time=[]
                DateTime=[]
                POP={}
                OPO={}
                OUTT=[]
                IN=[]
                loc=(r'REPORT.xls')
                wb = xlrd.open_workbook(loc)
                wb1=xlrd.open_workbook(loc)
                sheet = wb.sheet_by_name('Sheet1')
                sheet.cell_value(0,0)
                sh1 = wb.sheet_by_index(0)
                col = sh1.col_values(0)
                for i in range(sheet.nrows):
                    Trollies.append(str(sheet.row_values(i)[0]))
                    DateTime.append(str(sheet.row_values(i)[1]))
                Trollies.remove('TROLLEY NUMBER')
                DateTime.remove('DATETIME')
                NU=Trollies[::-1]
                TimeDate=DateTime[::-1]
                for i in range(1, len(NU)):
                    ARUN[NU[i]] = TimeDate[i]
                Trollyset=set(Trollies)
                K=len(Trollyset)

                under={}
                TROLLIES=[]
                AARUN={}
                DATETIME=[]

                LOC=(r'MAINTENANCE.xls')
                wb1=xlrd.open_workbook(LOC)
                Sheet=wb1.sheet_by_name('Sheet1')
                Sheet.cell_value(0,0)
                

                for i in range(Sheet.nrows):
                    TROLLIES.append(str(Sheet.row_values(i)[0]))
                    DATETIME.append(str(Sheet.row_values(i)[1]))
                TROLLIES.remove('MAINTENANCE TROLLEY NUMBER')
                DATETIME.remove('DATE AND TIME')
                NURA=TROLLIES[::-1]
                TIMEDATE=DATETIME[::-1]
                for i in range(1, len(NURA)):
                    AARUN[NURA[i]]=TIMEDATE[i]
                TROLLEYSET=set(TROLLIES)
                
                for trolly in Trollies:
                    if(Trollies.count(trolly)%2==0):
                        OUTT.append(trolly)
                    else:
                        IN.append(trolly)
                OutSET=set(OUTT)
                I=len(IN)
                O=len(OutSET)
                OUT=list(OutSET)

                
                for trolly in OUT:
                    if(trolly in ARUN):
                        POP[trolly]=ARUN[trolly]
                for trolly in IN:
                    if(trolly in ARUN):
                        OPO[trolly]=ARUN[trolly]

                for trolly in AARUN:
                    if trolly in POP:
                        under[trolly]=AARUN[trolly]
                    elif (trolly not in POP) and (trolly not in OPO):
                        under[trolly]=AARUN[trolly]
                    elif trolly in under and trolly in OPO:
                        under.pop(trolly)
                        
                xa=len(under)
                TROLLIESSS.config(text=f"TROLLIES UNDER MAINTENANCE \n {xa}")
                TROLLIESSS.after(1000,chec)
            chec()
        info()

if __name__=="__main__":
    root=Tk()
    obj=MAINTENANCEClass(root)
    root.mainloop()

