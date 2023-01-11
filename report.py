from tkinter import*
from tkinter import ttk
import tkinter as tk
from tkcalendar import DateEntry 
import numpy as np
import xlrd
import xlwt
from xlwt import Workbook
import openpyxl
from tkcalendar import Calendar
from datetime import datetime
class REPORTClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1900x700+0+250")
        self.root.title("REPORT GENERATOR")
        self.root.config(bg="white")
        self.root.focus_force()
        Trollies=[]
        Date=[]
        ARUN={}
        Time=[]
        QQQ=[]
        WWW=[]
        EEE=[]
        RRR=[]
        TTT=[]
        YYY=[]
        UUU=[]
        III=[]
        OOO=[]
        PPP=[]
        AAA=[]
        SSS=[]
        DDD=[]
        FFF=[]
        GGG=[]
        HHH=[]
        JJJ=[]
        KKK=[]
        LLL=[]
        ZZZ=[]
        DateTime=[]
        POP={}
        OPO={}
        DS=[]
        BI=[]
        CD=[]
        BS=[]
        DSTOUT=[]
        CDTOUT=[]
        BITOUT=[]
        BSTOUT=[]
        DSTIN=[]
        CDTIN=[]
        BITIN=[]
        BSTIN=[]
        OUTT=[]
        IN=[]
        loc=(r'C:\Users\Rajagopal\Desktop\BARCODE SCANNER PROJECT\REPORT.xls')
        wb = xlrd.open_workbook(loc)
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
        TROLLY=list(Trollyset)
        for Trolly in Trollyset:
            if(Trolly.startswith("DST")):
              DS.append(Trolly)
            elif(Trolly.startswith("CDT")):
              CD.append(Trolly)
            elif(Trolly.startswith("BIT")):
              BI.append(Trolly)
            elif(Trolly.startswith("BST")):
              BS.append(Trolly)
        q=len(DS)
        w=len(CD)
        e=len(BI)
        r=len(BS)
        for trollies in OutSET:
            if(trollies.startswith("DST")):
              DSTOUT.append(trollies)
            elif(trollies.startswith("CDT")):
              CDTOUT.append(trollies)
            elif(trollies.startswith("BST")):
              BSTOUT.append(trollies)
            elif(trollies.startswith("BIT")):
              BITOUT.append(trollies)
            else:
              print("INVALID BARCODE SCANNED")
        for trollies in IN:
            if(trollies.startswith("DST")):
              DSTIN.append(trollies)
            elif(trollies.startswith("CDT")):
              CDTIN.append(trollies)
            elif(trollies.startswith("BST")):
              BSTIN.append(trollies)
            else:
              BITIN.append(trollies)
        xx=len(DSTOUT)
        yy=len(CDTOUT)
        zz=len(BITOUT)
        ww=len(BSTOUT)
        lo=q-xx
        mo=w-yy
        po=e-zz
        ko=r-ww
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
        cal=Calendar(self.root,selectmode='day',year=datetime.now().year,month=datetime.now().month,day=datetime.now().day,locale='en_US', date_pattern='yyyy-MM-dd')
        cal.place(x=30,y=30,width=400,height=400)
        def grad_date():
            date.config(text = "SELECTED DAY IS: "+cal.get_date())
            dt=cal.get_date()
            print(dt)
            List=[]
            for x in ARUN:
                if(ARUN[x].startswith(dt) and x in IN):
                    a=(x,ARUN[x],"IN")
                    List.append(a)
                elif(ARUN[x].startswith(dt) and x in OUT):
                    a=(x,ARUN[x],"OUT")
                    List.append(a)
            for x in under:
                if(under[x].startswith(dt)):
                    a=(x,under[x],"MAINTENANCE")
                    List.append(a)
            style=ttk.Style()
            style.theme_use('clam')
            style.configure('mystyle.Treeview', rowheight=40)
            style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=("goudy old style",20,"bold"))
            style.configure("mystyle.Treeview.Heading", font=("goudy old style",20,"bold"))
            style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])
            treev=ttk.Treeview(self.root,selectmode ='browse',style="mystyle.Treeview",height=30)
            treev.place(x=800,y=30,width=800,height=650)
            verscrlbar = ttk.Scrollbar(treev,orient ="vertical",command = treev.yview)
            verscrlbar.pack(side ='right', fill ='y')
            treev.configure(xscrollcommand = verscrlbar.set)
            treev["columns"] = ("1", "2", "3")
            treev['show'] = 'headings'
            treev.column("1", width = 300, anchor ='c')
            treev.column("2", width = 200, anchor ='c')
            treev.column("3", width = 300, anchor ='c')
            treev.heading("1", text ="TROLLEY NUMBER")
            treev.heading("2", text ="TIMESTAMP")
            treev.heading("3", text ="STATUS")
            for i in range(len(List)):
                treev.insert("", 'end', text ="L{i}",values =List[i])
        Button(self.root, text = "SELECT DATE",command = grad_date,font=("goudy old style",20,"bold"),bg="cyan",fg="black",cursor="hand2").place(x=100,y=470,width=250)
        date = Label(self.root, text = "",font=("goudy old style",20,"bold"),fg="black")
        date.place(x=25,y=550,width=405)
if __name__=="__main__":
    root=Tk()
    obj=REPORTClass(root)
    while True:
        root.update()
