from tkinter import*
from tkinter import messagebox
from PIL import Image,ImageTk
from search import SEARCHClass
from categories import CATEGORIESClass
from scan import SCANClass
from report import REPORTClass
from maintenance import MAINTENANCEClass
import os.path
import time
import xlrd
import xlwt
import pandas
from xlwt import Workbook
import openpyxl
import tkinter as tk
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import threading
from PIL import ImageTk,Image
class IMS:
    def __init__(self,root):
                           
        self.root=root
        self.root.geometry("1600x900+0+0")
        self.root.title("TROLLEY MANAGEMENT SYSTEM")
        self.root.config(bg="white")
 
        title=Label(self.root,text="                                      TROLLEY MANAGEMENT SYSTEM",font=("times new roman",40,"bold"),bg="#010c48",fg="white",anchor="w",padx=20).place(x=0,y=0,relwidth=1,height=70)

        btn_logout=Button(self.root,text="LOGOUT",command=self.root.destroy,font=("times new roman",25,"bold"),bg="yellow",cursor="hand2").place(x=1600,y=10,height=50,width=300)

        def update_date_time(self):
            time_=time.strftime("%H:%M:%S")
            date_=time.strftime("%Y-%m-%d")
            self.lbl_clock=Label(text=f"Date: {str(date_)} \t\t Time: {str(time_)}",font=("times new roman",20,"bold"),bg="#4d636d",fg="white")
            self.lbl_clock.place(x=0,y=70,relwidth=1,height=40)
            self.lbl_clock.after(200,self.update_date_time)
        self.update_date_time()

        self.lbl_TROLLIES=Label(self.root,text="TOTAL TROLLIES\n {0}",bd=5,relief=RIDGE,bg="#607d8b",fg="white",font=("goudy old style",25,"bold"))
        self.lbl_TROLLIES.place(x=85,y=250,height=150,width=550)

        self.lbl_TROLLIESIN=Label(self.root,text="TOTAL TROLLIES IN\n {0}",bd=5,relief=RIDGE,bg="#33bbf9",fg="white",font=("goudy old style",25,"bold"))
        self.lbl_TROLLIESIN.place(x=685,y=250,height=150,width=550)

        self.lbl_TROLLIESOUT=Label(self.root,text="TOTAL TROLLIES OUT\n {0}",bd=5,relief=RIDGE,bg="red",fg="white",font=("goudy old style",25,"bold"))
        self.lbl_TROLLIESOUT.place(x=1285,y=250,height=150,width=550)

        self.update_dashboard()

        btn_SCAN=Button(self.root ,text="SCAN",command=self.SCAN,font=("times new roman",23,"bold"),bg="cyan",bd=3,cursor="hand2").place(x=100,y=150,height=60,width=325)
        btn_SEARCH=Button(self.root ,text="SEARCH",command=self.SEARCH,font=("times new roman",23,"bold"),bg="cyan",bd=3,cursor="hand2").place(x=450,y=150,height=60,width=325)
        btn_CATEGORIES=Button(self.root ,text="CATEGORIES",command=self.CATEGORIES,font=("times new roman",23,"bold"),bg="cyan",bd=3,cursor="hand2").place(x=800,y=150,height=60,width=325)
        btn_REPORT=Button(self.root ,text="REPORT",command=self.REPORT,font=("times new roman",23,"bold"),bg="cyan",bd=3,cursor="hand2").place(x=1150,y=150,height=60,width=325)
        btn_MAINTENANCE=Button(self.root ,text="MAINTENANCE",command=self.MAINTENANCE,font=("times new roman",23,"bold"),bg="cyan",bd=3,cursor="hand2").place(x=1500,y=150,height=60,width=325)
            
        lbl_footer=Label(self.root,text="DICV TROLLEY MANAGEMENT SYSTEM",font=("times new roman",20,"bold"),bg="#4d636d",fg="white").pack(side=BOTTOM,fill=X)

        self.menulogo=Image.open(r"C:\Users\Rajagopal\Desktop\BARCODE SCANNER PROJECT\unominda.jpg")
        self.menulogo=self.menulogo.resize((300,70),Image.ANTIALIAS)
        self.menulogo=ImageTk.PhotoImage(self.menulogo)

        leftmenu=Frame(self.root,bd=2,bg="#010c48")
        leftmenu.place(x=25,y=0,width=300,height=70)

        lbl_menu=Label(leftmenu,image=self.menulogo)
        lbl_menu.pack(side=TOP,fill=BOTH)

        def PIECHART():
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
            loc=(r'REPORT.xls')
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

            explode=(0,0.2)
            fig,(ax1,ax2,ax3,ax4,ax5) = plt.subplots(1,5)
            data1=np.array([K-O,O])
            mylabels = ["IN", "OUT"]
            ax1.set_title("TOTAL TROLLIES")
            data2=np.array([lo,xx])
            ax2.set_title("DRIVER SEAT TROLLIES")
            data3=np.array([mo,yy])
            ax3.set_title("CO-DRIVER SEAT TROLLIES")
            data4=np.array([po,ww])
            ax4.set_title("BERTH TROLLIES")
            data5=np.array([ko,zz])
            ax5.set_title("BUS SEAT TROLLIES")
            def absolutevalue(val):
                a=np.round(val/100.*data1.sum(),0)
                return a
            def absolutevaluee(val):
                a=np.round(val/100.*data2.sum(),0)
                return a
            def absolutevalueee(val):
                a=np.round(val/100.*data3.sum(),0)
                return a
            def absolutevalueeee(val):
                a=np.round(val/100.*data4.sum(),0)
                return a
            def absolutevalueeeee(val):
                a=np.round(val/100.*data5.sum())
                return a
            ax1.pie(data1,labels=mylabels,explode=explode,autopct=absolutevalue)
            ax2.pie(data2,labels=mylabels,explode=explode,autopct=absolutevaluee)
            ax3.pie(data3,labels=mylabels,explode=explode,autopct=absolutevalueee)
            ax4.pie(data4,labels=mylabels,explode=explode,autopct=absolutevalueeee)
            ax5.pie(data5,labels=mylabels,explode=explode,autopct=absolutevalueeeee)
            canvasbar = FigureCanvasTkAgg(fig, master=self.root)
            canvasbar.draw()
            canvasbar.get_tk_widget().place(x=-60,y=401,width=2000,height=500)
        PIECHART()

    def update_date_time(self):
        time_=time.strftime("%H:%M:%S")
        date_=time.strftime("%Y-%m-%d")
        self.lbl_clock=Label(text=f"Date: {str(date_)} \t\t Time: {str(time_)}",font=("times new roman",20,"bold"),bg="#4d636d",fg="white")
        self.lbl_clock.place(x=0,y=70,relwidth=1,height=40)
        self.lbl_clock.after(200,self.update_date_time)

    def update_dashboard(self):
        try:
            OUTT=[]
            IN=[]
            loc=(r'REPORT.xls')
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_name('Sheet1')
            sheet.cell_value(0,0)
            Trollies=[]
            ARUN={}
            DateTime=[]
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
            self.lbl_TROLLIES.config(text=f"TOTAL TROLLIES\n {str(K)}")
            self.lbl_TROLLIESIN.config(text=f"TOTAL TROLLIES IN\n {str(K-O)}")   
            self.lbl_TROLLIESOUT.config(text=f"TOTAL TROLLIES OUT\n {str(O)}")
            self.root.after(500, self.update_dashboard)
        except Exception as ex:
            messagebox.showerror(" ","ERROR")

    def SEARCH(self):
        self.new_win=Toplevel(self.root)
        self.new_obj=SEARCHClass(self.new_win)
    def SCAN(self):
        self.new_win=Toplevel(self.root)
        self.new_obj=SCANClass(self.new_win)
    def CATEGORIES(self):
        self.new_win=Toplevel(self.root)
        self.new_obj=CATEGORIESClass(self.new_win)
    def REPORT(self):
        self.new_win=Toplevel(self.root)
        self.new_obj=REPORTClass(self.new_win)
    def MAINTENANCE(self):
        self.new_win=Toplevel(self.root)
        self.new_obj=MAINTENANCEClass(self.new_win)

    def REPEAT():
        Trollies=[]
        ARUN={}
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
        p=0
        l=0
        LL=0
        PP=0
        OO=0
        II=0
        JJ=0
        DD=0
        FF=0
        KK=0
        BB=0
        DD=0
        INI=[]
        OU=[]
        loc=(r'REPORT.xls')
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
        NU=Trollies
        TimeDate=DateTime
        for i in range(1, len(NU)):
            ARUN[NU[i]] = TimeDate[i]
        Trollyset=set(Trollies)
        K=len(Trollyset)
        for trolly in NU:
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
        for x in POP:
            QQQ.append(x)
            WWW.append(POP[x])
        for x in OPO:
            AAA.append(x)
            SSS.append(OPO[x])
        for x in QQQ:
            if(x.startswith("DST")):
                EEE.append(x)
                UUU.append(POP[x])
            elif(x.startswith("CDT")):
                RRR.append(x)
                III.append(POP[x])
            elif(x.startswith("BIT")):
                TTT.append(x)
                OOO.append(POP[x])
            elif(x.startswith("BST")):
                YYY.append(x)
                PPP.append(POP[x])
        for x in AAA:
            if(x.startswith("DST")):
                DDD.append(x)
                FFF.append(OPO[x])
            elif(x.startswith("CDT")):
                GGG.append(x)
                HHH.append(OPO[x])
            elif(x.startswith("BIT")):
                JJJ.append(x)
                KKK.append(OPO[x])
            elif(x.startswith("BST")):
                LLL.append(x)
                ZZZ.append(OPO[x])      
        wb1=Workbook()
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.DASHED 
        sheet3=wb1.add_sheet('sheet1',cell_overwrite_ok = True)
        style = xlwt.easyxf('font: bold 1')
        sheet3.write(0,0,'TROLLIES OUT',style)
        sheet3.write(0,1,"TIMING",style)
        sheet3.write(0,2,'TROLLIES IN',style)
        sheet3.write(0,3,"TIMING",style)
        sheet3.write(0,4,'DST TROLLIES IN',style)
        sheet3.write(0,5,"TIMING",style)
        sheet3.write(0,6,'DST TROLLIES OUT',style)
        sheet3.write(0,7,"TIMING",style)
        sheet3.write(0,8,'CDT TROLLIES IN',style)
        sheet3.write(0,9,"TIMING",style)
        sheet3.write(0,10,'CDT TROLLIES OUT',style)
        sheet3.write(0,11,"TIMING",style)
        sheet3.write(0,12,'BIT TROLLIES IN',style)
        sheet3.write(0,13,"TIMING",style)
        sheet3.write(0,14,'BIT TROLLIES OUT',style)
        sheet3.write(0,15,"TIMING",style)
        sheet3.write(0,16,'BST TROLLIES IN',style)
        sheet3.write(0,17,"TIMING",style)
        sheet3.write(0,18,'BST TROLLIES OUT',style)
        sheet3.write(0,19,"TIMING",style)
        sheet3.col(0).width = 5500
        sheet3.col(1).width = 5500
        sheet3.col(2).width = 5500
        sheet3.col(3).width = 5500
        sheet3.col(4).width = 5500
        sheet3.col(5).width = 5500
        sheet3.col(6).width = 5500
        sheet3.col(7).width = 5500
        sheet3.col(8).width = 5500
        sheet3.col(9).width = 5500
        sheet3.col(10).width = 5500
        sheet3.col(11).width = 5500
        sheet3.col(12).width = 5500
        sheet3.col(13).width = 5500
        sheet3.col(14).width = 5500
        sheet3.col(15).width = 5500
        sheet3.col(16).width = 5500
        sheet3.col(17).width = 5500
        sheet3.col(18).width = 5500
        sheet3.col(19).width = 5500
        while(O>p):
            sheet3.write(p+1,0,QQQ[p])
            sheet3.write(p+1,1,WWW[p])
            p+=1
        while(len(OPO)>l):
            sheet3.write(l+1,2,AAA[l])
            sheet3.write(l+1,3,SSS[l])
            l+=1
        while(len(DDD)>LL):
            sheet3.write(LL+1,4,DDD[LL])
            sheet3.write(LL+1,5,FFF[LL])
            LL+=1
        while(len(EEE)>PP):
            sheet3.write(PP+1,6,EEE[PP])
            sheet3.write(PP+1,7,UUU[PP])
            PP+=1
        while(len(GGG)>OO):
            sheet3.write(OO+1,8,GGG[OO])
            sheet3.write(OO+1,9,HHH[OO])
            OO+=1
        while(len(RRR)>II):
            sheet3.write(II+1,10,RRR[II])
            sheet3.write(II+1,11,III[II])
            II+=1
        while(len(JJJ)>JJ):
            sheet3.write(JJ+1,12,JJJ[JJ])
            sheet3.write(JJ+1,13,KKK[JJ])
            JJ+=1
        while(len(TTT)>DD):
            sheet3.write(DD+1,14,TTT[DD])
            sheet3.write(DD+1,15,OOO[DD])
            DD+=1
        while(len(LLL)>FF):
            sheet3.write(FF+1,16,LLL[FF])
            sheet3.write(FF+1,17,ZZZ[FF])
            FF+=1
        while(len(YYY)>KK):
            sheet3.write(KK+1,18,YYY[KK])
            sheet3.write(KK+1,19,PPP[KK])
            KK+=1
        file=os.path.exists('DATA.xls')
        if(file==True):
            wb1.save("DATA.xls")
        else:
            wb1.save("DATA.xls")
    REPEAT()
    threading.Timer(5,REPEAT).start()
    

if __name__=="__main__":
    root=Tk()
    obj=IMS(root)
    while True:
        root.update()



