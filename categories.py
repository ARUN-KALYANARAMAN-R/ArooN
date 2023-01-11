from tkinter import*
from tkinter import messagebox
from PIL import Image,ImageTk
from tkinter import ttk
import time
import xlrd
import xlwt
from xlwt import Workbook
import openpyxl
class CATEGORIESClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1900x700+0+250")
        self.root.title("CATEGORIES")
        self.root.config(bg="white")
        self.root.focus_force()

        self.lbl_TROLLIES=Label(self.root,text="DRIVER SEAT TROLLIES\n[0]",bd=5,relief=RIDGE,bg="#607d8b",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIES.place(x=20,y=15,height=150,width=600)

        self.lbl_TROLLIESIN=Label(self.root,text="DRIVER SEAT TROLLIES IN\n[0]",bd=5,relief=RIDGE,bg="#33bbf9",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESIN.place(x=650,y=15,height=150,width=600)

        self.lbl_TROLLIESOUT=Label(self.root,text="DRIVER SEAT TROLLIES OUT\n[0]",bd=5,relief=RIDGE,bg="red",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESOUT.place(x=1270,y=15,height=150,width=600)

        self.lbl_TROLLIES1=Label(self.root,text="CO-DRIVER SEAT TROLLIES\n[0]",bd=5,relief=RIDGE,bg="#607d8b",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIES1.place(x=20,y=175,height=150,width=600)

        self.lbl_TROLLIESIN1=Label(self.root,text="CO-DRIVER SEAT TROLLIES IN\n[0]",bd=5,relief=RIDGE,bg="#33bbf9",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESIN1.place(x=650,y=175,height=150,width=600)

        self.lbl_TROLLIESOUT1=Label(self.root,text="CO-DRIVER SEAT TROLLIES OUT\n[0]",bd=5,relief=RIDGE,bg="red",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESOUT1.place(x=1270,y=175,height=150,width=600)

        self.lbl_TROLLIES2=Label(self.root,text="BERTH TROLLIES\n[0]",bd=5,relief=RIDGE,bg="#607d8b",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIES2.place(x=20,y=335,height=150,width=600)

        self.lbl_TROLLIESIN2=Label(self.root,text="BERTH TROLLIES IN\n[0]",bd=5,relief=RIDGE,bg="#33bbf9",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESIN2.place(x=650,y=335,height=150,width=600)

        self.lbl_TROLLIESOUT2=Label(self.root,text="BERTH TROLLIES OUT\n[0]",bd=5,relief=RIDGE,bg="red",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESOUT2.place(x=1270,y=335,height=150,width=600)

        self.lbl_TROLLIES3=Label(self.root,text="BUS SEAT TROLLIES\n[0]",bd=5,relief=RIDGE,bg="#607d8b",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIES3.place(x=20,y=495,height=150,width=600)

        self.lbl_TROLLIESIN3=Label(self.root,text="BUS SEAT TROLLIES IN\n[0]",bd=5,relief=RIDGE,bg="#33bbf9",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESIN3.place(x=650,y=495,height=150,width=600)

        self.lbl_TROLLIESOUT3=Label(self.root,text="BUS SEAT TROLLIES OUT\n[0]",bd=5,relief=RIDGE,bg="red",fg="white",font=("goudy old style",26,"bold"))
        self.lbl_TROLLIESOUT3.place(x=1270,y=495,height=150,width=600)
        self.update_dashboard()

    def update_dashboard(self):
        try:
            OUTT=[]
            OPO={}
            POP={}
            IN=[]
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
            for trolly in OUTT:
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
            self.lbl_TROLLIES.config(text=f"DRIVER SEAT TROLLIES\n {str(q)}")
            self.lbl_TROLLIESIN.config(text=f"DRIVER SEAT TROLLIES IN\n {str(lo)}")   
            self.lbl_TROLLIESOUT.config(text=f"DRIVER SEAT TROLLIES OUT\n {str(xx)}")
            self.lbl_TROLLIES1.config(text=f"CO-DRIVER SEAT TROLLIES\n {str(w)}")
            self.lbl_TROLLIESIN1.config(text=f"CO-DRIVER SEAT TROLLIES IN\n {str(mo)}")   
            self.lbl_TROLLIESOUT1.config(text=f"CO-DRIVER SEAT TROLLIES OUT\n {str(yy)}")
            self.lbl_TROLLIES2.config(text=f"BERTH TROLLIES\n {str(e)}")
            self.lbl_TROLLIESIN2.config(text=f"BERTH TROLLIES IN\n {str(po)}")   
            self.lbl_TROLLIESOUT2.config(text=f"BERTH TROLLIES OUT\n {str(zz)}")
            self.lbl_TROLLIES3.config(text=f"BUS SEAT TROLLIES\n {str(r)}")
            self.lbl_TROLLIESIN3.config(text=f"BUS SEAT TROLLIES IN\n {str(ko)}")   
            self.lbl_TROLLIESOUT3.config(text=f"BUS SEAT TROLLIES OUT\n {str(ww)}")
            self.root.after(500, self.update_dashboard)
        except Exception as ex:
            messagebox.showerror(" ","ERROR")
    
if __name__=="__main__":
    root=Tk()
    obj=CATEGORIESClass(root)
    root.mainloop()
