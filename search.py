from tkinter import*
from PIL import Image,ImageTk
from tkinter import ttk
from tkinter import messagebox
import time
import xlrd
import xlwt
from xlwt import Workbook
import openpyxl
class SEARCHClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1900x400+0+480")
        self.root.title("TROLLEY INFORMATION")
        self.root.config(bg="white")
        self.root.focus_force()

        self.var_searchtxt=StringVar()
         
        def search():
                aaaa=self.var_searchtxt.get()
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
                for trolly in under:
                    if trolly in IN:
                        under.pop(trolly)
                for trolly in AARUN:
                    if trolly in POP:
                        under[trolly]=AARUN[trolly]
                    elif (trolly not in POP) and (trolly not in OPO):
                        under[trolly]=AARUN[trolly]
                        
                lbl_trolleyid1=Label(self.root,text=aaaa,font=("goudy old style",15),bg="white").place(x=900,y=150)
                
                if(aaaa in IN and len(aaaa)==7):
                    lbl_status2=Label(self.root,text="IN                                                    ",font=("goudy old style",15),bg="white")
                    lbl_status2.place(x=900,y=180)
                    lbl_datetime3=Label(self.root,text=OPO[aaaa],font=("goudy old style",15),bg="white")
                    lbl_datetime3.place(x=900,y=210)
                elif(aaaa in OUT and aaaa not in under and len(aaaa)==7):
                    lbl_status2=Label(self.root,text="OUT                                                       ",font=("goudy old style",15),bg="white")
                    lbl_status2.place(x=900,y=180)
                    lbl_datetime3=Label(self.root,text=POP[aaaa],font=("goudy old style",15),bg="white")
                    lbl_datetime3.place(x=900,y=210)
                elif(aaaa in under and len(aaaa)==7):
                    lbl_status2=Label(self.root,text="MAINTENANCE",font=("goudy old style",15),bg="white")
                    lbl_status2.place(x=900,y=180)
                    lbl_datetime3=Label(self.root,text=under[aaaa],font=("goudy old style",15),bg="white")
                    lbl_datetime3.place(x=900,y=210)
                else:
                    messagebox.showerror("INVALID ENTRY","   NO TROLLEY FOUND  ")
                    txt_search.delete(0, END)

                if(aaaa.startswith("DST")):
                    lbl_type4=Label(self.root,text="DRIVER SEAT TROLLEY          ",font=("goudy old style",15),bg="white")
                    lbl_type4.place(x=900,y=240)
                elif(aaaa.startswith("CDT")):
                    lbl_type4=Label(self.root,text="CO-DRIVER SEAT TROLLEY       ",font=("goudy old style",15),bg="white")
                    lbl_type4.place(x=900,y=240)
                elif(aaaa.startswith("BIT")):
                    lbl_type4=Label(self.root,text="BERTH TROLLEY                ",font=("goudy old style",15),bg="white")
                    lbl_type4.place(x=900,y=240)
                elif(aaaa.startswith("BST")):
                    lbl_type4=Label(self.root,text="BUS SEAT TROLLEY             ",font=("goudy old style",15),bg="white")
                    lbl_type4.place(x=900,y=240)
                else:
                    lbl_type4=Label(self.root,text="UNKNOWN ENTRY                ",font=("goudy old style",15),bg="white")
                    lbl_type4.place(x=900,y=240)
                    txt_search.delete(0, END)
                    
        title=Label(self.root,text="TROLLEY DETAILS",font=("goudy old style",20),bg="#0f4d7d",fg="white",cursor="hand2").place(x=50,y=100,width=1800)
        SearchFrame=LabelFrame(self.root,text="SEARCH TROLLEY",font=("goudy old style",20,"bold"),bd=2,relief=RIDGE,bg="white")
        SearchFrame.place(x=600,y=10,width=600,height=80)
        lbl_trolleyid=Label(self.root,text="TROLLEY NUMBER",font=("goudy old style",15),bg="white").place(x=700,y=150)
        lbl_status=Label(self.root,text="STATUS",font=("goudy old style",15),bg="white").place(x=700,y=180)
        lbl_datetime=Label(self.root,text="DATE AND TIME",font=("goudy old style",15),bg="white").place(x=700,y=210)
        lbl_type=Label(self.root,text="TROLLEY TYPE",font=("goudy old style",15),bg="white").place(x=700,y=240)
        txt_search=Entry(SearchFrame,textvariable=self.var_searchtxt,font=("goudy old style",15),bg="lightyellow")
        txt_search.place(x=200,y=1)
        btn_search=Button(SearchFrame,text="SEARCH",command=search,font=("goudy old style",15),bg="cyan",fg="black",cursor="hand2").place(x=420,y=1,width=150,height=30)

        def clear():
            txt_search.delete(0, END)
            
        btn_clear=Button(self.root,text="CLEAR",command=clear,font=("goudy old style",15),bg="cyan",fg="black",cursor="hand2").place(x=850,y=300,width=170,height=40)
        root.bind("<Return>",lambda event:search())
                
if __name__=="__main__":
    root=Tk()
    obj=SEARCHClass(root)
    root.mainloop()
