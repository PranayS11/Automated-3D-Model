from tkinter import *
from tkinter import font
import openpyxl
import subprocess



root=Tk(className="Rivet Design")
font.families()


root.minsize(1000,450)
root.maxsize(1400,850)



'''-----------------------------------------'''

def calc():
    print("What the hell")
    
    
def show():
    print("display")

    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//op_gui.py",shell=True)
    if(j.get()==0):
        if(p.get()==0):
            if(r.get()==0):
                subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//lap.SLDASM",shell=True)
            elif(r.get()==1):
                subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//lap_2.SLDASM",shell=True)
            else:
                subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//lap_3.SLDASM",shell=True)
        else:
            if(r.get()==1):
                subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//lap_zig_2.SLDASM",shell=True)
            elif(r.get()==2):
                subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//lap_zig_3.SLDASM",shell=True)
    else:
        if(c.get()==0):       
            if(p.get()==0):
                if(r.get()==0):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt.SLDASM",shell=True)
                elif(r.get()==1):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_2.SLDASM",shell=True)
                else:
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_3.SLDASM",shell=True)
            else:
                if(r.get()==1):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_zig_2.SLDASM",shell=True)
                elif(r.get()==2):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_zig_3.SLDASM",shell=True)
        else:       
            if(p.get()==0):
                if(r.get()==0):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_cover.SLDASM",shell=True)
                elif(r.get()==1):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_cover_2.SLDASM",shell=True)
                else:
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_cover_3.SLDASM",shell=True)
            else:
                if(r.get()==1):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_cover_zig_2.SLDASM",shell=True)
                elif(r.get()==2):
                    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//MDPROJ//butt_cover_zig_3.SLDASM",shell=True)
                
                
    
    
def act():
    inp=openpyxl.load_workbook('C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//value.xlsx')
    sheet=inp.active
    print ("know me")
    b2=sheet['B2']
    b2.value=int(t.get())
    b3=sheet['B3']
    b3.value=int(j.get())
    b4=sheet['B4']
    b4.value=int(r.get())
    b5=sheet['B5']
    b5.value=int(p.get())
    b6=sheet['B6']
    b6.value=int(c.get())
    b7=sheet['B7']
    b7.value=op1.get()
    b8=sheet['B8']
    b8.value=op2.get()
    b9=sheet['B9']
    b9.value=float(ts.get())
    b10=sheet['B10']
    b10.value=float(ss.get())
    b11=sheet['B11']
    b11.value=float(cs.get())

    inp.save("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//value.xlsx")
    inp.close()
    
    th= int(t.get())
    print(th)
    print("display")
    
    subprocess.call("C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//value.xlsx",shell=True)
def lap():
    print(j.get())
    rad8.configure(state=DISABLED )
    rad9.configure(state=DISABLED )
    
def butt():
    print(j.get())
    rad8.configure(state=NORMAL )
    rad9.configure(state=NORMAL )
    
    
def single():
    print(r.get())
    
def double():
    print(r.get())
    
def triple():
    print(r.get())
    
def chain():
    print(p.get())
    rad3.configure(state=NORMAL )
    
def zig():
    print(p.get())
    r.set(1)
    rad3.configure(state=DISABLED)
    
    
def one():
    print(c.get())
    
def two():
    print(c.get())

def Matr():

    
    

    plate=op2.get()
    rivet=op1.get()
    if(plate=="Aluminium"):
        plate="al"
    elif(plate=="Aluminium Alloy"):
        plate="ala"
    elif(plate=="Bronze"):
        plate="bronze"
    elif(plate=="Copper"):
        plate="cu"
    elif(plate=="Brass"):
        plate="brass"
    elif(plate=="Mild Steel" or plate=="Stainless Steel"):
        plate="steel"
    
    if(rivet=="Aluminium"):
        rivet="al"
    elif(rivet=="Aluminium Alloy"):
        rivet="ala"
    elif(rivet=="Bronze"):
        rivet="bronze"
    elif(rivet=="Copper"):
        rivet="cu"
    elif(rivet=="Brass"):
        rivet="brass"
    elif(rivet=="Mild Steel" or rivet=="Stainless Steel"):
        rivet="steel"
        
    print(plate+", "+rivet)
    mwb=openpyxl.load_workbook('C://Users//Pps//Desktop//PROJECTMD//mat.xlsx')
    sheet=mwb.active
    w2=sheet['B2']
    w2.value=str(plate+", "+rivet)

    mwb.save("C://Users//Pps//Desktop//PROJECTMD//mat.xlsx")
    mwb.close()
    
    Matp()
    if(op1.get()==option[0]):
        
        cs.set("65")
        ss.set("160")
    elif(op1.get()==option[1]):
       
        cs.set("50.5")
        ss.set("97.5")
    elif(op1.get()==option[2]):
     
        cs.set("85.25")
        ss.set("165")
    elif(op1.get()==option[3]):
        
        cs.set("170.5")
        ss.set("293")
    elif(op1.get()==option[4]):
       
        cs.set("400")
        ss.set("620.4")
    elif(op1.get()==option[5]):
        
        cs.set("356.5")
        ss.set("612")
    elif(op1.get()==option[6]):
     
        cs.set("209.25")
        ss.set("378")
    else:
        cs.set("Mat!!")
        ss.set("Mat!!")
    
        

def Matp():
    if(op2.get()==option[0]):
        ts.set("80")
    
    elif(op2.get()==option[1]):
        ts.set("65")
        
    elif(op2.get()==option[2]):
        ts.set("110")
       
    elif(op2.get()==option[3]):
        ts.set("220")
        
    elif(op2.get()==option[4]):
        ts.set("517")
        
    elif(op2.get()==option[5]):
        ts.set("460")
        
    elif(op2.get()==option[6]):
        ts.set("270")
    else:
        ts.set("Mat!!")
    


'''-----------------------------------------'''


head=Label(root,text="DESIGN OF RIVETED JOINTS OF STRUCTURAL JOINTS",font='Times 20 bold underline',pady=10)
#Heading



t=StringVar()
nam1=Label(root,text="Enter Thickness value :-",padx = 20,pady = 20)
thick=Entry(root,textvariable=t)
th=IntVar()
th= t.get()
v=0

j=IntVar()
joint=Label(root,text="Type Of Joint :-",padx = 20,pady = 20)
rad1=Radiobutton(root, text="Lap",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=j, 
                  command=lap,
                  value=0)
rad2=Radiobutton(root, text="Butt",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=j, 
                  command=butt,
                  value=1)

r=IntVar()
rows=Label(root,text="No. of rows :-",padx = 20,pady = 20)
rad3=Radiobutton(root, text="Single",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=r, 
                  command=single,
                  value=0)



rad4=Radiobutton(root, text="Double",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=r, 
                  command=double,
                  value=1)
rad5=Radiobutton(root, text="Triple",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=r, 
                  command=triple,
                  value=2)

p=IntVar()
patt=Label(root,text="Type Of Pattern :-",padx = 20,pady = 20)
rad6=Radiobutton(root, text="Chain",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=p, 
                  command=chain,
                  value=0)
rad7=Radiobutton(root, text="Zig-Zag",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=p, 
                  command=zig,
                  value=1)

c=IntVar()
cover=Label(root,text="No Of Cover Plates :-",padx = 20,pady = 20)
rad8=Radiobutton(root, text="One",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=c, 
                  command=one,
                  value=0,
                 state=DISABLED )
rad9=Radiobutton(root, text="Two",font='Copperplate 12 bold',
                  padx = 20, 
                  variable=c, 
                  command=two,
                  value=1,
                 state=DISABLED )


op1=StringVar(root)
op1.set("Select Here")
option=["Mild Steel","Aluminium","Aluminium Alloy","Copper","Stainless Steel","Brass","Bronze"]

mat1=Label(root,text="Choose Material for Rivet:-",padx = 20,pady = 20)
matop1= OptionMenu(root, op1, *option)



op2=StringVar(root)
op2.set("Select Here")

mat2=Label(root,text="Choose Material for Plate:-",padx = 20,pady = 20)
matop2= OptionMenu(root, op2, *option)



ts=StringVar()
tens=Label(root,text="Tensile \nStrength(N/mm^2):-",font='Copperplate 12 bold',padx = 20,pady = 20,justify=LEFT)
ten=Entry(root,textvariable=ts,state=DISABLED,width=4)


ss=StringVar()

shea=Label(root,text="Shear \nStrength(N/mm^2):-",font='Copperplate 12 bold',padx = 20,pady = 20,justify=LEFT)
she=Entry(root,textvariable=ss,state=DISABLED,width=4)

cs=StringVar()
crus=Label(root,text="Crushing \nStrength(N/mm^2):-",font='Copperplate 12 bold',padx = 20,pady = 20,justify=LEFT)
cru=Entry(root,textvariable=cs,state=DISABLED,width=4)


cal=Button(root,text="Calculation And \nTable",fg="black",bg="yellow",font='Arial 18 bold',command=act)
#this is to start calculation


mod=Button(root,text="Output And \n3D model",fg="black",bg="cyan",font='Arial 18 bold',command=show)
#connect to os

matent=Button(root,text="Material \nConfirm",fg="Black",bg="Orange",font='Arial 18 bold',command=Matr)




head.grid(row=0,column=1,columnspan=3)
nam1.grid(row=1,column=0)
thick.grid(row=1,column=1)

joint.grid(row=2,column=0)
rad1.grid(row=2,column=1,sticky=W)
rad2.grid(row=2,column=2,sticky=W)

rows.grid(row=3,column=0)
rad3.grid(row=3,column=1,sticky=W)
rad4.grid(row=3,column=2,sticky=W)
rad5.grid(row=3,column=3,sticky=W)

patt.grid(row=4,column=0)
rad6.grid(row=4,column=1,sticky=W)
rad7.grid(row=4,column=2,sticky=W)

cover.grid(row=5,column=0)
rad8.grid(row=5,column=1,sticky=W)
rad9.grid(row=5,column=2,sticky=W)

mat1.grid(row=6,column=0)
matop1.grid(row=6,column=1)

mat2.grid(row=7,column=0)
matop2.grid(row=7,column=1)

tens.grid(row=8,column=0)
ten.grid(row=8,column=0,sticky=E)

shea.grid(row=8,column=1)
she.grid(row=8,column=1,sticky=E)


crus.grid(row=8,column=2)
cru.grid(row=8,column=2,sticky=E)


cal.grid(row=9,column=2,pady=10)
mod.grid(row=9,column=3,pady=10)
matent.grid(row=9,column=1,pady=10)








root.mainloop()
