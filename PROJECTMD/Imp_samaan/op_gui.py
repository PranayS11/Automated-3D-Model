from tkinter import *
import openpyxl 

root1=Tk(className="Rivet op Design")
root1.minsize(1000,450)
root1.maxsize(1400,850)

d=StringVar()

out=openpyxl.load_workbook('C://Users//Pps//Desktop//PROJECTMD//Imp_samaan//value.xlsx',read_only=True,data_only=True)
sheet=out.active



'''-----------------------------------------'''

head=Label(root1,text="OUTPUT DIMENSIONS",bg="yellow",font='Times 20 bold underline',)
#Heading




diam=Label(root1,text="Diameter(mm):-",padx = 20,pady = 20,justify=LEFT)
dia=Entry(root1,textvariable=d,state=DISABLED,width=8)

p=StringVar()
pitch=Label(root1,text="Pitch(mm):-",padx = 20,pady = 20,justify=LEFT)
pit=Entry(root1,textvariable=p,state=DISABLED,width=8)

l=StringVar()
lenoriv=Label(root1,text="Length of Rivet(mm):-",padx = 20,pady = 20,justify=LEFT)
lor=Entry(root1,textvariable=l,state=DISABLED,width=6)


dh=StringVar()
diamho=Label(root1,text="Diameter of Hole(mm):-",padx = 20,pady = 20,justify=LEFT)
diah=Entry(root1,textvariable=dh,state=DISABLED,width=6)

sr=StringVar()
shearrestp=Label(root1,text="Shearing resistance/pitch(N/mm):-",padx = 20,pady = 20,justify=LEFT)
srp=Entry(root1,textvariable=sr,state=DISABLED,width=10)

tr=StringVar()
tearrestp=Label(root1,text="Tearing resistance/pitch(N/mm):-",padx = 20,pady = 20,justify=LEFT)
trp=Entry(root1,textvariable=tr,state=DISABLED,width=10)

cr=StringVar()
crushrestp=Label(root1,text="Crushing resistance/pitch(N/mm):-",padx = 20,pady = 20,justify=LEFT)
crp=Entry(root1,textvariable=cr,state=DISABLED,width=10)

se=StringVar()
sheareff=Label(root1,text="Shearing Efficiency(%):-",padx = 20,pady = 20,fg='blue',justify=LEFT,font='Arial 14 bold')
shef=Entry(root1,textvariable=se,state=DISABLED,width=4)

ce=StringVar()
crusheff=Label(root1,text="Crushing Efficiency(%):-",padx = 20,pady = 20,fg='blue',justify=LEFT,font='Arial 14 bold')
cref=Entry(root1,textvariable=ce,state=DISABLED,width=6)

te=StringVar()
teareff=Label(root1,text="Tearing Efficiency(%):-",padx = 20,pady = 20,fg='blue',justify=LEFT,font='Arial 14 bold')
teef=Entry(root1,textvariable=te,state=DISABLED,width=6)

je=StringVar()
jointeff=Label(root1,text="Joint Efficiency(%):-",padx = 20,fg='blue',pady = 20,justify=LEFT,font='Arial 14 bold')
joef=Entry(root1,textvariable=je,state=DISABLED,width=6)

sd=StringVar()
stdiam=Label(root1,text="Std. Diameter(mm):-",padx = 20,pady = 20,justify=LEFT)
std=Entry(root1,textvariable=sd,state=DISABLED,width=6)

sp=StringVar()
stpitch=Label(root1,text="Std. Pitch(mm):-",padx = 20,pady = 20,justify=LEFT)
stp=Entry(root1,textvariable=sp,state=DISABLED,width=6)

sl=StringVar()
stlenoriv=Label(root1,text="Std. Length of Rivet (mm):-",padx = 20,pady = 20,justify=LEFT)
stl=Entry(root1,textvariable=sl,state=DISABLED,width=6)


st=StringVar()
strapth=Label(root1,text="Strap Thickness(mm):-",padx = 20,pady = 20,justify=LEFT)
stt=Entry(root1,textvariable=st,state=DISABLED,width=6)

'''-----------------------------------------------------------------------------------------------------'''

head.grid(row=0,column=1,columnspan=3)

diam.grid(row=1,column=0,sticky=W)
dia.grid(row=1,column=1,sticky=E)

pitch.grid(row=2,column=0,sticky=W)
pit.grid(row=2,column=1,sticky=E)

lenoriv.grid(row=3,column=0,sticky=W)
lor.grid(row=3,column=1,sticky=E)

diamho.grid(row=4,column=0,sticky=W)
diah.grid(row=4,column=1,sticky=E)

shearrestp.grid(row=5,column=0,sticky=W)
srp.grid(row=5,column=1,sticky=E)

tearrestp.grid(row=5,column=2,sticky=W)
trp.grid(row=5,column=3,sticky=E)

crushrestp.grid(row=5,column=4,sticky=W)
crp.grid(row=5,column=5,sticky=E)

sheareff.grid(row=6,column=0,sticky=W)
shef.grid(row=6,column=1,sticky=E)

teareff.grid(row=6,column=2,sticky=W)
teef.grid(row=6,column=3,sticky=E)

crusheff.grid(row=6,column=4,sticky=W)
cref.grid(row=6,column=5,sticky=E)

jointeff.grid(row=7,column=2,sticky=W)
joef.grid(row=7,column=3,sticky=E)

stdiam.grid(row=1,column=4,sticky=W)
std.grid(row=1,column=5,sticky=E)

stpitch.grid(row=2,column=4,sticky=W)
stp.grid(row=2,column=5,sticky=E)

stlenoriv.grid(row=3,column=4,sticky=W)
stl.grid(row=3,column=5,sticky=E)

strapth.grid(row=4,column=4,sticky=W)
stt.grid(row=4,column=5,sticky=E)




#--------------------------------------------------







#---------------------------------------------------


d2=sheet['d2']
print(d2.value)
d.set(d2.value)

d3=sheet['d3']
print(d3.value)
p.set(str(d3.value))

d4=sheet['d4']
print(d4.value)
sd.set(d4.value)

d5=sheet['d5']
print(d5.value)
sp.set(d5.value)

d6=sheet['d6']
print(d6.value)
l.set(d6.value)

d7=sheet['d7']
print(d7.value)
sl.set(d7.value)

d8=sheet['d8']
print(d8.value)
dh.set(d8.value)

d9=sheet['d9']
print(d9.value)
st.set(d9.value)

d10=sheet['d10']
print(d10.value)
sr.set(d10.value)

d11=sheet['d11']
print(d11.value)
tr.set(d11.value)

d12=sheet['d12']
print(d12.value)
cr.set(d12.value)

d13=sheet['d13']
print(d13.value)
se.set(d13.value)

d14=sheet['d14']
print(d14.value)
te.set(d14.value)

d15=sheet['d15']
print(d15.value)
ce.set(d15.value)

d16=sheet['d16']
print(d16.value)
je.set(d16.value)


out.close()







root1.mainloop()
