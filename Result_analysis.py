import openpyxl
import tkinter as tk
ra = tk.Tk()
ra.title('Result Analysis')
ra.configure(background='black')
tk.Label(ra, 
         text="Address",fg='green', bg='black', font=('comicsans', 20)).grid(row=0)
tk.Label(ra, 
         text="Number of students",fg='green', bg='black', font=('comicsans', 20)).grid(row=1)
tk.Label(ra, 
         text="Name of output file",fg='green', bg='black', font=('comicsans', 20)).grid(row=2)
tk.Label(ra, 
         text='''USER MANUAL:
1. Use only the .xlsx file provided with the software.(INPUT.xlsx)
2. Fill the absolute address of the input file provided with the software or if the file
       ispresent in the directory of the program mention the file name with the extension.
3. Mention the total number of student's data prived as it is mandatory.
4. Write the name of the output file without any extention(by default .xlsx format).
''',fg='CYAN', bg='black', font=('comicsans', 12)).grid(row=3,column=1)
img = tk.PhotoImage(file = r".\kvs-logo.png") 
img1 = img.subsample(12,12)
tk.Label(ra, image = img1).grid(row = 0, column = 2, 
       columnspan = 2, rowspan = 2, padx = 5, pady = 5)
e1 = tk.Entry(ra,width=55)
e2 = tk.Entry(ra,width=55)
e3=tk.Entry(ra,width=55)
e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2,column=1)
tk.Button(ra, 
          text='Get Results', 
          command=ra.quit).grid(row=7,column=1)
def close(): 
    ra.destroy()
tk.Button(ra, text = "Quit", command = close).grid(row=3,column=0)
tk.mainloop()
s=e1.get()
n=int(e2.get())
name=e3.get()
file=openpyxl.load_workbook(s)
sheet=file.active
############################################################
######################  English ############################
############################################################
engv=[]
for i in range (1,n+1):
    x='eng'+str(i)
    engv.append(x)
engsec=[]
for i1 in range (9,109):
    ab='C'+str(i1)
    ze=sheet[ab].value
    engsec.append(ze)
englishsec=dict(zip(engv,engsec))
engm=[]
for j in range (9,109):
    ab1='AC'+str(j)
    z=sheet[ab1].value
    engm.append(z)
english=dict(zip(engv,engm))
engseca=[]
engsecb=[]
aea=0
aeb=0
for j1 in range (1,n+1):
    xe=englishsec.get('eng'+str(j1))
    if xe=='a':
        y=english.get('eng'+str(j1))
        if y=='absent':
            y=0
            aea+=1
        engseca.append(y)
    else:
        y1=english.get('eng'+str(j1))
        if y1=='absent':
            y1=0
            aeb+=1
        engsecb.append(y1)
engm1=[]
for i2 in range(0,n):
    u=engm[i2]
    if u=='absent':
        u=0
    engm1.append(u)
taea=len(engseca)-aea
taeb=len(engsecb)-aeb
############################################################
######################  Hindi ##############################
############################################################
hndv=[]
for i3 in range (1,n+1):
    xy='hnd'+str(i3)
    hndv.append(xy)
hndsec=[]
for i4 in range (117,217):
    ab='C'+str(i4)
    zx=sheet[ab].value
    hndsec.append(zx)
hindisec=dict(zip(hndv,hndsec))
hndm=[]
for j2 in range (117,217):
    z1=sheet['AC'+str(j2)].value
    hndm.append(z1)
hindi=dict(zip(hndv,hndm))
hndseca=[]
hndsecb=[]
aha=0
ahb=0
for j3 in range(1,n+1):
    x4=hindisec.get('hnd'+str(j3))
    if x4=='a':
        y3=hindi.get('hnd'+str(j3))
        if y3=='absent':
            y3=0
            aha+=1
        hndseca.append(y3)
    else:
        y4=hindi.get('hnd'+str(j3))
        if y4=='absent':
            y4=0
            ahb+=1
        hndsecb.append(y4)
hndm1=[]
for i2 in range(0,n):
    u=hndm[i2]
    if u=='absent':
        u=0
    hndm1.append(u)
taha=len(hndseca)-aha
tahb=len(hndsecb)-ahb
#########################################################
######################  Sanskrit ########################
#########################################################
sktv=[]
for i in range (1,n+1):
    x='skt'+str(i)
    sktv.append(x)
sktsec=[]
for j in range (225,325):
    z=sheet['C'+str(j)].value
    sktsec.append(z)
sanskritsec=dict(zip(sktv,sktsec))
sktm=[]
for j in range (225,325):
    ab='AC'+str(j)
    z=sheet[ab].value
    sktm.append(z)
sanskrit=dict(zip(sktv,sktm))
sktseca=[]
sktsecb=[]
asa=0
asb=0
for j in range(1,n+1):
    x=sanskritsec.get('skt'+str(j))
    if x=='a':
        y=sanskrit.get('skt'+str(j))
        if y=='absent':
            y=o
            asa+=1
        sktseca.append(y)
    else:
        y=sanskrit.get('skt'+str(j))
        if y=='absent':
            y=o
            asb+=1
        sktsecb.append(y)
sktm1=[]
for i2 in range(0,n):
    u=sktm[i2]
    if u=='absent':
        u=0
    sktm1.append(u)
tasa=len(sktseca)-asa
tasb=len(sktsecb)-asb
#maths
mthv=[]
for i in range (1,n+1):
    x='mth'+str(i)
    mthv.append(x)
mthsec=[]
for j in range (333,433):
    z=sheet['C'+str(j)].value
    mthsec.append(z)
mathsec=dict(zip(mthv,mthsec))
mthm=[]
for j in range (333,433):
    z=sheet['AC'+str(j)].value
    mthm.append(z)
math=dict(zip(mthv,mthm))
mthseca=[]
mthsecb=[]
ama=0
amb=0
for j in range(1,n+1):
    x=mathsec.get('mth'+str(j))
    if x=='a':
        y=math.get('mth'+str(j))
        if y=='absent':
            y=o
            ama+=1
        mthseca.append(y)
    else:
        y=math.get('mth'+str(j))
        if y=='absent':
            y=o
            amb+=1
        mthsecb.append(y)
mthm1=[]
for i2 in range(0,n):
    u=mthm[i2]
    if u=='absent':
        u=0
    mthm1.append(u)
tama=len(mthseca)-ama
tamb=len(mthsecb)-amb
############################################################
######################  Science ############################
############################################################
sciv=[]
for i in range (1,n+1):
    x='sci'+str(i)
    sciv.append(x)
scisec=[]
for j in range (441,541):
    z=sheet['C'+str(j)].value
    scisec.append(z)
sciencesec=dict(zip(sciv,scisec))
scim=[]
for j in range (441,541):
    z=sheet['AC'+str(j)].value
    scim.append(z)
science=dict(zip(sciv,scim))
sciseca=[]
scisecb=[]
asca=0
ascb=0
for j in range (1,n+1):
    x=sciencesec.get('sci'+str(j))
    if x=='a':
        y=science.get('sci'+str(j))
        if y=='absent':
            y=0
            asca+=1
        sciseca.append(y)
    else:
        y=science.get('sci'+str(j))
        if y=='absent':
            y=0
            ascb+=1
        scisecb.append(y)
scim1=[]
for i2 in range(0,n):
    u=scim[i2]
    if u=='absent':
        u=0
    scim1.append(u)
tasca=len(sciseca)-asca
tascb=len(scisecb)-ascb
############################################################
################### Social science #########################
############################################################
ssciv=[]
for i in range (1,n+1):
    x='ssci'+str(i)
    ssciv.append(x)
sscisec=[]
for j in range (549,649):
    z=sheet['C'+str(j)].value
    sscisec.append(z)
social_sciencesec=dict(zip(ssciv,sscisec))
sscim=[]
for j in range (549,649):
    z=sheet['AC'+str(j)].value
    sscim.append(z)
social_science=dict(zip(ssciv,sscim))
ssciseca=[]
sscisecb=[]
assa=0
assb=0
for j in range (1,n+1):
    x=social_sciencesec.get('ssci'+str(j))
    if x=='a':
        y=social_science.get('ssci'+str(j))
        if y=='absent':
            y=0
            assa+=1
        ssciseca.append(y)
    else:
        y=social_science.get('ssci'+str(j))
        if y=='absent':
            y=0
            assb+=1
        sscisecb.append(y)
sscim1=[]
for i2 in range(0,n):
    u=sscim[i2]
    if u=='absent':
        u=0
    sscim1.append(u)
tassa=len(ssciseca)-assa
tassb=len(sscisecb)-assb
'''CALCULATION'''
##########################################################
############### calculation for 1st section ##############
##########################################################
#English
eaaa=0
eba=0
eca=0
eda=0
eea=0
for d in range(0,len(engseca)):
    i=engseca[d]
    if i<32.9:
        eea+=1
    elif i>=33 and i<=44.9:
        eda+=1
    elif i>=45 and i<=74.9:
        eca+=1
    elif i>=75 and i<=89.9:
        eba+=1
    elif i>=90:
        eaaa+=1
#Hindi
haaa=0
hba=0
hca=0
hda=0
hea=0
for d in range(0,len(hndseca)):
    i=hndseca[d]
    if i<32.9:
        hea+=1
    elif i>=33 and i<=44.9:
        hda+=1
    elif i>=45 and i<=74.9:
        hca+=1
    elif i>=75 and i<=89.9:
        hba+=1
    elif i>=90:
        haaa+=1
#Sanskrit
saaa=0
sba=0
sca=0
sda=0
sea=0
for d in range(0,len(sktseca)):
    i=sktseca[d]
    if i<32.9:
        sea+=1
    elif i>=33 and i<=44.9:
        sda+=1
    elif i>=45 and i<=74.9:
        sca+=1
    elif i>=75 and i<=89.9:
        sba+=1
    elif i>=90:
        saaa+=1
#Math
maaa=0
mba=0
mca=0
mda=0
mea=0
for d in range(0,len(mthseca)):
    i=mthseca[d]
    if i<32.9:
        mea+=1
    elif i>=33 and i<=44.9:
        mda+=1
    elif i>=45 and i<=74.9:
        mca+=1
    elif i>=75 and i<=89.9:
        mba+=1
    elif i>=90:
        maaa+=1
#Science
sciaaa=0
sciba=0
scica=0
scida=0
sciea=0
for d in range(0,len(sciseca)):
    i=sciseca[d]
    if i<32.9:
        sciea+=1
    elif i>=33 and i<=44.9:
        scida+=1
    elif i>=45 and i<=74.9:
        scica+=1
    elif i>=75 and i<=89.9:
        sciba+=1
    elif i>=90:
        sciaaa+=1
#Social Science
ssciaaa=0
ssciba=0
sscica=0
sscida=0
ssciea=0
for d in range(0,len(ssciseca)):
    i=ssciseca[d]
    if i<32.9:
        ssciea+=1
    elif i>=33 and i<=44.9:
        sscida+=1
    elif i>=45 and i<=74.9:
        sscica+=1
    elif i>=75 and i<=89.9:
        ssciba+=1
    elif i>=90:
        ssciaaa+=1
##################################################################################################
######## overall  pass percentage age for indivisual subjects for 1st section ####################
##################################################################################################
tqea=(eaaa+eba+eca+eda)
tqha=(haaa+hba+hca+hda)
tqsa=(saaa+sba+sca+sda)
tqma=(maaa+mba+mca+mda)
tqsca=(sciaaa+sciba+scica+scida)
tqssca=(ssciaaa+ssciba+sscica+sscida)
pea=int((tqea/taea)*100)
pha=int((tqha/taha)*100)
psa=int((tqsa/tasa)*100)
pma=int((tqma/tama)*100)
psca=int((tqsca/tasca)*100)
pssca=int((tqssca/tassa)*100)
#################################################################
############# Overall result fopr 1st section ###################
#################################################################
import numpy as np
oa1=np.add(engseca[0:n],hndseca[0:n])
oa2=np.add(oa1,sktseca[0:n])
oa3=np.add(oa2,sciseca[0:n])
oa4=np.add(oa3,ssciseca[0:n])
oa5=np.add(oa4,mthseca[0:n])
aaa=0
ba=0
ca=0
da=0
ea=0
for i in range(0,len(oa5)):
    i=oa5[i]
    x=(i/600)*100
    if x<32.9:
        ea+=1
    elif x>=33 and x<=44.9:
        da+=1
    elif x>=45 and x<=74.9:
        ca+=1
    elif x>=75 and x<=89.9:
        ba+=1
    elif x>=90:
        aaa+=1
tea=len(engseca)
taa=taea
tqa=aaa+ba+ca+da
tfa=ea
opa=(tqa/taa)*100
##################################################################
################ calculation for 2nd section #####################
##################################################################
#English
eaab=0
ebb=0
ecb=0
edb=0
eeb=0
for d in range(0,len(engsecb)):
    i=engsecb[d]
    if i<32.9:
        eeb+=1
    elif i>=33 and i<=44.9:
        edb+=1
    elif i>=45 and i<=74.9:
        ecb+=1
    elif i>=75 and i<=89.9:
        ebb+=1
    elif i>=90:
        eaab+=1
#Hindi
haab=0
hbb=0
hcb=0
hdb=0
heb=0
for d in range(0,len(hndsecb)):
    i=hndsecb[d]
    if i<32.9:
        heb+=1
    elif i>=33 and i<=44.9:
        hdb+=1
    elif i>=45 and i<=74.9:
        hcb+=1
    elif i>=75 and i<=89.9:
        hbb+=1
    elif i>=90:
        haab+=1
#Sanskrit
saab=0
sbb=0
scb=0
sdb=0
seb=0
for d in range(0,len(sktsecb)):
    i=sktsecb[d]
    if i<32.9:
        seb+=1
    elif i>=33 and i<=44.9:
        sdb+=1
    elif i>=45 and i<=74.9:
        scb+=1
    elif i>=75 and i<=89.9:
        sbb+=1
    elif i>=90:
        saab+=1
#Math
maab=0
mbb=0
mcb=0
mdb=0
meb=0
for d in range(0,len(mthsecb)):
    i=mthsecb[d]
    if i<32.9:
        meb+=1
    elif i>=33 and i<=44.9:
        mdb+=1
    elif i>=45 and i<=74.9:
        mcb+=1
    elif i>=75 and i<=89.9:
        mbb+=1
    elif i>=90:
        maab+=1
#Science
sciaab=0
scibb=0
scicb=0
scidb=0
scieb=0
for d in range(0,len(scisecb)):
    i=scisecb[d]
    if i<32.9:
        scieb+=1
    elif i>=33 and i<=44.9:
        scidb+=1
    elif i>=45 and i<=74.9:
        scicb+=1
    elif i>=75 and i<=89.9:
        scibb+=1
    elif i>=90:
        sciaab+=1
#Social Science
ssciaab=0
sscibb=0
sscicb=0
sscidb=0
sscieb=0
for d in range(0,len(sscisecb)):
    i=sscisecb[d]
    if i<32.9:
        sscieb+=1
    elif i>=33 and i<=44.9:
        sscidb+=1
    elif i>=45 and i<=74.9:
        sscicb+=1
    elif i>=75 and i<=89.9:
        sscibb+=1
    elif i>=90:
        ssciaab+=1
########################################################################################################
############### Overall  pass  percentage for indivisual subjects for 2nd section ######################
########################################################################################################
tqeb=(eaab+ebb+ecb+edb)
tqhb=(haab+hbb+hcb+hdb)
tqsb=(saab+sbb+scb+sdb)
tqmb=(maab+mbb+mcb+mdb)
tqscb=(sciaab+scibb+scicb+scidb)
tqsscb=(ssciaab+sscibb+sscicb+sscidb)
peb=int((tqeb/taeb)*100)
phb=int((tqhb/tahb)*100)
psb=int((tqsb/tasb)*100)
pmb=int((tqmb/tamb)*100)
pscb=int((tqscb/tascb)*100)
psscb=int((tqsscb/tassb)*100)
#####################################################################
################ overall result for 2nd section #####################
#####################################################################
import numpy as np
ob1=np.add(engsecb[0:n],hndsecb[0:n])
ob2=np.add(ob1,sktsecb[0:n])
ob3=np.add(ob2,scisecb[0:n])
ob4=np.add(ob3,sscisecb[0:n])
ob5=np.add(ob4,mthsecb[0:n])
aab=0
bb=0
cb=0
db=0
eb=0
for i in range(0,len(ob5)):
    i=ob5[i]
    x=(i/600)*100
    if x<32.9:
        eb+=1
    elif x>=33 and x<=44.9:
        db+=1
    elif x>=45 and x<=74.9:
        cb+=1
    elif x>=75 and x<=89.9:
        bb+=1
    elif x>=90:
        aab+=1
teb=len(engsecb)
tab=taeb
tqb=aab+bb+cb+db
tfb=eb
opb=(tqb/tab)*100
#################################################################
################## for overall class analysis ###################
#################################################################
#English
oae=taea+taeb
#Hindi
oah=taha+tahb
#Sankrit
oas=tasa+tasb
#Math
oam=tama+tamb
#Science
oasc=tasca+tascb
#Social science
oass=tassa+tassb
############################################################
##################### calculation ##########################
############################################################
#English
eaa=0
eb=0
ec=0
ed=0
ee=0
for d in range(0,n):
    i=engm1[d]
    if i<32.9:
        ee+=1
    elif i>=33 and i<=44.9:
        ed+=1
    elif i>=45 and i<=74.9:
        ec+=1
    elif i>=75 and i<=89.9:
        eb+=1
    elif i>=90:
        eaa+=1
#Hindi
haa=0
hb=0
hc=0
hd=0
he=0
for d in range(0,n):
    i=hndm1[d]
    if i<32.9:
        he+=1
    elif i>=33 and i<=44.9:
        hd+=1
    elif i>=45 and i<=74.9:
        hc+=1
    elif i>=75 and i<=89.9:
        hb+=1
    elif i>=90:
        haa+=1
#Sanskrit
saa=0
sb=0
sc=0
sd=0
se=0
for d in range(0,n):
    i=sktm1[d]
    if i<32.9:
        se+=1
    elif i>=33 and i<=44.9:
        sd+=1
    elif i>=45 and i<=74.9:
        sc+=1
    elif i>=75 and i<=89.9:
        sb+=1
    elif i>=90:
        saa+=1
#Math
maa=0
mb=0
mc=0
md=0
me=0
for d in range(0,n):
    i=mthm1[d]
    if i<32.9:
        me+=1
    elif i>=33 and i<=44.9:
        md+=1
    elif i>=45 and i<=74.9:
        mc+=1
    elif i>=75 and i<=89.9:
        mb+=1
    elif i>=90:
        maa+=1
#Science
sciaa=0
scib=0
scic=0
scid=0
scie=0
for d in range(0,n):
    i=scim1[d]
    if i<32.9:
        scie+=1
    elif i>=33 and i<=44.9:
        scid+=1
    elif i>=45 and i<=74.9:
        scic+=1
    elif i>=75 and i<=89.9:
        scib+=1
    elif i>=90:
        sciaa+=1
#Social Science
ssciaa=0
sscib=0
sscic=0
sscid=0
sscie=0
for d in range(0,n):
    i=sscim1[d]
    if i<32.9:
        sscie+=1
    elif i>=33 and i<=44.9:
        sscid+=1
    elif i>=45 and i<=74.9:
        sscic+=1
    elif i>=75 and i<=89.9:
        sscib+=1
    elif i>=90:
        ssciaa+=1
######################################################################
################ overall result for class ############################
######################################################################
import numpy as np
o1=np.add(engm1[0:n],hndm1[0:n])
o2=np.add(o1,sktm[0:n])
o3=np.add(o2,scim[0:n])
o4=np.add(o3,sscim[0:n])
o5=np.add(o4,mthm[0:n])
aa=0
b=0
c=0
d=0
e=0
for i in range(0,len(o5)):
    i=o5[i]
    x=(i/600)*100
    if x<32.9:
        e+=1
    elif x>=33 and x<=44.9:
        d+=1
    elif x>=45 and x<=74.9:
        c+=1
    elif x>=75 and x<=89.9:
        b+=1
    elif x>=90:
        aa+=1
#############################################################################
############# Overall  pass  %age for indivisual subjects ###################
#############################################################################
tqe=(eaa+eb+ec+ed)
tqh=(haa+hb+hc+hd)
tqs=(saa+sb+sc+sd)
tqm=(maa+mb+mc+md)
tqsc=(sciaa+scib+scic+scid)
tqssc=(ssciaa+sscib+sscic+sscid)
pe=int((tqe/oae)*100)
ph=int((tqh/oah)*100)
ps=int((tqs/oas)*100)
pm=int((tqm/oam)*100)
psc=int((tqsc/oasc)*100)
pssc=int((tqssc/oass)*100)
te=len(engm[0:n])
ta=oae
tq=aa+b+c+d
tf=e
op=(tq/ta)*100
################################################################
################### P.I CALCULATIONS ###########################
################################################################
#1st Section
#English
n1ea=0
n2ea=0
n3ea=0
n4ea=0
n5ea=0
n6ea=0
n7ea=0
for d in range(0,len(engseca)):
    i=engseca[d]
    if i>=33 and i<=40:
        n7ea+=1
    if i>=40.1 and i<=50:
        n6ea+=1
    if i>=50.1 and i<=60:
        n5ea+=1
    if i>=60.1 and i<=70:
        n4ea+=1
    if i>=70.1 and i<=80:
        n3ea+=1
    if i>=80.1 and i<=90:
        n2ea+=1
    if i>=90.1:
        n1ea+=1
piea=((n1ea*7+n2ea*6+n3ea*5+n4ea*4+n5ea*3+n6ea*2+n7ea*1)*100)/(taea*7)
#Hindi
n1ha=0
n2ha=0
n3ha=0
n4ha=0
n5ha=0
n6ha=0
n7ha=0
for d in range(0,len(hndseca)):
    i=hndseca[d]
    if i>=33 and i<=40:
        n7ha+=1
    if i>=40.1 and i<=50:
        n6ha+=1
    if i>=50.1 and i<=60:
        n5ha+=1
    if i>=60.1 and i<=70:
        n4ha+=1
    if i>=70.1 and i<=80:
        n3ha+=1
    if i>=80.1 and i<=90:
        n2ha+=1
    if i>=90.1:
        n1ha+=1
piha=((n1ha*7+n2ha*6+n3ha*5+n4ha*4+n5ha*3+n6ha*2+n7ha*1)*100)/(taha*7)
#Sanskrit
n1sa=0
n2sa=0
n3sa=0
n4sa=0
n5sa=0
n6sa=0
n7sa=0
for d in range(0,len(sktseca)):
    i=sktseca[d]
    if i>=33 and i<=40:
        n7sa+=1
    if i>=40.1 and i<=50:
        n6sa+=1
    if i>=50.1 and i<=60:
        n5sa+=1
    if i>=60.1 and i<=70:
        n4sa+=1
    if i>=70.1 and i<=80:
        n3sa+=1
    if i>=80.1 and i<=90:
        n2sa+=1
    if i>=90.1:
        n1sa+=1
pisa=((n1sa*7+n2sa*6+n3sa*5+n4sa*4+n5sa*3+n6sa*2+n7sa*1)*100)/(tasa*7)
#maths
n1ma=0
n2ma=0
n3ma=0
n4ma=0
n5ma=0
n6ma=0
n7ma=0
for d in range(0,len(mthseca)):
    i=mthseca[d]
    if i>=33 and i<=40:
        n7ma+=1
    if i>=40.1 and i<=50:
        n6ma+=1
    if i>=50.1 and i<=60:
        n5ma+=1
    if i>=60.1 and i<=70:
        n4ma+=1
    if i>=70.1 and i<=80:
        n3ma+=1
    if i>=80.1 and i<=90:
        n2ma+=1
    if i>=90.1:
        n1ma+=1
pima=((n1ma*7+n2ma*6+n3ma*5+n4ma*4+n5ma*3+n6ma*2+n7ma*1)*100)/(tama*7)
#science
n1sca=0
n2sca=0
n3sca=0
n4sca=0
n5sca=0
n6sca=0
n7sca=0
for d in range(0,len(sciseca)):
    i=sciseca[d]
    if i>=33 and i<=40:
        n7sca+=1
    if i>=40.1 and i<=50:
        n6sca+=1
    if i>=50.1 and i<=60:
        n5sca+=1
    if i>=60.1 and i<=70:
        n4sca+=1
    if i>=70.1 and i<=80:
        n3sca+=1
    if i>=80.1 and i<=90:
        n2sca+=1
    if i>=90.1:
        n1sca+=1
pisca=((n1sca*7+n2sca*6+n3sca*5+n4sca*4+n5sca*3+n6sca*2+n7sca*1)*100)/(tasca*7)
#Social science
n1ssca=0
n2ssca=0
n3ssca=0
n4ssca=0
n5ssca=0
n6ssca=0
n7ssca=0
for d in range(0,len(ssciseca)):
    i=ssciseca[d]
    if i>=33 and i<=40:
        n7ssca+=1
    if i>=40.1 and i<=50:
        n6ssca+=1
    if i>=50.1 and i<=60:
        n5ssca+=1
    if i>=60.1 and i<=70:
        n4ssca+=1
    if i>=70.1 and i<=80:
        n3ssca+=1
    if i>=80.1 and i<=90:
        n2ssca+=1
    if i>=90.1:
        n1ssca+=1
pissca=((n1ssca*7+n2ssca*6+n3ssca*5+n4ssca*4+n5ssca*3+n6ssca*2+n7ssca*1)*100)/(tassa*7)
piseca=(piea+piha+pima+pisa+pisca+pissca)/6
#2nd section
#English
n1eb=0
n2eb=0
n3eb=0
n4eb=0
n5eb=0
n6eb=0
n7eb=0
for d in range(0,len(engsecb)):
    i=engsecb[d]
    if i>=33 and i<=40:
        n7eb+=1
    if i>=40.1 and i<=50:
        n6eb+=1
    if i>=50.1 and i<=60:
        n5eb+=1
    if i>=60.1 and i<=70:
        n4eb+=1
    if i>=70.1 and i<=80:
        n3eb+=1
    if i>=80.1 and i<=90:
        n2eb+=1
    if i>=90.1:
        n1eb+=1
pieb=((n1eb*7+n2eb*6+n3eb*5+n4eb*4+n5eb*3+n6eb*2+n7eb*1)*100)/(taeb*7)
#Hindi
n1hb=0
n2hb=0
n3hb=0
n4hb=0
n5hb=0
n6hb=0
n7hb=0
for d in range(0,len(hndsecb)):
    i=hndsecb[d]
    if i>=33 and i<=40:
        n7hb+=1
    if i>=40.1 and i<=50:
        n6hb+=1
    if i>=50.1 and i<=60:
        n5hb+=1
    if i>=60.1 and i<=70:
        n4hb+=1
    if i>=70.1 and i<=80:
        n3hb+=1
    if i>=80.1 and i<=90:
        n2hb+=1
    if i>=90.1:
        n1hb+=1
pihb=((n1hb*7+n2hb*6+n3hb*5+n4hb*4+n5hb*3+n6hb*2+n7hb*1)*100)/(tahb*7)
#Sanskrit
n1sb=0
n2sb=0
n3sb=0
n4sb=0
n5sb=0
n6sb=0
n7sb=0
for d in range(0,len(sktsecb)):
    i=sktsecb[d]
    if i>=33 and i<=40:
        n7sb+=1
    if i>=40.1 and i<=50:
        n6sb+=1
    if i>=50.1 and i<=60:
        n5sb+=1
    if i>=60.1 and i<=70:
        n4sb+=1
    if i>=70.1 and i<=80:
        n3sb+=1
    if i>=80.1 and i<=90:
        n2sb+=1
    if i>=90.1:
        n1sb+=1
pisb=((n1sb*7+n2sb*6+n3sb*5+n4sb*4+n5sb*3+n6sb*2+n7sb*1)*100)/(tasb*7)
#Math
n1mb=0
n2mb=0
n3mb=0
n4mb=0
n5mb=0
n6mb=0
n7mb=0
for d in range(0,len(mthsecb)):
    i=mthsecb[d]
    if i>=33 and i<=40:
        n7mb+=1
    if i>=40.1 and i<=50:
        n6mb+=1
    if i>=50.1 and i<=60:
        n5mb+=1
    if i>=60.1 and i<=70:
        n4mb+=1
    if i>=70.1 and i<=80:
        n3mb+=1
    if i>=80.1 and i<=90:
        n2mb+=1
    if i>=90.1:
        n1mb+=1
pimb=((n1mb*7+n2mb*6+n3mb*5+n4mb*4+n5mb*3+n6mb*2+n7mb*1)*100)/(tamb*7)
#Science
n1scb=0
n2scb=0
n3scb=0
n4scb=0
n5scb=0
n6scb=0
n7scb=0
for d in range(0,len(scisecb)):
    i=scisecb[d]
    if i>=33 and i<=40:
        n7scb+=1
    if i>=40.1 and i<=50:
        n6scb+=1
    if i>=50.1 and i<=60:
        n5scb+=1
    if i>=60.1 and i<=70:
        n4scb+=1
    if i>=70.1 and i<=80:
        n3scb+=1
    if i>=80.1 and i<=90:
        n2scb+=1
    if i>=90.1:
        n1scb+=1
piscb=((n1scb*7+n2scb*6+n3scb*5+n4scb*4+n5scb*3+n6scb*2+n7scb*1)*100)/(tascb*7)
#Social Science
n1sscb=0
n2sscb=0
n3sscb=0
n4sscb=0
n5sscb=0
n6sscb=0
n7sscb=0
for d in range(0,len(sscisecb)):
    i=sscisecb[d]
    if i>=33 and i<=40:
        n7sscb+=1
    if i>=40.1 and i<=50:
        n6sscb+=1
    if i>=50.1 and i<=60:
        n5sscb+=1
    if i>=60.1 and i<=70:
        n4sscb+=1
    if i>=70.1 and i<=80:
        n3sscb+=1
    if i>=80.1 and i<=90:
        n2sscb+=1
    if i>=90.1:
        n1sscb+=1
pisscb=((n1sscb*7+n2sscb*6+n3sscb*5+n4sscb*4+n5sscb*3+n6sscb*2+n7sscb*1)*100)/(tassb*7)
pisecb=(pieb+pihb+pimb+pisb+piscb+pisscb)/6
#Overall
#English
n1e=0
n2e=0
n3e=0
n4e=0
n5e=0
n6e=0
n7e=0
for d in range(0,n):
    i=engm1[d]
    if i>=33 and i<=40:
        n7e+=1
    if i>=40.1 and i<=50:
        n6e+=1
    if i>=50.1 and i<=60:
        n5e+=1
    if i>=60.1 and i<=70:
        n4e+=1
    if i>=70.1 and i<=80:
        n3e+=1
    if i>=80.1 and i<=90:
        n2e+=1
    if i>=90.1:
        n1e+=1
pie=((n1e*7+n2e*6+n3e*5+n4e*4+n5e*3+n6e*2+n7e*1)*100)/(oae*7)
#Hindi
n1h=0
n2h=0
n3h=0
n4h=0
n5h=0
n6h=0
n7h=0
for d in range(0,n):
    i=hndm1[d]
    if i>=33 and i<=40:
        n7h+=1
    if i>=40.1 and i<=50:
        n6h+=1
    if i>=50.1 and i<=60:
        n5h+=1
    if i>=60.1 and i<=70:
        n4h+=1
    if i>=70.1 and i<=80:
        n3h+=1
    if i>=80.1 and i<=90:
        n2h+=1
    if i>=90.1:
        n1h+=1
pih=((n1h*7+n2h*6+n3h*5+n4h*4+n5h*3+n6h*2+n7h*1)*100)/(oah*7)
#Sanskrit
n1s=0
n2s=0
n3s=0
n4s=0
n5s=0
n6s=0
n7s=0
for d in range(0,n):
    i=sktm1[d]
    if i>=33 and i<=40:
        n7s+=1
    if i>=40.1 and i<=50:
        n6s+=1
    if i>=50.1 and i<=60:
        n5s+=1
    if i>=60.1 and i<=70:
        n4s+=1
    if i>=70.1 and i<=80:
        n3s+=1
    if i>=80.1 and i<=90:
        n2s+=1
    if i>=90.1:
        n1s+=1
pis=((n1s*7+n2s*6+n3s*5+n4s*4+n5s*3+n6s*2+n7s*1)*100)/(oas*7)
#Math
n1m=0
n2m=0
n3m=0
n4m=0
n5m=0
n6m=0
n7m=0
for d in range(0,n):
    i=mthm1[d]
    if i>=33 and i<=40:
        n7m+=1
    if i>=40.1 and i<=50:
        n6m+=1
    if i>=50.1 and i<=60:
        n5m+=1
    if i>=60.1 and i<=70:
        n4m+=1
    if i>=70.1 and i<=80:
        n3m+=1
    if i>=80.1 and i<=90:
        n2m+=1
    if i>=90.1:
        n1m+=1
pim=((n1m*7+n2m*6+n3m*5+n4m*4+n5m*3+n6m*2+n7m*1)*100)/(oam*7)
#Science
n1sc=0
n2sc=0
n3sc=0
n4sc=0
n5sc=0
n6sc=0
n7sc=0
for d in range(0,n):
    i=scim1[d]
    if i>=33 and i<=40:
        n7sc+=1
    if i>=40.1 and i<=50:
        n6sc+=1
    if i>=50.1 and i<=60:
        n5sc+=1
    if i>=60.1 and i<=70:
        n4sc+=1
    if i>=70.1 and i<=80:
        n3sc+=1
    if i>=80.1 and i<=90:
        n2sc+=1
    if i>=90.1:
        n1sc+=1
pisc=((n1sc*7+n2sc*6+n3sc*5+n4sc*4+n5sc*3+n6sc*2+n7sc*1)*100)/(oasc*7)
#Social Science
n1ssc=0
n2ssc=0
n3ssc=0
n4ssc=0
n5ssc=0
n6ssc=0
n7ssc=0
for d in range(0,n):
    i=sscim1[d]
    if i>=33 and i<=40:
        n7ssc+=1
    if i>=40.1 and i<=50:
        n6ssc+=1
    if i>=50.1 and i<=60:
        n5ssc+=1
    if i>=60.1 and i<=70:
        n4ssc+=1
    if i>=70.1 and i<=80:
        n3ssc+=1
    if i>=80.1 and i<=90:
        n2ssc+=1
    if i>=90.1:
        n1ssc+=1
pissc=((n1ssc*7+n2ssc*6+n3ssc*5+n4ssc*4+n5ssc*3+n6ssc*2+n7ssc*1)*100)/(oass*7)
piclass=(pie+pih+pim+pis+pisc+pissc)/6
####################################################
################### Output #########################
####################################################
AS='SHEET.xlsx'
file_output=openpyxl.load_workbook(AS)
ot=file_output.worksheets[0]
###################################################################
#################### For 1st section ##############################
###################################################################
#Total enrolled each subject
ot['D13']=len(engseca)
ot['D14']=len(hndseca)
ot['D15']=len(mthseca)
ot['D16']=len(sciseca)
ot['D17']=len(ssciseca)
ot['D18']=len(sktseca)
#Total appeared each subject
ot['E13']=taea
ot['E14']=taha
ot['E15']=tama
ot['E16']=tasca
ot['E17']=tassa
ot['E18']=tasa
#Total qualified each subject
ot['F13']=tqea
ot['F14']=tqha
ot['F15']=tqma
ot['F16']=tqsca
ot['F17']=tqssca
ot['F18']=tqsa
#TOTAL FAILED EACH SUBJECT
ot['G13']=eea
ot['G14']=hea
ot['G15']=mea
ot['G16']=sciea
ot['G17']=ssciea
ot['G18']=sea
#TOTAL COMPARTMENT EACH SUBJECT
ot['H13']=eea
ot['H14']=hea
ot['H15']=mea
ot['H16']=sciea
ot['H17']=ssciea
ot['H18']=sea
#PASS Percentage
ot['I13']=pea
ot['I14']=pha
ot['I15']=pma
ot['I16']=psca
ot['I17']=pssca
ot['I18']=psa
#P.I
ot['J13']=piea
ot['J14']=piha
ot['J15']=pima
ot['J16']=pisca
ot['J17']=pissca
ot['J18']=pisa
#D
ot['K13']=eda
ot['K14']=hda
ot['K15']=mda
ot['K16']=scida
ot['K17']=sscida
ot['K18']=sda
#C
ot['L13']=eca
ot['L14']=hca
ot['L15']=mca
ot['L16']=scica
ot['L17']=sscica
ot['L18']=sca
#B
ot['M13']=eba
ot['M14']=hba
ot['M15']=mba
ot['M16']=sciba
ot['M17']=ssciba
ot['M18']=sba
#AA
ot['N13']=eaaa
ot['N14']=haaa
ot['N15']=maaa
ot['N16']=sciaaa
ot['N17']=ssciaaa
ot['N18']=saaa
#OVERALL
#ENROLLED and rest
ot['D21']=tea
ot['E21']=taa
ot['F21']=tqa
ot['G21']=tfa
ot['H21']=0
ot['I21']=int(opa)
ot['J21']=piseca
ot['k21']=da
ot['L21']=ca
ot['M21']=ba
ot['N21']=aaa
#for 2st section
#total enrolled each subject
ot['D29']=len(engsecb)
ot['D30']=len(hndsecb)
ot['D31']=len(mthsecb)
ot['D32']=len(scisecb)
ot['D33']=len(sscisecb)
ot['D34']=len(sktsecb)
#total appeared each subject
ot['E29']=taeb
ot['E30']=tahb
ot['E31']=tamb
ot['E32']=tascb
ot['E33']=tassb
ot['E34']=tasb
#total qualified each subject
ot['F29']=tqeb
ot['F30']=tqhb
ot['F31']=tqmb
ot['F32']=tqscb
ot['F33']=tqsscb
ot['F34']=tqsb
#TOTAL FAILED EACH SUBJECT
ot['G29']=eeb
ot['G30']=heb
ot['G31']=meb
ot['G32']=scieb
ot['G33']=sscieb
ot['G34']=seb
#TOTAL COMPARTMENT EACH SUBJECT
ot['H29']=eeb
ot['H30']=heb
ot['H31']=meb
ot['H32']=scieb
ot['H33']=sscieb
ot['H34']=seb
#PASS %
ot['I29']=peb
ot['I30']=phb
ot['I31']=pmb
ot['I32']=pscb
ot['I33']=psscb
ot['I34']=psb
#P.I
ot['J29']=pieb
ot['J30']=pihb
ot['J31']=pimb
ot['J32']=piscb
ot['J33']=pisscb
ot['J34']=pisb
#D
ot['K29']=edb
ot['K30']=hdb
ot['K31']=mdb
ot['K32']=scidb
ot['K33']=sscidb
ot['K34']=sdb
#C
ot['L29']=ecb
ot['L30']=hcb
ot['L31']=mcb
ot['L32']=scicb
ot['L33']=sscicb
ot['L34']=scb
#B
ot['M29']=ebb
ot['M30']=hbb
ot['M31']=mbb
ot['M32']=scibb
ot['M33']=sscibb
ot['M34']=sbb
#AA
ot['N29']=eaab
ot['N30']=haab
ot['N31']=maab
ot['N32']=sciaab
ot['N33']=ssciaab
ot['N34']=saab
#OVERALL
#ENROLLED and rest
ot['D37']=teb
ot['E37']=tab
ot['F37']=tqb
ot['G37']=tfb
ot['H37']=0
ot['I37']=int(opb)
ot['J37']=pisecb
ot['k37']=db
ot['L37']=cb
ot['M37']=bb
ot['N37']=aab
#for overall class
#total enrolled each subject
ot['D45']=len(engseca)+len(engsecb)
ot['D46']=len(hndseca)+len(hndsecb)
ot['D47']=len(mthseca)+len(mthsecb)
ot['D48']=len(sciseca)+len(scisecb)
ot['D49']=len(ssciseca)+len(sscisecb)
ot['D50']=len(sktseca)+len(sktsecb)
#total appeared each subject
ot['E45']=oae
ot['E46']=oah
ot['E47']=oam
ot['E48']=oasc
ot['E49']=oass
ot['E50']=oas
#total qualified each subject
ot['F45']=tqe
ot['F46']=tqh
ot['F47']=tqm
ot['F48']=tqsc
ot['F49']=tqssc
ot['F50']=tqs
#TOTAL FAILED EACH SUBJECT
ot['G45']=ee
ot['G46']=he
ot['G47']=me
ot['G48']=scie
ot['G49']=sscie
ot['G50']=se
#TOTAL COMPARTMENT EACH SUBJECT
ot['H45']=ee
ot['H46']=he
ot['H47']=me
ot['H48']=scie
ot['H49']=sscie
ot['H50']=se
#PASS %
ot['I45']=pe
ot['I46']=ph
ot['I47']=pm
ot['I48']=psc
ot['I49']=pssc
ot['I50']=ps
#P.I
ot['J45']=pie
ot['J46']=pih
ot['J47']=pim
ot['J48']=pisc
ot['J49']=pissc
ot['J50']=pis
#D
ot['K45']=ed
ot['K46']=hd
ot['K47']=md
ot['K48']=scid
ot['K49']=sscid
ot['K50']=sd
#C
ot['L45']=ec
ot['L46']=hc
ot['L47']=mc
ot['L48']=scic
ot['L49']=sscic
ot['L50']=sc
#B
ot['M45']=eb
ot['M46']=hb
ot['M47']=mb
ot['M48']=scib
ot['M49']=sscib
ot['M50']=sb
#AA
ot['N45']=eaa
ot['N46']=haa
ot['N47']=maa
ot['N48']=sciaa
ot['N49']=ssciaa
ot['N50']=saa
#OVERALL
#ENROLLED and rest
ot['D53']=te
ot['E53']=ta
ot['F53']=tq
ot['G53']=tf
ot['H53']=0
ot['I53']=int(op)
ot['J53']=piclass
ot['k53']=d
ot['L53']=c
ot['M53']=b
ot['N53']=aa
file_output.save(name+'.xlsx')
print("The program has been executed successfully")

# made my Debajyoti Sarkar
