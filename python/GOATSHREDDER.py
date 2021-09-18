# -*- coding: utf-8 -*-
"""
Spyder Editor
this script was written by Jackson Perry POC jackson.d.perry.mil@mail.mil. It was written for government use by a government employee and is therefore completly in the public domain.
"""
from xml.etree import ElementTree as ET
from datetime import datetime
import zipfile
from openpyxl import load_workbook
from os import rename, path, startfile
from glob import iglob
from tkinter import filedialog, Tk, Label, Button, StringVar, mainloop
########################### set the path for shredder file #################################
def browse_button():
    global folder_path
    global path2
    path2 = ()
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    path2=filename
    print (filename)
    
root = Tk()
folder_path=StringVar()
lbl2 = Label(root, text=" Select the Folder containing your TMS Reports").grid(row=0, column=3)
button2 = Button(root,text="Browse", command=browse_button).grid(row=3, column=3)
lbl1 = Label(root, textvariable=folder_path).grid(row=5, column=3)
button1 = Button(root, text="Continue", command=root.destroy).grid(row=7, column=3)
mainloop()
############## convert each docx file in zip file so XML files can be read  ###########################
#path= "C:\\Users\\jacks\\Desktop\\SHREDDER\\"
for filename in iglob(path.join(path2, '*.docx')):
    rename(filename, filename[:-5] + '.zip')
zlist=[]
for filename in iglob(path.join(path2, '*.zip')):
    zlist.append(filename)
zlistCount= len(zlist)
##################make a list of zip files in folder #########################################
zlist=[]
for filename in iglob(path.join(path2, '*.zip')):
    zlist.append(filename)
zlistCount= len(zlist)

############################### GOAT SHRED EACH docx(now a zip) filein folder #################################
while zlistCount > 0:

    docxfile=zlist[zlistCount-1]
    SDT = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'  #structured document tag
    SDTPR= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr' #structured document tag Properties
    SDTCON= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent' #structured document Tag content
    PARA= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p' #paragraph 
    RUN= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r' #run of text in the same font, size and boldness etc.
    TXT= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t' # text element
##################word document opened and XML root tree built  ####################################
    z = zipfile.ZipFile(docxfile)
    f =z.open('word/document.xml')
    tree= ET.parse(f)
    root= tree.getroot()
##################   this function will bring you from a final sdt elemtent to the first text it holds in its first run####################################
    def GetStrucText(x):
        runlist=x.find(SDTCON).findall(RUN)
        result=[]
        for run in runlist:
            result.append( run.find(TXT).text)
        return ' '.join(result)
#################  find each top level structured document tags and name the content subpart(ignoring formating subpart)#############

    body= list(root)[0]
    body=body.findall(SDT) #ignores all free typed paragraphs only looks for structured elements
    TitleBlock= list(body)[0].find(SDTCON)
    ExecSum=list(body)[1].find(SDTCON)
    POC=list(body)[2].find(SDTCON)
    ODR=list(body)[4].find(SDTCON)
    OIV=list(body)[6].find(SDTCON)
###################################  Title Block #############################
    x=[]
    y=[]
    x = TitleBlock.find(PARA).findall(SDT)[0]
    y = TitleBlock.find(PARA).findall(SDT)[1]
    CONOP = "CONOP"+" "+GetStrucText(x)+"-"+GetStrucText(y)
    x= TitleBlock.findall(PARA)[1]
    x=x.findall(SDT)[1]
    EVENTDATE=GetStrucText(x)
    EVENTDATE= datetime.strptime(EVENTDATE,'%m/%d/%Y') # convert EVENTDATE form string to datetime
    x=[]
    y=[]
############################  Executive Summary  ######################
    x=ExecSum.findall(PARA)[3].findall(SDT)
    EDClass=GetStrucText(x[0])
    ED=GetStrucText(x[1])
    x=[]
######################### Point of Contact         ########################
    x=POC.find(SDT).find(SDTCON).find(PARA).findall(SDT)
    RANK=GetStrucText(x[0])
    FirstName=GetStrucText(x[1])
    LastName=GetStrucText(x[2])
    PHONE=GetStrucText(x[4])
    EMAIL=GetStrucText(x[5])
    x=[]
##########################   create excel sheet with date and CONOP number  ###############################
    wb= load_workbook(path2+"\\index.xlsx")

    wb['Sheet1']['AJ2']=EVENTDATE
    wb['Sheet1']['AJ3']=CONOP
############################### ODR section ######################
    ODRlist=[]
    ODRcount=()    
    ODRlist= ODR.findall(SDT)
    ODRcount=len(ODRlist)
    while ODRcount > 0:
        ROW=str(ODRcount+6)
        x=ODR.findall(SDT)[ODRcount-1].find(SDTCON).findall(PARA)
        wb['Sheet1']['D'+ROW]=GetStrucText(x[1].findall(SDT)[1])
        if GetStrucText(x[1].find(SDT).find(SDTCON).find(SDT))=='(U//FOUO)':
            wb['Sheet1']['B'+ROW]='FOUO' 
            wb['Sheet1']['D1']='FOUO'
        wb['Sheet1']['H'+ROW]=GetStrucText(x[2].findall(SDT)[1])
        if GetStrucText(x[1].find(SDT).find(SDTCON).find(SDT))=='(U//FOUO)':
            wb['Sheet1']['F'+ROW]='FOUO'
            wb['Sheet1']['D1']='FOUO'
        wb['Sheet1']['L'+ROW]=GetStrucText(x[3].findall(SDT)[1])
        if GetStrucText(x[1].find(SDT).find(SDTCON).find(SDT))=='(U//FOUO)':
            wb['Sheet1']['J'+ROW]='FOUO'
            wb['Sheet1']['D1']='FOUO'
        wb['Sheet1']['P'+ROW]=GetStrucText(x[4].findall(SDT)[1])
        if GetStrucText(x[1].find(SDT).find(SDTCON).find(SDT))=='(U//FOUO)':
            wb['Sheet1']['N'+ROW]='FOUO'
            wb['Sheet1']['D1']='FOUO'
        wb['Sheet1']['X'+ROW]=ED
        if EDClass =='(U//FOUO)':
            wb['Sheet1']['V'+ROW]='FOUO'
            wb['Sheet1']['D1']='FOUO'
        wb['Sheet1']['T'+ROW]=GetStrucText(x[5].findall(SDT)[1])+' '+GetStrucText(x[6].findall(SDT)[1])+' '+GetStrucText(x[7].findall(SDT)[1])+' '+GetStrucText(x[8].findall(SDT)[1])         
        ODRcount=ODRcount-1
############################ save the excel file   ###################################################################################
    wb.save(zlist[zlistCount-1][:-4]+'.xlsx')
    zlistCount=zlistCount-1
################################ convert all the zip back to docx format ########################################################
z.close()
f.close()
for filename in iglob(path.join(path2, '*.zip')):
    rename(filename, filename[:-4] + '.docx')
exit
import webbrowser
webbrowser.open('https://www.jllis.mil/index.cfm?disp=../admin/import/Import.cfm', new = 2)
startfile(path2)