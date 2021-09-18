# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 21:28:20 2019

@author: jacks
"""
from xml.etree import ElementTree as ET
from datetime import datetime
import zipfile
from openpyxl import load_workbook
from os import rename, path, startfile, listdir
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
#while zlistCount > 0:

#    docxfile=zlist[zlistCount-1]
SDT = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'  #structured document tag
SDTPR= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr' #structured document tag Properties
SDTCON= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent' #structured document Tag content
PARA= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p' #paragraph 
RUN= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r' #run of text in the same font, size and boldness etc.
TXT= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t' # text element
##################word document opened and XML root tree built  ####################################
docxfile='CONOP 19 - Copy.zip'
z = zipfile.ZipFile(docxfile)
f =z.open('word/document.xml')
tree= ET.parse(f)
root= tree.getroot()

A_tag=root.tag
A_attrib=root.attrib
print (A_tag)
    
