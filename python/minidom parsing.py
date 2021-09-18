# -*- coding: utf-8 -*-
"""
Created on Sat Apr 20 18:19:58 2019

@author: jacks
"""

import xml.dom.minidom as MD
import zipfile
SDT = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'  #structured document tag
SDTPR= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr' #structured document tag Properties
SDTCON= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}stdContent' #structured document Tag content
PARA= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p' #paragraph 
RUN= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r' #run of text in the same font, size and boldness etc.
TXT= '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t' # text element
path= "C:\\Users\\jacks\\Desktop\\SHREDDER\\"
docxfile= path+'CONOP 19-320 TMS.zip'
z = zipfile.ZipFile(docxfile)
f =z.open('word/document.xml')
doc= MD.parse(f)

mapping ={} #create dictionary to map XML Tree 
for nodeBook in doc.getElementsByTagName(SDT):
    