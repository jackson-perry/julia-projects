# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 08:19:17 2019

@author: jacks
"""
from openpyxl import load_workbook
import zipfile
import pandas as pd
from bs4 import BeautifulSoup as BS
import os
import tkinter
import winsound
from tkinter import filedialog, Tk, Label, Button, mainloop


########################### set the path for shredder file #################################
def browse_button():  # command for slecting the TMs imput file
    global filename
    filename = filedialog.askopenfilename()
    print (filename)
def browse_button2():  # command for slecting the excel output file
    global XLfilename
    XLfilename = filedialog.askopenfilename()
    print (XLfilename)
def play():  #ommand to continue script 
    winsound.PlaySound("bleat.wav", winsound.SND_FILENAME)
    root.destroy()
############################## builds the pop up window ##################################         
root= Tk()
pic=tkinter.PhotoImage(file= "GOAT.gif")  #adds GOAT logo
button3=Button(root, text="Browse", command=browse_button2).grid(row=9,column=5) #creates button for selecting excel output file  output file pat to variable XLfilename
lbl3 = Label(root, text="Select the Excel Template from JLLIS").grid(row=7,column=5) #creats label 
lbl2 = Label(root, text=" Select the TMS Report").grid(row=0, column=5)  #creates button label
button2 = Button(root,text="Browse", command=browse_button).grid(row=3, column=5)   #creates a button to select TMS inputs out put file path to variable filename
lbl1= Label(root, image=pic).grid(row=13, column=5) # puts goat image on window
button1 = Button(root, text="Continue", command=play).grid(row=13, column=5) #continue button closes window and continues script while playing wav file
mainloop()
############## convert each docx file in zip file so XML files can be read  ###########################
docxfile= filename[:-5] + '.zip'
pre, ext = os.path.splitext(filename)
os.rename(filename, pre + '.zip')
###################################### Open File ########################################################
z = zipfile.ZipFile(docxfile)
f =z.open('word/document.xml')
################################ beutiful soup is used to read and navigate teh xml file################################
soup = BS(f, 'xml')  # entire xml; file is called soup
soupdocument=soup.document  #document only child node of soup contains MS office schemas used in the document not 
soupbody=soupdocument.body  # body is the only child node of document contains all the paragraphs (p) and structured document tags(sdt)
############################### PROBLEM found ONe TMS where The ODR is soupbody[7] probably due to some paragraphs or something figue out a way to find the ODR section wihtout hardcoding the number.
#ODR=soupbody.contents[5]   # ODR contains the Observation discusion recomendation section structured document tag (the sixth xml element in the documents body counting from 0) 
ODR=soupbody.find('tag',{'w:val' : "ODR Repeating Section"}).parent.parent #finds the parent node of the ODR Repeating Section Header
CONTACT=soupbody.find('tag',{'w:val':"Repeating POC Section"}).parent.parent ## this contains the AWG POC section structured documetn tag 

######################### pulls all the tags and text form each descendent node of the structured document tags, overwrites the ODR section, only produces the last ODr repeating section##############################
sdt=soup.find_all("sdt")  #make a list of all structured documetn tags
xdict={}   #make a dictionary
for i in sdt:        #iterate through each structured document tag  and append the tag and the text to a dictionary
    xdict.update({str(i.tag) : i.get_text()}) #make sure to store the tag a string or it will be an xml tag
########################## creates a pandas dataframe to hold all the data extracted key is the tag index is the text string #######################
CHILDcount=len(ODR.contents[len(ODR.contents)-1].contents)-1
df = pd.DataFrame(data=xdict, index=[CHILDcount])
########################## this section captures all but the last ODr and adds it to the dataframe ######################  
ODRdict ={}     #make another dictionary for the ODR and section labels
#CHILDcount= CHILDcount-1 # The last observation is already done we will work form last to first below
while CHILDcount >= 0:  #0 is the indes location of the first odr section we are working form last to first
    CHILD= ODR.contents[len(ODR.contents)-1].contents[CHILDcount].find_all('sdt') #find all Structured Document Tag in the give ODR repeating section
    for i in CHILD: #iterate through that list of sdt
        ODRdict.update({str(i.tag):i.get_text()}) #add each tag and text to a dictionary make sure to store the tag a string or it will be an xml tag
        
    df_ODR=pd.DataFrame(ODRdict, index=[CHILDcount]) #convert the dictionary into a dataframe
    df=pd.concat([df,df_ODR], axis=0,sort=True) #append that dataframe into the main datafram that already holds the header info and the last ODr section
    CHILDcount=CHILDcount-1 # take one from our counter so we will begin the loop on the next closest to first ODR section
###################### this section handels the repeating contacts section  #############################
CONTACTcount=len(CONTACT.contents[len(CONTACT.contents)-1].contents)-1
CONTACTdict ={}  #make another dictionary for the repeating contacts section and thier labels
df_CONTACT=pd.DataFrame(CONTACTdict, index=[CONTACTcount]) #convert the dictionary into a dataframe
while CONTACTcount >=0:
    CONTACTchild = CONTACT.contents[len(CONTACT.contents)-1].contents[CONTACTcount].findAll('sdt')
    for i in CONTACTchild:
        CONTACTdict.update({str(i.tag):i.get_text()})
    df_CONTACTloop=pd.DataFrame(CONTACTdict, index=[CONTACTcount]) #convert the dictionary into a dataframe
    df_CONTACT=pd.concat([df_CONTACT,df_CONTACTloop], axis=0,sort=True) #append that dataframe into the main datafram that already holds the header info and the last ODr section
    CONTACTcount=CONTACTcount-1
ROW=len(ODR.contents[len(ODR.contents)-1].contents)-1 # the ODR variable has the number of repeating sections in it, we subtract one because the computer counts from zero 
df_CONTACT=pd.DataFrame(CONTACTdict, index=[CONTACTcount]) #convert the dictionary into a dataframe
df2 =list(df.columns.values)     # pull out all the tags from the first column
df2= [l.strip('<w:tag w:val="') for l in df2]  # remove the XML stuff form the front of the tag
df2= [l.strip('"/>') for l in df2] # remove the XML stuff form the back of the tag
df.columns = df2 # rename all the column headers thier tag names without all the XML markup
######################################## strip all ###############################################

#######now again for the CONTACTS Dataframe ####################
df2 =list(df_CONTACT.columns.values)     # pull out all the tags from the first column
df2= [l.strip('<w:tag w:val="') for l in df2]  # remove the XML stuff form the front of the tag
df2= [l.strip('"/>') for l in df2] # remove the XML stuff form the back of the tag
df_CONTACT.columns = df2 # rename all the column headers thier tag names without all the XML markup
del df['ODR Repeating Section'] #this is a duplicate holding all the lasT ODR in one long string, it is no longer needed
df.to_excel('test3.xlsx')  # creates an excel file containing the data frame  
##########################   create excel sheet with date and CONOP number  ###############################

wb = load_workbook(XLfilename)  #click on the excel templae from jllis
wb['Sheet1']['B4']=df.iloc[0]['End Date']            #sends enddate from first frow of dataframe to excel file cell AJ2
wb['Sheet1']['AJ3']='CONOP'+' '+df.iloc[0]['CONOP_Year']+'-'+df.iloc[0]['CONOP #']       #concatenats CONOP year and number and passes to spreadsheet AJ32

while ROW >=0: # check if we have any observations left
        wb['Sheet1']['A'+ str(ROW+7)] = "Unclassified"        
        if df.iloc[ROW]['TOI Portion Mark']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['B'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['C'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO'  #these four rows do portion marking of the topic
        else:
            wb['Sheet1']['B'+str(ROW+7)]=''
        CONCAT=()
        if "TOI 1" in df:
            CONCAT= str(df.iloc[ROW]['TOI 1'])+', '+str(df.iloc[ROW]['TOI 2'])+', '+str(df.iloc[ROW]['TOI 3'])+', '+str(df.iloc[ROW]['TOI 4'])+'.'
            CONCAT = CONCAT.replace('Choose an item.', '') # concatenate topics of interest and send them to the excel
        elif  "TOIs" in df:
            CONCAT= str(df.iloc[ROW]['TOIs'])
        wb['Sheet1']['D'+ str(ROW+7)] = CONCAT
        wb['Sheet1']['E'+ str(ROW+7)] = "Unclassified"
        if df.iloc[ROW]['Obs Portion Mark']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['F'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['G'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO' 
        else:
            wb['Sheet1']['F'+str(ROW+7)]=''              # does observation portion marking
        if 'observation' in df:
            wb['Sheet1']['H'+ str(ROW+7)] = df.iloc[ROW]['observation'] # observation to excel
        elif 'Observation' in df:
            wb['Sheet1']['H'+ str(ROW+7)] = df.iloc[ROW]['Observation'] # observation to excel
        else:
            wb['Sheet1']['H'+ str(ROW+7)] = ''
        wb['Sheet1']['I'+ str(ROW+7)] = "Unclassified"
        if df.iloc[ROW]['Discussion Portion Mark']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['J'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['K'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO'                    
        else:
             wb['Sheet1']['J'+str(ROW+7)]=''# does discussion portion marking
        if 'Discussion' in df:
            wb['Sheet1']['L'+ str(ROW+7)] = df.iloc[ROW]['Discussion'] # discusiion to excel
        wb['Sheet1']['M'+ str(ROW+7)] = "Unclassified"
        if df.iloc[ROW]['Recomendation Portion Mark']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['N'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['O'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO'                      
        else:
            wb['Sheet1']['N'+str(ROW+7)]=''# does Recommendation portion marking
        if 'Recommendation' in df:
            wb['Sheet1']['P'+ str(ROW+7)] = df.iloc[ROW]['Recommendation'] #recommendation to excel
        wb['Sheet1']['Q'+ str(ROW+7)] = "Unclassified"
        if df.iloc[ROW]['WfF PM']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['R'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['S'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO'           
        elif df.iloc[ROW]['DOTMLPF Portion Mark']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['R'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['S'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO'
        elif df.iloc[ROW]['AWfC PM']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['R'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['S'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO'     
        else:
            wb['Sheet1']['R'+str(ROW+7)]='' ### this portion marks implecations as the highest classification of WfF, DOTMLPF or AWfC
        if 'Hasty DOTMLPF-P2' in df:
            CONCAT= str(df.iloc[ROW]['AWfC'])+', '+str(df.iloc[ROW]['Hasty DOTMLPF-P'])+', '+str(df.iloc[ROW]['Hasty DOTMLPF-P2'])+', '+str(df.iloc[ROW]['Hasty DOTMLPF-P3'])+' '+str(df.iloc[ROW]['Hasty DOTMLPF-P4'])+', '+str(df.iloc[ROW]['WfF'])+', '+str(df.iloc[ROW]['WfF2'])+', '+str(df.iloc[ROW]['WfF3'])+'.'
            CONCAT = CONCAT.replace('Choose an item.', '')
        elif 'Hasty DOTMLPF-P' in df:
            CONCAT= str(df.iloc[ROW]['AWfC'])+', '+str(df.iloc[ROW]['Hasty DOTMLPF-P'])+', '+str(df.iloc[ROW]['WfF'])+'.'
        else:
            CONCAT=""
        wb['Sheet1']['T'+ str(ROW+7)] = CONCAT# concatenates DOTMLPF x 4 WfF x 3 and Awfc free text block then removes and choose an items left in
        wb['Sheet1']['U'+ str(ROW+7)] = "Unclassified"
        if df.iloc[ROW]['Mission Statement PM']=='(U//FOUO)' or 'U//FOUO':
            wb['Sheet1']['Z'+str(ROW+7)]='FOUO' 
            wb['Sheet1']['AA'+str(ROW+7)]='USA'
            wb['Sheet1']['D1']='FOUO'                #mission statement portion marking
            wb['Sheet1']['AB'+ str(ROW+7)] = df.iloc[0]['Mission']
        if 'First Name' in df_CONTACT:
            wb['Sheet1']['AC'+ str(ROW+7)] = df_CONTACT.iloc[0]['First Name']
        if 'Last Name' in df_CONTACT:
            wb['Sheet1']['AD'+ str(ROW+7)] = df_CONTACT.iloc[0]['Last Name']
        if 'Rank' in df_CONTACT:
            wb['Sheet1']['AE'+ str(ROW+7)] = df_CONTACT.iloc[0]['Rank']
        if "Email" in df_CONTACT:
            wb['Sheet1']['AG'+ str(ROW+7)] = df_CONTACT.iloc[0]['Email']
        if 'Phone Number' in df_CONTACT:
            wb['Sheet1']['AH'+ str(ROW+7)] = df_CONTACT.iloc[0]['Phone Number']
        if 'AWG Unit' in df:
            wb['Sheet1']['AI'+ str(ROW+7)] = df.iloc[0]['AWG Unit']       # adds authotr contact info 
        ROW = ROW-1 #restart while loop next observation     
############################ save the excel file   ###################################################################################
#wb.save(filename[:-4]+'.xlsx')
wb.save(XLfilename)
################################ convert all the zip back to docx format ########################################################
z.close()
f.close()
filename=docxfile
#rename(filename, filename[:-4] + '.docx')
pre, ext = os.path.splitext(filename)
os.rename(filename, pre + '.docx')
exit
