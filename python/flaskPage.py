# -*- coding: utf-8 -*-
"""
Created on Sat Jun 22 21:38:42 2019

@author: jacks
"""


from openpyxl import load_workbook
import zipfile
import pandas as pd
from bs4 import BeautifulSoup as BS
from os import rename
import os
from flask import Flask, request, redirect, url_for
from werkzeug import secure_filename
import sys
sys.path.insert(0, '/home/jacks/Desktop/SHREDDER/')

UPLOAD_FOLDER = '/home/jacks/Desktop/SHREDDER/'


app = Flask(__name__, template_folder='C:\\Users\\jacks\\Desktop\\SHREDDER\\python\\templates')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
filename1=()
filename2=()

def shred():
    ############## convert each docx file in zip file so XML files can be read  ###########################
    rename(filename1, filename1[:-5] + '.zip')
    docxfile= filename1[:-5] + '.zip'
###################################### Open File ########################################################
    z = zipfile.ZipFile(docxfile)
    f =z.open('word/document.xml')
    ################################ beutiful soup is used to read and navigate teh xml file################################
    soup = BS(f, 'xml')  # entire xml; file is called soup
    soupdocument=soup.document  #document only child node of soup contains MS office schemas used in the document not 
    soupbody=soupdocument.body  # body is the only child node of document contains all the paragraphs (p) and structured document tags(sdt)
    ODR=soupbody.contents[5].contents[1]   # ODR contains the Observation discusion recomendation section structured document tag (the sixth xml element in the documents body counting from 0)
    CHILDcount=len(ODR.contents)-1
    ROW=CHILDcount # the ODR variable has the number of repeating sections in it, we subtract one because the computer counts from zero 
    CONTACT=soupbody.contents[3].contents[1] # this contains the AWg POC section structured documetn tag ( the 4teh xml elemetn in the documents body counting from 0)
    CONTACTcount=len(CONTACT.contents)-1
    ######################### pulls all the tags and text form each descendent node of the structured document tags, overwrites the ODR section, only produces the last ODr repeating section##############################
    sdt=soup.find_all("sdt")  #make a list of all structured documetn tags
    xdict={}   #make a dictionary
    for i in sdt:        #iterate through each structured document tag  and append the tag and the text to a dictionary
        xdict.update({str(i.tag) : i.get_text()}) #make sure to store the tag a string or it will be an xml tag
    ########################## creates a pandas dataframe to hold all the data extracted key is the tag index is the text string #######################
    df = pd.DataFrame(data=xdict, index=[CHILDcount])
    ########################## this section captures all but the last ODr and adds it to the dataframe ######################  
    ODRdict ={}     #make another dictionary for the ODR and section labels
    CHILDcount= CHILDcount-1 # The last observation is already done we will work form last to first below
    while CHILDcount >= 0:  #0 is the indes location of the first odr section we are working form last to first
        CHILD= ODR.contents[CHILDcount].find_all('sdt') #find all Structured Document Tag in the give ODR repeating section
        for i in CHILD: #iterate through that list of sdt
            ODRdict.update({str(i.tag):i.get_text()}) #add each tag and text to a dictionary make sure to store the tag a string or it will be an xml tag
            
        df_ODR=pd.DataFrame(ODRdict, index=[CHILDcount]) #convert the dictionary into a dataframe
        df=pd.concat([df,df_ODR], axis=0,sort=True) #append that dataframe into the main datafram that already holds the header info and the last ODr section
        CHILDcount=CHILDcount-1 # take one from our counter so we will begin the loop on the next closest to first ODR section
    ###################### this section handels the repeating contacts section  #############################
    CONTACTdict ={}  #make another dictionary for the repeating contacts section and thier labels
    CONTACTchild = CONTACT.contents[0].find('sdt')
    CONTACTdict.update({str(i.tag):i.get_text()})
    
    df_CONTACT=pd.DataFrame(ODRdict, index=[ROW]) #convert the dictionary into a dataframe
    df=pd.concat([df,df_CONTACT], axis=0,sort=True) #append that dataframe into the main datafram that already holds the header info and the last ODr section
      
        
        
    df2 =list(df.columns.values)     # pull out all the tags from the first column
    df2= [l.strip('<w:tag w:val="') for l in df2]  # remove the XML stuff form the front of the tag
    df2= [l.strip('"/>') for l in df2] # remove the XML stuff form the back of the tag
    df.columns = df2 # rename all the column headers thier tag names without all the XML markup
    del df['ODR Repeating Section'] #this is a duplicate holding all the lasT ODR in one long string, it is no longer needed
    df.to_excel('test3.xlsx')  # creates an excel file containing the data frame  
    ##########################   create excel sheet with date and CONOP number  ###############################
    
    wb = load_workbook(filename2)  #click on the excel templae from jllis
    wb['Sheet1']['B4']=df.iloc[0]['End Date']            #sends enddate from first frow of dataframe to excel file cell AJ2
    wb['Sheet1']['AJ3']='CONOP'+' '+df.iloc[0]['CONOP_Year']+'-'+df.iloc[0]['CONOP #']       #concatenats CONOP year and number and passes to spreadsheet AJ32
    
    while ROW >=0: # check if we have any observations left
            wb['Sheet1']['A'+ str(ROW+7)] = "Unclassified"        
            if df.iloc[ROW]['TOI Portion Mark']=='(U//FOUO)' or 'U//FOUO':
                wb['Sheet1']['B'+str(ROW+7)]='FOUO' 
                wb['Sheet1']['C'+str(ROW+7)]='USA'
                wb['Sheet1']['D1']='FOUO'  #these four rows do portion marking of the topic
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
                wb['Sheet1']['D1']='FOUO'                     # does observation portion marking
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
                wb['Sheet1']['D1']='FOUO'                     # does discussion portion marking
            if 'Discussion' in df:
                wb['Sheet1']['L'+ str(ROW+7)] = df.iloc[ROW]['Discussion'] # discusiion to excel
            wb['Sheet1']['M'+ str(ROW+7)] = "Unclassified"
            if df.iloc[ROW]['Recomendation Portion Mark']=='(U//FOUO)' or 'U//FOUO':
                wb['Sheet1']['N'+str(ROW+7)]='FOUO' 
                wb['Sheet1']['O'+str(ROW+7)]='USA'
                wb['Sheet1']['D1']='FOUO'                       # does Recommendation portion marking
            if 'Recommendation' in df:
                wb['Sheet1']['P'+ str(ROW+7)] = df.iloc[ROW]['Recommendation'] #recommendation to excel
            wb['Sheet1']['Q'+ str(ROW+7)] = "Unclassified"
            if df.iloc[ROW]['WfF PM']=='(U//FOUO)' or 'U//FOUO':
                wb['Sheet1']['R'+str(ROW+7)]='FOUO' 
                wb['Sheet1']['S'+str(ROW+7)]='USA'
                wb['Sheet1']['D1']='FOUO'           
            if df.iloc[ROW]['DOTMLPF Portion Mark']=='(U//FOUO)' or 'U//FOUO':
                wb['Sheet1']['R'+str(ROW+7)]='FOUO' 
                wb['Sheet1']['S'+str(ROW+7)]='USA'
                wb['Sheet1']['D1']='FOUO'
            if df.iloc[ROW]['AWfC PM']=='(U//FOUO)' or 'U//FOUO':
                wb['Sheet1']['R'+str(ROW+7)]='FOUO' 
                wb['Sheet1']['S'+str(ROW+7)]='USA'
                wb['Sheet1']['D1']='FOUO'                ### this portion marks implecations as the highest classification of WfF, DOTMLPF or AWfC
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
                wb['Sheet1']['V'+str(ROW+7)]='FOUO' 
                wb['Sheet1']['W'+str(ROW+7)]='USA'
                wb['Sheet1']['D1']='FOUO'                #mission statement portion marking
                wb['Sheet1']['X'+ str(ROW+7)] = df.iloc[0]['Mission']
                if 'First Name' in df:
                    wb['Sheet1']['Y'+ str(ROW+7)] = df.iloc[0]['First Name']
                if 'Last name' in df:
                    wb['Sheet1']['Z'+ str(ROW+7)] = df.iloc[0]['Last Name']
                if 'Rank' in df:
                    wb['Sheet1']['AA'+ str(ROW+7)] = df.iloc[0]['Rank']
                if "Email" in df:
                    wb['Sheet1']['AC'+ str(ROW+7)] = df.iloc[0]['Email']
                if 'Phone Number' in df:
                    wb['Sheet1']['AD'+ str(ROW+7)] = df.iloc[0]['Phone Number']
                if 'AWG Unit' in df:
                    wb['Sheet1']['AE'+ str(ROW+7)] = df.iloc[0]['AWG Unit']       # adds authotr contact info 
                ROW = ROW-1 #restart while loop next observation     
    ############################ save the excel file   ###################################################################################
    #wb.save(filename[:-4]+'.xlsx')
    wb.save(filename2)
    ################################ convert all the zip back to docx format ########################################################
    z.close()
    f.close()
    filename=docxfile
    rename(filename, filename[:-4] + '.docx')
    exit

@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file1 = request.files['file1']
        file2 = request.files['file2']
        if file1 > 0:
            filename1 = secure_filename(file1.filename)
            file1.save('/var/www/', filename1)
        if file2 > 0:
            filename2 = secure_filename(file2.filename)
            file2.save(os.path.join('/var/www/', filename2))
            shredd()
        return render_template('index.html') 
        
        

    return """
    <!doctype html>
    <title>Upload new File</title>
    <h1>Upload new File</h1>
    <form action="" method=post enctype=multipart/form-data>
      <p><input type=file name=file1>
          <input type=file name=file2>
         <input type=submit value=Upload>
    </form>
    <p>%s</p>
    """ % "<br>".join(os.listdir(app.config['UPLOAD_FOLDER'],))
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
