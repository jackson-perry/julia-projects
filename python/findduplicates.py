# -*- coding: utf-8 -*-
"""
Created on Wed May 13 21:17:27 2020

@author: jacks
"""
from __future__ import print_function
import os
import hashlib
import datetime
import time
import shutil

 
#%% build archive and faster duplicate search
''' this function will find all files that have the same contents regardless of name, it checks all files over minSize
'''
def FindDuplicateFiles(pth, minSize = 0, hashName = "md5"):
    knownFiles = {}
    
	#Analyse files
    for root, dirs, files in os.walk(pth):
        for filename in files:
            fullFiname = os.path.join(root, filename)
            try:
                isSymLink = os.path.islink(fullFiname)
                if isSymLink:
                    continue # Skip symlinks
                size = os.path.getsize(fullFiname)
            
                if size < minSize:
                    continue
                if size not in knownFiles:
                    knownFiles[size] = {}
                    h = hashlib.new(hashName)
                    h.update(open(fullFiname, "rb").read())
                    hashed = h.digest()
            except (OSError,):
                continue         #skip files that can't be read due to permissions etc.
            if hashed in knownFiles[size]:
                fileRec = knownFiles[size][hashed]
                fileRec.append(fullFiname)
            else:
                knownFiles[size][hashed] = [fullFiname]
 
	#Print result
    sizeList = list(knownFiles.keys())
    sizeList.sort(reverse=True)
    outFile=open(DESTINATION+'/duplicateList.txt','w')
    deleteFile=open(DESTINATION+'/erasethese.txt',"w")
    for size in sizeList:
        filesAtThisSize = knownFiles[size]
        for hashVal in filesAtThisSize:
            if len(filesAtThisSize[hashVal]) < 2:
                continue
            fullFinaList = filesAtThisSize[hashVal]
            outFile.write("=======Duplicate=======\n")
            count=-1
            for fullFiname in fullFinaList:
                count=count+1
                st = os.stat(fullFiname)
                isHardLink = st.st_nlink > 1 
                infoStr = []
                if isHardLink:
                    infoStr.append("(Hard linked)")
                fmtModTime = datetime.datetime.utcfromtimestamp(st.st_mtime).strftime('%Y-%m-%dT%H:%M:%SZ')
                outFile.write(fmtModTime+" "+str(size)+os.path.relpath(fullFiname, pth)+" ".join(infoStr)+"\n")
                if count > 0:
                    deleteFile.write(os.path.relpath(fullFiname, pth)+'\n')
    
    outFile.close()
    deleteFile.close()
'''
this function will scan the target directory pth and all subdirectories and copy all files of the filetype and copy them to the destination after assigning them a 7digit serial number
the serial number prevents files with the same name from overwriting each other 
'''
start_time = time.time()
def BuildArchive(pth,destination, filetype):
    sn= 0
    for root, dirs, files in os.walk(pth):
        for filename in files:
            try:
                if filename[-len(filetype):] == filetype:
                    fullFiname= os.path.join(root, filename)
                    shutil.copy(fullFiname, destination+"/"+f'{sn}'.zfill(7)+'  '+filename) 
                    sn=sn+1
            except(OSError,):
                    continue
 

#%%  erase duplicates lited in a textfile
def RmDuplicates(textlist):
    with open(textlist) as f:
     data=f.readlines()
     for i in data:
         os.remove(i.strip())
#%% 100% hash search   

def flatDup(parentFolder):
    # Dups in format {hash:[names]}
    dups = {}
    for dirName, subdirs, fileList in os.walk(parentFolder):
        print('Scanning %s...' % dirName)
        for filename in fileList:
            print('Scanning %s' % filename)
            # Get the path to the file
            path = os.path.join(dirName, filename)
            # Calculate hash
            file_hash = hashfile(path)
            # Add or append the file path
            if file_hash in dups:
                dups[file_hash].append(path)
            else:
                dups[file_hash] = [path]
    printResults(dups, parentFolder)
 
 
def hashfile(path, blocksize = 65536):
    afile = open(path, 'rb')
    hasher = hashlib.md5()
    buf = afile.read(blocksize)
    while len(buf) > 0:
        hasher.update(buf)
        buf = afile.read(blocksize)
    afile.close()
    return hasher.hexdigest()
 
 
def printResults(dict1, parentFolder):
    outFile=open(parentFolder+'/duplicateList.txt','w')
    deleteFile=open(parentFolder+'/erasethese.txt','w')
    results = list(filter(lambda x: len(x) > 1, dict1.values()))
    if len(results) > 0:
        outFile.write('Duplicates Found:\n')
        outFile.write('The following files are identical. The name could differ, but the content is identical\n')
        outFile.write('___________________\n')
        for result in results:
            counter =0
            for subresult in result:
                counter= counter+1
                outFile.write('\t\t%s\n' % subresult)
                if counter >1:
                    deleteFile.write('\t\t%s\n' % subresult)
            outFile.write('___________________\n')
 
    else:
        outFile.write('No duplicate files found.')
        deleteFile.write('No duplicate files found.')
 
#%% copy folder and run quick duplicate search
start_time = time.time()
FILE_EXT=".PNG"
SOURCE= "D:/amy photo backup"
DESTINATION= "Y:/Amy/photo backup"
if not os.path.exists(DESTINATION):
    os.makedirs(DESTINATION)
BuildArchive(SOURCE,DESTINATION,filetype=FILE_EXT.upper())
BuildArchive(SOURCE,DESTINATION,filetype=FILE_EXT.lower())
print("files copied in --- %s seconds ---" % (time.time() - start_time))
FindDuplicateFiles(DESTINATION)
print("duplicates found in --- %s seconds ---" % (time.time() - start_time))     