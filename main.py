# -*- coding: utf-8 -*-
"""
Created on Wed Feb 26 20:47:41 2020

@author: Thomas Østgulen

SUMMARY:

Script for archiving and keeping track of versions of files in a given 
folder. It will also send an e-mail to a list of recipients if any 
models has been updated.     
"""

import os
import shutil
import datetime
import subprocess
import pymsteams

def htmlTable(folderPath, modelList):
    htmlTableStr = ("<table align='left'>"
                "<colgroup>"
                "<col style='width:100%'>"
                "</colgroup>"
                "<col width='120'>")
    htmlTableHead = ("<tr>"
                    "<th align='left'>{0}</th></tr>").format(
        "Filnavn")
    htmlTableStr += htmlTableHead

    ### HTML Tabel content ### 
    for model in modeList:
        
        htmlTableRow = '<tr><td><a href="{0}">{1}</a></td></tr>'.format(folderPath ,model[2]) 
        htmlTableStr += htmlTableRow

    # Close the HTML table
    htmlTableStr += "</table>"

    return htmlTableStr




# Marvin @ Gr_Ugly Digital Group
myTeamsMessage = pymsteams.connectorcard("https://outlook.office.com/webhook/fc8b11e8-2a57-4b9a-abfa-6ee266b746f7@b7872ef0-9a00-4c18-8a4a-c7d25c778a9e/IncomingWebhook/ada2d7a0599b4208ad13feb32f64250e/1d19ccf3-fe7c-4342-9af5-9b17074ac4b1")
# Create the two individual messages in the Teams message
myMessageSectionFM = pymsteams.cardsection()
myMessageSectionGM = pymsteams.cardsection()


project = r'\\sweco.se\NO\Oppdrag\SVG\35218\10215682_E39_Sykkelstamveien_Schancheholen-_Sørmarka\000\07 Modeller - Tegninger'
#project = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium'

folderFM = os.path.join(project, '02 Fagmodeller')
folderGM = os.path.join(project, '01 Grunnlagsmodeller')

#folder = r'\\sweco.se\NO\Oppdrag\SVG\35218\10215682_E39_Sykkelstamveien_Schancheholen-_Sørmarka\000\07 Modeller - Tegninger\02 Fagmodeller'
#archive = r'\\sweco.se\NO\Oppdrag\SVG\35218\10215682_E39_Sykkelstamveien_Schancheholen-_Sørmarka\000\07 Modeller - Tegninger\02 Fagmodeller\_Arkiv'

#folder = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test'
#archive = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test\_Arkiv'

folders = [folderFM, folderGM]

bimboi = r'C:\Scripts\SSV\Python'
#varslingFolder = os.path.join(project, 'RIB\_Generelt grunnlag\_Python\Varsling')
htmldoc = os.path.join(bimboi, 'htmldoc.html')
fagmodellEpost = os.path.join(bimboi, 'fagmodellEpost.vbs')

#logFile = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test\_Arkiv\logg.txt'




# List of file extentions to include in the archiving process
include = ['.ifc','.dwg','.nwc']
# List of folder that will be ignored in the archiving process
exclude = set(['_Arkiv', '_old'])

# Get yesterdays date
today = datetime.datetime.today()
yesterday = (today - datetime.timedelta(1)).strftime('%Y-%m-%d')
#print(yesterday)

#print(htmldoc)


j = 0
listFagmodeller = []
listGrunnlagsmodeller = []
# Walk throuhg alle files and directories in given folder
for folder in folders:
    n = 0
    archive = os.path.join(folder, '_Arkiv')
    for root, dirs, files in os.walk(folder):
        # Do not walk throuhg folders given in the exclude set 
        dirs[:] = [d for d in dirs if d not in exclude]    
        
        # Check all files in the given folder of a the filetypes given in 
        # the include list. Get last updated date from each file
        for file in files:
            #print(file)
            base, extention = os.path.splitext(file)
            if extention in include:
            
                orignFile = root + '\\' + file
                
                (mode, ino, dev, nlink, uid, gid,
                size, atime, mtime, ctime) = os.stat(orignFile)
            
                orignDateTime = str(datetime.datetime.fromtimestamp(mtime))
                fileDate = orignDateTime.split(" ")[0]
            
                # Files updatet yestarday or today will be archived to keep track of
                # an eventual new version
                if fileDate >= yesterday:
                    # Folder to archive this file
                    arcDir = os.path.join(archive, file) 
                    # Archived file will have the date first in its file name
                    newFileName = fileDate + '_' + file
                    newFile = os.path.join(arcDir, newFileName)

                    # Create a folder for each file if its not already there
                    if not os.path.isdir(arcDir):                        
                        os.mkdir(arcDir) 
                    # Try to make a copy of the original file to the archive directory             
                    if not os.path.isfile(newFile):
                        shutil.copy(orignFile, newFile)
                        n += 1
                        # Deside if its fagmodeller or grunnlagsmodeller
                        if j == 0:
                            listFagmodeller.append((base,extention, file, orignFile))
                        elif j == 1:
                            listGrunnlagsmodeller.append((base, extention, file, orignFile))

    logg = fileDate + "\t|\t" + str(n) + '\n'
    logFile = os.path.join(archive, 'logg.txt')
    with open(logFile, 'a') as f:
        f.write(logg)


# Teams message
myTeamsMessage.title("")
myTeamsMessage.text(f"I går, {yesterday}, ble følgende filer oppdatert")
myTeamsMessage.color('red')

# Make Teams tables
for fm in listFagmodeller:
    myMessageSectionFM.addFact(f"{fm[0]}", f"{fm[1]}")

for gm in listGrunnlagsmodeller:
    myMessageSectionGM.addFact(f"{gm[0]}", f"{gm[1]}")

myMessageSectionFM.text("Fagmodeller")
myMessageSectionGM.text("Grunnlagsmodeller")

##### Create HTML #####

### HTML Header ###
htmlFile = ("<!DOCTYPE html><html><head>"
            "<title>Filer som er endret siden sist:</title>"
            "</head><body>")
# Table Fagmodeller
htmlFile += htmlTable(folders[0], listFagmodeller)
# Table Grunnlagsmodeller
htmlFile += htmlTable(folders[0], listFagmodeller)    

### HTML summary ###
htmlSummary = "<p>Det er blitt oppdatert {0} fagmodeller og {2} grunnlagsmodeller siden {1}.</p>".format(
    len(listFagmodeller),
    yesterday,
    len(listGrunnlagsmodeller))

##### Structure the HTML document #####
# Summary
htmlFile += htmlSummary

# Table 
htmlFile += htmlTable

# Add easter egg
htmlGif = ('<iframe src="https://giphy.com/embed/gjgWQA5QBuBmUZahOP" '
           'width="480" height="480" frameBorder="0" '
           'class="giphy-embed" allowFullScreen>'
           '</iframe>')
htmlFile += htmlGif

# Close the HTML code
htmlFile += "</body></html>"

#### End HTML #### 

#### Create HTML file ####
f = open(htmldoc, 'w')
f.write(htmlFile)
f.close()

##### Send email if new dicipline models detected #####
#print(len(listFagmodeller))
if len(listFagmodeller) > 0 or len(listGrunnlagsmodeller) > 0:
    #print("Sending...")
    subprocess.call(["cscript", fagmodellEpost])
    myTeamsMessage.send()

    #print("Sent!")
#else:
#    print("Ingen fagmodeller --> ingen epost")
