# -*- coding: utf-8 -*-
"""
Created on Wed Feb 26 20:47:41 2020

@author: tostg

SUMMARY:

Script for archiving and keeping track of versions of files in a given 
folder.     
"""

import os
import shutil
import datetime

       

#folder = r'\\sweco.se\NO\Oppdrag\SVG\35218\10215682_E39_Sykkelstamveien_Schancheholen-_Sørmarka\000\07 Modeller - Tegninger\02 Fagmodeller'
#archive = r'\\sweco.se\NO\Oppdrag\SVG\35218\10215682_E39_Sykkelstamveien_Schancheholen-_Sørmarka\000\07 Modeller - Tegninger\02 Fagmodeller\_Arkiv'

folder = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test'
archive = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test\_Arkiv'

logFile = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test\_Arkiv\logg.txt'

# List of file extentions to include in the archiving process
include = ['.ifc','.dwg','.nwc']
# List of folder that will be ignored in the archiving process
exclude = set(['_Arkiv', '_old'])

# Get yesterdays date
today = datetime.datetime.today()
yesterday = (today - datetime.timedelta(100)).strftime('%Y-%m-%d')
print(yesterday)



n = 0

# Walk throuhg alle files and directories in given folder
for root, dirs, files in os.walk(folder):
    # Do not walk throuhg folders given in the exclude set 
    dirs[:] = [d for d in dirs if d not in exclude]    
    
    # Check all files in the given folder of a the filetypes given in 
    # the include list. Get last updated date from each file
    for file in files:
        
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
                print()
                print("HEY! I found a new file:")
                print(file)
                print("Lets see.. What should I do with you, my friend?")
                # Folder to archive this file
                arcDir = os.path.join(archive, file) 
                # Archived file will have the date first in its file name
                newFileName = fileDate + '_' + file
                newFile = os.path.join(arcDir, newFileName)

                # Create a folder for each file if its not already there
                if not os.path.isdir(arcDir):
                    print("You poor fellow, why have you not a home?")
                    print("...")
                    print("Let me make it for you")
                    os.mkdir(arcDir)
                    print("See. Its warm and cozy, your welcome")
        
                    # Try to make a copy of the original file to the archive directory
                if os.path.isfile(newFile):
                    print("All good here")
                    print("Next")
                    print()
                
                if not os.path.isfile(newFile):
            
                    print("Oh shait")
                    print("Haven't you moved in yet?")
                    print("...")
                    shutil.copy(orignFile, newFile)
                    n += 1
                    print("Ahh... There you go, enjoy you stay")
                    print("From now and till the end of times")
                    print("mohahahahahaha...")

logg = fileDate + "\t|\t" + str(n) + '\n'

with open(logFile, 'a') as f:
    f.write(logg)
