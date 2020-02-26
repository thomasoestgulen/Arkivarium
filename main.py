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
       

folder = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test'
archive = r'C:\Users\tostg\Documents\Python Scripts\Arkivarium\Test\_Arkiv'

# List of file extentions to include in the archiving process
include = ['.ifc','.dwg']
# List of folder that will be ignored in the archiving process
exclude = set(['_Arkiv'])

# Get yesterdays date
today = datetime.datetime.today()
yesterday = (today - datetime.timedelta(1)).strftime('%Y-%m-%d')


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
            # Folder to archive this file
            arcDir = os.path.join(archive, base)
            # Archived file will have the date first in its file name
            newFileName = fileDate + '_' + file
            newFile = os.path.join(arcDir, newFileName)

        # Create a folder for each file if its not already there
        if not os.path.isdir(arcDir):
            os.mkdir(arcDir)
        
        # Try to make a copy of the original file to the archive directory
        try:
            shutil.copy(orignFile, newFile)
            
        except:
            pass
            
                
