import os
import win32com.client
from Tkinter import *
import tkFileDialog

# visio2VSDX.py - Requires Visio 2013+
#
# This script simply renames all .vsd files found that have whitespace in the filename, then launches
# Visio and loads and saves them as .vsdx files. On startup a dialog pops asking for the src files
# folder where it loads and saves everything.
#
# The Companion script, visioVSDXCompare.py, compares same named .vsdx files and reports
# differences found between them in an html diff report that can be loaded and viewed in your favorite
# browser.


# Ensures none of the filenames have spaces
def RenameFiles(srcpath):
    files = os.listdir(srcpath)
    visiofiles = [f for f in files if f.endswith('.vsd')]

    count=0

    print "Renaming files with spaces..."

    for file in visiofiles:
        if (' ' in file) == True:
            count+=1
            newfile = file.replace(' ', '')
            os.rename(srcpath + '/' + file, srcpath + '/' + newfile)

    print "Done. %d files renamed" % count


# Save all VSD files found as VSDX files
def ConvertFiles(srcpath):

    # Visio 2013 won't save with Unix folder separators, must use DOS style separators
    srcpath = srcpath.replace('/', '\\')

    print "Saving Visio files to VSDX format... please wait"
    count=0
    files = os.listdir(srcpath)
    visiofiles = [i for i in files if i.endswith('.vsd')]

    visio = win32com.client.gencache.EnsureDispatch("Visio.InvisibleApp")
    visio.AlertResponse = 1

    for file in visiofiles:
        count+=1
        doc = visio.Documents.Open(srcpath + "\\" + file)
        destfile = os.path.splitext(file)[0] + ".vsdx"
        doc.SaveAs(srcpath + "\\" + destfile)

    visio.Quit()
    print "Done. %d Visio files save to VSDX format" % count


# Main routine
def main():
    root=Tk()
    root.withdraw()  #Get rid of visible default TK base window
    srcpath = tkFileDialog.askdirectory(initialdir='/', title='Select Visio Documents Folder')

    RenameFiles(srcpath)
    ConvertFiles(srcpath)

if __name__ == "__main__":
    main()

