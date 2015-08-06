import os
from datetime import date
import zipfile
import difflib
from Tkinter import *
import tkFileDialog

# visioVSDXCompare.py
#
# This script compares Visio VSDX files at the XML level and reports differences, if any,
# in an HTML output file. On startup it asks for the old files folder, the new files folder, and the
# diff report destination folder.
#
# The Companion script, visio2VSDX.py, prepares a folder of VSD files by renaming
# to exclude filename whitespace and then saving as Visio .vsdx format

def getOldFiles(oldPath):
    files = os.listdir(oldPath)
    oldFiles = [f for f in files if f.endswith('.vsdx')]
    return oldFiles


def getNewFiles(newPath):
    files = os.listdir(newPath)
    newFiles = [f for f in files if f.endswith('.vsdx')]
    return newFiles


def compareFiles(oldPath, newPath, destPath, targetfile):
    # Visio .vsdx files are basically zipped files, and must opened as such scanning the
    # visio/pages virtual folder for diagram pages. Pages start at page1

    # prepare the new .vsdx file for compare
    new = zipfile.ZipFile(newPath + '/' + targetfile, "r")
    newPages = [f for f in new.namelist() if 'visio/pages/page' in f and 'visio/pages/pages' not in f]

    old = zipfile.ZipFile(oldPath + '/' + targetfile, "r")
    oldPages = [f for f in old.namelist() if 'visio/pages/page' in f and 'visio/pages/pages' not in f]

    pagenum = 1

    for page in newPages:
        if page in oldPages:
            newdata = new.read(page)
            moddata = newdata.replace('/><', '/>\n<')  # broke the lines on the xml tag ends to keep line length shorter
            newlines = moddata.splitlines(0)
            newFiltered = [line for line in newlines if '<Cell' not in line]  #strip <Cell/> noise

            olddata = old.read(page)
            moddata = olddata.replace('/><', '/>\n<')
            oldlines = moddata.splitlines(0)
            oldFiltered = [line for line in oldlines if '<Cell' not in line]  #strip <Cell/> noise

            d = difflib.HtmlDiff(wrapcolumn=100)

            # use HtmlDiff to compare, numlines=1 shows 1 line of context above/below the difference found
            diff = d.make_file(oldFiltered, newFiltered, fromdesc='Old File', todesc='New File', context=True, numlines=1)

            # write out the diff html file only if differences were found
            if ("No Differences Found" not in diff):
                diffFile = os.path.splitext(targetfile)[0] + '_page' + str(pagenum) + '.html'
                outfile = open(destPath + '/' + diffFile, 'w')
                outfile.write(diff)
                outfile.close()

            pagenum += 1

    new.close()
    old.close()


# Main routine
def main():
    root=Tk()
    root.withdraw() #Get rid of visible default TK base window
    newPath = tkFileDialog.askdirectory(initialdir='/', title='Select NEW VSDX Documents Folder')
    oldPath = tkFileDialog.askdirectory(initialdir='/', title='Select OLD VSDX Documents Folder')
    destPath = tkFileDialog.askdirectory(initialdir='/', title='Select Comparison Results Folder')

    oldFiles = getOldFiles(oldPath)
    newFiles = getNewFiles(newPath)

    # Open report.txt file
    rptFilename = destPath + '/' + 'ComparisonReport_' + str(date.today()) + '.txt'

    with open(rptFilename, 'w') as rptFile:
        rptFile.write('Comparison Report for %s\n\n' % str(date.today()))
        rptFile.write('Total OLD Files: %d\n' % len(oldFiles))
        rptFile.write('Total NEW Files: %d\n\n' % len(newFiles))

        # Compare the files
        for newFile in newFiles:
            if newFile in oldFiles:
                compareFiles(oldPath, newPath, destPath, newFile)
            else:
                rptFile.write('New file added: %s\n' % newFile)

        rptFile.write('\n\n')

        for oldFile in oldFiles:
            if oldFile not in newFiles:
                rptFile.write('Old file dropped: %s\n' % oldFile)

if __name__ == "__main__":
    main()