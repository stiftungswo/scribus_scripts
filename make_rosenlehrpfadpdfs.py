#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ABOUT THIS SCRIPT:
Import CSV data files as tables into Scribus
"""
 
 
import sys
 
try:
    # Please do not use 'from scribus import *' . If you must use a 'from import',
    # Do so _after_ the 'import scribus' and only import the names you need, such
    # as commonly used constants.
    import scribus
except ImportError,err:
    print "This Python script is written for the Scribus scripting interface."
    print "It can only be run from within Scribus."
    sys.exit(1)
 
#########################
# YOUR IMPORTS GO HERE  #
#########################
import csv
 
def getCSVdata():
    """opens a csv file, reads it in and returns a 2 dimensional list with the data"""
    #csvfile = scribus.fileDialog("csv2table :: open file", "*.csv")
    csvfile = "/home/toast/diskstation/intern/Publikationen/Info-Tafeln/Rosenlehrpfad/data.csv"
    if csvfile != "":
        try:
            reader = csv.reader(file(csvfile))
            datalist=[]
            for row in reader:
                rowlist=[]
                for col in row:
                    rowlist.append(col)
                datalist.append(rowlist)
            return datalist
        except Exception,  e:
            scribus.messageBox("csv2table", "Could not open file %s"%e)
    else:
        sys.exit
 
 
def create(argv):
    """
        - create a page using the "badge" master page.
        - add the text field with the name
        - add the text with the role in the project
    """
    #########################
    #  YOUR CODE GOES HERE  #
    #########################
    userdim=scribus.getUnit() #get unit and change it to mm
    scribus.setUnit(scribus.UNIT_POINTS)
 
    data = getCSVdata()
 
    scribus.progressTotal(len(data)+1)
    scribus.setRedraw(False)

    i = 0;
    
    #Select Output Directory and define some Vars based on it.
    #baseDirName = scribus.fileDialog("Rosenlehrpfad - Wahl des Ausgabeverzeichnisses", "*","./" ,False, False, True)
    baseDirName = os.path.abspath("./output")
    tmpFileName = os.path.join(baseDirName, "tmp.sla")
    idtextstring = "Rosenid"
    gertextstring = "RosentextDeutsch"
    lattextstring = "RosentextLatein"
    
    if not os.path.exists(baseDirName):
      os.makedirs(baseDirName)
    else:
      for the_file in os.listdir(baseDirName):
	file_path = os.path.join(baseDirName, the_file)
	try:
	  if os.path.isfile(file_path):
	    os.unlink(file_path)
	except Exception, e:
	  print e
    
    
    #Make Document Temporary
    scribus.saveDoc()
    scribus.saveDocAs(tmpFileName)
    scribus.messagebarText("Dokument in Bearbeitung, bitte warten")
    scribus.closeDoc()
    scribus.messagebarText("Dokument in Bearbeitung, bitte warten")
    scribus.openDoc(tmpFileName)
    
    
    scribus.selectText(0, 1, idtextstring)
    idfont = scribus.getFont(idtextstring)
    idfonSize = scribus.getFontSize(idtextstring)
    idtextcolor = scribus.getTextColor(idtextstring)
    idtextshade = scribus.getTextShade(idtextstring)
    
    scribus.selectText(0, 1, gertextstring)
    gerfont = scribus.getFont(gertextstring)
    gerfonSize = scribus.getFontSize(gertextstring)
    gertextcolor = scribus.getTextColor(gertextstring)
    gertextshade = scribus.getTextShade(gertextstring)
    
    scribus.selectText(0, 1, lattextstring)
    latfont = scribus.getFont(lattextstring)
    latfonSize = scribus.getFontSize(lattextstring)
    lattextcolor = scribus.getTextColor(lattextstring)
    lattextshade = scribus.getTextShade(lattextstring)
    
    for row in data:
        #if i > 0 :
	  #sys.exit(1)
	#scribus.messageBox("csv2table", row)
	roseid = row[0]
	rosenameger = row[1]
	rosenamelat = row[2]
        #scribus.messageBox("csv2table", roseid+' '+rosenameger+' ('+rosenamelat+')')
        
        #Insert the Right Text into the Right Position
        
        #Write ID into ID Field
	scribus.deleteText(idtextstring)
	scribus.setText(roseid, idtextstring)
	scribus.setFont(idfont, idtextstring)
	scribus.setFontSize(idfonSize, idtextstring)
	scribus.setTextColor(idtextcolor, idtextstring)
	scribus.setTextShade(idtextshade, idtextstring)
	
	#Write German Name into german Name Field
	scribus.deleteText(gertextstring)
	scribus.setText(rosenameger, gertextstring)
	scribus.setFont(gerfont, gertextstring)
	scribus.setFontSize(gerfonSize, gertextstring)
	scribus.setTextColor(gertextcolor, gertextstring)
	scribus.setTextShade(gertextshade, gertextstring)
	
	#Write Latin Name into Latin Name field
	scribus.deleteText(lattextstring)
	scribus.setText(rosenamelat, lattextstring)
	scribus.setFont(latfont, lattextstring)
	scribus.setFontSize(latfonSize, lattextstring)
	scribus.setTextColor(lattextcolor, lattextstring)
	scribus.setTextShade(lattextshade, lattextstring)
	
	#Make a PDF of Edited Page
	pdfname = roseid+'. '+rosenameger+' ('+rosenamelat+')'
	pdf = scribus.PDFfile()
	pdf.info = pdfname
	pdf.pages = [1]
        pdf.file = os.path.join(baseDirName, pdfname+'.pdf')
        pdf.save()
        
        #Incarese Rowcount and update Progress
        i = i + 1;
        scribus.progressSet(i)
    
    #Done Close tmpfile and remove it
    scribus.closeDoc()
    os.remove(tmpFileName)
    
    scribus.progressReset()
    scribus.statusMessage("Done")
 
 
def main(argv):
    """The main() function disables redrawing, sets a sensible generic
    status bar message, and optionally sets up the progress bar. It then runs
    the main() function. Once everything finishes it cleans up after the create()
    function, making sure everything is sane before the script terminates."""
    currentDoc = scribus.getDocName()
    try:
        scribus.statusMessage("Importing .csv table...")
        scribus.progressReset()
        create(argv)
        if scribus.haveDoc():
	  scribus.closeDoc()
	scribus.openDoc(currentDoc)
    finally:
        # Exit neatly even if the script terminated with an exception,
        # so we leave the progress bar and status bar blank and make sure
        # drawing is enabled.
        if scribus.haveDoc():
            scribus.setRedraw(True)
            scribus.closeDoc()
	scribus.openDoc(currentDoc)
        scribus.statusMessage("")
        scribus.progressReset()
 
# This code detects if the script is being run as a script, or imported as a module.
# It only runs main() if being run as a script. This permits you to import your script
# and control it manually for debugging.
if __name__ == '__main__':
    main(sys.argv)
