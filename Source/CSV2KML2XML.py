# Imports
import base64
import xml.etree.ElementTree as ET
import simplekml
import tkinter 
import os
import sys
import csv
import array as arr
import IPython
import time
import logging

from io import StringIO
from tkinter import Label, messagebox, Button, Listbox, END, Scrollbar, Frame
from tkinter.filedialog import askopenfilename, askopenfile
from tkinter.font import Font

# Initialize global variables
dataList = []
matchingGroup = []
matchedGroup = []
failedGroup = []
genFiles = ''

# Method to generate KML files from CSV
def CSV2KML():
    # Initialize variables
    global genCaseNumList, dataList, genFiles
    logger = logging.getLogger('CSV to KML')
    listbox.delete(0, END)

    # Define acceptable file types
    ftypes = [
        ('Microsoft Excel Comma Separated Values File', '*.csv'), 
    ]
    
    # Open file
    try:
        inputfile = askopenfile(mode = 'rt', filetypes = ftypes)
        logger.info('Successfully opened ' + str(inputfile.name))
    except Exception as e:
        logger.warning('ERROR opening ' + str(inputfile.name))
        logger.warning(e)
    
    # Convert file contents to list
    inputfile = list(csv.reader(inputfile))

    # Initialize variables for KML generation
    prev = [""]*7
    last = inputfile[-1]
    genFiles = True

    # If "KML Files" folder does not already exist, creates file
    os.makedirs(os.getcwd() + "/KML Files", exist_ok=True)

    # Index through input file and generate KML elements
    if (inputfile):
        for row in inputfile[1:]: 
            # Catches if CSV is missing a column 
            if(len(row) < 7):
                csvFormattingError()
                os._exit(0)
            
            # Calls method to generate KML elements
            if(prev[0] !=  row[0] or prev[0] == last[0]):
                generateKML(prev, row)
            prev = row
            dataList.append(row)
            if(prev[0] == last[0]):
                generateKML(prev, row)
            
        # Insert newly generated KML file names into the list box if they did not exist previously
        genCaseNumList = list(dict.fromkeys(genCaseNumList))    
        for case in genCaseNumList:
            listbox.insert(END, case)
        messagebox.showinfo("Done", "Your KML Files Have Been Generated!")
    
    logger.info("Completed KML File Generation")

# Method to generate KML elements
def generateKML(prev, row):
    # Initialize variables
    logger = logging.getLogger('CSV to KML')
    global genCaseNumList, dataList, matchingGroup, matchedGroup, failedGroup
    kml = simplekml.Kml()
    caseNum = prev[0]

    # Index through data under same case number
    for data in dataList:
        # Check if data is part of a group or a failed case number
        if(data[5] != "" and data[5] not in matchedGroup and data[0] not in failedGroup):
            # Index through all rows under same case number 
            for row2 in dataList:
                # If group number matches add to list
                if(row2[5] == data[5]):
                    matchingGroup.append(row2)
        
            # Index through all rows under same group number
            if (len(matchingGroup) > 1):
                # Create line string for KML with point overlay    
                firstRow = matchingGroup[0]
                secondRow = matchingGroup[1]
                line = kml.newlinestring(name = firstRow[1] + "; " + secondRow[1],coords = [(firstRow[4],firstRow[3],35),(secondRow[4],secondRow[3],35)])
                line.style.linestyle.width = 5
                line.altitudemode = simplekml.AltitudeMode.relativetoground
                kml.newpoint(name = firstRow[1], description = firstRow[2], coords = [(firstRow[4],firstRow[3])])
                kml.newpoint(name = secondRow[1], description = secondRow[2], coords = [(secondRow[4],secondRow[3])])
                matchedGroup.append(firstRow[5])
                # Empty contents of group matching for
                del matchingGroup[0:]
            elif (len(matchingGroup) == 1):
                newpoint = kml.newpoint(name = data[1], description = data[2], coords = [(data[4],data[3])])
                # Styling for outcome points
                if("Outcome" not in data[1]):
                    newpoint.style.iconstyle.color = simplekml.Color.red
                del matchingGroup[0:]

        # If data is not part of group or failed case number, generate point
        elif (data[5] == "" and data[0] not in failedGroup):
            newpoint = kml.newpoint(name = data[1], description = data[2], coords = [(data[4],data[3])])
            # Styling for outcome points
            if("Outcome" not in data[1]):
                newpoint.style.iconstyle.color = simplekml.Color.red

    # Save generated KML if validation passes
    if(prev[0] != ""):
        if (data[6] == "Pass"):
            caseFile = prev[0] + ".kml"
            genCaseNumList += [caseFile]
            kml.save("KML Files/" + prev[0] + ".kml")
            logger.info("Generated KML for Case #: " + str(prev[0]))
        elif (data[6] == "Fail"):
            genCaseNumList += ["Data validation failed: ..." + caseNum[-4:]]
            failedGroup += caseNum
            logger.warning("ERROR in Generating KML for Case #: " + str(prev[0]))
        del dataList[0:]

        
# Method to insert encoded KML files into the XML 
def KML2XML():
    # Initialize variables
    failedCases = []
    logger = logging.getLogger('KML to XML')
    
    # Define accepted file types
    ftypes = [
        ('XML Document', '*.xml'), 
    ]

    # Open XML file
    try:
        inputfile = askopenfile(mode = 'rt', filetypes = ftypes)
        logger.info('Successfully opened ' + inputfile.name)
        inputfile = open(inputfile.name, encoding = "ANSI")
    except Exception as e:
        logger.warning('ERROR opening ' + inputfile.name)
        logger.warning(e)

    # Parse through XML tree
    try:
        tree = ET.parse(inputfile)
    except ET.ParseError:
        xmlFormattingError()
        os._exit(0)

    root = tree.getroot()

    # Check for required child element
    if(root.findall("./ProjectSubmission") == []):
        xmlFormattingError()
        os._exit(0)

    # Parse through and find all cases
    for case in root.findall("./ProjectSubmission"):
        # Parse through and find required XML elements
        General_PS_General = case.find("General_PS_General")
        if(General_PS_General == None):
            xmlFormattingError()
            os._exit(0)

        submissionPTProjectIDCreate = General_PS_General.find("submissionPTProjectIDCreate")
        if(submissionPTProjectIDCreate == None):
            xmlFormattingError()
            os._exit(0)

        # Define case number from XML
        caseNum = submissionPTProjectIDCreate.text
    
        logger.info("Case Number: " + str(caseNum))

        # Open KML file, encode data, and insert encoded data along with required tags into XML
        try:
            kmldata = open("KML Files/" + caseNum + ".kml", "rb").read()
            logger.info("Opened KML File: " + str(caseNum))

            # Encode data
            kmlencoded = base64.b64encode(kmldata)
            kmlencoded = str(kmlencoded)
            kmlencoded = kmlencoded[2:-1]

            logger.info("Base64 Encoded KML File: " + str(kmlencoded))
            
            locationFileElement = case.find("LocationFile")

            # Insert KML elements if tags do not already exist
            if (locationFileElement.find("KLMdescription") == None):
                KLMdescriptionElement = ET.SubElement(locationFileElement, "KLMdescription")
                KLMdescriptionElement.text = "KML_" + caseNum
                KLMdescriptionElement.tail = "\n      "

                KLMSubjectElement = ET.SubElement(locationFileElement, "KLMSubject")
                KLMSubjectElement.text = "Location"
                KLMSubjectElement.tail = "\n      "

                KLMuploadtxtElement = ET.SubElement(locationFileElement, "KLMuploadtxt")
                KLMuploadtxtElement.text = str(kmlencoded)
                KLMuploadtxtElement.tail = "\n      "

                file_name_fldLocationElement = ET.SubElement(locationFileElement, "file_name_fldLocation")
                file_name_fldLocationElement.text = caseNum + ".kml"
                file_name_fldLocationElement.tail = "\n      "

                mimetypeKLMuploadElement = ET.SubElement(locationFileElement, "mimetypeKLMupload")
                mimetypeKLMuploadElement.text = "application/vnd.google-earth.kml+xml"
                mimetypeKLMuploadElement.tail = "\n      "

                isdocumentKLMuploadElement = ET.SubElement(locationFileElement, "isdocumentKLMupload") 
                isdocumentKLMuploadElement.text = "KML"
                isdocumentKLMuploadElement.tail = "\n      "
            # Insert KML elements if tags already exist
            else:
                KLMdescriptionElement = locationFileElement.find("KLMdescription")
                KLMdescriptionElement.text = "KML_" + caseNum
                KLMdescriptionElement.tail = "\n      "

                KLMSubjectElement = locationFileElement.find("KLMSubject")
                KLMSubjectElement.text = "Location"
                KLMSubjectElement.tail = "\n      "

                KLMuploadtxtElement = locationFileElement.find("KLMuploadtxt")
                KLMuploadtxtElement.text = str(kmlencoded)
                KLMuploadtxtElement.tail = "\n      "

                file_name_fldLocationElement = locationFileElement.find("file_name_fldLocation")
                file_name_fldLocationElement.text = caseNum + ".kml"
                file_name_fldLocationElement.tail = "\n      "

                mimetypeKLMuploadElement = locationFileElement.find("mimetypeKLMupload")
                mimetypeKLMuploadElement.text = "application/vnd.google-earth.kml+xml"
                mimetypeKLMuploadElement.tail = "\n      "

                isdocumentKLMuploadElement = locationFileElement.find("isdocumentKLMupload") 
                isdocumentKLMuploadElement.text = "KML"
                isdocumentKLMuploadElement.tail = "\n      "
            
            # Update XML with changed elements
            tree.write(inputfile.name, encoding = "ANSI")
            logger.info("Successfully wrote to " + str(inputfile.name)) 
            changeXmlEncoding(inputfile.name)
        except Exception as e: # Exception to catch if KML files are not found
            failedCases.append(str(caseNum))
            logger.warning("File Not Found: " + str(caseNum))  
            logger.warning(e)
    
    # Display different message depending on whether or not all/some of the KMLs were added into the XML
    if(len(failedCases) == 0):
        messagebox.showinfo("XML File Modified", "All of the KML files have successfully been added to the XML")
    else:
        failedCasesStr = '\n'.join(failedCases)
        messagebox.showinfo("XML File Modified", "Some KML files were successfully added to the XML. \n\nThe following KMLs were not found:\n" + failedCasesStr)

def changeXmlEncoding(filename):
    logger = logging.getLogger('Replace ANSI with UTF-8')
    try:
        with open(filename, 'r', encoding = "ANSI") as fin:
            data = fin.read().splitlines(True)
            logger.info("Successfully read " + filename)
        with open(filename, 'w', encoding = "ANSI") as fout:
            fout.writelines(data[1:])
            logger.info("Successfully deleted line " + filename)
    except Exception as e:
        logger.warning("ERROR reading/writing to " + filename)
        logger.warning(e)
    # with open(filename, 'w') as f:
    #     f.writelines([1:])

# Method to open file if double clicked in listbox
def selectOpenFile():
    fileName = listbox.get("active")
    os.startfile(os.getcwd() + "\\KML Files\\" + fileName)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Method to pull up "More Information" document
def displayMoreInfo():
    logger = logging.getLogger('More Info')
    logger.info('Instructions.pdf Opened')
    file = "Instructions.pdf"
    os.startfile(resource_path(file))

# Method that displays error message if CSV is not formatted properly
def csvFormattingError():
    logger = logging.getLogger('CSV Formatting Error')
    logger.warning('The CSV is missing a required column. Please see the formatting guidelines under the More Info button.')
    messagebox.showinfo("ERROR", "The CSV is missing a required column. Please see the formatting guidelines under the 'More Info' button.")

# Method that displays error message if XML is not formatted properly
def xmlFormattingError():
    logger = logging.getLogger('XML Formatting Error')
    logger.warning('The XML does not follow the correct format. Please see the formatting guidelines under the More Info button.')
    messagebox.showinfo("ERROR", "The XML does not follow the correct format. Please see the formatting guidelines under the 'More Info' button.")

# GUI Elements 
# Main window
tk = tkinter.Tk()
tk.geometry("400x650") 
tk.title("CSV2KML2XML")
tk.resizable(False, False)
tk.configure(background = "#332f31")
tk.iconbitmap(resource_path('S&C Icon.ico'))

# Fonts
titleFont = Font(family = "Verdana", weight = "bold", size = "25")
buttonFont = Font(family = "Verdana", weight = "bold", size = "10")

# Title
titleLabel = Label(tk, text = "CSV 2 KML 2 XML", height = 1, width = 18)
titleLabel.config(font = titleFont, background = "#332f31", fg = "#ffffff")
titleLabel.pack(anchor = "center", padx = (10, 10), pady = (20, 15))

# Listbox frame
frame = Frame(tk)
frame.pack(anchor = "center", padx = "15")

# Listbox description
listboxDesc = Label(tk, text = "Generated KML files will be displayed above", height = 1, width = 40)
listboxDesc.config(font = ("Verdana", 10), background = "#332f31", fg = "#ffffff")
listboxDesc.pack(anchor = "center", pady = (10, 0))

# Listbox
listbox = Listbox(frame, selectmode = "SINGLE", width = 30, height = 20, font = ("Helvetica", 12))
listbox.bind('<Double-1>', lambda x: selectOpenFile())
listbox.pack(side = "left", fill = "y")

# Scrollbar for list box
scrollbar = Scrollbar(frame, orient = "vertical")
scrollbar.config(command = listbox.yview)
scrollbar.pack(side = "right", fill = "y")

listbox.config(yscrollcommand = scrollbar.set)

# Button for "More Information"
moreInfoButton = Button(tk, text = "More Info", font = buttonFont, command = displayMoreInfo, height = 2, width = 32)
moreInfoButton.pack(side = "bottom", pady = (10, 25))

# Button for CSV 2 KML
csvToKmlButton = Button(tk, text = "CSV to KML", font = buttonFont, command = CSV2KML, height = 2, width = 15)
csvToKmlButton.pack(side = "left", pady = (25,10), padx = (50, 0))

# Button for KML 2 XML 
kmlToXmlButton = Button(tk, text = "KML to XML", font = buttonFont, command = KML2XML, height = 2, width = 15)
kmlToXmlButton.pack(side = "right", pady = (25,10), padx = (0, 50))


# Create logger with 'spam_application'
logger = logging.getLogger('')
logger.setLevel(logging.DEBUG)

# Create file handler which logs even debug messages
fh = logging.FileHandler('Logger.log', mode = 'w')
fh.setLevel(logging.DEBUG)

# Create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.ERROR)

# Create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)

# Add the handlers to the logger
logger.addHandler(fh)
logger.addHandler(ch)

# Array to display all previously existing files, if any
if(os.path.exists(os.getcwd() + "\\KML Files") and genFiles != True):
    genCaseNumList = os.listdir(os.getcwd() + "\\KML Files")
    for case in genCaseNumList:
        # case = case[:-4]
        listbox.insert(END, case)
else:
    genCaseNumList = []

# Keeps GUI running
tk.mainloop()