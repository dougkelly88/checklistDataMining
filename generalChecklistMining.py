import os, time, string, datetime, tkFileDialog
from openpyxl import Workbook, load_workbook
from Tkinter import *
from tkMessageBox import *

MODE_APPEND = 0
MODE_NEW = 1
CHECKLIST_FILE_FORMAT = '.xlsm'

# Error messages
NOT_YET_SUPPORTED = "Sorry, this feature is not yet supported!"
INVALID_PATH = "No (valid) path selected!"
FILES_LENGTH_ZERO = "No (valid) files found in this location!"

class ChecklistBase(object):
    """ Base class for checklist-level common properties"""
    def __init__(self):
        self.sOP = ""
        self.sOPVersion = ""
        self.date = ""
        self.qCPerson = ""
        self.tasks = []
    
class TaskBase(object):
    """ Base class for task-level common properties"""
    def __init__(self):
        self.taskNumber = ""
        self.taskLabel = ""
        self.taskCategory = ""
        self.doneBy = ""
        self.startedAt = ""
        self.timeTaken = ""           
        self.notes = ""

class Printing(ChecklistBase):
    """Class containing all variables and methods relating to the printing checklist"""
    def __init__(self):
        # Initialise general checklist class
        super(Printing, self).__init__()

        # Initialise specialised print checklist variables
        self.sampleName = ""
        self.experimenter = ""
        
        # Initialise specialised print checklist variables that are not explicitly in checklist but may be derived from checklist entries
        self.printDate = ""
        self.printRig = ""

    class PrintTask(TaskBase):
        def __init__(self):
            self.dwell = ""
            self.step = ""
            self.voltage = ""
            self.freq = ""
            self.pressure = ""

    class BulksTask(TaskBase):
        def __init__(self):
            self.type = ""
            self.claire700 = ""
            self.claire655 = ""
            self.claire594 = ""
            self.claire532 = ""
            self.claire488 = ""

    class TipTask(TaskBase):
        def __init__(self):
            self.size = ""
            self.batch = ""
            self.ID = ""

    class SlideTask(TaskBase):
        def __init__(self):
            self.CA = ""
            self.batch = ""
            self.ID = ""

    class HumidityTask(TaskBase):
        def __init__(self):
            self.humidity = ""

    class TemperatureTask(TaskBase):
        def __init__(self):
            self.temperature = ""

    def populatePrintingClass(self, ws):
        """Populate checklist-level data for the Printing case"""
        self.sOP = ws['E2'].value
        self.sOPVersion = ws['E4'].value
        self.date = ws['E6'].value
        self.sampleName = ws['E8'].value
        self.experimenter = ws['E10'].value
        self.qCPerson = ws['E12'].value
        self.printDate = self.parseSampleID(self.sampleName)[1]
        self.printRig = "Rig %s" % self.parseSampleID(self.sampleName)[0]

    def parseSampleID(self, sampleID):
        """ Outputs rig #, date """
        splitString = string.split(sampleID, "-")
        output = [""]*2
        if len(splitString)==3:
            output[0] = splitString[0][1]
            d=splitString[1][4:]
            m=splitString[1][2:4]
            y=splitString[1][0:2]
            output[1] = "%s/%s/20%s" % (d, m, y)
        else:
            #TODO: log warning about unexpected sample ID format, taking best guess
            output[0] = sampleID[1]
            y = sampleID[2:4]
            m = sampleID[4:6]
            d = sampleID[6:8]
            output[1] = "%s/%s/20%s" % (d, m, y)
        return output

    def populatePrintingTasks(self, ws):
        print('Populating...')
        category = ""
        for row in ws.iter_rows('B1:B100'):
            for cell in row:
                if (isinstance(cell.value, long)):
                    r = row[0].row
                    taskNumber = cell.value
                    if (ws.cell(row = r, column = 3).value != None):
                        category = ws.cell(row = r, column = 3).value
                    taskLabel = ws.cell(row = r, column = 4).value
                    t = TaskBase()
                    t.taskLabel = taskLabel
                    t.taskCategory = category
                    t.taskNumber  = taskNumber
                    print(t.taskNumber)
                    print(t.taskLabel)
                    print(t.taskCategory)

                    #print(cell.value)
                    #print(row[0].row)
                    #print(type(row[0]))

#class PrintheadPrinting(ChecklistBase):

def chooseFolder(initialdir, prompt):
    import tkFileDialog

    try:
        options = {}
        options['initialdir']=initialdir
        file = tkFileDialog.askdirectory(**options)
        print(file)
        
        if file == "":
            errorHandler(INVALID_PATH)

    except IOError:
        errorHandler(INVALID_PATH)

    return file

def chooseOutputFile(initialdir, initialfile):
    try:
        options = {}
        options['initialdir']=initialdir
        options['initialfile'] = initialfile
        options['defaultextension'] = '.xls'
        options['filetypes'] = [('Excel spreadsheet', '.xlsx'), ('Comma separated variable file', '.csv')]
        options['title'] = 'Choose an output file to save the summary data to...'
        file = tkFileDialog.asksaveasfilename(**options)
        print(file)
        
        if file == "":
            errorHandler(INVALID_PATH)

    except IOError:
        errorHandler(INVALID_PATH)

    return file

def errorHandler(message):
    print(message)
    showerror("Error!", message)
    # TODO: log file? send email?
    exit()

if __name__ == "__main__":

    mode = MODE_NEW
    xml_export = False;

    if (mode == MODE_APPEND):
        print("Mode = APPEND")
        # prompt user to choose type of checklist being summarised, or do so automatically?
        # prompt user for output file, or take from arguments to allow scheduled running?
        # identify output file
        # identify date of last entry
        # identify checklists since last entry
        # loop through these checklists and add a checklist class for each case to a List
        # from output file column titles identify the fields to export - if fields don't exist, add warning text to these entries?
        # from output file, determine list of ID fields that can be used to confirm that new entries don't already exist
        # perform check for exisiting entries and delete those classes from the List
        # loop through classes and fields and parse to output format
        # append to output file - first checking that write is possible and warning user (email?) if not
        # TODO: add option for googlesheets export
        # OPTION: re-export checklist data in XML format?

    if (mode == MODE_NEW):
        print("Mode = NEW")
        # prompt user for output file name/location
        # TODO: add googledocs option?
        initialDir = "Z:\\SOPs\\Completed Checklists\\Data"
        initialFile = "%s printing data summary.xlsx" % time.strftime('%Y-%m-%d')
        outputFile = chooseOutputFile(initialDir, initialFile)
        
        # prompt user for place to look for checklists
        prompt = "Please choose a location in which to look for checklist spreadsheets..."
        inputPath = chooseFolder(initialDir, prompt)

        # generate list of checklist paths
        checklistList = []
        for root, dirs, files in os.walk(inputPath):
            for basename in files:
                if CHECKLIST_FILE_FORMAT in basename:
                    checklistList.append(os.path.join(root, basename))
        print(checklistList)
        if (len(checklistList) == 0):
            errorHandler(FILES_LENGTH_ZERO)
        
        # prompt user for checklist type - or get from place to look for checklists?
        # for now, assume that data is in Z:\SOPs\Completed Checklists\Data and figure out which class to use from which folder is selected immediately below
        internalDataList = []
        dummy = os.path.split(checklistList[0])
        testString = dummy[0].split('/Data/')
        print(testString)
        if (testString[1] == 'Printing'):
            print("Use Printing class")
            internalDataList.append(Printing())
        elif (testString[1] == 'Slide coating - fluorinated silane'):
            print("Use SlideCoating class")
        else:
            errorHandler(NOT_YET_SUPPORTED)

        # populate class tasks based on checklist spreadsheet
        wb = load_workbook(checklistList[0])
        ws = wb.active
        internalDataList[0].populatePrintingTasks(ws)

        # prompt user for fields to include in summary      
        inclusionList = vars(internalDataList[0])
        print(inclusionList)
        print('Run to this point')

        # loop through all checklists in this location and add a checklist class for each case to a List
        #for cl in checklistList:
        # loop through classes and fields and parse to output format
        # write to output file (googledoc?)

    #wb = load_workbook('Z:\\SOPs\\Completed Checklists\\Data\\Printing\\Printing 1 2015-10-09 1130.xlsm')
    #ws = wb.active

    #printing = Printing()
    #printing.populatePrintingClass(ws)
    #v = vars(printing)
    #print(', '.join("%s: %s" % item for item in v.items()))
    #print(printing.sampleName)


