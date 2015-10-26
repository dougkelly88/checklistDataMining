import os, time, string, datetime, tkFileDialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from Tkinter import *
from tkMessageBox import *
import tkSimpleDialog
import inspect

MODE_APPEND = 0
MODE_NEW = 1
CHECKLIST_FILE_FORMAT = '.xlsm'

# Error messages
NOT_YET_SUPPORTED = "Sorry, this feature is not yet supported!"
INVALID_PATH = "No (valid) path selected!"
FILES_LENGTH_ZERO = "No (valid) files found in this location!"

class outputSelectionDialog(object):
    """ Dialog box allowing selection of output params"""
    def __init__(self, internalData):
        self.data = internalData
        print(self.data)
        root = self.root = Tk()
        self.all = IntVar()
        self.sizex = 400
        self.sizey = 600
        self.posx  = 100
        self.posy  = 100
        root.wm_geometry("%dx%d+%d+%d" % (self.sizex, self.sizey, self.posx, self.posy))
        root.title('Select data to export...')
        frm_1 = Frame(root)
        frm_1.pack(ipadx=2, ipady=2)
        self.canvas = Canvas(frm_1)
        frm_2 = Frame(self.canvas)
        scrlbar = Scrollbar(frm_1,orient="vertical",command=self.canvas.yview)
        scrlbar.pack(side="right",fill="y")
        self.canvas.configure(yscrollcommand=scrlbar.set)
        self.canvas.pack(side="left")
        self.canvas.create_window((0,0),window=frm_2,anchor='nw')
        frm_2.bind("<Configure>",self.myfunction)

        i = 0
        self.vars = []
        chkAll = Checkbutton(frm_2, text = 'All', variable = self.all)
        chkAll.pack(padx = 0, pady = 5, anchor = W)
        chkAll['command'] = self.c2_action
        for t in self.data.tasks:
            txt = "%s: %s" % (t.taskCategory, t.taskLabel)
            var = IntVar()
            chk = Checkbutton(frm_2, text = txt, variable = var)
            i = i+1
            chk.pack(padx = 0, pady = 5, anchor = W)
            chk['command'] = self.c1_action
            self.vars.append(var)
        btn_1 = Button(frm_2, width=8, text='OK')
        btn_1['command'] = self.b1_action
        btn_1.pack(anchor = CENTER)
        
    def myfunction(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"),width=self.sizex,height=self.sizey)
    
    def b1_action(self, event=None):
        #print('quitting...')
        self.root.quit()

    def c2_action(self, event = None):
        i = 0
        for v in self.vars:
            if (self.all.get() == 1):
                self.data.tasks[i].output = True
                v.set(1)
                print(self.data.tasks[i].taskLabel)
                print(self.data.tasks[i].output)
            else:
                self.data.tasks[i].output = False
                v.set(0)
                print(self.data.tasks[i].taskLabel)
                print(self.data.tasks[i].output)
            i = i+1

    def c1_action(self, event = None):
        i = 0
        for v in self.vars:
            if (v.get() == 1):
                self.data.tasks[i].output = True
            else:
                self.data.tasks[i].output = False
            #print(self.data.tasks[i].taskLabel)
            #print(self.data.tasks[i].output)
            i = i+1

class ChecklistBase(object):
    """ Base class for checklist-level common properties"""
    def __init__(self):
        self.sOP = ""
        self.sOPVersion = ""
        self.date = ""
        self.qCPerson = ""
        self.tasks = []
        self.path = ""
    
class TaskBase(object):
    """ Base class for task-level common properties"""
    def __init__(self):
        self.taskNumber = ""
        self.taskLabel = ""
        self.taskCategory = ""
        self.doneBy = ""
        self.startedAt = ""
        self.timeTaken = ""    
        self.output = False       
        self.notes = []

    def populate(self, ws, r):
        self.doneBy = ws.cell(row = r, column = 5).value
        self.startedAt = ws.cell(row = r, column = 7).value
        self.timeTaken = ws.cell(row = r, column = 9).value
        self.output = False

def identifyCorrectClass(description):
        description = description.lower()
        if ("humidity" in description) or ("humidifier" in description):
            return Printing.HumidityTask()
        if ("temperature" in description):
            return Printing.TemperatureTask()
        if (description == "slide"):
            return Printing.SlideTask()
        if (description == "tip") or ("tip size" in description):
            return Printing.TipTask()
        if (description == "mix") or ("push-through" in description):
            return Printing.BulksTask()
        if (description == "print"):
            return Printing.PrintTask()
        else:
            return TaskBase()

class Printing(ChecklistBase):
    """Class containing all variables and methods relating to the printing checklist"""
    def __init__(self, path):
        # Initialise general checklist class
        super(Printing, self).__init__()
        self.path = path
        # Initialise specialised print checklist variables
        self.sampleName = ""
        self.experimenter = ""
        self.output = False
        
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

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for col in range(26):
                c = ws.cell(row = r+1, column = col).value
                if (isinstance(c, basestring)):
                    if ("dwell" in c.lower()):
                        self.dwell = ws.cell(row = r+1, column = col+1).value
                    if ("step" in c.lower()):
                        self.step = ws.cell(row = r+1, column = col+1).value
                    if ("voltage" in c.lower()):
                        self.voltage = ws.cell(row = r+1, column = col+1).value
                    if ("freq" in c.lower()):
                        self.freq = ws.cell(row = r+1, column = col+1).value
                    if ("pressure" in c.lower()):
                        self.pressure = ws.cell(row = r+1, column = col+1).value

    class BulksTask(TaskBase):
        def __init__(self):
            self.type = ""
            self.claire700 = ""
            self.claire655 = ""
            self.claire594 = ""
            self.claire532 = ""
            self.claire488 = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)
            for col in range(26):
                c = ws.cell(row = r, column = col).value
                if (isinstance(c, basestring)):
                    if ("633nm" in c) or ("655nm" in c):
                        self.claire655 = ws.cell(row = r, column = col+1).value
                    if ("700nm" in c):
                        self.claire700 = ws.cell(row = r, column = col+1).value
                    if ("594nm" in c):
                        self.claire594 = ws.cell(row = r, column = col+1).value
                    if ("532nm" in c):
                        self.claire532 = ws.cell(row = r, column = col+1).value
                    if ("488nm" in c):
                        self.claire488 = ws.cell(row = r, column = col+1).value
                    if ("type" in c.lower()):
                        self.type = ws.cell(row = r, column = col+1).value

    class TipTask(TaskBase):
        def __init__(self):
            self.size = ""
            self.batch = ""
            self.ID = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for col in range(26):
                c = ws.cell(row = r, column = col).value
                if (isinstance(c, basestring)):
                    if ("Size" in c):
                        self.size = ws.cell(row = r, column = col+1).value
                    if ("batch" in c.lower()):
                        self.batch = ws.cell(row = r, column = col+1).value
                    if ("ID" in c):
                        self.ID = ws.cell(row = r, column = col+1).value

    class SlideTask(TaskBase):
        def __init__(self):
            self.CA = ""
            self.batch = ""
            self.ID = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for col in range(26):
                c = ws.cell(row = r, column = col).value
                if (isinstance(c, basestring)):
                    if ("CA" in c):
                        self.CA = ws.cell(row = r, column = col+1).value
                    if ("batch" in c.lower()):
                        self.batch = ws.cell(row = r, column = col+1).value
                    if ("#" in c):
                        self.ID = ws.cell(row = r, column = col+1).value

    class HumidityTask(TaskBase):
        def __init__(self):
            self.humidity = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)
            self.humidity = ws.cell(row = r, column = 12).value
            self.notes = ws.cell(row = r, column = 14).value
            #print("Humidity task")
            #print(vars(self))

    class TemperatureTask(TaskBase):
        def __init__(self):
            self.temperature = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)
            #print(self.output)
            self.temperature = ws.cell(row = r, column = 12).value
            self.notes = ws.cell(row = r, column = 14).value
            #print("Temperature task")
            #print(vars(self))

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
        output = [""]*2
        if isinstance(sampleID, basestring):
            dashes = sampleID.count('-')
            if (dashes == 2):
                splitString = string.split(sampleID, "-")
            #if len(splitString)==3:
                output[0] = splitString[0][1]
                d=splitString[1][4:]
                m=splitString[1][2:4]
                y=splitString[1][0:2]
                
            elif (dashes == 1):
                output[0] = sampleID[1]
                y = sampleID[3:5]
                m = sampleID[5:7]
                d = sampleID[7:9]

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
                    t = identifyCorrectClass(taskLabel)
                    t.taskLabel = taskLabel
                    t.taskCategory = category
                    t.taskNumber  = taskNumber
                    t.populate(ws, r)
                    self.tasks.append(t)




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
        #initialDir = "Z:\\SOPs\\Completed Checklists\\Data"
        #initialFile = "%s printing data summary.xlsx" % time.strftime('%Y-%m-%d')
        #outputFile = chooseOutputFile(initialDir, initialFile)
        
        # prompt user for place to look for checklists
        #prompt = "Please choose a location in which to look for checklist spreadsheets..."
        #inputPath = chooseFolder(initialDir, prompt)
        inputPath = "Z:/SOPs/Completed Checklists/Data/Printing"
        #inputPath = "C:/Users/d.kelly/Desktop/Python/DKBase4PythonScripts/checklistDataMining/sampleData/test/Data/Printing"

        # generate list of checklist paths
        checklistList = []
        for root, dirs, files in os.walk(inputPath):
            for basename in files:
                if CHECKLIST_FILE_FORMAT in basename:
                    checklistList.append(os.path.join(root, basename))
        #print(checklistList)
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
            for checklist in checklistList:
                print(checklist)
                wb = load_workbook(checklist)
                ws = wb.active
                p = Printing(checklist)
                p.populatePrintingClass(ws)
                p.populatePrintingTasks(ws)
                if (p.sampleName != None):
                    if ("ppl" not in p.sampleName.lower()):
                        internalDataList.append(p)
        elif (testString[1] == 'Slide coating - fluorinated silane'):
            print("Use SlideCoating class")
        else:
            errorHandler(NOT_YET_SUPPORTED)

        no_tasks = []
        for internalData in internalDataList:
            no_tasks.append(len(internalData.tasks))


        # prompt user for fields to include in summary      
        m = outputSelectionDialog(internalDataList[no_tasks.index(max(no_tasks))])
        m.root.mainloop()
        m.root.destroy()
        print('run past dialog, data to save:')
        for t in m.data.tasks:
            if (t.output):
                print(t.taskLabel)

        # loop through classes and fields and parse to output format
        # set up file to output to
        outputFile = "C:/Users/d.kelly/Desktop/Python/DKBase4PythonScripts/checklistDataMining/sampleData/test/output.xlsx"
        wb = Workbook()
        ws = wb.active

        exclude_vars = ['taskLabel', 'taskCategory', 'taskNumber', 'notes', 'output']
        
        # set up invariant section
        r = 1
        ws.cell(row = r, column = 1).value = 'Source checklist'
        ws.cell(row = r, column = 2).value = 'Print date'
        ws.cell(row = r, column = 3).value  = 'Sample name'
        ws.cell(row = r, column = 4).value  = 'Print rig'
        ws.cell(row = r, column = 5).value  = 'Experimenter'
        ws.cell(row = r, column = 6).value  = 'QC'
        ws.cell(row = r, column = 7).value  = 'SOP version'

        for internalData in internalDataList:
            r = r + 1
            ws.cell(row = r, column = 1).value = internalData.path
            ws.cell(row = r, column = 2).value = internalData.printDate
            ws.cell(row = r, column = 3).value  = internalData.sampleName
            ws.cell(row = r, column = 4).value  = internalData.printRig
            ws.cell(row = r, column = 5).value  = internalData.experimenter
            ws.cell(row = r, column = 6).value  = internalData.qCPerson
            ws.cell(row = r, column = 7).value  = internalData.sOPVersion

        col = 8
        for mt in m.data.tasks:
            if mt.output:
                v = vars(mt)
                for key in v:
                    if key not in exclude_vars:
                        r = 1
                        col_filled = False
                        ws.cell(row = r, column = col).value = "%s: %s" % (mt.taskLabel, key)
                        for internalData in internalDataList:
                            print(internalData.path)
                            r = r + 1
                            for task in internalData.tasks:
                                if (task.taskLabel == mt.taskLabel) and (task.taskCategory == mt.taskCategory):
                                    vv = vars(task)
                                    if (vv[key] != None) and (vv[key] != ""):
                                        col_filled = True
                                    ws.cell(row = r, column = col).value = vv[key]
                        if col_filled:
                            col = col + 1

                        #for internalData in internalDataList:
                            

        
        #for internalData in internalDataList:
        #    print("SampleName: %s" % internalData.sampleName)
        #    print("Experimenter: %s" % internalData.experimenter)
        #    print("QC person: %s" % internalData.qCPerson)
        #    print("Print date: %s" % internalData.printDate) 
        #    print("")
        #    for mt in m.data.tasks:
        #        if (mt.output):
        #            for task in internalData.tasks:
        #                if (task.taskLabel == mt.taskLabel):
        #                    print("")
        #                    print(mt.taskLabel)
        #                    v = vars(task)
        #                    for key in v:
        #                        exclude_vars = ['taskLabel', 'taskCategory', 'taskNumber', 'notes', 'output']
        #                        if (key not in exclude_vars):
        #                            print("%s: " % key)
        #                            print(v[key])

                #i = 1 + i    

        # write to output file (googledoc?)
        wb.save(outputFile)



