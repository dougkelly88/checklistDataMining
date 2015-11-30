import os, time, string, datetime, tkFileDialog, ast, webbrowser, sys, traceback
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from Tkinter import *
from tkMessageBox import *
import tkSimpleDialog
import requests, gspread
from oauth2client.client import SignedJwtAssertionCredentials

# Constants
MODE_APPEND = 0
MODE_NEW = 1
MODE_UPDATEWEBFROMCL = 2
CHECKLIST_FILE_FORMAT = '.xlsm'

# Error messages
NOT_YET_SUPPORTED = "Sorry, this feature is not yet supported!"
INVALID_PATH = "No (valid) path selected!"
FILES_LENGTH_ZERO = "No (valid) files found in this location!"

class outputSelectionDialog(object):
    """ Dialog box allowing selection of output params"""
    def __init__(self, internalData):
        self.data = internalData
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
                #print(self.data.tasks[i].taskLabel)
                #print(self.data.tasks[i].output)
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
            print(self.data.tasks[i].taskLabel)
            print(self.data.tasks[i].output)
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

    def returnTaskByLabel(self, label):
        for task in self.tasks:
            if (task.taskLabel == label):
                return task

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

    def returnTaskByLabel(self, label):
        if (taskLabel == label):
            return self

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
        if (description == "oil"):
            return Printing.OilTask()
        if ("box" in description):
            return Printing.BoxTask()
        if ("oven" in description):
            return Printing.OvenTask()
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
            self.dcOffset = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for rowind in range(2):
                for col in range(26):
                    c = ws.cell(row = r+rowind, column = col).value
                    if (isinstance(c, basestring)):
                        if ("dwell" in c.lower()):
                            self.dwell = ws.cell(row = r+rowind, column = col+1).value
                        if ("step" in c.lower()):
                            self.step = ws.cell(row = r+rowind, column = col+1).value
                        if ("voltage" in c.lower()):
                            self.voltage = ws.cell(row = r+rowind, column = col+1).value
                        if ("freq" in c.lower()):
                            self.freq = ws.cell(row = r+rowind, column = col+1).value
                        if ("pressure" in c.lower()):
                            self.pressure = ws.cell(row = r+rowind, column = col+1).value
                        if ("offset" in c.lower()):
                            self.dcOffset = ws.cell(row = r+rowind, column = col+1).value

    class OvenTask(TaskBase):
        def __init__(self):
            self.type = ""
            self.temperature = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for col in range(26):
                c = ws.cell(row = r, column = col).value
                if (isinstance(c, basestring)):
                    print(c)
                    if ("oven type" in c.lower()):
                        self.type = ws.cell(row = r, column = col+1).value
                    if ("temp" in c.lower()):
                        self.temperature = ws.cell(row = r, column = col+1).value

    class OilTask(TaskBase):
        def __intit__(self):
            self.aliquote = ""
            self.id = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for col in range(26):
                c = ws.cell(row = r, column = col).value
                if (isinstance(c, basestring)):
                    if ("ID" in c):
                        self.id = ws.cell(row = r, column = col+1).value
                    if ("Aliquote" in c):
                        self.aliquote = ws.cell(row = r, column = col+1).value

    class BoxTask(TaskBase):
        def __init__(self):
            self.bottom = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for col in range(26): 
                c = ws.cell(row = r, column = col).value
                if (isinstance(c, basestring)):
                    if ("Bottom" in c):
                        self.bottom = ws.cell(row = r, column = col+1).value

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
                        self.batch = '%s-%s' % (ws.cell(row = r, column = col+1).value, ws.cell(row = r, column = col+3).value)
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
            self.oilVolume = ""

        def populate(self, ws, r):
            TaskBase.populate(self, ws, r)

            for col in range(26):
                c = ws.cell(row = r, column = col).value
                if (isinstance(c, basestring)):
                    if ("Volume" in c):
                        self.oilVolume = ws.cell(row = r, column = col+1).value
                    if ("Humidity" in c):
                        self.humidity = ws.cell(row = r, column = col+1).value

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
        self.printRig = "P%s" % self.parseSampleID(self.sampleName)[0]

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
                
            output[1] = "%s%s%s" % (y, m, d)
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

def authenticate_google_docs():
    #f = file(os.path.join('C:/Users/d.kelly/Desktop/Python/bulkSummariser-f4e730f107d4.p12'), 'rb')
    f = file(os.path.join('//base4share/share/Doug/checklistDataMining/bulkSummariser-f4e730f107d4.p12'), 'rb')
    SIGNED_KEY = f.read()
    f.close()
    scope = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds', 'https://www.googleapis.com/auth/gmail.send']
    credentials = SignedJwtAssertionCredentials('d.kelly@base4.co.uk', SIGNED_KEY, scope)

    data = {
        'refresh_token' : '1/J9DflNyNnx2J9WnWtHsKnH9YDTTBTuoz3tJtnSXRtLc',
        'client_id' : '326340397307-lgicfbcjmu9863pjjkfrn0c5dviqsb42.apps.googleusercontent.com',
        'client_secret' : 'xeNtki_QWyoLaGEjYqtzF5si',
        'grant_type' : 'refresh_token',
    }

    r = requests.post('https://accounts.google.com/o/oauth2/token', data = data)
    credentials.access_token = ast.literal_eval(r.text)['access_token']

    gc = gspread.authorize(credentials)
    return gc

def chooseFolder(initialdir, prompt):
    try:
        import tkFileDialog
    except:
        from tkinter import filedialog as tkFileDialog

    root = Tk()
    try:
        options = {}
        options['initialdir']=initialdir
        file = tkFileDialog.askdirectory(**options)
        print(file)
        
        if file == "":
            errorHandler(INVALID_PATH)

    except IOError:
        errorHandler(INVALID_PATH)

    root.destroy()
    return file

def chooseOutputFile(initialdir, initialfile):
    try:
        root = Tk()
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

    root.destroy()
    return file

def errorHandler(message):
    print(message)
    showerror("Error!", message)
    # TODO: log file? send email?
    exit()

def numberToLetters(q):
    q = q - 1
    result = ''
    while q >= 0:
        remain = q % 26
        result = chr(remain+65) + result;
        q = q//26 - 1
    return result
                
def formatToGS(p, gws):

    # find row, or make new row
    sampleNamesFromGS = gws.col_values(1)
    if p.sampleName in sampleNamesFromGS:
        row = sampleNamesFromGS.index(p.sampleName) + 1
    else:
        row = len(sampleNamesFromGS) + 1
        gws.update_cell(row, 1, p.sampleName)

    headings = gws.row_values(1)
    splitPath = os.path.split(p.path)
    filename = os.path.splitext(splitPath[1])
    fmt = ('%Y\%B\%d')
    date = datetime.datetime.strptime(p.printDate, '%y%m%d')
    datePath = date.strftime(fmt)
    pdfPath = os.path.join("\\\\base4share\share\SOPs\Completed Checklists", datePath, (filename[0]).replace("Printing 1", "Printing") +  '.pdf')
    gws.update_cell(row, headings.index("Completed Checklist Link") + 1, pdfPath)

    first_col = numberToLetters(headings.index("Protocol version") + 1)
    last_col = numberToLetters(headings.index("Pressure [kPa]") + 1) 
    cellrange = '%s%d:%s%d' % (first_col, row, last_col, row)
    cells = gws.range(cellrange)
    headings = headings[headings.index("Protocol version"):(headings.index("Pressure [kPa]") + 1)]

    for heading in headings:
        if (heading == "Protocol version"):
            cells[headings.index(heading) ].value = 'v %1.2f' % p.sOPVersion
        if (heading == "Rig"):
            cells[headings.index(heading) ].value = p.printRig
        if (heading == "Printer"):
            cells[headings.index(heading) ].value = p.experimenter
        if (heading == "Slide CA"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Slide").CA
        if (heading == "Slide batch"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Slide").batch
        if (heading == "Tip size"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Tip").size
        if (heading == "Tip Batch"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Tip").batch
        if (heading == "Tip ID"):
            cells[headings.index(heading) ].value = "%s%s" % (p.printDate, p.returnTaskByLabel("Tip").ID)
        if (heading == "Room Temperature"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Note room temperature").temperature
        if (heading == "Room Humidity"):
            if (p.returnTaskByLabel("Note room humidity").humidity != None):
                cells[headings.index(heading) ].value = p.returnTaskByLabel("Note room humidity").humidity * 100
        if (heading == "Humidity low"):
            if (p.returnTaskByLabel("Position Oil and turn humidifier to low").humidity != None):
                cells[headings.index(heading) ].value = p.returnTaskByLabel("Position Oil and turn humidifier to low").humidity * 100
        if (heading == "Humidity high"):
            try:
                if (p.returnTaskByLabel("Move tip into oil and turn humidifier to high").humidity != None):
                    cells[headings.index(heading) ].value = p.returnTaskByLabel("Move tip into oil and turn humidifier to high").humidity * 100
            except:
                if (p.returnTaskByLabel("Turn humidifier to high").humidity != None):
                    cells[headings.index(heading) ].value = p.returnTaskByLabel("Turn humidifier to high").humidity * 100
        if (heading == "Print time"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Print").timeTaken
        if (heading == "Print voltage (AC) / V x 100"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Print").voltage
        if (heading == "Voltage frequency (sine) /Hz"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Print").freq
        if (heading == "Dwell time"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Print").dwell
        if (heading == "Step size"):
            cells[headings.index(heading) ].value = p.returnTaskByLabel("Print").step
        if (heading == "Pressure [kPa]"):
            cells[headings.index(heading)].value = p.returnTaskByLabel("Print").pressure

    #print(p.printDate)
    #for cell in cells:
    #    print(cell)

    gws.update_cells(cells)
    webbrowser.open('https://docs.google.com/spreadsheets/d/1W_S4NUgCKchcfokpm7cbRywsh-AIfmzk5ksjX05vKII/edit#gid=1149687153', 2, True)
            
if __name__ == "__main__":

    try:
        #mode = MODE_NEW
        print(len(sys.argv))
        if len(sys.argv) == 1:
            mode = MODE_NEW
            print('Mode = generate new Excel spreadsheet summary')
        elif len(sys.argv) == 2 :
            mode = MODE_UPDATEWEBFROMCL
            checklistPath = sys.argv[1]
            print('Mode = update google sheet')
        elif len(sys.argv) == 3:
            mode = MODE_APPEND
            checklistPath = sys.argv[1]
            outputFile = sys.argv[2]
            print('Mode = append to existing Excel sheet')
        else:
            errorHandler('Too many arguments passed!')

        mode = MODE_UPDATEWEBFROMCL
        mode = MODE_NEW

        xml_export = False;

        if (mode == MODE_UPDATEWEBFROMCL):
        
            gc = authenticate_google_docs()
            gsh = gc.open("Sample register")
            gws = gsh.worksheet("Sample register")
            #gsh = gc.open("Dummy sample register")
            #gws = gsh.worksheet("Sheet1")

            # Replace this with argument input from command line - get from excel using Application.ActiveWorkbook.Path or Application.ActiveWorkbook.FullName 
            #checklistPath = '//base4share/share/SOPs/Completed Checklists/Data/Printing/Printing 1 2015-10-30 0909.xlsm'
            #checklistPath = '//base4share/share/SOPs/Completed Checklists/Data/Printing/Printing 1 2015-11-04 1732.xlsm'
            #checklistPath = '//base4share/share/Doug/P1-151102-A - Copy.xlsm'

            wb = load_workbook(checklistPath)
            ws = wb.active
            p = Printing(checklistPath)
            p.populatePrintingClass(ws)
            p.populatePrintingTasks(ws)
            formatToGS(p, gws)

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
            print("nonsense")


        if (mode == MODE_NEW):
            print("Mode = NEW")
            # prompt user for output file name/location
            # TODO: add googledocs option?
            initialDir = "//base4share/share/SOPs/Completed Checklists/Data/Printing"
        
            # prompt user for place to look for checklists
            prompt = "Please choose a location in which to look for checklist spreadsheets..."
            inputPath = chooseFolder(initialDir, prompt)
            #inputPath = "//base4share/share/SOPs/Completed Checklists/Data/Printing"

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
            
            # TODO: better to establish correct number of tasks by adding tasks to a dictionary so that
            # there is a list of unique tasks. 
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
            initialFile = "%s printing data summary.xlsx" % time.strftime('%Y-%m-%d')
            outputFile = chooseOutputFile(initialDir, initialFile)
            #outputFile = "C:/Users/d.kelly/Desktop/Python/DKBase4PythonScripts/checklistDataMining/sampleData/test/output.xlsx"
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
                            print(mt.taskLabel)
                            for internalData in internalDataList:
                                r = r + 1
                                for task in internalData.tasks:
                                    if (task.taskLabel == mt.taskLabel) and (task.taskCategory == mt.taskCategory):
                                        vv = vars(task)
                                        if (vv[key] != None) and (vv[key] != ""):
                                            col_filled = True
                                        ws.cell(row = r, column = col).value = vv[key]
                            if col_filled:
                                col = col + 1

            # write to output file (googledoc?)
            wb.save(outputFile)
    except:
        errorHandler(traceback.format_exc())


