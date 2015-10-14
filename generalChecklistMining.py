import os, time, string, datetime
from openpyxl import Workbook, load_workbook

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


#class PrintheadPrinting(ChecklistBase):


if __name__ == "__main__":

    wb = load_workbook('Z:\\SOPs\\Completed Checklists\\Data\\Printing\\Printing 1 2015-10-09 1130.xlsm')
    ws = wb.active

    printing = Printing()
    printing.populatePrintingClass(ws)
    v = vars(printing)
    print(', '.join("%s: %s" % item for item in v.items()))
    print(printing.sampleName)
