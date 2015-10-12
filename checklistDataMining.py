import os, time
from openpyxl import Workbook, load_workbook

class Base(object):
    """ Base class for common properties"""
    def __init__(self):
        self.doneBy = ""
        self.startedAt = ""
        self.timeTaken = ""

    def basePopulate(self, ws, r):
        self.doneBy = ws.cell(row = r-1, column = 4).value
        self.startedAt = ws.cell(row = r-1, column = 6).value
        self.startedAt = ws.cell(row = r-1, column = 8).value

class PrintSampleData:

    def __init__(self, sampleName):

        self.sampleName = sampleName
        self.generalTasks = GeneralTasks()
        self.electrodePrep = ElectrodePrep()
        self.printingSetup = PrintingSetup()

class GeneralTasks:

    def __init__(self):

        self.roomTemperature = RoomTemperature()
        self.roomHumidity = RoomHumidity()
        self.tip = Tip()
        self.slide = Slide()
        self.oil = Oil()
        self.mix = Mix()

    def populate(self, ws):
        self.roomTemperature.populate(ws)
        self.roomHumidity.populate(ws)
        self.tip.populate(ws)
        self.slide.populate(ws)
        self.oil.populate(ws)
        self.mix.populate(ws)

class RoomTemperature(Base):

    def __init__(self):

        super(RoomTemperature, self).__init__()
        self.roomTemperature = ""

    def populate(self, ws):
        self.doneBy = ws['E15'].value
        self.startedAt = ws['G15'].value
        self.roomTemperature = ws['L15'].value

class RoomHumidity(Base):

    def __init__(self):

        super(RoomHumidity, self).__init__()
        self.roomHumidity = ""

    def populate(self, ws):
        self.doneBy = ws['E17'].value
        self.startedAt = ws['G17'].value
        self.roomHumidity = ws['L17'].value

class Tip(Base):

    def __init__(self):

        super(Tip, self).__init__()
        self.size = ""
        self.tipBatch = ""
        self.tipID = ""

    def populate(self, ws):
        self.doneBy = ws['E19'].value
        self.size = ws['L19'].value
        self.tipBatch = ws['O19'].value
        self.tipID = ws['S19'].value

class Slide(Base):

    def __init__(self):

        super(Slide, self).__init__()
        self.CA = ""
        self.batch = ""
        self.slideID = ""

    def populate(self, ws):
        self.doneBy = ws['E21'].value
        self.CA = ws['L21'].value
        self.batch = ws['O21'].value
        self.slideID = ws['S21'].value

class Oil(Base):

    def __init__(self):

        super(Oil, self).__init__()
        self.oilID = ""

    def populate(self, ws):
        self.doneBy = ws['E23'].value
        self.oilID = ws['L23'].value

class Mix(Base):

    def __init__(self):

        super(Mix, self).__init__()
        self.type = ""
        self.claire633 = ""
        self.claire594 = ""
        self.claire532 = ""
        self.claire488 = ""

    def populate(self, ws):
        self.doneBy = ws['E25'].value
        self.type = ws['L25'].value
        self.claire633 = ws['N25'].value
        self.claire594 = ws['P25'].value
        self.claire532 = ws['R25'].value
        self.claire488 = ws['S25'].value
        
class ElectrodePrep():

    def __init__(self):

        self.sonication = Sonication()
        self.washing = Washing()
        self.drying = Drying()
        self.flaming = Flaming()

    def populate(self, ws):
        self.sonication.populate(ws)
        self.washing.populate(ws)
        self.drying.populate(ws)
        self.flaming.populate(ws)

class Sonication(Base):

    def __init__(self):

        super(Sonication, self).__init__()

    def populate(self, ws):
        self.basePopulate(ws, 27)
        

class Washing(Base):

    def __init__(self):

        super(Washing, self).__init__()

    def populate(self, ws):
        self.basePopulate(ws, 29)

class Drying(Base):

    def __init__(self):

        super(Drying, self).__init__()

    def populate(self, ws):
        self.basePopulate(ws, 31)

class Flaming(Base):

    def __init__(self):

        super(Flaming, self).__init__()

    def populate(self, ws):
        self.basePopulate(ws, 33)

class PrintingSetup():

    def __init__(self):

        self.filling = Filling()
        self.mounting = Mounting()
        self.positionSlide = PositionSlide()
        self.positionOil = PositionOil()

    def populate(self, ws):
        self.filling.populate(ws)
        self.mounting.populate(ws)
        self.positionSlide.populate(ws)
        self.positionOil.populate(ws)

class Filling(Base):

    def __init__(self):

        super(Filling, self).__init__()

    def populate(self, ws):
        self.basePopulate(ws, 35)

class Mounting(Base):

    def __init__(self):

        super(Mounting, self).__init__()

    def populate(self, ws):
        self.basePopulate(ws, 37)

class PositionSlide(Base):

    def __init__(self):

        super(PositionSlide, self).__init__()

    def populate(self, ws):
        self.basePopulate(ws, 39)

class PositionOil(Base):

    def __init__(self):

        super(PositionOil, self).__init__()
        self.lowHumidity = ""

    def populate(self, ws):
        self.basePopulate(ws, 33)
        self.lowHumidity = ws['L41'].value
        

if __name__ == "__main__":

    wb = load_workbook('Z:\\SOPs\\Completed Checklists\\Data\\Printing\\Printing 1 2015-10-09 1130.xlsm')
    ws = wb.active
    
    psd = PrintSampleData(ws['E8'].value)

    print(psd.sampleName)
    psd.generalTasks.populate(ws)
    psd.electrodePrep.populate(ws)
    psd.printingSetup.populate(ws)

    #testing
    print(psd.generalTasks.roomTemperature.roomTemperature)
    print(psd.generalTasks.roomTemperature.doneBy)
    print(psd.generalTasks.mix.doneBy)
    print(psd.printingSetup.positionOil.doneBy)
    print(psd.printingSetup.positionOil.startedAt)
    print(psd.printingSetup.positionOil.lowHumidity)

