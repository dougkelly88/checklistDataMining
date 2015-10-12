import os, time

class Base(object):

    def __init__(self):
        self.doneBy = "ME"
        self.startedAt = ""

class PrintSampleData:

    def __init__(self):

        self.generalTasks = GeneralTasks()

class GeneralTasks:

    def __init__(self):

        self.roomTemperature = RoomTemperature()
        self.roomHumidity = RoomHumidity()
        self.tip = Tip()
        self.slide = Slide()
        self.oil = Oil()
        self.mix = Mix()

class RoomTemperature(Base):

    def __init__(self):

        super(RoomTemperature, self).__init__()
        self.roomTemperature = ""

class RoomHumidity(Base):

    def __init__(self):

        super(RoomHumidity, self).__init__()
        self.roomHumidity = ""

class Tip(Base):

    def __init__(self):

        super(Tip, self).__init__()
        self.size = ""
        self.tipBatch = ""
        self.tipID = ""

class Slide(Base):

    def __init__(self):

        super(Slide, self).__init__()
        self.CA = ""
        self.batch = ""
        self.slideID = ""

class Oil(Base):

    def __init__(self):

        super(Oil, self).__init__()
        self.oilID = ""

class Mix(Base):

    def __init__(self):

        super(Mix, self).__init__()
        self.type = ""
        self.claire633 = ""
        self.claire594 = ""
        self.claire532 = ""
        self.claire488 = "500000"
        
    
    



if __name__ == "__main__":

    psd = PrintSampleData()

    print(psd.generalTasks.mix.claire488)
    psd.generalTasks.mix.claire488 = "15"
    print(psd.generalTasks.mix.claire488)
    print(psd.generalTasks.roomTemperature.doneBy)
    

    
