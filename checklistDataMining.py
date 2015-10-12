import os, time

class Base(object):

    def __init__(self):
        self.doneBy = "ME"
        self.startedAt = ""

class RoomTemperature(Base):

    def __init__(self):

        super(RoomTemperature, self).__init__()
        #self.doneBy = ""
        #self.startedAt = ""
        self.roomTemperature = ""

class PrintSampleData:

    def __init__(self):

        self.generalTasks = self.GeneralTasks()

    class GeneralTasks:

        def __init__(self):

#            self.roomTemperature = self.RoomTemperature()
            self.roomTemperature = RoomTemperature()
            self.roomHumidity = self.RoomHumidity()
            self.tip = self.Tip()
            self.slide = self.Slide()
            self.oil = self.Oil()
            self.mix = self.Mix()

##        class RoomTemperature(Base):
##
##            def __init__(self):
##
##                super(RoomTemperature, self).__init__()
##                #self.doneBy = ""
##                #self.startedAt = ""
##                self.roomTemperature = ""

        class RoomHumidity:

            def __init__(self):

                self.doneBy = ""
                self.startedAt = ""
                self.roomHumidity = ""

        class Tip:

            def __init__(self):

                self.doneBy = ""
                self.size = ""
                self.tipBatch = ""
                self.tipID = ""

        class Slide:

            def __init__(self):

                self.doneBy = ""
                self.CA = ""
                self.batch = ""
                self.slideID = ""

        class Oil:

            def __init__(self):

                self.doneBy = ""
                self.oilID = ""

        class Mix:

            def __init__(self):

                self.doneBy = ""
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
    

    
