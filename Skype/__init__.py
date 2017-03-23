import eg

eg.RegisterPlugin(
    name = "Skype",
    author = "Daniel J. Dunn II (orbitaldan@gmail.com)",
    version = "0.0.1",
    kind = "program",
    guid = "{950A9379-87D1-4981-99DF-AB727D53A0EB}",
    url = "https://github.com/OrbitalDan/EventghostSkypePlugin",
    description = (
        'Adds events and actions to interact with <a href="http://www.skype.com/">Skype</a>.'
    ),
)

# TODO: Add icon like so:
#icon = (
#    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAACDElEQVR42pWTT0gUcRTH"
#    "P7OQemvKQ7pSbkQgHWLm0k2cgxpUsLOBFUS4UmCBZEuHLsHugBAeLPca4U4RmILsinno"
#    "j+SCYLTETiUsRsFuBOpSMoE2SGzbONg6q6vR7/J+vx/v832/937vCeyygsFgcd3qui7s"
#    "5CPsBttgYmOvukSCxxtF//ucGdhR4C9cLBZVx0kQEusiyxnd0HrPSuPDM0Qm8kJFga1w"
#    "NpsleidES8NL1I6rkBzj8tAXYyhVkLcJuGHTNNEfRPmR1ghfOwo1p2B2hO7HefP+m8Jh"
#    "290sE3DDtuXdZIjwuVXE2gNwpAuexdCeVxMZ+5xw10Rww+l0Wn14t4vOEx+Q5CY7ah0U"
#    "bJc1D+Z3Ae1ehs7oOLIsl0QEd2RNs597+iPkFsH6DScvwNfX9nkZMnNMr/yCtkcoilIq"
#    "7HaBphFY+Qmt3TY4D1Mz4PXAnjUGJ/Io4VkkSdoU2JpCIHCG+HAP0twU5D/BoSrMhW9o"
#    "SZGWKwOoqlqCnRTcRTSMFLF+P1L6KeyznKh60iJXf5HeW2FEUSyDy37B690fm7zdHpSs"
#    "t3Cwiuz8IlGjEX/PgJOzu6Hcre1sfHU1kXj0fFiyUlBcRRtdYm9bHzdCNzd7vgJcEuho"
#    "rn01er1eMYwFQk/MaV/zJaVSK1caKuei/Vh1XG7wKP0vLM0+Dv5jmCpOo2/DZv93nP8A"
#    "opkfXpsJ2wUAAAAASUVORK5CYII="
#),

from threading import Event, Thread
import win32com.client
import pythoncom

# Event handler object to relay events back to main plugin class
class SkypeCallStatusHandler(object):
    
    def SetPluginParent(self,parent):
        self.parent = parent
        
    def SetSkypeClient(self,skype):
        self.skype = skype
        
    def OnCallStatus(self, theCall, callStatus):
        self.parent.SetStatus(self.skype.Convert.CallStatusToText(callStatus))
    
#End SkypeCallStatusHandler

class Skype(eg.PluginBase):

    #--- Lifecycle Management -----------------------------------------
    
    def __init__(self):
        self.status = ""
        print "Skype Plugin is initialized."
    
    def __start__(self):
        self.stopThreadEvent = Event()
        thread = Thread(
            target=self.ThreadLoop,
            args=(self.stopThreadEvent, )
        )
        thread.start()    
        print "Skype Plugin is started."
    
    def __stop__(self):
        self.stopThreadEvent.set()
        print "Skype Plugin is stopped."
    
    def __close__(self):
        print "Skype Plugin is closed."
    
    #--- External COM Thread ------------------------------------------
    
    def ThreadLoop(self, stopThreadEvent):
        pythoncom.CoInitialize()
        skype = win32com.client.Dispatch("Skype4COM.Skype")
        skype.Attach()
        handler = win32com.client.WithEvents(skype, SkypeCallStatusHandler)
        handler.SetSkypeClient(skype)
        handler.SetPluginParent(self)
        while not stopThreadEvent.isSet():
            pythoncom.PumpWaitingMessages()
            stopThreadEvent.wait(1.0)
    
    def SetStatus(self, status):
        if ( self.status != status ):
            self.status = status
            self.TriggerEvent(status)
    
    #------------------------------------------------------------------
    
#End Skype Plugin