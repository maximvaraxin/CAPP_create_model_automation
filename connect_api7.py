import pythoncom
from win32com.client import Dispatch, gencache

class ConnectApi7():

    def __init__(self):
        module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.api7 = module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
        self.const7 = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        self.connect7 = module

    def getApplication(self):
        return self.api7.Application

    def getActiveDocument(self):
        return self.api7.Application.ActiveDocument

    def getTypeActiveDocument(self):
        return self.api7.Application.ActiveDocument.DocumentType

