from connect_api7 import ConnectApi7
import pythoncom
from win32com.client import Dispatch, gencache

class ConnectApi5(ConnectApi7):

    def __init__(self):
        module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        self.kompas_object = module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(module.KompasObject.CLSID, pythoncom.IID_IDispatch))
        self.kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        self.kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
        self.kompas6_api5_module = module

    def createModelAutomation(self, activeDoc, *args):
        kompas_document = activeDoc
        kompas_document_3d =  ConnectApi7().connect7.IKompasDocument3D(activeDoc)
        iDocument3D = self.kompas_object.ActiveDocument3D()

        iPart7 = kompas_document_3d.TopPart
        iPart = iDocument3D.GetPart(self.kompas6_constants_3d.pTop_Part)

        iSketch = iPart.NewEntity(self.kompas6_constants_3d.o3d_sketch)
        iDefinition = iSketch.GetDefinition()
        iPlane = iPart.GetDefaultEntity(self.kompas6_constants_3d.o3d_planeXOY)
        iDefinition.SetPlane(iPlane)
        iSketch.Create()
        iDocument2D = iDefinition.BeginEdit()
        kompas_document_2d = ConnectApi7().connect7.IKompasDocument2D(activeDoc)
        iDocument2D = self.kompas_object.ActiveDocument2D()

        iDefinition.EndEdit()
        iPart7 = kompas_document_3d.TopPart
        iPart = iDocument3D.GetPart(self.kompas6_constants_3d.pTop_Part)

        iSketch = iPart.NewEntity(self.kompas6_constants_3d.o3d_sketch)
        iDefinition = iSketch.GetDefinition()
        iPlane = iPart.GetDefaultEntity(self.kompas6_constants_3d.o3d_planeXOY)
        iDefinition.SetPlane(iPlane)
        iSketch.Create()
        iDocument2D = iDefinition.BeginEdit()
        kompas_document_2d = ConnectApi7().connect7.IKompasDocument2D(activeDoc)
        iDocument2D = self.kompas_object.ActiveDocument2D()

        # create 3D model --start

        obj = iDocument2D.ksCircle(0, 0, args[0], 1)
        obj = iDocument2D.ksCircle(0, 0, args[1], 1)
        iDefinition.EndEdit()
        iDefinition.angle = 180
        iSketch.Update()
        iPart7 = kompas_document_3d.TopPart
        iPart = iDocument3D.GetPart(self.kompas6_constants_3d.pTop_Part)

        obj = iPart.NewEntity(self.kompas6_constants_3d.o3d_bossExtrusion)
        iDefinition = obj.GetDefinition()
        iCollection = iPart.EntityCollection(self.kompas6_constants_3d.o3d_edge)
        iCollection.SelectByPoint(args[0], 0, 0)
        iEdge = iCollection.Last()
        iEdgeDefinition = iEdge.GetDefinition()
        iSketch = iEdgeDefinition.GetOwnerEntity()
        iDefinition.SetSketch(iSketch)
        iExtrusionParam = iDefinition.ExtrusionParam()
        iExtrusionParam.direction = self.kompas6_constants_3d.dtNormal
        iExtrusionParam.depthNormal = args[2]
        iExtrusionParam.depthReverse = 0
        iExtrusionParam.draftOutwardNormal = False
        iExtrusionParam.draftOutwardReverse = False
        iExtrusionParam.draftValueNormal = 0
        iExtrusionParam.draftValueReverse = 0
        iExtrusionParam.typeNormal = self.kompas6_constants_3d.etBlind
        iExtrusionParam.typeReverse = self.kompas6_constants_3d.etBlind
        iThinParam = iDefinition.ThinParam()
        iThinParam.thin = False
        obj.name = "Элемент выдавливания:1"
        iColorParam = obj.ColorParam()
        iColorParam.ambient = 0.5
        iColorParam.color = 9474192
        iColorParam.diffuse = 0.6
        iColorParam.emission = 0.5
        iColorParam.shininess = 0.8
        iColorParam.specularity = 0.8
        iColorParam.transparency = 1
        obj.Create()

        # create 3D model --end