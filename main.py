from connect_api7 import ConnectApi7
from connect_api5 import ConnectApi5

if __name__ == "__main__":
    try:
        param_a = 60
        param_b = 35
        param_c = 10

        api7 = ConnectApi7()
        application = api7.getApplication()
        activeDoc = api7.getActiveDocument()
        activeDocType = api7.getTypeActiveDocument()

        if activeDocType == 4:
            api5 = ConnectApi5()
            IKompasDocument3D = api7.connect7.IKompasDocument3D(activeDoc)
            IKompasDocument2D = api7.connect7.IKompasDocument2D(activeDoc)

            if param_a >= param_b:
                api5.createModelAutomation(activeDoc, param_a, param_b, param_c)
            else:
                api7.getApplication().MessageBoxEx('Размер отверстия больше чем размер обоймы', 'Ошибка размеров', False)
        else:
            application.MessageBoxEx('Необходимо открыть документ типа "Деталь"', 'Ошибка типа файла', False )
    except Exception as e:
        api7.getApplication().MessageBoxEx(e, 'Ошибка', False )
