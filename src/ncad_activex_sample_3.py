"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 3: Понятие nanoCAD.Document и OdaX.AcadDatabase, доступ к набору интерфейсов Utility
"""
from ncad_activex_tools import ncad_tools
import win32com.client

sample_doc_path = ncad_tools.get_sample_drawing("nCAD. Модель-лист (оформление).dwg")

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc = ncad_doc.Open(sample_doc_path)
    """
    Получаемый интерфейс ncad_doc, он же nanoCAD.Document предоставляет доступ к методам открытия нового чертежа, 
    сохранения/экспорта текущего, управлением Свойств документа, единиц измерения и главное -- к Базе данных чертежа.
    
    База данных в свою очередь (интерфейс OdaX.IAcadDatabase) предоставляет доступ к чтению и изменению данных модели: 
    слои, листы, информация о документе SummaryInfo, размерные стили, текстовые стили и т.д.
    Рассмотрим далее некоторые сценарии работ со свойствами документа и стилями:
    """
    ncad_db = ncad_doc.Database
    """
    Работа с интерфейсом OdaX.AcadSummaryInfo
    Доступ к нему есть также и в NRX, и в .NET API. Это ряд стандартных и пользовательских свойств, которые можно 
    прописывать в документ и "таскать" вслед за ним. Это актуально, например, при работе с документом и записью в него 
    неких параметров, хранящихся отдельно от документа. Тогда можно будет всегда сопоставить данный документ с набором 
    параметров для него посредством введенного CustomInfo что и рассмотрено ниже.
    """
    ncad_doc_info = ncad_db.SummaryInfo
    test_key = "for_activex_learning"
    test_value = "Hello, ActiveX!"
    was_found = False
    for CustomInfo_counter in range(0, ncad_doc_info.NumCustomInfo(), 1):
        p_Key = ""
        p_Value = ""
        CustomInfo = ncad_doc_info.GetCustomByIndex(CustomInfo_counter, p_Key, p_Value)
        if CustomInfo[0] != "":
            print(CustomInfo[0] + " ", CustomInfo[1])
            if CustomInfo[0] == test_key:
                was_found = True
    if was_found == False:
        """
        Теперь если исполнить команду DWGPROPS и зайти в группу "Пользовательский" мы увидим там добавленное свойство
        """
        ncad_doc_info.AddCustomInfo(test_key, test_value)

    ncad_doc_Utility = ncad_doc.Utility
    """
    Всякий документ имеет доступ к внутреннему набору методов - интерфейсу OdaX.Utility. Часть методов направлена на 
    приведение типов объектов, например AngleToReal (Конвертирует значение угла из строки в тип Real) или 
    GetObjectIdString (Преобразует ObjectId в строку). Часть методов предоставляет доступ к вспомогательным процедурам, 
    ради которых в иных случаях пришлось бы подключать сторонние пакеты -- например, меню выбора файла ChooseFile в 
    альтернативу System.Windows.Forms.OpenFileDialog из .NET API или IronPython и похожей процедуры из-пол Windows-api у 
    С++. 
    Часть методов предоставляет дополнительные инструменты работы с документом -- например опция вывода сообщений в 
    командную строку Prompt()
    
    Отдельная группа методов предназначена для обработки пользовательского ввода - строк GetString(), чисел и объектов
    геометрии: GetPoint, GetPolyline, GetAngle. Специфика "геометрических" методов в том, что Пользователь их не 
    выбирает, а именно рисует, и как результат метода -- возвращается нарисованная временная геометрия в виде чисел. 
    Это может быть удобно при работе с объектами модели - для разных процедур генерации геометрии.
    """
    ncad_doc_Utility.Prompt("Это вывод в командную строку :)")
    """
    Работа с базой данных чертежа (интерфейс OdaX.IAcadDatabase) сводится к получению аннотативных стилей (при 
    необходимости на стороне кода создавать элементы оформления с нужным стилем, делая его на момент создания активным), 
    получению доступа к Слоям, Материалам документа - для их просмотра/изменения.
    
    К примеру, создадим новый слой "Hello ActiveX", сделаем его текущим с проверкой, нет ли такого уже в документе
    """
    ncad_doc_layers = ncad_db.Layers
    was_found = False
    need_layer_name = "Hello Active"
    for AcadLayer_counter in range(0,ncad_doc_layers.Count, 1 ):
        AcadLayer = ncad_doc_layers.Item(AcadLayer_counter)
        if AcadLayer.Name == need_layer_name:
            was_found = True
            #Делаем этот слой текущим активным
            ncad_doc.ActiveLayer = AcadLayer
            break
    if was_found == False:
        AcadLayer = ncad_doc_layers.Add(need_layer_name)
        ncad_doc.ActiveLayer = AcadLayer
    print("Имя активного слоя: " + ncad_doc.ActiveLayer.Name)
print("End!")