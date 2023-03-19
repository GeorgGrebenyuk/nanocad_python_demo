"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 2: Понятие AcadBlock и AcadEntity, приведение интерфейсов
"""
from ncad_activex_tools import ncad_tools
import win32com.client

sample_doc_path = ncad_tools.get_sample_drawing("nCAD. Модель-лист (оформление).dwg")

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc = ncad_doc.Open(sample_doc_path)
    """
    В объектной модели nanoCAD (и AutoCAD) существует устойчивое понятие "Блок", которым принятно называть 2 вещи:
    1. Объект из "таблицы записей блоков" (BlockTableRecord), к которым относятся пространство модели и листов, то есть 
    в составе которых расположены иные объекты модели
    2. Пользовательский "массив элементов" с возможностью наличия атрибутов и вставкой экземпляров такого блока в 
    модель, и вхождения будут называться "Вхождения блока"
    3*. Также к Блокам сводятся все внешние ссылки, поэтому их тоже будем так называть
    
    Вне зависимости от трактовки (1 или 2 варианты), интерфейс у блока будет одинаковый -- это OdaX.IAcadBlock, который 
    соответственно наследуют 2 интерфейса -- OdaX.IAcadModelSpace и OdaX.IAcadPaperSpace. ModelSpace являеется самым 
    частым используемым пространством через любые API: будь то текущий ActiveX, .NET API или ObjectARX (NRX)
    
    В то же время, напрямую с Блоками для листов не работают. Как правило, получают перечень всех листов документа как 
    ncad_doc.Layouts и перебирая эту коллекцию, получают связанный с листом Блок как поле ".Block", что и показано ниже:
    """
    ncad_modespace = ncad_doc.ModelSpace
    print("В модели " + str(ncad_modespace.Count) + " объектов")
    ncad_layouts = ncad_doc.Layouts
    for layout_counter in range(0, ncad_layouts.Count, 1):
        ncad_layout = ncad_layouts.Item(layout_counter)
        ncad_block_for_layout = ncad_layout.Block
        print("На листе " + ncad_layout.Name + " "
              + str(ncad_block_for_layout.Count) + " объектов")
    """
    Каждый объект Блока -- это интерфейс OdaX.IAcadEntity. Строго говоря, это "материальный" объект, то есть имеющий 
    геометрическое отображение. Есть ряд объектов, не имеющих отображение но являющихся частью документа -- это,
     к примеру, слои, размерные и текстовые стили и т.д.
     
     Подобный объект получается запросом через метод Block.Item(i) в теле цикла-перебора объектов блока от 0 до Count.
     Заметим, что возможен и другой сценарий перебора -- не только через цикл с запосом Item но и прямым перебором.
     То есть записи ниже равноценны.
     
    for AcadEntity in ncad_modespace:
        print(AcadEntity.ObjectName)
    for AcadEntity_counter in range(0, ncad_modespace.Count, 1):
        print(ncad_modespace.Item(AcadEntity_counter).ObjectName)
    """

    """
    Получаемые объекты AcadEntity можно рассматривать как объекты модели -- в этом случае необходимо привести интерфейс 
    к интерфейсу соответствующего объекта модели. Узнать, к какому интерфейсу надо приводить данный можно по значению 
    поля ObjectName. Например, для Отрезка (ObjectName = AcDbLine) наименование целевого интерфейса будет AcadLine 
    (или IAcadLine).
    Сама операция приведения будет осуществляться через стандартную механику win32com.client.CastTo()
    """
    for AcadEntity_counter in range(0, ncad_modespace.Count, 1):
        AcadEntity = ncad_modespace.Item(AcadEntity_counter)
        if AcadEntity.ObjectName == "AcDbLine":
            object_as_AcDbLine = win32com.client.CastTo(AcadEntity, "IAcadLine")
            if object_as_AcDbLine is not None:
                print("Длина отрезка = " + str(object_as_AcDbLine.Length))
                break
    """
    Если у чертежа есть внешние ссылки, то они также рассматриваются как интерфейсы AcadBlock. Рассмотрим далее 
    доступ к перечню внешних ссылок для данного чертежа:
    """
    for AcadBlock in ncad_doc.Blocks:
        if AcadBlock.IsXRef:
            print("Внешняя ссылка " + AcadBlock.Name)

print("End!")