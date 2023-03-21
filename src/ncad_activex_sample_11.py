"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 11: Работа с таблицами
"""
from ncad_activex_tools import ncad_tools
import win32com.client

sample_doc_path = ncad_tools.get_sample_drawing("Чертеж из ГОСТ 21.501-2011 (nanoCAD).dwg")

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc = ncad_doc.Open(sample_doc_path)
    ncad_doc_ms = ncad_doc.ModelSpace
    """
    В качестве примера работы с таблицами рассмотрим процесс подсчета числа объектов чертежа с сортировкой по классам, 
    то есть переберём все объекты из ModelSpace, у каждого возьмем ObjectName и посчитаем количество одинаковых
    """
    model_objects_data = dict()
    for one_acad_entity in ncad_doc_ms:
        o_name = one_acad_entity.ObjectName
        if o_name not in model_objects_data.keys():
            model_objects_data[o_name] = 1
        else:
            model_objects_data[o_name] += 1
        #С удалением других таблиц
        if o_name == "AcDbTable":
            one_acad_entity.Delete()
    print(model_objects_data)

    """
    Теперь переходим к генерации таблицы. Метод по ее добавлению также находится у AcadBlock (у нас, ModelSpace)
    При этом сперва создается "болванка" таблицы с пустыми полями, а последующее заполнение уже происходит при работе с 
    создавшимся объектом
    
    Обращаем внимание, что первая строка таблицы идёт без колонок, она = заголовку, у нее статичный индекс 0,0
    Ширину колонок можно изменить отдельно для каждой колонки, то есть аргумент в конструкторе 
    таблицы не является не-редактируемым для ОБЩЕЙ длины. 
    """
    AcadTable_object = ncad_doc_ms.AddTable([400, 410, 0], len(model_objects_data.keys()) + 2, 2, 1, 70)
    AcadTable_object.SetText(0, 0, "Спецификация количества объектов модели")
    AcadTable_object.SetText(1, 0, "Объектный класс")
    AcadTable_object.SetText(1, 1, "Число, шт.")
    #Установка ширины колонок
    AcadTable_object.SetColumnWidth(0, 100)
    AcadTable_object.SetColumnWidth(1, 25)
    counter_rows = 2
    for class_name, obj_count in model_objects_data.items():
        #Устаналиваем выравнивание для первой колонки
        AcadTable_object.SetCellAlignment(counter_rows, 0, 7)
        AcadTable_object.SetText(counter_rows, 0, class_name)
        AcadTable_object.SetText(counter_rows, 1, obj_count)
        counter_rows += 1

    #Обновляем прорисовку графики и центрируем экран в месте объектов
    ncad_doc.Application.Update()
    ncad_doc.Application.ZoomExtents()