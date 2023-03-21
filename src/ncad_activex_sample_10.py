"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 10: Общие операции по работе с геометрией объектов и ПСК
"""
from ncad_activex_tools import ncad_tools
import win32com.client
import os

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc_ms = ncad_doc.ModelSpace
    for acad_entity in ncad_doc_ms:
        acad_entity.Delete()
    #Создадим несколько новых объектов для дальнейшей работы
    AcadCircle_object = ncad_doc_ms.AddCircle([0, 0, 0], 5)
    AcadPoint_object = ncad_doc_ms.AddPoint([1, 2, 0])
    """
    Всякий "материальный" объект модели наследует интерфейс AcadEntity, методы которого позволяют рассматривать объект 
    как часть документа, и проводить над объектом ряд геометрических операций. Отдельная группа методов находиться у 
    солидов (Acad3DSolid), но большая часть из них полностью не реализована.
    
    Наиболее полезный метод -- это GetBoundingBox(), и пожалуй всё. Метод IntersectWith() к примеру не работает, а 
    метода Contains вообще нет.
    """
    #Ограничивающая рамка для круга
    bbox_min_point = []
    bbox_max_point = []
    """
    Из-за особенности реализации указателей в Python, процедура GetBoundingBox() фактически имеет возвращаемое значение
    """
    bbox_circle = AcadCircle_object.GetBoundingBox(bbox_min_point, bbox_max_point)
    print(bbox_circle)

    """
    Касательно ПСК и преобразований между ними:
    
    Доступ к перечню ПСК чертежа осуществляется из-под Базы данных чертежа Database.UserCoordinateSystems
    Пересчёт между разными системами координат осуществляется через метод doc.Utility.TranslateCoordinates()
    """

    doc_UCS_group = ncad_doc.Database.UserCoordinateSystems
    for UCS in doc_UCS_group:
        UCS.Delete()
    #Создаем 2 ПСК
    #UCS_new_1 = doc_UCS_group.Add([0,0,0], [1,0,0], [0,1,0], "UCS_World")
    UCS_new_2 = doc_UCS_group.Add([5, 5, 0], [6, 5, 0], [5, 6, 0], "UCS_Users")

    #Выводим информацию по созданным ПСК
    for one_UCS in doc_UCS_group:
        print(one_UCS.Name + " " + str(one_UCS.GetUCSMatrix()))
    #Осуществляем пересчет для одной точки
    """
    Для того, чтобы узнать, какая координата будет у точки в Новой ПСК необходимо:
    1. Сделать у текущего документа активной целевую ПСК ncad_doc.ActiveUCS = UCS_new_2
    2. В метод TranslateCoordinate() первым аргументом занести координаты точки в текущей мировой СК файла, затем вторым 
    членом указать "0", это не что иное как enum AcCoordinateSystem = acWorld (0)
    3. Третий аргумент = 1, это тот же enum что в п. 2, только = acUCS (1)
    4. Четвертый аргумент всегда = False для точек и True для векторов.
    5. Пятый аргумент опционален и обозначает вектор нормали
    
    Полученную координату можно проверить у точки чертежа. Это та самая точка, которую мы создавали выше.
    """
    ncad_doc.ActiveUCS = UCS_new_2
    point_in_Users_coords = ncad_doc.Utility.TranslateCoordinates([1, 2, 0], 0, 1, False)
    print(point_in_Users_coords)


    #Обновляем прорисовку графики и центрируем экран в месте объектов
    ncad_doc.Application.Update()
    ncad_doc.Application.ZoomExtents()