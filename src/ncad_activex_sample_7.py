"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 7: Работа с площадными объектами: штриховки
"""
from ncad_activex_tools import ncad_tools
import win32com.client

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc_ms = ncad_doc.ModelSpace
    for ncad_object in ncad_doc_ms:
        ncad_object.Delete()

    """
    Штриховки - это комплексные объекты, состоящие из Определения и Геометрии. При желании создать новую штриховку, 
    сперва создается её нематериальное поределение, а потом к нему добавляются контуры с заполнением. Рассмотрим далее
     эту механику на примере штриховки типа SOLID для прямоугольной области, а затем получим параметры этой штриховки:
     
     Первый параметр: PatternType это enum AcPatternType Энумератор, причем он не может быть изменен после создания 
     штриховки, только  получен как свойство.
     Вторйо параметр: PatternName - это наименование штриховки. Видимый список в Диспетчере штриховок -- это они и есть
     Третий параметр: Associativity установка ассоциативности штриховки. Можно сменить на готовом объекте. Задает изменение 
     штриовки от изменения родительский сущностей-границ
     Четвертый параметр (опциональный): HatchObjectType это enum AcHatchObjectType, то есть определение является ли это 
     непосредственно штриховкой или градиентом
    """
    print("Creation")
    AcadHatch_definition = ncad_doc_ms.AddHatch(1, "SOLID", True, 0)
    """
    После создания Определения штриховки, необходимо добавить к ней контуры. Здесь процедура похожа на логику в .NET API:
    у нас имеется метод InsertLoopAt(),который и добавляет (или заменяет). Альтернативынй вариант, при формировании только 
    внешней границы можно воспользоваться методом AppendOuterLoop(). Мы же воспользуемся общим случаем InsertLoopAt и 
    сделаем 2 непересекающих друг друга контура:
    """
    points_outer1 = [20, 10, 50, 0, 60, -20, 20, -10, 20, 10]
    points_outer2 = [50, 10, 100, 20, 60, 30, 50, 10]
    AcadLWPline_as_outer_loop1 = ncad_doc_ms.AddLightWeightPolyline(points_outer1)
    AcadLWPline_as_outer_loop2 = ncad_doc_ms.AddLightWeightPolyline(points_outer2)
    AcadHatch_definition.InsertLoopAt(0, 0, [AcadLWPline_as_outer_loop1])
    AcadHatch_definition.InsertLoopAt(1, 0, [AcadLWPline_as_outer_loop2])
    hatch_props = {
        "Area": AcadHatch_definition.Area,
        "AssociativeHatch": AcadHatch_definition.AssociativeHatch,
        "GradientAngle": AcadHatch_definition.GradientAngle,
        "GradientCentered": AcadHatch_definition.GradientCentered,
        #"GradientColor1": AcadHatch_definition.GradientColor1,
        #"GradientColor2": AcadHatch_definition.GradientColor2,
        "GradientName": AcadHatch_definition.GradientName,
        "HatchObjectType": AcadHatch_definition.HatchObjectType,
        "HatchStyle": AcadHatch_definition.HatchStyle,
        "ISOPenWidth": AcadHatch_definition.ISOPenWidth
        #и т.д.
    }
    print(hatch_props)

    #Обновляем прорисовку графики и центрируем экран в месте объектов
    ncad_doc.Application.Update()
    ncad_doc.Application.ZoomExtents()