"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 6: Работа с линейными объектами: отрезки, полилинии, слпайны, окружности, эллипсы
"""
from ncad_activex_tools import ncad_tools
import win32com.client

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc_ms = ncad_doc.ModelSpace
    """
    Под линейными объектами будем понимать все такие, для задания которых нужно > 1 точки, и которые не образуют 
    замкнутуню область (кроме замкнутых полилиний)
    К подобным объектам можно отнести дуги (AcadArc), отрезки (AcadLine), полилинии (AcadLWPolyline), окружности 
    (AcadCircle), эллипсы (AcadEllipse), сплайны (AcadSpline).
    Также линейным можно назвать "рудиментный" AcadTrace, но его лучше вообще не использовать, так как со стороны 
    Пользователя нет инструментов работы с ним.
    
    Рассмотрим далее механику создания таких примитивов и, после создания, получения по ним информации.
    """
    #сперва удалим всё имеющееся
    for object_Acad in ncad_doc_ms:
        object_Acad.Delete()

    def circle_work():
        AcadCircle_object = ncad_doc_ms.AddCircle([0, 0, 0], 10)
        props = {
            "Area": AcadCircle_object.Area,
            "Center": AcadCircle_object.Center,
            "Circumference": AcadCircle_object.Circumference,
            "Diameter": AcadCircle_object.Diameter,
            "Radius": AcadCircle_object.Radius
        }
        print(props)
        pass


    def ellipse_work():
        """
        В параметрах создания эллипса фигурирует второй параметр = MajorAxis, он же "Большая полуось". Задается точкой,
         через которую проходит эллипс, а не размером (типизация VARIANT). Последнее число, 0 < RadiusRatio < 1.0
        """
        AcadEllipse_object = ncad_doc_ms.AddEllipse([30, 0, 0], [15, 0, 0], 0.3)
        props = {
            "Area": AcadEllipse_object.Area,
            "Center": AcadEllipse_object.Center,
            "EndAngle": AcadEllipse_object.EndAngle,
            "EndParameter": AcadEllipse_object.EndParameter,
            "EndPoint": AcadEllipse_object.EndPoint,
            "MajorAxis": AcadEllipse_object.MajorAxis,
            "MajorRadius": AcadEllipse_object.MajorRadius,
            "MinorAxis": AcadEllipse_object.MinorAxis,
            "MinorRadius": AcadEllipse_object.MinorRadius,
            "RadiusRatio": AcadEllipse_object.RadiusRatio,
            "StartAngle": AcadEllipse_object.StartAngle
        }
        print(props)
        pass
    def line_work():
        AcadLine_object = ncad_doc_ms.AddLine([50, -5, 0], [50, 5, 0])
        props = {
            "Angle": AcadLine_object.Angle,
            "Delta": AcadLine_object.Delta,
            "EndPoint": AcadLine_object.EndPoint,
            "Length": AcadLine_object.Length,
            "StartPoint": AcadLine_object.StartPoint
        }
        print(props)
        pass
    def lw_pline_work():
        """
        Стоит внести ясность насчет существования двух различных процедурных функций, создающих полилинии -- с одной
        стороны, это AddPolyline(), с другой -- AddLightWeightPolyline(); обе принимают на вход список образующих точек.
        Неудобство их в том, что оба имеют ObjectName = AcDbPolyline, при этом Polyline считаются более "старой"
        реализацией (уже как минимум 12 лет судя по https://forums.autodesk.com/t5/vba/make-acadentity-out-of-acadobject/m-p/2727771#M94547)
        То есть практично создавать методом AddLightWeightPolyline() и не задаваться подобным фопросом в будущем.

        Обращаем также внимание, что список образующих точек VerticesList - это не список скписков 2 координат, а
        именно единый список координат вида [x1, y1, x2, y2 .... xn, yn]. Также образаем внимание, что это "плоская
        геометия", то есть потрубны 2 координаты - только X и Y
        """
        AcadLWPolyline_object = ncad_doc_ms.AddLightWeightPolyline([60, 0, 70, 10, 70, 0])
        props = {
            "Area": AcadLWPolyline_object.Area,
            "Closed": AcadLWPolyline_object.Closed,
            "ConstantWidth": AcadLWPolyline_object.ConstantWidth,
            "Coordinates": AcadLWPolyline_object.Coordinates,
            "Elevation": AcadLWPolyline_object.Elevation,
            "Length": AcadLWPolyline_object.Length,
            "LinetypeGeneration": AcadLWPolyline_object.LinetypeGeneration
        }
        print(props)
        pass

    def spline_work():
        #https://github.com/therealhadron/svg2autocad/blob/443c5db2b0ff3a8a2a7c92f43b94dc9c6be0b949/main.py
        """
        В параметрах создания сплайна есть аргументы StartTangent и EndTangent. По типизации -- это векторы, они же
        3 координаты. Здесь для простоты они = 0,0,0

        По умолчанию, сплайн создается через режим "Определяющие точки". При желании сменить на "Управляющие вершины"
        добавьте позицию AcadSpline_object.ControlPoints = points
        """
        points = [-10, -15, 0, -20, -40, -10, 20, -20, 0]
        AcadSpline_object = ncad_doc_ms.AddSpline(points, [0, 0, 0], [0, 0, 0])

        props = {
            "Closed": AcadSpline_object.Closed,
            "ControlPoints": AcadSpline_object.ControlPoints,
            "Degree": AcadSpline_object.Degree,
            #"EndTangent": AcadSpline_object.EndTangent,
            "FitPoints": AcadSpline_object.FitPoints,
            "FitTolerance": AcadSpline_object.FitTolerance,
            "IsPeriodic": AcadSpline_object.IsPeriodic,
            "IsPlanar": AcadSpline_object.IsPlanar,
            "IsRational": AcadSpline_object.IsRational,
            "Knots": AcadSpline_object.Knots,
            "NumberOfControlPoints": AcadSpline_object.NumberOfControlPoints,
            "NumberOfFitPoints": AcadSpline_object.NumberOfFitPoints,
            "Weights": AcadSpline_object.Weights
        }
        print(props)
        pass

    circle_work()
    ellipse_work()
    line_work()
    lw_pline_work()
    spline_work()

#Обновляем прорисовку графики и центрируем экран в месте объектов
ncad_doc.Application.Update()
ncad_doc.Application.ZoomExtents()
