"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 8: Работа с трехмерными объектами: 3d face, полигональные сети, солиды
"""
from ncad_activex_tools import ncad_tools
import win32com.client
import random

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc_ms = ncad_doc.ModelSpace
    for model_object in ncad_doc_ms:
        model_object.Delete()
    """
    На практике, что в ActiveX, что в NET/NRX работа с 3D-объектами как правило ограничена 3д-полилиниями и, как максимум, 
    работой с 3д-солидами. Полигональные сети и грани являются уже более узкими объектными классами, которые тем не менее
     будет полезно рассмотреть.
     
     Пример полигональной сети -- это визуализация одноканального растрового снимка типа DEM (с высотами).
     Также обращаем внимание, что для работы с функциями 3д-тел необходима лицензия на 3D-модуль у Пользователя!!! 
     Иначе, будут выбиваться ошибки
    """
    def faces_work():
        """
        Обращаем внимание, что одиночная грань не обязательно должна быть планарной, то есть 3 точки грани могут
        образовывать плоскость, а четвертая не обязательно будет ей принадлежать. При желании задать грань по 3 точкам,
        одна из вершин будет задублированной.
        :return:
        """
        Acad3dFace_object = ncad_doc_ms.Add3DFace([0, 0, 0], [0, 1, 1], [1, 1, 2], [1, 0, 3])
        props = {
            "Coordinates": Acad3dFace_object.Coordinates
        }
        print(props)
        pass


    def mesh3d_work():
        """
        Функционально равен методу AddPolyfaceMesh(), с теми же аргументами
        """
        width = 15
        height = 12
        points_matrix = []
        for i in range(1, width, 1):
            for j in range(1, height, 1):
                points_matrix.append(i)
                points_matrix.append(j)
                points_matrix.append(random.random()*2)
        print(points_matrix)
        Acad3dMesh_object = ncad_doc_ms.Add3DMesh(width, height, points_matrix)
        props = {
            "Coordinates": Acad3dMesh_object.Coordinates,
            "MClose": Acad3dMesh_object.MClose,
            "NClose": Acad3dMesh_object.NClose,
            "MDensity": Acad3dMesh_object.MDensity,
            "NDensity": Acad3dMesh_object.NDensity,
            "MVertexCount": Acad3dMesh_object.MVertexCount,
            "NVertexCount": Acad3dMesh_object.NVertexCount
        }
        print(props)
        pass
    faces_work()
    mesh3d_work()
    """
    Касательно 3D Solid'ов, есть ряд методов создающих простые тела:
    AddBox и AddSolid -- параллелепипед
    AddCone и AddEllipticalCone -- конус
    AddCylinder и AddEllipticalCylinder - цилиндр
    AddExtrudedSolid, AddExtrudedSolidAlongPath и AddRevolvedSolid - выдавливание солида вдоль траектории
    AddTorus -- торус (бублик, грубо говоря)
    AddSphere -- сфера
    Helix -- спираль, без опции создания, только чтение
    
    Всё перечисленное выше многообразие методов имеет единый интерфейс -- Acad3DSolid, при это почти все геометрические 
    методы по работе с ними по анализу зафиксированы в документации, как нереализованные "Not implemented", при этом 
    отдельные операции, например, взятие центроида работают, что позволяет сделать вывод о частичной реализации. 
    """
    AcadSphere_object = ncad_doc_ms.AddSphere([-10, -10, 0], 5)
    props = {
        "Centroid": AcadSphere_object.Centroid,
        "Volume": AcadSphere_object.Volume
    }
    print(props)
    #Обновляем прорисовку графики и центрируем экран в месте объектов
    ncad_doc.Application.Update()
    ncad_doc.Application.ZoomExtents()