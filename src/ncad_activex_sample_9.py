"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 9: Вставка растровых данных и работа с OLE-объектами
"""
from ncad_activex_tools import ncad_tools
import win32com.client
import os

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc_ms = ncad_doc.ModelSpace
    test_image_path = os.path.join(os.getcwd(), 'sample_image_8.png')
    for model_object in ncad_doc_ms:
        model_object.Delete()
    """
    Для работы с растровыми данными в ActiveX предусмотрены 2 подхода: с одной стороны, это работа с внешними данными - 
    интерфейсом AcadRasterImage (метод AddRaster()) и внутренними преобразованными данными - интерфейсом IAcadOle 
    (метод AddEmbeddedRaster()). Параметры у методов одинаковые, различие только в способе хранения.
    При первом варианте при смене файлового пути, изображение перестает отображаться в документе, а при втором -- оно 
    внедряется в чертёж и существует как OLE-объект
    
    При этом отметим, что вставка OLE в nanoCAD ActiveX не реализована
    """
    AcadRasterImage_object = ncad_doc_ms.AddRaster(test_image_path, [0, 0, 0], 1.0, 0.0)
    AcadRasterImage_props = {
        "Brightness": AcadRasterImage_object.Brightness,
        "ClippingEnabled": AcadRasterImage_object.ClippingEnabled,
        "Contrast": AcadRasterImage_object.Contrast,
        "Fade": AcadRasterImage_object.Fade,
        "Height": AcadRasterImage_object.Height,
        "ImageFile": AcadRasterImage_object.ImageFile,
        "ImageHeight": AcadRasterImage_object.ImageHeight,
        "ImageVisibility": AcadRasterImage_object.ImageVisibility,
        "ImageWidth": AcadRasterImage_object.ImageWidth,
        "Origin": AcadRasterImage_object.Origin,
        "ShowRotation": AcadRasterImage_object.ShowRotation,
    }
    """
    Отметим ещё важный момент. Точка вставки снимка (origin) по документации - lower left corner (левый нижний угол), 
    в то время как растровые данные из категории ГИС как правило имеют точку привязки (начало снимка) в верхнем левом 
    углу, и для получения нижней границы нужны преобразования (с использованием GDAL-информации о растре, 6 параметрах), 
    они же прописываются в файлы привязки одноименные с файлом изображения.
    """
    print(AcadRasterImage_props)

    """
    Касательно работы с OLE. В ActiveX не работает механика их вставки. В ознакомительных целях код прикладывается 
    комментариями
    
    AcadOle_object = ncad_doc_ms.AddEmbeddedRaster(test_image_path, [-10, 0, 0], 0.56, 1.0)
    AcadOle_props = {
        "Height": AcadOle_object.Height,
        "LockAspectRatio": AcadOle_object.LockAspectRatio,
        "OleItemType": AcadOle_object.OleItemType,
        "OlePlotQuality": AcadOle_object.OlePlotQuality,
        "OleSourceApp": AcadOle_object.OleSourceApp,
        "ScaleHeight": AcadOle_object.ScaleHeight
    }
    print(AcadOle_props)
    
    """
    #Обновляем прорисовку графики и центрируем экран в месте объектов
    ncad_doc.Application.Update()
    ncad_doc.Application.ZoomExtents()
