"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 5: Работа со стилями текста (аналогично про другие стили) и объектами AcadText и AcadMText
"""
from ncad_activex_tools import ncad_tools
import win32com.client

sample_doc_path = ncad_tools.get_sample_drawing("Чертеж из ГОСТ 21.501-2011 (nanoCAD).dwg")

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc = ncad_doc.Open(sample_doc_path)
    """
    Получаем набор действующих текстовых стилей БД документа
    """
    ncad_db = ncad_doc.Database
    doc_text_styles = list()
    for one_TextStyle in ncad_db.TextStyles:
        TextStyle_props = {
            "Name": one_TextStyle.Name,
            "fontFile": one_TextStyle.fontFile,
            "BigFontFile": one_TextStyle.BigFontFile,
            "TextGenerationFlag": one_TextStyle.TextGenerationFlag,
            "Height": one_TextStyle.Height,
            " Width": one_TextStyle. Width
        }
        print(TextStyle_props)
    """
    Создадим ниже новый текстовый стиль для шрифта Times New Roman
    Примечание: в аргументах присвоения стилю шрифта есть параметры PitchAndFamily, CharSet описанные здесь 
    https://help.autodesk.com/view/ACD/2023/ENU/?guid=GUID-DB668114-2395-43C6-858C-2F2514C4BF46
    После создания текстового стиля делаем его активным в документе для автоматического связывания с новыми создаваемыми
    объектами текста (AcadText и AcadMText)
    """
    TNR_TestStyle_name = "TimesNewRoman TEXT STYLE"
    TNR_TextStyle = ncad_db.TextStyles.Add(TNR_TestStyle_name)
    TNR_TextStyle.SetFont("Times New Roman", False, False, 1, 64)
    """
    Примечание: выше мы создавали стиль для TTF шрифта, предполагая что он подгружен в nanoCAD (здесь, из стандартной 
    папки шрифтов Windows C:\Windows\Fonts). В общем случае это надо дополнительно проверять.
    В теории, аналогично можно "перезадать" стилю шрифт через ту же команду SetFont если у него установлен другой шрифт
    """
    TNR_TextStyle.Height = 0.05
    ncad_doc.ActiveTextStyle = TNR_TextStyle


    """
    Далее рассмотрим, как работать с текстом в чертеже.Покажем это на примере получения текстовых строк 
    имеющегося однострочного текста и создание нового объекта многострочного текста в одном из полей штампа в чертеже.
    """
    #Работа с однострочным текстом
    all_text_data = []
    for one_AcadEntity in ncad_doc.ModelSpace:
        if one_AcadEntity.ObjectName == "AcDbText":
            #В аргументе у нас используется IAcadText вместо AcadText в силу того, что таковое приведение невозможно
            object_as_AcadText = win32com.client.CastTo(one_AcadEntity, "IAcadText")
            all_text_data.append(object_as_AcadText.TextString)
    print(all_text_data)
    #Работа с многострочным текстом
    mtext_point_insertion = [207, 17, 0]
    AcadMText_object = ncad_doc.ModelSpace.AddMText(mtext_point_insertion, 30, "Надпись, созданная через ActiveX")
    AcadMText_object.Height = 2
    #Примечательно, что для MText нет возможности установить выравнивание текста как для однострочного текста
ncad_doc.Application.Update()