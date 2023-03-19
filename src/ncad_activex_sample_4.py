"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 4: Работа с точечными объектами: AcadPoint, Вхождения блоков, динамические блоки, анонимные блоки, атрибуты блоков
"""
from ncad_activex_tools import ncad_tools
import win32com.client

sample_doc_path = ncad_tools.get_sample_drawing("Чертеж из ГОСТ 21.501-2011 (nanoCAD).dwg")

ncad_doc = ncad_tools.get_active_document()
if ncad_doc is not None:
    ncad_doc = ncad_doc.Open(sample_doc_path)
    """
    Под точечными объектами будем далее понимать объекты, описываемые интерфейсами AcadPoint и AcadBlockReference.
    Несмотря на то, что AcadText и AcadMText также задаются одной точкой, мы рассмотрим их в другом контексте.
    
    Стоит отметить, что создаются объекты только из-под интерфейса родительского Блока (здесь, пространства модели 
    или пространства листа). Ниже создается точка с координатами 500,500. Стилистика точки настраивается вручную через
    диалоговое окно DDPTYPE.
    """
    ncad_doc_ms = ncad_doc.ModelSpace
    AcadPoint_object = ncad_doc_ms.AddPoint([500, 500, 0])
    print(AcadPoint_object.Coordinates)
    """
    Если же требуется перебрать объекты модели до совпадения с нужным типом, можно воспользоваться сценарием, 
    приведенным в предыдущем Уроке следующего вида:
    """
    def points_querry():
        for one_AcadEntity in ncad_doc_ms:
            if one_AcadEntity.ObjectName == "AcDbPoint":
                object_as_AcDbPoint = win32com.client.CastTo(one_AcadEntity, "IAcadPoint")
                if object_as_AcDbPoint.Coordinates == [500, 500, 0.0]:
                    one_AcadEntity.Delete()
        pass
    #points_querry()
    """
    Вхождения блоков, в отличие от точек, требуют для размещения в чертеже наименование родительского Блока в своей 
    функции InsertBlock(). Рассмотрим далее процесс получения наименований блоков, для которых можно создавать вхождения:
    """
    doc_blocks_instances = []
    for one_AcadBlock in ncad_doc.Blocks:
        #Отсеиваем блоки внешних ссылок и листов, а также "анонимные" блоки с символа "*"
        if one_AcadBlock.IsXRef == False and \
            one_AcadBlock.IsLayout == False:
            if "*" != str(one_AcadBlock.Name[0]):
                doc_blocks_instances.append([one_AcadBlock.Name, one_AcadBlock])
            if one_AcadBlock.IsDynamicBlock:
                print("Dynamic block name: " + one_AcadBlock.Name)
    print(doc_blocks_instances)
    """
    Теперь произведем создание вхождения блока для первого Блока в списке и выведем в консоль его наименование.
    Обратим внимание, что для Вхождения блока существует также EffectiveName, что позволяет получать действительные 
    имена вхождения блоков в случае, если они стали, например, по какой-то причине анонимными (если стандартное поле 
    "Name" будет возвращать наименование через "*") или преобразовывать его в статический -- ConvertToStaticBlock()
    """
    AcadBlockReference_object = ncad_doc_ms.InsertBlock([500, 500, 0], doc_blocks_instances[0][0], 1.0, 1.0, 1.0, 0.0)
    block_red_names = {
        "Name": AcadBlockReference_object.Name,
        "EffectiveName": AcadBlockReference_object.EffectiveName
    }
    print(str(block_red_names))

    """
    При желании добавить блоку атрибут, необходимо сделать это для Родительского блока
    И после добавления атрибута обязательно обновляем Вхождение блока -- для регенрации у него этой позиции
    """
    parent_block = dict(doc_blocks_instances)[AcadBlockReference_object.EffectiveName]
    attr = parent_block.AddAttribute(5, 8, "Test_attr", [0,0,0], "Test_attr", "Hello, ActiveX!")
    AcadBlockReference_object.Update
    ncad_doc.Application.Update()

    """
    Доступ к имеющимся атрибутам блока осуществляется следующим образом: сперва проверяется, есть ли у блока атрибуты методом 
    HasAttributes() и если он возвратит True -- то начнем перебирать атрибуты методом GetAttributes(), где каждый объект
     будет иметь интерфейс OdaX.AcadAttribute.
    """
    if AcadBlockReference_object.HasAttributes:
        for one_attr in AcadBlockReference_object.GetAttributes():
            print(one_attr.TagString + " " + one_attr.TextString)

print("End!")