"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Урок 1: Работа с nanoCAD.Application
"""
import win32com.client

"""
Получение запущенной сессии nanoCAD (23), либо запуск этой сессии.

Работа с интерфейсом приложения сводится, в основном, к получению текущего -- ActiveDocument 
или открытию нужного документа Open(). Также через интерфейс можно загружать модули (LoadArx), макросы (RunMacro), 
управлять размерами окна nanoCAD (Width, Height).
"""
nanocad_app = win32com.client.Dispatch("nanoCAD.Application")
if nanocad_app is not None:
    """
    При работе через COM особо следует отметить возможность получения метаданных приложения: 
    Version (версия программы, например 22 или 23), 
    Caption - заголовок приложения (например, с какой базовой конфигурацией он запущен), 
    LocaleId -- номер локализации системы, в котр=орой запущена программа - например, для авто-конвертации единиц. 
        Подробнее см. https://limagito.com/list-of-locale-id-lcid-values/
        
    На основе этих данных можно из своего внешнего приложения получать список всех экземпляров nanoCAD и 
    давать пользователю право выбрать нужный
    """
    app_params = {
        "LocaleId": nanocad_app.LocaleId,
        "Version": nanocad_app.Version,
        "Caption":  nanocad_app.Caption
    }
    print(str(app_params))
    ncad_doc = nanocad_app.ActiveDocument
    if ncad_doc is not None:
        print("Документ обнаружен. Имя = " + ncad_doc.Name)
        #Дальнейшая работа с активным документом
    else:
        print("Отсутствуют активные документы")
else:
    print("Не найден nanoCAD для запуска")