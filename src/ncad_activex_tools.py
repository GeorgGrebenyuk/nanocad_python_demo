"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Общие инструменты для работы
"""
import win32com.client

class ncad_tools:
    @staticmethod
    #Получает активный документ с проверкой на ошибки
    def get_active_document():
        nanocad_app = win32com.client.Dispatch("nanoCAD.Application")
        if nanocad_app is not None:
            ncad_doc = nanocad_app.ActiveDocument
            if ncad_doc is not None:
                return ncad_doc
            else:
                print("Отсутствуют активные документы")
                return None
        else:
            print("Не найден nanoCAD для запуска")
            return None
