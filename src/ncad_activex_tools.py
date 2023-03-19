"""
Демонстрационные материалы работы в Python через технологию COM (ActiveX® Automation) с nanoCAD
Опубликованы на https://github.com/GeorgGrebenyuk/nanocad_python_demo
Общие инструменты для работы
"""
import win32com.client
import os

ncad_app_name = "nanoCAD.Application"
class ncad_tools:
    @staticmethod
    #Получает активный документ с проверкой на ошибки
    def get_active_document():
        nanocad_app = win32com.client.Dispatch(ncad_app_name)
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
    @staticmethod
    #Возвращает путь до чертежа из папки Samples
    def get_sample_drawing(name):
        nanocad_app = win32com.client.Dispatch(ncad_app_name)
        if nanocad_app is not None:
            """
            Получаем файловый путь до вспомогательных данных nanoCAD в папке AppData/Roaming для данного Пользователя
            """
            ncad_appdata_path = nanocad_app.CurUserAppData
            ncad_file_path = os.path.join(ncad_appdata_path, "Samples", name)
            if os.path.exists(ncad_file_path):
                return ncad_file_path
            else:
                print("Указанный путь не существует " + ncad_file_path)
                return None
        else:
            return None