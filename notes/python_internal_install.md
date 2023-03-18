# Указания к развёртыванию окружения Python на ПК для работы с nanoCAD 23.0+

Настоящая статья является расширением [базовой инструкции по установке](https://habr.com/ru/company/nanosoft/blog/718362/) и предназначена для пользователей, у кого стоит несколько Python-интерпретаторов в системе.

Указать, какой именно интерпретатор будет использоваться, можно только управлением иерархией других путей в числе системной (не пользовательской) переменной окружения PATH, где нужный интепретатор указывается как файловый пути к нему (файлу `python.exe`). 

В таком случае, дальнейшие действия будут следующими:

1. Запуск `cmd.exe` от Администратора
2. Перемещение среды исполнения в целевую папку установки python, в моём случае - следующий путь: `cd /d C:\Users\Georg\AppData\Local\Programs\Python\Python310`
3. Исполнение команд:

```
python -m pip install --upgrade pywin32
python -m pip install pywin32 --upgrade
python Scripts/pywin32_postinstall.py -install
```
И тогда консоль "здорового человека" будет иметь вид (с поправкой на другие файлоавые пути у Вас):
```
Microsoft Windows [Version 10.0.19044.1889]
(c) Корпорация Майкрософт (Microsoft Corporation). Все права защищены.

C:\Windows\system32>cd /d C:\Users\Georg\AppData\Local\Programs\Python\Python310

C:\Users\Georg\AppData\Local\Programs\Python\Python310>python -m pip install --upgrade pywin32
Requirement already satisfied: pywin32 in c:\users\georg\appdata\local\programs\python\python310\lib\site-packages (305)

C:\Users\Georg\AppData\Local\Programs\Python\Python310>python -m pip install pywin32 --upgrade
Requirement already satisfied: pywin32 in c:\users\georg\appdata\local\programs\python\python310\lib\site-packages (305)

C:\Users\Georg\AppData\Local\Programs\Python\Python310>python Scripts/pywin32_postinstall.py -install
Parsed arguments are: Namespace(install=True, remove=False, wait=None, silent=False, quiet=False, destination='C:\\Users\\Georg\\AppData\\Local\\Programs\\Python\\Python310\\Lib\\site-packages')
Copied pythoncom310.dll to C:\Windows\system32\pythoncom310.dll
Copied pywintypes310.dll to C:\Windows\system32\pywintypes310.dll
Registered: Python.Interpreter
Registered: Python.Dictionary
Registered: Python
-> Software\Python\PythonCore\3.10\Help[None]=None
-> Software\Python\PythonCore\3.10\Help\Pythonwin Reference[None]='C:\\Users\\Georg\\AppData\\Local\\Programs\\Python\\Python310\\Lib\\site-packages\\PyWin32.chm'
Registered help file
Pythonwin has been registered in context menu
Shortcut for Pythonwin created
Shortcut to documentation created
The pywin32 extensions were successfully installed.

C:\Users\Georg\AppData\Local\Programs\Python\Python310>
```
А при запуске без прав будет в логе что-то типа `The file 'C:\Windows\system32\pythoncom310.dll' exists, but can not be replaced due to insufficient permissions. You must reinstall this software as an Administrator`
