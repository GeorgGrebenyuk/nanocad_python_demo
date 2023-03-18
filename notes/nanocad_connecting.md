# Отличия при связывании Python и nanoCAD (внешнее и внутреннее использование)

## Логика связывания скрипта с nanoCAD

Если работа с nanoCAD осуществляется через внутреннюю среду (загрузкой в nanoCAD через команды `PY` или `-PY`), то текущий документ получается инструкцией `ThisDrawing`, а текущее приложение (интерфейс `nanoCAD.Application`) -- запросом поля `ThisDrawing.Application`. 

Если же работа с nanoCAD осуществляется из-под стороннего запушенного приложения, то см. следующие положения. В силу того, что nanoCAD использует логику связывания через ActiveX (COM), при которой в текущей системе регистрируются так называемые COM Server'ы, каждый из которых имеет одинаковый базовый строковый идентификатор `nanoCAD.Application` различие между разными версиями осуществляется путем конкретизации, какое именно приложение нужно, например при установленном на ПК nanoCAD 22 и 23 стандартное название (ProgID) `nanoCAD.Application`  будет применено к запуску 23 версии (он был установлен позднее). При желании запускать именно 22 версию можно воспользоваться строкой `nanoCADx64.Application.22.0`.

Тогда запуск нужной версии nanoCAD будет выглядеть следующим образом:
```python
import win32com.client
nanocad_app = win32com.client.Dispatch("nanoCADx64.Application.22.0")
```

## О выводе сообщений

Если предполагается внутреннее использование скрипта, то для вывода текстовых сообщений вместо стандартной командлеты `print()` необходимо использовать метод nanoCAD ActiveX API: `document.Utility.Prompt()`