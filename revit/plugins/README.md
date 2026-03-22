# Revit Plugins (C# Addins)

Плагины для Revit на C# — расчёты ОВ и ЭОМ.

---

## HvacCalc — расчёт теплопотерь (ОВ)

**Файл:** `HvacCalc_plugin.zip`

Revit addin для расчёта теплопотерь помещений и подбора оборудования ОВ.

**Архитектура:**
```
HvacCalc.Addin/          ← Revit addin (.addin + команды + UI)
HvacCalc.Core/           ← Бизнес-логика
  Models/                ← RoomData, HvacSettings, EquipmentItem, HeatLossResult
  Services/              ← HeatLossCalculator, EquipmentSelector, ReportGenerator
HvacCalc.RevitBridge/    ← Чтение/запись данных Revit API
  RevitDataProvider.cs
  RevitWriter.cs
```

**Функции:**
- Чтение помещений из модели Revit (площадь, объём, ограждения)
- Расчёт теплопотерь по СП 50.13330
- Подбор оборудования отопления
- Генерация отчёта в Excel

---

## HvacCalc EOM Module — расчёт нагрузок (ЭОМ)

**Файл:** `HvacCalc_EOM_module.zip`

Модуль расширения HvacCalc для раздела ЭОМ. Устанавливается поверх основного плагина.

**Состав:**
```
HvacCalc.Core/Models/
  EomModels.cs           ← Модели данных ЭОМ
  EomSettings.cs         ← Настройки расчёта
HvacCalc.Core/Services/
  TrnCalculator.cs       ← Расчёт ТРН нагрузок
  TrnExcelExporter.cs    ← Экспорт в Excel
HvacCalc.RevitBridge/
  RevitEomProvider.cs    ← Чтение электрооборудования из Revit
HvacCalc.Addin/
  Commands/EomCommands.cs
  UI/EomSettingsWindow.cs
  App.cs                 ← Обновлённый App с регистрацией команд ЭОМ
```

**Функции:**
- Чтение электрооборудования из модели Revit
- Расчёт трансформаторно-реакторных нагрузок (ТРН)
- Экспорт нагрузок в Excel

---

## LoadCalcPlugin — расчёт нагрузок (ТРН)

**Файл:** `LoadCalcPlugin_TRN.zip`

Самостоятельный плагин расчёта электрических нагрузок с диалоговым окном.

**Архитектура:**
```
LoadCalcPlugin/
  Commands/CmdLoadCalc.cs    ← Команда запуска
  Core/
    RevitDataReader.cs        ← Чтение данных из Revit
    LoadCalculator.cs         ← Расчёт нагрузок
  Excel/ExcelExporter.cs     ← Экспорт результатов
  UI/LoadCalcDialog.cs       ← Диалог настройки
  LoadCalcPlugin.csproj
```

---

## Установка плагинов

1. Собрать solution в Visual Studio (Release)
2. Скопировать `.dll` и `.addin` в:
   ```
   %AppData%\Autodesk\Revit\Addins\<версия Revit>\
   ```
3. Перезапустить Revit
4. Плагины появятся во вкладке **Add-ins**

## Требования

- Visual Studio 2022
- Revit API SDK (версия соответствует вашему Revit)
- .NET Framework 4.8
- EPPlus или ClosedXML (для Excel-экспорта)
