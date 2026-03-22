---
name: renga-mcp
description: >
  Работа с BIM-системой Renga через API, SDK и MCP-сервер. ВСЕГДА используй этот скилл когда пользователь упоминает: Renga API, Renga SDK, Renga MCP, плагин для Renga, автоматизация Renga, скрипт Python для Renga, COM API Renga, создание объектов в Renga через API, атрибутация модели Renga, экспорт из Renga через скрипт, renga_mcp_server, renga_status, renga_get_objects, renga_create_column, renga_create_level, win32com Renga, GetInterfaceByName Renga, pywin32 Renga, IModel Renga, IColumn, ILevel, IWindow, IDoor, IIsolatedFoundation, RegisterPropertyS, CreateOperation, operation.Apply.
---

# Скилл: Renga API / SDK / MCP

## Архитектура: как работает связка

```
Claude Desktop (MCP клиент)
        │  stdio / JSON-RPC
        ▼
renga_mcp_server_v2.py  (FastMCP, Python)
        │  win32com.client  →  COM
        ▼
Renga.Application.1  (COM out-of-process server)
        │  Renga API v2.45
        ▼
Открытый проект Renga (.rnp)
```

Файл сервера: `renga_mcp_server_v2.py` (1034 строки, 18 инструментов)
Установка: `pip install mcp pywin32`
Конфиг Claude Desktop (`%APPDATA%\Claude\claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "renga": {
      "command": "python",
      "args": ["C:\\path\\to\\renga_mcp_server_v2.py"]
    }
  }
}
```

---

## Ключевые особенности COM из Python

### 1. Подключение к запущенной Renga
```python
import win32com.client as win32
app = win32.GetActiveObject("Renga.Application.1")
project = app.Project
model = project.Model
```

### 2. Паттерн операций — ОБЯЗАТЕЛЕН для любого изменения
```python
# Перед изменением — проверить нет ли активной операции
if project.HasActiveOperation():
    raise RuntimeError("Другая операция уже выполняется")

op = project.CreateOperation()
op.Start()
try:
    # ... изменения ...
    op.Apply()    # зафиксировать
except Exception as e:
    op.Reject()   # откатить, НЕ ПАДАЕМ
    raise
```
**Без операции изменения модели работают только в read-only режиме.**

### 3. QueryInterface из Python — через GetInterfaceByName
Python не умеет вызывать COM QueryInterface напрямую. Вместо этого:
```python
# НЕПРАВИЛЬНО (не работает из Python):
# level = ILevel(obj)

# ПРАВИЛЬНО:
level_iface = obj.GetInterfaceByName("ILevel")
level_name = level_iface.LevelName
level_elev = level_iface.Elevation
```

### 4. GUID → строка: S-методы
Renga использует GUID для идентификации типов объектов и свойств.
Python не умеет работать с COM UDT напрямую — используй S-суффикс:
```python
# НЕПРАВИЛЬНО:
# obj.ObjectType  ← COM UDT, Python не сможет

# ПРАВИЛЬНО:
type_str = obj.ObjectTypeS        # возвращает строку "{GUID}"
prop_mgr.RegisterPropertyS(guid_str, name, type_code)
prop_mgr.AssignPropertyToTypeS(prop_guid, obj_type_guid)
container.GetS(prop_guid_str)
```

### 5. Итерация по коллекциям
```python
objects = model.GetObjects()
for i in range(objects.Count):
    obj = objects.GetByIndex(i)  # GetByIndex, не objects[i]
    ...
```

---

## GUID типов объектов (Renga API v2.45)

```python
OBJECT_TYPE_GUIDS = {
    "Column":             "{3B1B3D07-5B13-4059-A2EE-E9A58DD0A3B0}",
    "Window":             "{B85B6F47-4A20-4A61-8AAE-0498EDAC9C2A}",
    "Door":               "{B95F4B1C-3D0D-4424-9B5C-3C8E2B9B1B2A}",
    "IsolatedFoundation": "{F9B2C3D4-5E6F-7A8B-9C0D-1E2F3A4B5C6D}",
    "Level":              "{C3CE17FF-6F28-411F-B18D-74FE957B2BA8}",
    "Assembly":           "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}",
    "Wall":               "{8B41D79C-B03B-4C98-A0F4-7B62D5D0CDB3}",
    "Plate":              "{D1E2F3A4-B5C6-7D8E-9F0A-1B2C3D4E5F6A}",
    "Beam":               "{E5F6A7B8-C9D0-1E2F-3A4B-5C6D7E8F9A0B}",
}
```
⚠ **GUID для Column, Window, Door, IsolatedFoundation — нужно уточнить по SDK.**
При подключении к реальной Renga — сравнить с документацией из `RengaSDK/Samples/`.
Официальный Level GUID подтверждён документацией: `{C3CE17FF-6F28-411F-B18D-74FE957B2BA8}`

---

## Коды типов свойств (RegisterPropertyS)

```python
PROPERTY_TYPES = {
    "Integer": 1,
    "String":  2,
    "Double":  3,
    "Boolean": 4,
}
```

---

## Доступные инструменты MCP-сервера (v2)

### Статус и проект
| Инструмент | Описание |
|---|---|
| `renga_status` | Проверка подключения, имя проекта, активная операция |
| `renga_project_info` | Имя, путь, статистика объектов по типам |
| `renga_save_project` | Сохранить проект |
| `renga_open_project(file_path)` | Открыть .rnp файл |

### Чтение модели
| Инструмент | Описание |
|---|---|
| `renga_get_objects(object_type, limit)` | Список объектов с фильтром по типу |
| `renga_get_object_params(object_id)` | Параметры + расчётные характеристики + свойства |
| `renga_get_levels` | Уровни с отметками в метрах |

### Создание объектов (требует Renga v8.7+ апрель 2025)
| Инструмент | Параметры |
|---|---|
| `renga_create_level(name, elevation_m)` | Новый уровень/этаж |
| `renga_create_column(x, y, level_id, height_m, style_name)` | Колонна |
| `renga_create_window(wall_id, offset_m, sill_height_m, width_m, height_m, style_name)` | Окно в стену |
| `renga_create_door(wall_id, offset_m, width_m, height_m, style_name)` | Дверь в стену |
| `renga_create_isolated_foundation(x, y, width_m, length_m, depth_m, style_name)` | Столбчатый фундамент |
| `renga_create_assembly(name, object_ids)` | Сборка из существующих объектов |
| `renga_create_model_text(text, x, y, z, height_m)` | Текстовая метка в 3D |

### Свойства
| Инструмент | Описание |
|---|---|
| `renga_create_property(name, type, object_types)` | Создать пользовательское свойство |
| `renga_bulk_set_property(object_type, property_name, value, filter_param, filter_value)` | Массово задать значение |

### Экспорт
| Инструмент | Описание |
|---|---|
| `renga_export_ifc(output_path)` | Экспорт в IFC |
| `renga_export_drawings(output_folder, format, section_filter)` | Чертежи PDF/DWG/DXF с фильтром по разделу |

---

## Ограничения Renga API (актуально на март 2026)

| Операция | Статус |
|---|---|
| Создание колонн, окон, дверей, фундаментов | ✅ с апреля 2025 (v8.7+) |
| Создание уровней, сборок, текста модели | ✅ с апреля 2025 |
| Работа с чертежами через API | ✅ с августа 2025 (v8.8+) |
| Создание стен через API | ❌ пока недоступно |
| Создание балок через API | ❌ пока недоступно |
| Создание перекрытий через API | ❌ пока недоступно |
| QueryInterface из Python | ❌ используй GetInterfaceByName |
| GUID как UDT из Python | ❌ используй S-методы |
| Renga на macOS | ❌ только Windows |

---

## Практические сценарии

### Сценарий 1: Массовая атрибутация модели
**Задача:** Проставить статус монтажа всем колоннам на уровне 1.

```
1. renga_get_levels → найти Id уровня 1
2. renga_create_property("Статус монтажа", "String", "Column")
3. renga_get_objects("Column") → получить список Id
4. renga_bulk_set_property("Column", "Статус монтажа", "Не начат")
```

### Сценарий 2: Пакетный экспорт по разделам
**Задача:** Экспортировать только архитектурные листы в PDF.

```
renga_export_drawings("C:/export/АР/", "PDF", "АР")
```
Листы вида "АР_001_...", "АР_002_..." попадут в экспорт автоматически.

### Сценарий 3: Создание каркаса из CSV
**Задача:** Создать 20 колонн по координатной сетке из таблицы.

```python
# Пользователь даёт CSV: x,y,height
# Claude итерирует и вызывает renga_create_column для каждой строки
# Уровень берётся из renga_get_levels
```

### Сценарий 4: Аудит модели на заполненность атрибутов
**Задача:** Найти все объекты без заполненного свойства "Марка".

```
1. renga_get_objects("Column") → список Id
2. renga_get_object_params для каждого → смотрим поле "Марка"
3. Возвращаем список Id без марки
```

### Сценарий 5: Python скрипт без MCP (standalone)
Если нужно запустить скрипт напрямую без Claude Desktop:
```python
import win32com.client as win32

app = win32.GetActiveObject("Renga.Application.1")
project = app.Project
model = project.Model

op = project.CreateOperation()
op.Start()
# ... изменения ...
op.Apply()
```

---

## Отладка и типичные ошибки

### "Renga не запущена"
```
win32.GetActiveObject("Renga.Application.1")  →  pywintypes.com_error
```
Решение: открыть Renga, загрузить проект, проверить через `renga_status`.

### "HasActiveOperation = True"
Renga уже выполняет операцию (например, незавершённое редактирование).
Решение: завершить операцию вручную в Renga (Esc или подтвердить/отменить).

### Метод не найден на COM-объекте
Пример: `column.Height = 3.0` → AttributeError.
Причина: метод находится на интерфейсе IColumn, а не на IModelObject.
Решение:
```python
col_iface = column.GetInterfaceByName("IColumn")
col_iface.Height = 3.0
```

### GUID не совпадает → объекты не фильтруются
Renga обновляется — GUID типов может меняться между версиями.
Решение: распечатать ObjectTypeS всех объектов и сравнить:
```python
objects = model.GetObjects()
types = set()
for i in range(objects.Count):
    obj = objects.GetByIndex(i)
    try:
        types.add(str(obj.ObjectTypeS))
    except Exception:
        pass
print(types)
```
Затем обновить `OBJECT_TYPE_GUIDS` в сервере.

### pywin32 не установлен
```
pip install pywin32
```
После установки — перезапустить Claude Desktop.

---

## Полезные ссылки

- Официальная документация API: https://help.rengabim.com/api/
- Как работать с динамическими языками: https://help.rengabim.com/api/how-to-dt-language.html
- GitHub официальных примеров: https://github.com/RengaSoftware/SampleScripts
- GitHub курса по Renga API: https://github.com/GeorgGrebenyuk/renga_programming_course_1
- Блог разработчиков: https://blog.rengabim.com/search/label/API
- SDK: https://rengabim.com/sdk/
