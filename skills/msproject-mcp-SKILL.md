---
name: msproject-mcp
description: >
  Работа с MS Project через MCP-сервер, COM API и Aspose.Tasks. ВСЕГДА используй этот скилл когда
  пользователь упоминает: MS Project MCP, Microsoft Project API, автоматизация MS Project, Python MSProject,
  msproject_mcp_server, msproject_status, msproject_get_tasks, msproject_add_task, msproject_update_task,
  msproject_get_resources, msproject_assign_resource, msproject_set_baseline, msproject_get_earned_value,
  msproject_export_pdf, msproject_export_excel, win32com MSProject, MSProject.Application, .mpp файл Python,
  базовый план MSP, EVM MS Project, критический путь MSP, перегрузка ресурсов MSP, экспорт Ганта, EVM SPI CPI.
---

# Скилл: MS Project MCP

## Архитектура

```
Claude Desktop (MCP клиент)
        │  stdio / JSON-RPC
        ▼
msproject_mcp_server.py  (FastMCP, Python)
        │
        ├── Режим "com"  →  win32com.client  →  COM  →  MSProject.Application  →  .mpp
        │
        └── Режим "file" →  aspose.tasks.Project  →  .mpp (без запущенного MSP)
```

**Файл сервера:** `msproject_mcp_server.py` (~38 инструментов)

**Установка зависимостей:**
```bash
pip install mcp pywin32 openpyxl          # минимум (COM-режим)
pip install aspose-tasks                   # для file-режима (без MSP)
```

**Конфиг Claude Desktop** (`%APPDATA%\Claude\claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "msproject": {
      "command": "python",
      "args": ["C:\\path\\to\\msproject_mcp_server.py"],
      "env": { "MSP_MODE": "com" }
    }
  }
}
```

---

## Два режима работы

| Параметр        | Режим `com`                            | Режим `file`                          |
|----------------|----------------------------------------|---------------------------------------|
| Требования      | Установлен MS Project Professional    | Только `pip install aspose-tasks`     |
| Как подключить  | `msproject_open(path, mode="com")`     | `msproject_open(path, mode="file")`   |
| Запись          | ✅ Полная (задачи, ресурсы, базовый)   | ⚠ Ограниченная                       |
| Скорость записи | Медленно при >5000 задач               | Быстро                                |
| PDF-экспорт     | ✅ Через MS Project                    | ✅ Через Aspose                       |
| EVM / Базовый   | ✅ Полноценно                          | ❌ Только через XML                   |
| Рекомендуется   | Для управления проектом                | Для чтения/анализа .mpp файлов        |

**Переключение режимов в диалоге:**
```
msproject_open("C:\\stroy.mpp", mode="com")   # открыть через MS Project
msproject_open("C:\\stroy.mpp", mode="file")  # открыть напрямую (без MSP)
```

---

## COM-объектная модель MS Project (Python)

### Подключение
```python
import win32com.client as win32

# К запущенному MSP:
app = win32.GetActiveObject("MSProject.Application")

# Запустить новый экземпляр:
app = win32.Dispatch("MSProject.Application")
app.Visible = True

proj = app.ActiveProject
```

### Итерация — через индекс от 1, не через []
```python
# ПРАВИЛЬНО:
for i in range(1, proj.Tasks.Count + 1):
    t = proj.Tasks(i)      # (i) — вызов как метод, не [i]
    if t is None: continue  # удалённые задачи возвращают None

# НЕПРАВИЛЬНО:
for t in proj.Tasks: ...   # не работает стабильно
```

### Длительность — строковый формат
```python
t.Duration = "5d"     # 5 рабочих дней
t.Duration = "2w"     # 2 недели
t.Duration = "8h"     # 8 часов
t.Duration = "0d"     # веха (milestone)
```

### Работа (Work) — в минутах
```python
hours = float(t.Work) / 60.0     # перевод минут → часы
t.ActualWork = hours * 60.0       # запись часов → минуты
```

### Затраты (Cost) — float, рубли/валюта проекта
```python
cost = float(t.Cost)
```

### Тип связи TaskDependencies
```python
# FS=1, FF=2, SS=3, SF=4
t.TaskDependencies.Add(predecessor, 1, "0")  # FS без лага
t.TaskDependencies.Add(predecessor, 1, "2d") # FS + 2 дня лага
```

### Ограничения (Constraint)
```python
# 0=ASAP, 1=ALAP, 4=SNET, 5=FNLT, 6=SNLT, 7=FNET
t.ConstraintType = 4              # Start No Earlier Than
t.ConstraintDate = "2026-06-01"
```

### Типы ресурсов
```python
# 0=Work (трудовой), 1=Material (материал), 2=Cost (затратный)
r.Type = 0
r.MaxUnits = 1.0   # 100% загрузки
```

---

## Все инструменты MCP-сервера

### Статус и проект
| Инструмент | Параметры | Описание |
|---|---|---|
| `msproject_status` | — | Режим, имя проекта, кол-во задач/ресурсов |
| `msproject_open` | `file_path`, `mode` | Открыть .mpp файл |
| `msproject_save` | `file_path?` | Сохранить (или SaveAs) |
| `msproject_project_info` | — | Даты, бюджет, % выполнения, статистика |

### Задачи — чтение
| Инструмент | Параметры | Описание |
|---|---|---|
| `msproject_get_tasks` | `filter_by`, `filter_value`, `limit`, `include_summary` | Список задач с фильтрами |
| `msproject_get_task` | `task_id` | Одна задача + назначения |
| `msproject_get_critical_path` | — | Задачи критического пути |
| `msproject_get_task_tree` | `parent_id`, `depth` | WBS-дерево |
| `msproject_get_milestones` | — | Все вехи |
| `msproject_find_tasks` | `query` | Поиск по названию |

### Задачи — запись
| Инструмент | Параметры | Описание |
|---|---|---|
| `msproject_add_task` | `name`, `duration_days`, `start`, `parent_task_id`, `predecessor_ids`, `notes`, `is_milestone` | Добавить задачу |
| `msproject_update_task` | `task_id`, `name?`, `duration_days?`, `start?`, `finish?`, `notes?`, `predecessor_ids?`, `constraint_type?`, `constraint_date?` | Обновить задачу |
| `msproject_delete_task` | `task_id` | Удалить задачу |
| `msproject_link_tasks` | `predecessor_id`, `successor_id`, `link_type`, `lag_days` | Создать связь |
| `msproject_set_task_percent` | `task_id`, `percent` | Установить % выполнения |
| `msproject_bulk_update_tasks` | `updates` (JSON-строка) | Массовое обновление |

### Ресурсы
| Инструмент | Параметры | Описание |
|---|---|---|
| `msproject_get_resources` | `resource_type` | Список ресурсов |
| `msproject_get_resource` | `resource_id` | Ресурс + его задачи |
| `msproject_add_resource` | `name`, `resource_type`, `max_units_pct`, `standard_rate`, `email` | Добавить ресурс |
| `msproject_update_resource` | `resource_id`, `name?`, `max_units_pct?`, `standard_rate?` | Обновить ресурс |
| `msproject_delete_resource` | `resource_id` | Удалить ресурс |
| `msproject_get_overallocated` | — | Перегруженные ресурсы |

### Назначения
| Инструмент | Параметры | Описание |
|---|---|---|
| `msproject_get_assignments` | `task_id?`, `resource_id?` | Назначения с фильтром |
| `msproject_assign_resource` | `task_id`, `resource_id`, `units_pct` | Назначить ресурс |
| `msproject_remove_assignment` | `task_id`, `resource_id` | Снять назначение |
| `msproject_get_resource_workload` | `resource_id` | Нагрузка ресурса по задачам |

### Базовый план и статус
| Инструмент | Параметры | Описание |
|---|---|---|
| `msproject_set_baseline` | `baseline_number` (0..10) | Сохранить базовый план |
| `msproject_clear_baseline` | `baseline_number` | Очистить базовый план |
| `msproject_get_baseline_comparison` | `limit` | Отклонения факт/план |
| `msproject_update_progress` | `status_date`, `task_updates` (JSON) | Ввод факта выполнения |
| `msproject_get_earned_value` | — | EVM: BCWS, BCWP, ACWP, SPI, CPI |
| `msproject_get_late_tasks` | `days_threshold` | Просроченные задачи |
| `msproject_get_summary` | — | Сводка для статус-репорта |

### Экспорт
| Инструмент | Параметры | Описание |
|---|---|---|
| `msproject_export_xml` | `output_path` | Экспорт в MS Project XML |
| `msproject_export_csv` | `output_path`, `fields?` | Задачи в CSV |
| `msproject_export_pdf` | `output_path`, `view?` | PDF (Ганта или другой вид) |
| `msproject_export_excel` | `output_path` | Excel: 3 листа (Tasks, Resources, Assignments) |
| `msproject_export_html` | `output_path` | HTML-отчёт с прогресс-барами |

---

## Практические сценарии

### Сценарий 1: Еженедельный статус-репорт
```
1. msproject_open("C:\\project.mpp", mode="com")
2. msproject_get_summary()                          → сводка
3. msproject_get_late_tasks()                       → просроченные задачи
4. msproject_get_overallocated()                    → перегруженные ресурсы
5. msproject_export_excel("C:\\reports\\week_12.xlsx")
```

### Сценарий 2: Ввод факта выполнения
```
1. msproject_open("C:\\project.mpp")
2. msproject_get_tasks("percent_complete", "0")    → незавершённые задачи
3. msproject_update_progress("2026-03-22",
     '[{"task_id": 5, "percent_complete": 75},
       {"task_id": 7, "percent_complete": 100}]')
4. msproject_save()
```

### Сценарий 3: Анализ EVM для заказчика
```
1. msproject_open("C:\\project.mpp")
2. msproject_get_earned_value()
   → SPI < 1.0 = опережение графика
   → CPI < 1.0 = превышение бюджета
3. msproject_get_baseline_comparison()             → таблица отклонений
4. msproject_export_pdf("C:\\reports\\evm.pdf")
```

### Сценарий 4: Создание структуры проекта из списка работ
```
1. msproject_open("C:\\new_project.mpp")
2. msproject_add_task("1. Подготовительный период", is_milestone=False, duration_days=5)
3. msproject_add_task("1.1 Оформление разрешений", duration_days=3, predecessor_ids="1")
4. msproject_add_task("1.2 Завоз оборудования", duration_days=2, predecessor_ids="2")
...
5. msproject_set_baseline()                        → сохранить исходный план
6. msproject_save()
```

### Сценарий 5: Анализ файла без установленного MS Project
```
1. msproject_open("C:\\archive\\project_2025.mpp", mode="file")
2. msproject_export_excel("C:\\out\\analysis.xlsx")
3. msproject_export_html("C:\\out\\gantt.html")
```

### Сценарий 6: Управление ресурсами
```
1. msproject_get_overallocated()                   → кто перегружен
2. msproject_get_resource_workload(resource_id=3)  → нагрузка по задачам
3. msproject_remove_assignment(task_id=15, resource_id=3)
4. msproject_assign_resource(15, 5, 50)            → другой ресурс, 50%
```

---

## Формат `msproject_bulk_update_tasks`
```json
[
  {"task_id": 5, "percent_complete": 100},
  {"task_id": 7, "duration_days": 3, "start": "2026-04-01"},
  {"task_id": 10, "notes": "Перенесено из-за погоды"},
  {"task_id": 12, "name": "Бетонирование фундамента (ред.)"}
]
```

## Формат `msproject_update_progress`
```json
[
  {"task_id": 5, "percent_complete": 75, "actual_start": "2026-03-01"},
  {"task_id": 7, "percent_complete": 100, "actual_finish": "2026-03-18"},
  {"task_id": 9, "actual_work_hours": 24}
]
```

---

## Типовые ошибки и решения

### "MS Project не запущен"
```
win32.GetActiveObject("MSProject.Application") → pywintypes.com_error
```
Решение: открыть MS Project, загрузить проект, повторить `msproject_status`.

### `Tasks(i)` возвращает `None`
Норма: удалённые задачи дают `None`. Всегда проверяй `if t is None: continue`.

### `t.Duration` — странные значения
MS Project хранит длительность в минутах с суффиксом типа ("4800 min+"). Используй `_duration_days()` для преобразования. При записи — строковый формат: `t.Duration = "5d"`.

### `t.Work` / `t.Cost` — None или COM-ошибка
Некоторые задачи (суммарные, вехи) могут не иметь трудозатрат. Оборачивай в `try/except`.

### PDF-экспорт падает
Причина: MS Project не показывает все задачи в текущем представлении.
Решение: перед экспортом применить вид `app.ViewApply("Gantt Chart")`, убрать фильтры.

### `aspose-tasks` требует лицензии
Без лицензии работает с ограничением: первые 25 задач / водяной знак в экспорте.
Для снятия ограничения: `aspose.tasks.License().SetLicense("license.lic")`.

### openpyxl не установлен
```
pip install openpyxl
```
Нужен для `msproject_export_excel` в COM-режиме.

---

## Полезные ссылки

- VBA-справочник MS Project (объектная модель): https://learn.microsoft.com/en-us/office/vba/api/overview/project
- Константы типов связей, ограничений, видов: https://learn.microsoft.com/en-us/office/vba/api/project.pjconstrainttype
- Aspose.Tasks for Python: https://docs.aspose.com/tasks/python-net/
- CData MCP Server for MS Project: https://github.com/CDataSoftware/microsoft-project-mcp-server-by-cdata
- COM-автоматизация MS Project примеры: https://win32com.goermezer.de/microsoft/ms-office/automating-microsoft-project.html
