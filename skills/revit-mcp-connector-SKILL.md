---
name: revit-mcp-connector
description: "Работа с Revit через MCP — настройка, подключение, написание кода и отладка. ВСЕГДА используй этот скилл когда пользователь упоминает: MCP для Revit, revit-mcp, pyRevit Routes, execute_revit_code, Revit Connector, подключение Claude к Revit, настройка MCP-сервера для Revit, LLM + Revit, автоматизация Revit через AI, C# Revit addin для MCP, revit-mcp-plugin, SocketService, ExternalEventHandler в контексте MCP. Различай два стека: Python/pyRevit (IronPython, порт 48884) и C#/.NET (addin + TypeScript/Node сервер, WebSocket)."
---

# Скилл: Работа с Revit через MCP

## Главное различие: два независимых стека

| | **Python-стек** | **C#-стек** |
|---|---|---|
| Движок внутри Revit | pyRevit Routes (IronPython) | C# Addin (.NET, ExternalEventHandler) |
| Порт внутри Revit | **48884** (HTTP REST) | **WebSocket** (настраивается) |
| MCP-сервер снаружи | Python (`uvx revit-mcp` / `uv run main.py`) | TypeScript/Node.js (`npx mcp-server-for-revit`) |
| Код от Claude | IronPython, выполняется через `execute_revit_code` | Нет — только фиксированные команды из CommandSet |
| Расширяемость | Любой код на лету | Только через компиляцию нового командсета |
| Установка | pyRevit extension + Python/uv | .addin файл + Node/npm |
| Реконнект при сбое | `taskkill python.exe` + `uv run main.py --combined` | Перезапуск плагина или Revit |

**Правило выбора:**
- Нужно писать произвольный код под задачу → **Python-стек**
- Нужны стабильные повторяемые команды в команде/библиотеке → **C#-стек**

---

## PYTHON-СТЕК (pyRevit Routes)

### Архитектура
```
Claude Desktop / Cursor
       │ MCP Protocol (stdio)
       ▼
  main.py (MCP Server, Python/uv)
       │ HTTP requests → localhost:48884
       ▼
  pyRevit Routes (REST API внутри Revit)
       │ Revit API
       ▼
  Revit Application
```

### Установка

**1. pyRevit** — установить с https://github.com/pyrevitlabs/pyRevit/releases

**2. pyRevit extension** — добавить путь к расширению в настройках pyRevit:
```
pyRevit Settings → Custom Extension Directories → добавить папку расширения
```
Перезагрузить pyRevit. Проверить: расширение загружено, Routes API активно.

**3. MCP-сервер** — варианты установки:

**Вариант A — uvx (самый простой):**
```json
{
  "mcpServers": {
    "Revit Connector": {
      "command": "uvx",
      "args": ["revit-mcp"],
      "timeout": 30
    }
  }
}
```

**Вариант B — uv run из локальной папки:**
```json
{
  "mcpServers": {
    "Revit Connector": {
      "command": "uv",
      "args": ["run", "--with", "mcp[cli]", "mcp", "run", "C:/path/to/main.py"]
    }
  }
}
```

**Вариант C — python напрямую (RevitMCP / oakplank):**
```json
{
  "mcpServers": {
    "revitmcp": {
      "command": "python",
      "args": ["C:/path/to/server.py"]
    }
  }
}
```

**Конфиг-файл** (Windows): `%APPDATA%\Claude\claude_desktop_config.json`

### Проверка подключения

```python
# Инструмент: Revit Connector:get_revit_status
# Ожидаемый ответ:
{
  "status": "active",
  "health": "healthy",
  "revit_available": true,
  "document_title": "имя_файла.rvt",
  "api_name": "revit_mcp"
}
```

Если нет ответа — см. раздел «Реконнект».

### Основные инструменты Python-стека

| Инструмент | Назначение |
|-----------|-----------|
| `Revit Connector:get_revit_status` | Проверка живости коннектора |
| `Revit Connector:get_revit_model_info` | Инфо о документе (уровни, виды, листы, элементы) |
| `Revit Connector:execute_revit_code` | Выполнение произвольного IronPython кода в Revit |

### Правила написания кода для execute_revit_code

**Обязательные импорты:**
```python
from pyrevit import revit, DB
doc = revit.doc
```

**Константа единиц:**
```python
ft = 0.3048  # Revit API работает в футах, всё остальное в метрах
# Перевод: метры → футы: x / 0.3048
# Перевод: футы → метры: x * 0.3048
```

**Транзакции (любое изменение модели):**
```python
with revit.Transaction("Имя транзакции"):
    # код изменения
```
Или классический вариант:
```python
t = DB.Transaction(doc, "Имя")
t.Start()
# код
t.Commit()
```

**⚠️ Кодировка кириллицы (IronPython):**
```python
# В print() — обязательно .encode('utf-8'):
print("Категория:", fam_cat.Name.encode('utf-8'))

# В return и в строках без print — кириллица работает напрямую:
return {"name": fam_cat.Name}  # OK

# НИКОГДА не делать:
print(u"кириллица")  # → UnicodeEncodeError в MCP
```

**Структура типичного скрипта:**
```python
from pyrevit import revit, DB
doc = revit.doc
ft = 0.3048

try:
    # чтение данных
    elements = list(DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
        .ToElements())
    print("Walls:", len(elements))
    
    # изменение (если нужно)
    with revit.Transaction("My Change"):
        pass  # код изменения

except Exception as e:
    print("ERROR:", str(e)[:200])
```

### Реконнект при потере связи

**Симптом:** `No result received` / таймаут / `revit_available: false`

**Последовательность восстановления:**
```
1. taskkill /f /im python.exe          (Windows, убить зависший MCP-процесс)
2. В папке расширения:
   uv run main.py --combined           (или через UI кнопку Launch в Revit)
3. Перезапустить Claude Desktop
4. Проверить: get_revit_status
```

Если Routes не стартует после перезапуска Revit — убедиться что extension path добавлен в pyRevit Settings.

---

## C#-СТЕК (Revit Addin + TypeScript Server)

### Архитектура
```
Claude Desktop / Cursor
       │ MCP Protocol (stdio)
       ▼
  MCP Server (TypeScript/Node.js)
       │ WebSocket
       ▼
  Revit Plugin (C# .dll addin)
  SocketService + CommandExecutor
  ExternalEventHandler
       │ Revit API
       ▼
  Revit Application
```

### Три компонента

**1. revit-mcp-plugin** — C# addin, устанавливается в Revit:
```
%AppData%\Autodesk\Revit\Addins\<версия>\
├── revit-mcp-plugin.addin
└── revit_mcp_plugin/
    ├── revit-mcp-plugin.dll
    └── Commands/RevitMCPCommandSet/
        ├── command.json
        └── <версия>/RevitMCPCommandSet.dll
```

**2. revit-mcp** — TypeScript MCP-сервер (сторона Claude):
```json
{
  "mcpServers": {
    "mcp-server-for-revit": {
      "command": "cmd",
      "args": ["/c", "npx", "-y", "mcp-server-for-revit"]
    }
  }
}
```
Или через node напрямую:
```json
{
  "mcpServers": {
    "revit-mcp": {
      "command": "node",
      "args": ["C:/path/to/build/index.js"]
    }
  }
}
```

**3. revit-mcp-commandset** — CommandSet: `.dll` с реализацией команд на C#.

### Активация в Revit

После установки:
```
Add-ins → Revit MCP Plugin → Settings → отметить нужные команды
Add-ins → Revit MCP Plugin → Revit MCP Switch → включить
```

### Доступные инструменты C#-стека (фиксированные)

Инструменты определены в `command.json` CommandSet. Типичный набор:

| Команда | Что делает |
|---------|-----------|
| `get_project_info` | Информация о проекте |
| `get_elements_in_view` | Элементы текущего вида |
| `get_family_types` | Доступные типы семейств |
| `get_selected_elements` | Выбранные элементы |
| `get_current_view` | Информация о текущем виде |
| `read_element_parameters` | Параметры элемента по Id |
| `update_element_parameter` | Изменить параметр элемента |
| `place_view_on_sheet` | Разместить вид на листе |
| `execute_code` | Отправить код на выполнение (не всегда надёжно) |

### Добавление своей команды (C#-стек)

1. В `command.json` добавить запись:
```json
{
  "commandName": "my_command",
  "description": "Что делает команда",
  "assemblyPath": "RevitMCPCommandSet.dll"
}
```

2. Реализовать `IExternalEventHandler` в C#:
```csharp
public class MyCommandHandler : IExternalEventHandler
{
    public void Execute(UIApplication app)
    {
        var doc = app.ActiveUIDocument.Document;
        using (var t = new Transaction(doc, "My Command"))
        {
            t.Start();
            // логика
            t.Commit();
        }
    }
    public string GetName() => "MyCommand";
}
```

3. Пересобрать solution → перезапустить Revit.

---

## ГИБРИДНЫЙ СТЕК (Python MCP Server + C# Bridge)

Существует третий вариант (Sam-AEC / Autodesk-Revit-MCP-Server):
```
Claude → Python MCP Server → HTTP :3000 → C# Bridge Addin → Revit API
```
- MCP-сервер на Python (инструменты описаны в Python)
- Addin на C# работает как HTTP-сервер внутри Revit
- Позволяет комбинировать удобство Python-описания инструментов с надёжностью C# в Revit

Конфиг:
```json
{
  "mcpServers": {
    "revit": {
      "command": "python",
      "args": ["-m", "revit_mcp_server.mcp_server"],
      "env": {
        "MCP_REVIT_BRIDGE_URL": "http://127.0.0.1:3000",
        "MCP_REVIT_MODE": "bridge"
      }
    }
  }
}
```

---

## ДИАГНОСТИКА И ОШИБКИ

| Симптом | Стек | Причина | Решение |
|---------|------|---------|---------|
| `No result received` / таймаут | Python | pyRevit Routes завис | `taskkill /f /im python.exe` + рестарт MCP |
| `revit_available: false` | Python | Revit закрыт или Routes не загружен | Открыть Revit, проверить extension |
| `UnicodeEncodeError` | Python | Кириллица в `print()` без encode | `.encode('utf-8')` для строк в print() |
| `Plugin process exited with code 1` | Оба | Неправильная команда запуска | Проверить пути и команды в конфиге |
| `hammer icon` не появляется | Оба | MCP-сервер не стартовал | Проверить конфиг, пути, Node/Python |
| Команды не видны | C# | CommandSet не отмечен | Settings → отметить команды → рестарт Revit |
| WebSocket disconnected | C# | Plugin завис | Переключить MCP Switch Off/On в Revit |
| `LM Studio` + MCP конфиг | — | LM Studio не поддерживает MCP | Использовать Claude Desktop или Cursor |
| `ft/м путаница` | Python | Неверная единица в координатах | `m(x) = x/0.3048` → Revit; `p.X*0.3048` → метры |

---

## БЫСТРЫЙ СТАРТ (Python-стек, самый частый случай)

**Минимальный рабочий конфиг** (`%APPDATA%\Claude\claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "Revit Connector": {
      "command": "uvx",
      "args": ["revit-mcp"],
      "timeout": 30
    },
    "mcp-filesystem-server": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-filesystem", "C:/Users/Имя/Desktop"]
    }
  }
}
```

**Проверка после запуска:**
1. Открыть Revit с проектом
2. Убедиться, что pyRevit загружен (вкладка pyRevit в ленте)
3. Открыть Claude Desktop → должен появиться 🔨 (hammer icon)
4. Написать: «Проверь статус Revit» → Claude вызовет `get_revit_status`

---

## Справочные файлы

- `references/python-snippets.md` — Готовые сниппеты IronPython для частых задач
- `references/csharp-patterns.md` — Паттерны C# для ExternalEventHandler и CommandSet
