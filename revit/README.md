# 🏢 Revit MCP

Интеграция Revit с Claude/Qwen через MCP. Два независимых стека на выбор.

---

## Два стека

| | Python-стек | C#-стек |
|---|---|---|
| Движок внутри Revit | pyRevit Routes (IronPython) | C# Addin + ExternalEventHandler |
| MCP-сервер | `uvx revit-mcp` | `npx mcp-server-for-revit` |
| Произвольный код | ✅ Любой IronPython на лету | ❌ Только фиксированные команды |
| Установка | pyRevit + uv | .addin + Node.js |
| Рекомендуется | Для разработки и автоматизации | Для стабильных повторяемых команд |

---

## Python-стек (рекомендуется)

### Архитектура
```
Claude Desktop / Qwen Desktop
       │ MCP (stdio)
       ▼
  uvx revit-mcp  (Python MCP Server)
       │ HTTP → localhost:48884
       ▼
  pyRevit Routes (REST API внутри Revit)
       │ Revit API
       ▼
  Revit + открытый .rvt проект
```

### Установка

**1. pyRevit** → https://github.com/pyrevitlabs/pyRevit/releases

**2. Конфиг** (`%APPDATA%\Claude\claude_desktop_config.json`):
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

**3. Проверка:**
```
Открыть Revit → pyRevit загружен → Claude Desktop → 🔨 → «Проверь статус Revit»
```

### Основные инструменты

| Инструмент | Описание |
|---|---|
| `get_revit_status` | Проверка подключения |
| `get_revit_model_info` | Уровни, виды, листы, элементы |
| `execute_revit_code` | Выполнить IronPython код в Revit |

### Шаблон кода для execute_revit_code

```python
from pyrevit import revit, DB
doc = revit.doc
ft = 0.3048  # Revit API в футах

try:
    walls = list(DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType().ToElements())
    print("Walls:", len(walls))

    with revit.Transaction("My change"):
        pass  # изменения здесь

except Exception as e:
    print("ERROR:", str(e)[:200])
```

⚠️ **Кириллица в print() — обязательно `.encode('utf-8')`**

---

## C#-стек

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

Addin устанавливается в `%AppData%\Autodesk\Revit\Addins\<версия>\`

---

## Сценарии использования

### Аудит семейств по BIM-стандарту 2.0
```
→ Открыть .rfa в Revit
→ execute_revit_code → скрипт аудита (SKILL: revit-family-creator)
→ Отчёт: ADSK_-параметры, LOD, коннекторы, импорт
```

### Моделирование из DWG-подложки
```
→ Открыть .rvt с DWG-подложкой
→ execute_revit_code → анализ слоёв → оси → стены → помещения
→ (SKILL: revit-dwg-modeler)
```

### Визуализация через nano-banana
```
→ execute_revit_code → ImageExportOptions → PNG
→ filesystem MCP → прочитать PNG
→ nano-banana MCP → фотореалистичный рендер
```

---

## Диагностика

| Симптом | Причина | Решение |
|---------|---------|---------|
| `revit_available: false` | Revit закрыт / pyRevit не загружен | Открыть Revit, проверить вкладку pyRevit |
| `No result received` | Routes завис | `taskkill /f /im python.exe` → перезапустить Claude |
| `UnicodeEncodeError` | Кириллица в `print()` | `.encode('utf-8')` |
| 🔨 не появляется | MCP не стартовал | Проверить конфиг, установлен ли `uv` |

---

## Ссылки

- [pyRevit releases](https://github.com/pyrevitlabs/pyRevit/releases)
- [revit-mcp на PyPI](https://pypi.org/project/revit-mcp/)
- [mcp-server-for-revit](https://www.npmjs.com/package/mcp-server-for-revit)
- [Revit API документация](https://www.revitapidocs.com/)
