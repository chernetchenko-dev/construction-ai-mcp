# 🏗️ construction-ai-mcp

MCP-серверы и скиллы для автоматизации строительного проектирования и управления проектами через Claude / Qwen Desktop.

> **MCP (Model Context Protocol)** — стандарт интеграции LLM с внешними инструментами.

---

## 📦 Серверы

| Сервер | Инструментов | Статус | Описание |
|--------|:---:|:---:|---------|
| [renga/](./renga/) | 18 | ✅ | BIM-система Renga через COM API |
| [msproject/](./msproject/) | 38 | ✅ | Microsoft Project через COM / Aspose |
| [revit/](./revit/) | 3+ | ✅ | Revit через pyRevit Routes / C# Addin |

**В планах:** AutoCAD MCP, ArchiCAD MCP

---

## 🧠 Скиллы для Claude Projects

| Скилл | Назначение |
|-------|-----------|
| `renga-mcp-SKILL.md` | Renga COM API, GUID объектов, паттерны, отладка |
| `msproject-mcp-SKILL.md` | MS Project, EVM, ресурсы, базовый план |
| `revit-mcp-connector-SKILL.md` | Подключение Revit MCP, два стека, диагностика |
| `revit-family-creator-SKILL.md` | Создание семейств по BIM-стандарту 2.0, ФОП 2021 |
| `revit-dwg-modeler-SKILL.md` | Подъём BIM-модели из DWG-подложки |
| `normcontrol-pdf-SKILL.md` | Нормоконтроль ПД по ГОСТ 21.101, 21.110 |
| `notebooklm-mcp-SKILL.md` | NotebookLM через MCP, поиск по каталогам |

---

## 🚀 Быстрый старт

```bash
git clone https://github.com/chernetchenko-dev/construction-ai-mcp.git
cd construction-ai-mcp
```

Зависимости:
```bash
pip install mcp pywin32           # Renga
pip install mcp pywin32 openpyxl  # MS Project
pip install uv                    # Revit
```

Конфиг: `%APPDATA%\Claude\claude_desktop_config.json` → см. [examples/](./examples/)

Скиллы: Claude.ai → Projects → загрузить нужные `.md` из папки `skills/`

---

## 📁 Структура

```
construction-ai-mcp/
├── renga/            ← сервер Renga (18 инструментов)
├── msproject/        ← сервер MS Project (38 инструментов)
├── revit/            ← настройка Revit MCP (два стека)
├── skills/           ← 7 SKILL.md файлов для Claude Projects
├── examples/         ← конфиг и примеры скриптов
└── LICENSE
```

---

## 🖥️ Требования

- Windows (COM API — только Windows)
- Python 3.10+
- Claude Desktop или Qwen Desktop
- Renga v8.7+, MS Project Professional, Revit 2021+ с pyRevit

---

## 📄 Лицензия

[MIT](./LICENSE)
