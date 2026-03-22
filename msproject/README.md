# MS Project MCP Server

38 инструментов для управления проектами MS Project через Claude/Qwen.

**Файл сервера:** `msproject_mcp_server.py` (добавить вручную)

## Два режима
- **COM-режим** — полная функциональность, нужен MS Project
- **File-режим** — только чтение, без MS Project (`pip install aspose-tasks`)

## Установка
```bash
pip install mcp pywin32 openpyxl   # COM-режим
```

## Конфиг
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

## Группы инструментов

**Статус:** `msproject_status` · `msproject_open` · `msproject_save` · `msproject_project_info`  
**Задачи:** get / add / update / delete / link / bulk_update / critical_path / milestones  
**Ресурсы:** get / add / update / delete / overallocated / workload  
**EVM:** `msproject_get_earned_value` · `msproject_set_baseline` · `msproject_get_baseline_comparison`  
**Экспорт:** PDF · Excel · CSV · HTML · XML

Подробнее → [skills/msproject-mcp-SKILL.md](../skills/msproject-mcp-SKILL.md)
