# Renga MCP Server

18 инструментов для управления BIM-моделями Renga через Claude/Qwen.

**Файл сервера:** `renga_mcp_server_v2.py` (добавить вручную)

## Установка
```bash
pip install mcp pywin32
```

## Конфиг
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

## Инструменты

`renga_status` · `renga_project_info` · `renga_save_project` · `renga_open_project`  
`renga_get_objects` · `renga_get_object_params` · `renga_get_levels`  
`renga_create_level` · `renga_create_column` · `renga_create_window` · `renga_create_door`  
`renga_create_isolated_foundation` · `renga_create_assembly` · `renga_create_model_text`  
`renga_create_property` · `renga_bulk_set_property`  
`renga_export_ifc` · `renga_export_drawings`

## Требования
- Windows, Python 3.10+
- Renga v8.7+ (апрель 2025)
- Запущенная Renga с открытым проектом

Подробнее → [skills/renga-mcp-SKILL.md](../skills/renga-mcp-SKILL.md)
