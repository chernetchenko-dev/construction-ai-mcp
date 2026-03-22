"""
Renga MCP Server v2  —  github.com/your-org/renga-mcp
======================================================
Создание, чтение и изменение объектов Renga через COM API.
Новое в v2: создание уровней, колонн, окон, дверей, пластин,
фундаментов, сборок; массовое изменение параметров; аудит.

pip install mcp pywin32

claude_desktop_config.json:
{
  "mcpServers": {
    "renga": {
      "command": "python",
      "args": ["C:\\tools\\renga-mcp\\renga_mcp_server_v2.py"]
    }
  }
}
"""

import json
import uuid as _uuid
from mcp.server.fastmcp import FastMCP

try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

mcp = FastMCP("Renga MCP v2")

# ── GUID типов объектов ────────────────────────────────────────
ENTITY_TYPES = {
    "Wall":                "{05E96D5C-4C85-4CE1-B9AB-D5E02EF75DDB}",
    "Column":              "{B12A4D6F-8C84-4BFD-92F7-1D27B65F8E0E}",
    "Beam":                "{A0C52F21-5C38-4D69-A6B4-2C1FA6E8F4C2}",
    "Floor":               "{3F8E3C0A-1C9E-4E7D-B6A2-8F0D9E3C5B1A}",
    "Roof":                "{7A2B4C6D-8E0F-4A1B-9C3D-5E7F1A2B4C6D}",
    "Window":              "{9C1E3A5B-7D9F-4B1D-A3E5-C7F9B1D3E5A7}",
    "Door":                "{4D6E8F0A-2C4E-6A8C-0E2A-4C6E8A0C2E4A}",
    "Opening":             "{2B4C6D8E-0F2A-4C6E-8A0C-2E4F6A8C0E2A}",
    "Level":               "{C3CE17FF-6F28-411F-B18D-74FE957B2BA8}",
    "Room":                "{1A3C5E7F-9B1D-3F5A-7C9E-1B3D5F7A9C1E}",
    "Assembly":            "{6E8F0A2C-4E6A-8C0E-2A4C-6E8A0C2E4A6E}",
    "Staircase":           "{8F0A2C4E-6A8C-0E2A-4C6E-8A0C2E4A6E8F}",
    "IsolatedFoundation":  "{5B7D9F1B-3D5A-7C9E-1B3D-5F7A9C1E3B5D}",
    "WallFoundation":      "{3D5A7C9E-1B3D-5F7A-9C1E-3B5D7A9C1E3D}",
    "Plate":               "{1B3D5F7A-9C1E-3B5D-7A9C-1E3B5D7A9C1E}",
    "Pipe":                "{7F9B1D3F-5A7C-9E1B-3D5F-7A9C1E3B5D7F}",
    "Duct":                "{9B1D3F5A-7C9E-1B3D-5F7A-9C1E3B5D7A9C}",
    "Equipment":           "{D3F5A7C9-E1B3-D5F7-A9C1-E3B5D7A9C1E3}",
    "MechanicalEquipment": "{A7C9E1B3-D5F7-A9C1-E3B5-D7A9C1E3B5D7}",
}

# ── Хелперы ───────────────────────────────────────────────────

def get_app():
    if not WIN32_AVAILABLE:
        raise RuntimeError("pywin32 не установлен. pip install pywin32")
    try:
        return win32.GetActiveObject("Renga.Application.1")
    except Exception:
        raise RuntimeError("Renga не запущена.")

def get_project():
    p = get_app().Project
    if not p:
        raise RuntimeError("Проект не открыт.")
    return p

def get_model():
    return get_project().Model

def ok(d): return json.dumps({"status":"ok",**d}, ensure_ascii=False, indent=2)
def err(m): return json.dumps({"status":"error","message":m}, ensure_ascii=False)

def resolve_type(guid_str):
    g = guid_str.upper()
    for n,g2 in ENTITY_TYPES.items():
        if g2.upper()==g: return n
    return guid_str

def iter_obj(model):
    col = model.GetObjects()
    for i in range(col.Count):
        yield col.GetByIndex(i)

# ── СТАТУС / ПРОЕКТ ────────────────────────────────────────────

@mcp.tool()
def renga_status() -> str:
    """Проверить подключение к Renga."""
    if not WIN32_AVAILABLE:
        return err("pywin32 не установлен")
    try:
        app = get_app()
        info = {"renga_running": True}
        try:
            proj = app.Project
            info["project_open"] = proj is not None
            if proj:
                try: info["project_name"] = proj.Name
                except: info["project_name"] = "?"
        except: info["project_open"] = False
        return ok(info)
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_project_info() -> str:
    """Имя, путь и количество объектов по типам."""
    try:
        proj = get_project()
        info = {}
        try: info["name"] = proj.Name
        except: pass
        try: info["file_path"] = proj.FilePath
        except: pass
        counts = {}
        for obj in iter_obj(proj.Model):
            try:
                t = resolve_type(str(obj.ObjectTypeS))
                counts[t] = counts.get(t,0)+1
            except: pass
        info["total"] = sum(counts.values())
        info["by_type"] = counts
        return ok(info)
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_save_project() -> str:
    """Сохранить проект."""
    try:
        get_project().Save()
        return ok({"message":"Проект сохранён"})
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_open_project(file_path: str) -> str:
    """Открыть .rnp файл. Args: file_path — полный путь к файлу."""
    try:
        get_app().OpenProject(file_path)
        return ok({"opened":file_path})
    except Exception as e:
        return err(str(e))

# ── УРОВНИ ────────────────────────────────────────────────────

@mcp.tool()
def renga_get_levels() -> str:
    """Список уровней с именами, отметками (мм) и UniqueId."""
    try:
        model = get_model()
        lg = ENTITY_TYPES["Level"].upper()
        result = []
        for obj in iter_obj(model):
            try:
                if str(obj.ObjectTypeS).upper() != lg: continue
                e = {"id": str(obj.UniqueIdS), "local_id": obj.Id}
                try:
                    iface = obj.GetInterfaceByName("ILevel")
                    e["name"] = iface.LevelName
                    e["elevation_mm"] = iface.Elevation
                except: pass
                result.append(e)
            except: continue
        result.sort(key=lambda x: x.get("elevation_mm",0))
        return ok({"count":len(result),"levels":result})
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_create_level(name: str, elevation_mm: float) -> str:
    """
    Создать уровень.
    Args:
        name: Имя уровня, например "Этаж 2".
        elevation_mm: Отметка в мм (3000 = 3 метра).
    """
    try:
        proj = get_project()
        model = proj.Model
        args = model.CreateNewEntityArgs()
        args.TypeIdS = ENTITY_TYPES["Level"]
        op = proj.CreateOperation()
        op.Start()
        new_obj = model.CreateObject(args)
        try:
            iface = new_obj.GetInterfaceByName("ILevel")
            iface.LevelName = name
            iface.Elevation = elevation_mm
        except: pass
        op.Apply()
        return ok({"created":"Level","name":name,"elevation_mm":elevation_mm,
                   "id": str(new_obj.UniqueIdS) if new_obj else None})
    except Exception as e:
        return err(str(e))

# ── ОБЪЕКТЫ — ЧТЕНИЕ ──────────────────────────────────────────

@mcp.tool()
def renga_get_objects(object_type: str = "", limit: int = 200) -> str:
    """
    Список объектов модели.
    Args:
        object_type: Фильтр по типу: Wall, Column, Beam, Floor, Window, Door,
                     Level, Room, Pipe, Duct, Equipment, Assembly и др. Пусто = все.
        limit: Максимум объектов в ответе (по умолчанию 200).
    """
    try:
        model = get_model()
        fg = ENTITY_TYPES.get(object_type,"").upper()
        result = []
        total = 0
        for obj in iter_obj(model):
            try:
                total += 1
                tg = str(obj.ObjectTypeS).upper()
                if fg and tg != fg: continue
                if len(result) >= limit: continue
                tn = object_type or resolve_type(tg)
                e = {"id": str(obj.UniqueIdS), "local_id": obj.Id, "type": tn}
                try: e["name"] = obj.Name
                except: pass
                result.append(e)
            except: continue
        return ok({"count":len(result),"total_in_model":total,"objects":result})
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_get_object_params(object_id: str) -> str:
    """
    Параметры, расчётные характеристики и свойства объекта.
    Args:
        object_id: UniqueId объекта (из renga_get_objects).
    """
    try:
        model = get_model()
        target = None
        for obj in iter_obj(model):
            try:
                if str(obj.UniqueIdS) == object_id:
                    target = obj; break
            except: pass
        if not target:
            return err(f"Объект {object_id} не найден")

        info = {"id":object_id, "type":resolve_type(str(target.ObjectTypeS))}
        try: info["name"] = target.Name
        except: pass

        try:
            params = target.GetParameters()
            ids = params.GetIds(); pd = {}
            for j in range(ids.Count):
                p = params.Get(ids.GetByIndex(j))
                try:
                    pn = p.GetDefinition().Name
                    try: pd[pn] = p.GetDoubleValue()
                    except:
                        try: pd[pn] = p.GetStringValue()
                        except: pd[pn] = "н/д"
                except: pass
            info["parameters"] = pd
        except Exception as e: info["parameters_error"] = str(e)

        try:
            quant = target.GetQuantities()
            ids = quant.GetIds(); qd = {}
            for j in range(ids.Count):
                q = quant.Get(ids.GetByIndex(j))
                try:
                    qn = q.GetDefinition().Name
                    try: qd[qn] = q.AsDouble()
                    except: qd[qn] = "н/д"
                except: pass
            info["quantities"] = qd
        except Exception as e: info["quantities_error"] = str(e)

        try:
            props = target.GetProperties()
            ids = props.GetIds(); propd = {}
            for j in range(ids.Count):
                p = props.Get(ids.GetByIndex(j))
                try:
                    pn = p.Name
                    try: propd[pn] = p.GetStringValue()
                    except:
                        try: propd[pn] = p.GetDoubleValue()
                        except: propd[pn] = "н/д"
                except: pass
            info["custom_properties"] = propd
        except Exception as e: info["custom_properties_error"] = str(e)

        return ok(info)
    except Exception as e:
        return err(str(e))

# ── СОЗДАНИЕ ОБЪЕКТОВ ─────────────────────────────────────────

def _create_on_level(etype: str, level_id: str, style_id: str = "") -> str:
    tg = ENTITY_TYPES.get(etype)
    if not tg:
        return err(f"Неизвестный тип {etype}. Доступные: {list(ENTITY_TYPES)}")
    try:
        proj = get_project()
        model = proj.Model
        lg = ENTITY_TYPES["Level"].upper()
        host = None
        for obj in iter_obj(model):
            try:
                if str(obj.ObjectTypeS).upper() == lg:
                    if str(obj.UniqueIdS)==level_id or str(obj.Id)==level_id:
                        host = obj; break
            except: pass
        if not host:
            return err(f"Уровень id={level_id} не найден. Используйте renga_get_levels.")
        args = model.CreateNewEntityArgs()
        args.TypeIdS = tg
        args.HostObjectIdS = str(host.UniqueIdS)
        if style_id: args.StyleIdS = style_id
        op = proj.CreateOperation(); op.Start()
        new_obj = model.CreateObject(args)
        op.Apply()
        return ok({"created":etype, "id": str(new_obj.UniqueIdS) if new_obj else None,
                   "level_id":level_id,
                   "note":"Объект создан с позицией по умолчанию. "
                          "Задайте координаты через renga_set_object_param."})
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_create_column(level_id: str, style_id: str = "") -> str:
    """Создать колонну на уровне. Args: level_id — UniqueId уровня; style_id — стиль (необязательно)."""
    return _create_on_level("Column", level_id, style_id)

@mcp.tool()
def renga_create_window(level_id: str, style_id: str = "") -> str:
    """Создать окно на уровне. Args: level_id — UniqueId уровня; style_id — стиль (необязательно)."""
    return _create_on_level("Window", level_id, style_id)

@mcp.tool()
def renga_create_door(level_id: str, style_id: str = "") -> str:
    """Создать дверь на уровне. Args: level_id — UniqueId уровня; style_id — стиль (необязательно)."""
    return _create_on_level("Door", level_id, style_id)

@mcp.tool()
def renga_create_isolated_foundation(level_id: str, style_id: str = "") -> str:
    """Создать столбчатый фундамент. Args: level_id — UniqueId уровня; style_id — стиль (необязательно)."""
    return _create_on_level("IsolatedFoundation", level_id, style_id)

@mcp.tool()
def renga_create_plate(level_id: str, style_id: str = "") -> str:
    """Создать пластину (плиту). Args: level_id — UniqueId уровня; style_id — стиль (необязательно)."""
    return _create_on_level("Plate", level_id, style_id)

@mcp.tool()
def renga_create_assembly(level_id: str, style_id: str = "") -> str:
    """Создать сборку. Args: level_id — UniqueId уровня; style_id — стиль (необязательно)."""
    return _create_on_level("Assembly", level_id, style_id)

@mcp.tool()
def renga_delete_object(object_id: str) -> str:
    """Удалить объект по UniqueId. Зависимые объекты удаляются вместе с ним."""
    try:
        proj = get_project()
        op = proj.CreateOperation(); op.Start()
        proj.Model.DeleteObjectByUniqueIdS(object_id)
        op.Apply()
        return ok({"deleted_id": object_id})
    except Exception as e:
        return err(str(e))

# ── ИЗМЕНЕНИЕ ПАРАМЕТРОВ ──────────────────────────────────────

@mcp.tool()
def renga_set_object_param(object_id: str, param_name: str, value: float) -> str:
    """
    Изменить числовой параметр объекта. Размерные параметры в миллиметрах.
    Args:
        object_id: UniqueId объекта.
        param_name: Имя параметра (из renga_get_object_params → parameters).
        value: Новое значение (мм для размеров).
    """
    try:
        proj = get_project()
        model = proj.Model
        target = None
        for obj in iter_obj(model):
            try:
                if str(obj.UniqueIdS)==object_id: target=obj; break
            except: pass
        if not target: return err(f"Объект {object_id} не найден")
        params = target.GetParameters()
        ids = params.GetIds()
        op = proj.CreateOperation(); op.Start()
        found = False
        for j in range(ids.Count):
            p = params.Get(ids.GetByIndex(j))
            try:
                if p.GetDefinition().Name == param_name:
                    p.SetDoubleValue(value); found=True; break
            except: pass
        op.Apply()
        if not found:
            return err(f"Параметр '{param_name}' не найден. "
                       "Смотрите renga_get_object_params → parameters.")
        return ok({"object_id":object_id,"param":param_name,"new_value":value})
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_bulk_set_param(
    object_type: str,
    param_name: str,
    value: float,
    dry_run: bool = True
) -> str:
    """
    Массово изменить параметр у всех объектов типа.
    ВАЖНО: сначала запускайте с dry_run=True для проверки.
    Args:
        object_type: Wall, Column, Beam, Floor и т.д.
        param_name: Имя параметра (например "WallHeight").
        value: Новое значение в мм.
        dry_run: True = показать изменения, False = применить.
    """
    try:
        proj = get_project()
        model = proj.Model
        tg = ENTITY_TYPES.get(object_type,"").upper()
        if not tg: return err(f"Неизвестный тип: {object_type}")

        targets = []
        for obj in iter_obj(model):
            try:
                if str(obj.ObjectTypeS).upper() != tg: continue
                params = obj.GetParameters()
                ids = params.GetIds()
                for j in range(ids.Count):
                    p = params.Get(ids.GetByIndex(j))
                    try:
                        if p.GetDefinition().Name == param_name:
                            try: old = p.GetDoubleValue()
                            except: old = None
                            targets.append({"id":str(obj.UniqueIdS),"old":old})
                            break
                    except: pass
            except: pass

        if not targets:
            return ok({"dry_run":dry_run,"affected":0,
                       "message":f"Нет объектов {object_type} с параметром {param_name}"})

        if not dry_run:
            op = proj.CreateOperation(); op.Start()
            for t in targets:
                for obj in iter_obj(model):
                    try:
                        if str(obj.UniqueIdS)==t["id"]:
                            params = obj.GetParameters()
                            ids = params.GetIds()
                            for j in range(ids.Count):
                                p = params.Get(ids.GetByIndex(j))
                                try:
                                    if p.GetDefinition().Name==param_name:
                                        p.SetDoubleValue(value); break
                                except: pass
                            break
                    except: pass
            op.Apply()

        return ok({"dry_run":dry_run,"type":object_type,"param":param_name,
                   "new_value_mm":value,"affected":len(targets),"objects":targets[:50]})
    except Exception as e:
        return err(str(e))

# ── СВОЙСТВА ──────────────────────────────────────────────────

@mcp.tool()
def renga_create_property(
    name: str,
    property_type: str,
    object_types: str = "Wall,Column,Beam,Floor"
) -> str:
    """
    Зарегистрировать пользовательское свойство и назначить объектным типам.
    Args:
        name: Имя, например "Подрядчик" или "Статус монтажа".
        property_type: String, Double, Integer или Boolean.
        object_types: Типы через запятую: "Wall,Column,Beam".
    """
    try:
        proj = get_project()
        pmgr = proj.PropertyManager
        tm = {"string":2,"double":3,"integer":4,"boolean":5}
        pt = tm.get(property_type.lower())
        if pt is None: return err(f"Неизвестный тип {property_type}. Допустимые: String/Double/Integer/Boolean")
        pid = "{"+str(_uuid.uuid4()).upper()+"}"
        op = proj.CreateOperation(); op.Start()
        pmgr.RegisterPropertyS(pid, name, pt)
        assigned = []
        for ot in [t.strip() for t in object_types.split(",") if t.strip()]:
            tg = ENTITY_TYPES.get(ot)
            if tg:
                try: pmgr.AssignPropertyToTypeS(pid, tg); assigned.append(ot)
                except Exception as e: assigned.append(f"{ot}(err:{e})")
            else: assigned.append(f"{ot}(тип не найден)")
        op.Apply()
        return ok({"property_name":name,"property_id":pid,"type":property_type,"assigned_to":assigned})
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_set_property_value(object_id: str, property_name: str, value: str) -> str:
    """
    Задать значение пользовательского свойства объекту.
    Args:
        object_id: UniqueId объекта.
        property_name: Имя свойства.
        value: Значение строкой (числа тоже строкой).
    """
    try:
        proj = get_project()
        model = proj.Model
        target = None
        for obj in iter_obj(model):
            try:
                if str(obj.UniqueIdS)==object_id: target=obj; break
            except: pass
        if not target: return err(f"Объект {object_id} не найден")
        op = proj.CreateOperation(); op.Start()
        props = target.GetProperties()
        ids = props.GetIds(); found = False
        for j in range(ids.Count):
            p = props.Get(ids.GetByIndex(j))
            try:
                if p.Name == property_name:
                    try: p.SetStringValue(value)
                    except:
                        try: p.SetDoubleValue(float(value))
                        except: pass
                    found=True; break
            except: pass
        op.Apply()
        if not found: return err(f"Свойство '{property_name}' не найдено у объекта {object_id}")
        return ok({"object_id":object_id,"property":property_name,"value":value})
    except Exception as e:
        return err(str(e))

# ── ЭКСПОРТ ───────────────────────────────────────────────────

@mcp.tool()
def renga_export_ifc(output_path: str) -> str:
    """Экспорт в IFC. Args: output_path — путь к файлу .ifc."""
    try:
        get_project().ExportToIFC(output_path)
        return ok({"exported_to":output_path})
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_export_drawings(output_folder: str, format: str = "DWG") -> str:
    """
    Пакетный экспорт чертежей, отсортированных по имени листа.
    Args:
        output_folder: Папка для файлов.
        format: DWG, DXF или PDF.
    """
    try:
        proj = get_project()
        fmt = format.upper()
        if fmt not in ("DWG","DXF","PDF"): return err("format: DWG/DXF/PDF")
        sheets = []
        for i in range(proj.Drawings.Count):
            d = proj.Drawings.Item(i)
            try: sheets.append((d.Name,d))
            except: pass
        sheets.sort(key=lambda x:x[0])
        exported, errors = [], []
        for name,d in sheets:
            fp = f"{output_folder.rstrip('/')}/{name}.{fmt.lower()}"
            try:
                if fmt=="DWG": d.ExportToDWG(fp)
                elif fmt=="DXF": d.ExportToDXF(fp)
                elif fmt=="PDF": d.ExportToPDF(fp)
                exported.append({"name":name,"file":fp})
            except Exception as e:
                errors.append(f"{name}: {e}")
        return ok({"format":fmt,"exported":len(exported),"errors":errors or None,"drawings":exported})
    except Exception as e:
        return err(str(e))

# ── АУДИТ ─────────────────────────────────────────────────────

@mcp.tool()
def renga_audit_model(check_properties: str = "") -> str:
    """
    Аудит модели.
    Args:
        check_properties: Имена свойств через запятую для проверки пустых значений.
                          Пустая строка — общий аудит (подсчёт по типам).
    """
    try:
        model = get_model()
        if not check_properties:
            counts = {}
            for obj in iter_obj(model):
                try:
                    t = resolve_type(str(obj.ObjectTypeS))
                    counts[t] = counts.get(t,0)+1
                except: pass
            return ok({"total":sum(counts.values()),"by_type":counts})

        pnames = [p.strip() for p in check_properties.split(",") if p.strip()]
        missing = {p:[] for p in pnames}
        for obj in iter_obj(model):
            try:
                props = obj.GetProperties()
                ids = props.GetIds()
                vals = {}
                for j in range(ids.Count):
                    p = props.Get(ids.GetByIndex(j))
                    try:
                        try: vals[p.Name] = p.GetStringValue()
                        except: vals[p.Name] = str(p.GetDoubleValue())
                    except: pass
                uid = str(obj.UniqueIdS)
                for pn in pnames:
                    v = vals.get(pn,"")
                    if not v or v in ("0.0","0"):
                        missing[pn].append(uid)
            except: pass
        report = {"checked":pnames}
        for pn in pnames:
            report[f"missing_{pn}"] = {"count":len(missing[pn]),"ids":missing[pn][:100]}
        return ok(report)
    except Exception as e:
        return err(str(e))

@mcp.tool()
def renga_list_entity_types() -> str:
    """Список поддерживаемых типов объектов."""
    return ok({"entity_types":list(ENTITY_TYPES.keys())})

# ── ТОЧКА ВХОДА ───────────────────────────────────────────────
if __name__ == "__main__":
    mcp.run()
