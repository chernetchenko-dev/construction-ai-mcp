---
name: revit-dwg-modeler
description: "Моделирование строительных элементов в Revit на основе DWG-подложки через MCP. Используй этот скилл ВСЕГДА когда пользователь хочет поднять модель из DWG, создать стены/помещения/двери/окна по чертежу AutoCAD, проанализировать DWG-план в Revit, определить типы стен по штриховке, нанести или проверить оси, замоделить здание или помещение по подложке. Скилл работает с любым DWG и любым проектом Revit через pyRevit MCP на порту 48884."
---

# Revit DWG Modeler

Скилл для пошагового подъёма BIM-модели из DWG-подложки через Revit MCP API.

## Инструменты
- `Revit Connector:execute_revit_code` — выполнение IronPython в Revit
- `Revit Connector:get_revit_status` — проверка коннекта
- `Revit Connector:get_revit_model_info` — информация о модели

## Технические константы
```python
ft = 0.3048   # Revit API работает в футах, DWG в метрах
# Транзакции: with revit.Transaction("name"):
# Кириллица: .encode('utf-8') только в print(), НЕ в return
```

---

## ПОРЯДОК РАБОТЫ С НОВЫМ ПРОЕКТОМ

```
0. Проверить коннект (get_revit_status)
1. ОСИ — всегда первый шаг:
   а) Есть оси в модели? → уточнить правильные ли
   б) Нет осей → найти в DWG → создать → подтвердить с пользователем
2. Найти DWG-подложку → определить архитектурный слой
3. Проанализировать геометрию → найти стены
4. Показать схему пользователю → подтвердить до моделирования
5. Запросить типы стен из документа → назначить по справочнику
6. Создать стены пакетом
7. Проверить замкнутость помещений → исправить
8. Создать помещения по экспликации
```

---

## ЭТАП 0 — Проверка коннекта

```python
from pyrevit import revit, DB
doc = revit.doc
print("Doc:", doc.Title)
walls  = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
grids  = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
print("Walls:", len(list(walls)), "Grids:", len(list(grids)))
```

При потере коннекта:
```
taskkill /f /im python.exe
uv run main.py --combined   (в папке расширения)
→ перезапустить Claude Desktop
```

---

## ЭТАП 1 — ОСИ (первый шаг всегда)

### 1.1 Проверить оси в модели

```python
from pyrevit import revit, DB
doc = revit.doc
ft = 0.3048

grids = list(DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_Grids)
    .WhereElementIsNotElementType().ToElements())

print("Grids:", len(grids))
for g in grids:
    c = g.Curve
    p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
    print("  {} ({:.2f},{:.2f})->({:.2f},{:.2f})".format(
        g.Name.encode('utf-8'),
        p0.X*ft, p0.Y*ft, p1.X*ft, p1.Y*ft))
```

- **Оси есть** → показать пользователю, спросить "Оси правильные?"
  - Да → идём к ЭТАПУ 2
  - Нет → удалить, создать новые по DWG (п. 1.2–1.4)
- **Осей нет** → идём к 1.2

### 1.2 Найти маркеры осей в DWG

Признаки осей в DWG:
- Кружки (Arc) с R=0.15–0.60м **за пределами зоны плана**
- Слои: `ОСЬ`, `AXIS`, `GRID`, `РАЗБ`, `АХ`, `OS`
- Длинные одиночные линии через весь план

```python
def find_axis_markers(doc, dwg_id, y_plan_min, x_plan_min):
    """Кружки маркеров осей в зоне за пределами плана"""
    ft = 0.3048
    imp = doc.GetElement(DB.ElementId(dwg_id))
    geo = imp.get_Geometry(DB.Options())
    markers = []
    for obj in geo:
        if type(obj).__name__ != 'GeometryInstance': continue
        t = obj.Transform
        for sub in obj.GetSymbolGeometry():
            try:
                if type(sub).__name__ != 'Arc': continue
                c = t.OfPoint(sub.Center)
                r = sub.Radius * ft
                cx, cy = c.X*ft, c.Y*ft
                if not (0.15 < r < 0.60): continue
                if cy < y_plan_min - 0.3 or cx < x_plan_min - 0.3:
                    markers.append({'x': round(cx,3), 'y': round(cy,3), 'r': round(r,3)})
            except: pass
    return sorted(markers, key=lambda m: m['x'])
```

### 1.3 Создать оси в Revit

```python
def create_grid(doc, x0, y0, x1, y1, name):
    ft = 0.3048
    def m(v): return v/ft
    p0 = DB.XYZ(m(x0), m(y0), 0)
    p1 = DB.XYZ(m(x1), m(y1), 0)
    if p0.DistanceTo(p1) < 0.5/ft: return None
    grid = DB.Grid.Create(doc, DB.Line.CreateBound(p0, p1))
    grid.Name = name
    return grid

# Вертикальные оси (по X-координатам маркеров):
# from x_coord → ось идёт y_from..y_to (чуть за пределы плана)
# Горизонтальные оси (по Y-координатам маркеров):
# from y_coord → ось идёт x_from..x_to

with revit.Transaction("Create grids"):
    for x, name in vert_axes:
        create_grid(doc, x, y_from, x, y_to, name)
    for y, name in horiz_axes:
        create_grid(doc, x_from, y, x_to, y, name)
```

### 1.4 Подтвердить с пользователем

После создания — показать список осей и спросить:
- Правильные имена?
- Правильное расположение?
- Нет лишних / пропущенных?

**Только после подтверждения → переходить к стенам.**

---

## ЭТАП 2 — Разведка DWG

### 2.1 Найти подложку
```python
imports = list(DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance).ToElements())
for imp in imports:
    print("id:{} Linked:{} {}".format(
        imp.Id.IntegerValue, imp.IsLinked,
        imp.get_Parameter(DB.BuiltInParameter.IMPORT_SYMBOL_NAME).AsString().encode('utf-8')))
```

### 2.2 Слои DWG → архитектурный слой
```python
ARCH_KEYWORDS = [u'АРХ', u'ARCH', u'WALL', u'СТЕН', u'AR-', u'A-WALL']

# Собрать все слои с количеством элементов
# Выбрать слой по ключевым словам (или самый нагруженный)
# Проверить: есть ли парные линии → архитектурный слой
```

---

## ЭТАП 3 — Анализ стен

### 3.1 Извлечь линии архитектурного слоя
```python
def get_arch_segments(doc, dwg_id, layer_name):
    import math
    ft = 0.3048
    imp = doc.GetElement(DB.ElementId(dwg_id))
    geo = imp.get_Geometry(DB.Options())
    segs = []
    for obj in geo:
        if type(obj).__name__ != 'GeometryInstance': continue
        t = obj.Transform
        for sub in obj.GetSymbolGeometry():
            try:
                gs = doc.GetElement(sub.GraphicsStyleId)
                if not gs or gs.GraphicsStyleCategory.Name != layer_name: continue
                pts = []
                tn = type(sub).__name__
                if tn == 'Line':
                    pts = [(t.OfPoint(sub.GetEndPoint(0)), t.OfPoint(sub.GetEndPoint(1)))]
                elif tn == 'PolyLine':
                    coords = list(sub.GetCoordinates())
                    pts = [(t.OfPoint(coords[i]), t.OfPoint(coords[i+1]))
                           for i in range(len(coords)-1)]
                for p0,p1 in pts:
                    x0,y0 = p0.X*ft, p0.Y*ft
                    x1,y1 = p1.X*ft, p1.Y*ft
                    L = ((x1-x0)**2+(y1-y0)**2)**0.5
                    if L < 0.1: continue
                    segs.append((x0,y0,x1,y1,L))
            except: pass
    # Дедупликация с допуском 3см
    dd = []
    for s in segs:
        if not any(all(abs(s[i]-d[i])<0.03 for i in range(4)) for d in dd):
            dd.append(s)
    return dd
```

### 3.2 Найти парные линии (включая переменную толщину)

```python
def find_walls(segs, t_min=0.07, t_max=0.60, overlap=0.25):
    walls = []
    used = set()
    for i,(x0a,y0a,x1a,y1a,La) in enumerate(segs):
        if i in used: continue
        best = None; best_sc = 0
        for j,(x0b,y0b,x1b,y1b,Lb) in enumerate(segs):
            if j<=i or j in used: continue
            # Расстояние между серединами
            dp = (((x0a+x1a)/2-(x0b+x1b)/2)**2+((y0a+y1a)/2-(y0b+y1b)/2)**2)**0.5
            if not (t_min < dp < t_max): continue
            sc = min(La,Lb)/max(La,Lb)
            if sc < overlap: continue
            if sc > best_sc: best_sc=sc; best=(j,x0b,y0b,x1b,y1b,Lb,dp)
        if best:
            j,x0b,y0b,x1b,y1b,Lb,dp = best
            # Толщина в начале и конце (для переменной толщины)
            t_s = ((x0a-x0b)**2+(y0a-y0b)**2)**0.5
            t_e = ((x1a-x1b)**2+(y1a-y1b)**2)**0.5
            t_max_v = max(t_s,t_e)
            delta = abs(t_s-t_e)/t_max_v if t_max_v>0 else 0
            walls.append({
                'x0':(x0a+x0b)/2,'y0':(y0a+y0b)/2,
                'x1':(x1a+x1b)/2,'y1':(y1a+y1b)/2,
                't': t_max_v,           # всегда максимум
                't_start': t_s, 't_end': t_e,
                'status': 'ok' if delta<=0.10 else 'variable',
                'L': (La+Lb)/2
            })
            used.add(i); used.add(j)
    return [w for w in walls if w['L']>=0.4]
```

### 3.3 Правила дополнения

```
• Линия касается границы чертежа → наружная стена (даже одиночная)
• Нет контура между зонами → стена-замыкатель
• Переменная толщина delta ≤ 10% → берём max, моделируем прямой
• Переменная толщина delta > 10% → лог + уточнить у пользователя
```

**Всегда показать схему пользователю и получить подтверждение перед созданием.**

---

## ЭТАП 4 — Типы стен

См. `references/wall-materials.md` — полный справочник по ГОСТ 2.306-68.

Быстро по толщине:
```
< 100мм  → ГКЛ 100      |  250–350мм → Бетон 300 / Кирпич 250
100–150  → ГКЛ 125/ПГП  |  350–430мм → Кирпич 380
150–200  → Газобетон     |  > 430мм   → Газобетон 400 / Наружная
200–250  → Бетон 200     |  Неизвестно → Условная 200мм
```

Переменные стены — вывести отдельно для подтверждения:
```python
for w in [w for w in walls if w['status']=='variable']:
    print("ПЕРЕМЕННАЯ: ({:.2f},{:.2f})->({:.2f},{:.2f}) "
          "t_start={:.0f}мм t_end={:.0f}мм delta={:.0%}".format(
          w['x0'],w['y0'],w['x1'],w['y1'],
          w['t_start']*1000, w['t_end']*1000,
          abs(w['t_start']-w['t_end'])/max(w['t_start'],w['t_end'])))
```

---

## ЭТАПЫ 5–7 — Создание стен, проверка, помещения

(Код в предыдущих сессиях — добавить в следующей версии скилла)

### Создание стены
```python
def create_wall(doc, level, x0,y0,x1,y1, wtype_id, h=3.2):
    ft=0.3048; m=lambda v: v/ft
    p0=DB.XYZ(m(x0),m(y0),0); p1=DB.XYZ(m(x1),m(y1),0)
    if p0.DistanceTo(p1)<0.04/ft: return None
    return DB.Wall.Create(doc,DB.Line.CreateBound(p0,p1),
                          DB.ElementId(wtype_id),level.Id,m(h),0,False,False)
```

### Линия разделения помещений
```python
def add_separation(doc, view, x0,y0,x1,y1):
    ft=0.3048; m=lambda v: v/ft
    plane=DB.SketchPlane.Create(doc,
        DB.Plane.CreateByNormalAndOrigin(DB.XYZ(0,0,1),DB.XYZ(0,0,0)))
    arr=DB.CurveArray()
    arr.Append(DB.Line.CreateBound(DB.XYZ(m(x0),m(y0),0),DB.XYZ(m(x1),m(y1),0)))
    doc.Create.NewRoomBoundaryLines(plane,arr,view)
```

---

## Частые ошибки

| Ошибка | Причина | Решение |
|--------|---------|---------|
| `No result received` | Routes завис | `taskkill /f /im python.exe` + `uv run main.py --combined` |
| JSON encode error | Кириллица в return | `.encode('utf-8')` только в `print()` |
| Помещение OPEN | Незамкнут контур | Стена-замыкатель или линия разделения |
| Ось не создаётся | Дублирующееся имя | Проверить уникальность `grid.Name` |
| Неверные координаты | ft/м путаница | `m(x) = x/0.3048` → Revit, `p.X*0.3048` → метры |

---

## Дополнительные справочники

- `references/wall-materials.md` — ГОСТ 2.306-68, штриховки, толщины, переменная толщина

---

## ЭТАП 0б — Экспликация помещений

Экспликация в DWG/на плане содержит: номер, название, площадь, доп. параметры.

### Как найти экспликацию в DWG
```python
# Слой "Экспликация помещений" — содержит PolyLine (рамки таблицы) и Line
# Тексты в AutoCAD DWG при импорте в Revit → NurbSpline (не читаются напрямую)
# Поэтому экспликацию вводим вручную по изображению/скриншоту от пользователя
```

### Формат экспликации (из скриншотов)
```python
# Структура: Номер по плану | Номер помещения | Наименование | Площадь м² | Доп.
# Пример корпус 6:
rooms_corp6 = [
    # (num_on_plan, room_code, name, area_m2)
    ("61-11",  "6160",   u"Офисное помещение ЗЛП",  33.8),
    ("61-11/1","",       u"",                        0),
    ("61-12",  "6161",   u"Комната отдыха",          8.9),
    ("61-13",  "6100",   u"Технический коридор",     11.4),
    ("61-13/1","6100",   u"Электрощитовая",          2.6),
    ("61-13/2","6100",   u"Тамбур",                  1.4),
    ("61-14",  "61038",  u"Вентикамера",             22.1),
    ("61-14/1","61048",  u"Вентикамера",             3.3),
    ("61-15",  "6102м",  u"Техническое помещение",   27.2),
    ("61-16",  "6101м",  u"Пожарный пост",           9.9),
    ("61-20",  "6162c",  u"С/у",                     2.6),
    ("7ЛК-1",  "Л6",     u"Лестничная клетка",       14.8),
]
```

### Алгоритм работы с экспликацией
1. Пользователь присылает скриншот экспликации → извлечь данные вручную
2. Сопоставить номера помещений с планом (слой "Экспликация помещений" в DWG)
3. Найти центральные точки помещений по их номерам на плане
4. Создать Room в Revit с правильным именем и номером
5. Проверить площадь Room vs площадь из экспликации (допуск ±5%)

```python
# Создание помещений по экспликации
with revit.Transaction("Create rooms from explikation"):
    for num, code, name, area in rooms_data:
        try:
            r = doc.Create.NewRoom(level, DB.UV(m(cx), m(cy)))
            r.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).Set(num)
            r.get_Parameter(DB.BuiltInParameter.ROOM_NAME).Set(name)
            revit_area = r.Area * ft * ft
            delta = abs(revit_area - area) / area if area > 0 else 0
            status = "OK" if delta < 0.05 else "CHECK area={:.1f} vs {:.1f}".format(revit_area, area)
            print("{}: {} {}".format(num, status, name.encode('utf-8')))
        except Exception as e:
            print("ERR {}: {}".format(num, str(e)[:60]))
```

### Важные замечания
- Номера помещений вида `61-14` = корпус 6, помещение 14
- Суффиксы `/1`, `/2` = подпомещения (ниши, тамбуры)
- Суффикс `c` = санузел
- `ЛК-N` = лестничная клетка
- Площадь в экспликации = нормативная, Revit считает по оси стен → возможно расхождение


---

## КРИТИЧЕСКИЕ ПРАВИЛА — ОСИ (из опыта)

### Правило 1: Слой осей
Слой называется **не всегда "ОСЬ" или "AXIS"**. Примеры реальных имён:
- `ОСИ ЛОФТ` — жилой комплекс
- `АХ`, `A-GRID`, `РАЗБИВКА` — другие проекты
Если не найден по ключевым словам — **спросить у пользователя** имя слоя.

### Правило 2: Оси строятся "по линии" DWG
НЕ вручную по координатам — а **точно копируя линии из слоя осей**.
```python
# ПРАВИЛЬНО: берём p0,p1 из GetInstanceGeometry() — мировые координаты
for sub in obj.GetInstanceGeometry():
    if type(sub).__name__ == 'Line':
        p0, p1 = sub.GetEndPoint(0), sub.GetEndPoint(1)
        g = DB.Grid.Create(doc, DB.Line.CreateBound(p0, p1))
```
GetSymbolGeometry() даёт локальные координаты блока — **неверно**.
GetInstanceGeometry() даёт мировые координаты — **правильно**.

### Правило 3: Оси могут быть наклонными
Здания с ромбовидной/повёрнутой сеткой имеют оси под углом.
Не фильтровать линии по "горизонтальные/вертикальные" — брать ВСЕ из слоя осей.

### Правило 4: Пересечение первых осей ≈ (0,0)
Начало координат DWG обычно совпадает с пересечением первых осей.
Если ось 1 не проходит через X≈0 или Y≈0 — проверить трансформ DWG.

### Правило 5: Дедупликация
Линии в DWG часто дублируются (блоки вставлены несколько раз).
Перед созданием осей — дедупликация с допуском 5-10 см по всем 4 координатам.
```python
uniq = []
for s in raw:
    if not any(all(abs(s[i]-u[i])<0.1 for i in range(4)) for u in uniq):
        uniq.append(s)
```

### Алгоритм поиска слоя осей
```python
# Сначала ищем по ключевым словам
OSI_KEYWORDS = [u'ОСИ', u'AXIS', u'GRID', u'РАЗБ', u'АХ', u'A-GRID', u'ЛОФТ']
# Если не нашли → показать пользователю список всех слоёв → он укажет нужный
# НИКОГДА не угадывать — только по точному имени или подтверждению пользователя
```


### Правило 6: Не все оси могут быть в слое
Оси 4–11 (промежуточные) часто отсутствуют в слое — нарисованы только маркеры.
В DWG есть только "несущие" оси разбивочной сетки, промежуточные — по размерным цепочкам.

### Правило 7: Нумерация по размерным цепочкам
Сопоставление по расстояниям между осями:
```python
# Шаг 1: измерить Δ между соседними осями из DWG (в мм)
# Шаг 2: найти эту последовательность в размерной цепочке скриншота
# Шаг 3: определить номера от совпадения Δ
# Пример: Δ(x5→x6)=4140 совпадает с цепочкой "4140" между 13→14 → x5=13, x6=14
```

### Правило 8: Двойная нумерация осей
На плане с несколькими корпусами оси могут иметь двойную нумерацию:
- Числовые оси правой части: 12, 13, 14...22
- Буквенные оси правой части: А, Б1, В1, Д1, Е1, Ж1, И
- Буквенные оси левой части: Б, В, Г, Д (короткие, наклонные)
Всегда спрашивать у пользователя: "Правильные имена осей?"


### Правило 9: Горизонтальные оси — правильные имена
Правая часть здания (снизу вверх): А, Б1, В1, Д1, Е1, Ж1, И (наклонная)
Левая часть здания (короткие, снизу вверх): Б, В, Г, Д
Имена со скриншота — единственный достоверный источник. Всегда показывать их пользователю.

### Правило 10: Часть осей может отсутствовать в слое DWG
Промежуточные оси 4-7, 8-11 нарисованы только маркерами без линий.
Создавать расчётно по цепочке размеров. Нумерацию подтверждать у пользователя.


### Правило 11: Привязка стен к осям
Перед моделированием стен определить тип привязки по правилам:
- Внутренние несущие/колонны → ось элемента = разбивочная ось
- Наружные несущие → внутренняя грань = ось ± t_внутр_стены/2
- Наружные самонесущие → внутренняя грань = ось (нулевая)
- Каркас, крайние колонны → наружная грань = ось (нулевая)
- **Старые здания могут нарушать правила** — всегда проверять по DWG

Подробнее: references/wall-materials.md раздел "ПРАВИЛА ПРИВЯЗКИ"

