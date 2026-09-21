# SapfirLib COM

Плагин работает с САПФИР через COM-библиотеку `SapfirLib`. Ссылка объявлена в проекте:

```xml
<COMReference Include="SapfirLib">
  <Guid>{94863C7A-F77E-418F-9F50-3C0AD68A509A}</Guid>
  <VersionMajor>19</VersionMajor>
  <VersionMinor>0</VersionMinor>
  <Lcid>0</Lcid>
  <WrapperTool>tlbimp</WrapperTool>
  <Isolated>False</Isolated>
  <EmbedInteropTypes>True</EmbedInteropTypes>
</COMReference>
```

`EmbedInteropTypes=True` означает, что типы взаимодействия **встраиваются** в сборку плагина: отдельный interop-DLL не требуется, но COM-сервер САПФИР должен быть зарегистрирован в системе.

## Используемый объектный API

### `SapfirLib.Application`

| Метод/свойство | Где используется | Назначение |
|---|---|---|
| `new Application()` | конструктор `SapfirDrafter` | подключение к запущенному САПФИР |
| `Visible = 1` | конструктор | показать окно приложения |
| `GetActiveSapfirView()` | `ResolveActiveContext` | активное окно САПФИР → документ |
| `GetActiveDoc()` | `ResolveActiveContext` | активный документ (fallback) |
| `NewDocument()` | `ResolveActiveContext` | создание документа, если ничего не открыто |
| `GetDocByID(i)` | `MaterialSelectorForm` | перебор документов (1…100) при сканировании материалов |

### `SapfirDoc`

`GetApplication()`, `Title`, `CountProjects`, `GetActiveProject()`, `NewProject()`, `GetProjectByIndex(i)`, `CountStorey`, `GetStoreyByIndex(i)`, `GetActiveStorey()`.

### `AutoProject` (проект/здание)

`ID`, `CountStorey`, `NewStorey(name)`, `GetStoreyByIndex(i)`, `DelStoreyByIndex(0)` — последнее используется для удаления всех этажей перед созданием новых.

### `AutoStorey` (этаж)

- `CountModel` — число моделей на этаже;
- `NewModel((int)ModelsTypes.…)` — создание модели (стена, плита, колонна, балка, нагрузка, осевой размер, линия);
- `GetModelByIndex(i)`, `GetModelByID(id)` — поиск модели (последнее — по ID из XData);
- `Parameter["M_LEVEL"]`, `Parameter["M_HEIGHT"]` — отметка и высота этажа.

### `AutoModel` (модель элемента)

| Член | Назначение |
|---|---|
| `ID` | идентификатор модели в САПФИР — сохраняется в XData |
| `TypeModel` | числовой тип модели (сравнивается с `ModelsTypes`) |
| `Parameter[name]` | чтение/запись именованных параметров (`M_THICKNESS`, `M_MATERIAL`, `M_LOAD_1`, …) |
| `SetPosition(x, y, z)` | позиционирование (точечные нагрузки, отверстия, этажи) |
| `RegenModel()` | перегенерация модели после изменения параметров |
| `GetAxisLine()` | осевая линия модели (`AutoPolyLine`) |
| `GetMultiContour()` | мультиконтур (сечения колонн и балок) |
| `NewHole((int)ModelsTypes.TM_DOOR / TM_WINDOW / TM_HOLE)` | создание проёма/отверстия в стене или плите |

### `AutoObjDim` (осевой размер)

`SetDimParam((int)Models3dTypes.DIM_AXIS, (int)Models3dTypes.DIM_DIR_XY, marca, points)` + параметр `M_MARK` — так создаются координационные оси.

### `AutoPolyLine` / `AutoLine`

| Член | Назначение |
|---|---|
| `SetPoints(object[] coords)` | задать вершины (`x, y, z, …`) |
| `AddLine((int)Models3dTypes.TM3_LINE, object[] coords)` | добавить линейный сегмент 3D-представления |
| `GetLine(0)` | получить первый сегмент (`AutoLine`) |
| `Closed` | признак замкнутости контура |
| `AutoLine.GetPoints(ref buf)` | чтение координат сегмента обратно (в метрах) |

### `AutoMultiContour` / `AutoContour`

- `GetContour()` — получить основной контур;
- `cont.NewPolyLine()` + `SetPoints(...)` + `Closed = 0` — задать сечение колонны/балки;
- `Parameter["M_SIZE_GRO"] = 0`, `Parameter["M_SIZE_GRC"] = 0` — отключить автоматическую подгонку;
- `AddToLibPrj()` — добавить контур в библиотеку проекта.

### `AutoParametersDlg`

`DoModal()` — родной диалог параметров САПФИР, вызывается в `Sapfir_Line_LOADS` перед созданием каждой линейной нагрузки.

## Зеркальные перечисления

Чтобы не тянуть константы из COM в код, плагин объявляет собственные перечисления с теми же значениями (`SapfirDrafter.cs`):

- `ModelsTypes` — типы параметрических объектов (`TM_WALL = 10`, `TM_SLAB = 11`, `TM_COLUMN = 12`, `TM_BEAM = 13`, `TM_LINE = 18`, `TM_DIMENSION = 20`, `TM_HOLE = 30`, `TM_WINDOW = 31`, `TM_DOOR = 32`, `TM_LOAD_1 = 65557`, `TM_LOAD_2 = 131093`, `TM_SLAB_1 = 65547`, …);
- `Models3dTypes` — типы 3D-геометрии и размеров (`TM3_LINE = 1`, `DIM_AXIS = 8`, `DIM_DIR_XY = 3`, …);
- `ParameterOptionsTypes` — константы управления именованными параметрами (в текущем коде не используется).

Полные списки значений — в разделе [Структура проекта](../development/structure.md).

!!! note "О `MathServLib`"
    В старых заметках проекта упоминается COM-библиотека `MathServLib`, однако в `AcSapfir.csproj` она **не подключена** — используется только `SapfirLib`.
