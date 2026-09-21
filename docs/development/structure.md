# Структура проекта

## Файлы репозитория

| Файл | Строк | Роль |
|---|---|---|
| `SapfirDrafter.cs` | ~1116 | команды AutoCAD, создание моделей САПФИР, перечисления типов |
| `AcUtilites.cs` | ~336 | выбор примитивов, итерация с транзакциями, XData, обработка `MText` |
| `AxesWizardForm.cs` | ~350 | мастер координационных осей (группировка, правила имён) |
| `StoreyGeneratorForm.cs` | ~258 | генератор этажей (таблица, шаблоны) |
| `MaterialSelectorForm.cs` | ~146 | выбор материала из проекта САПФИР |
| `Properties/AssemblyInfo.cs` | 36 | метаданные сборки |
| `AcSapfir.csproj` | 78 | проект сборки (.NET Framework 4.8) |
| `AcSapfir.sln` | — | решение Visual Studio |
| `README.md` | — | краткое описание и список команд |
| `mkdocs.yml` | — | конфигурация этого сайта документации |
| `docs/` | — | исходники документации (markdown) |

## Параметры сборки (`AcSapfir.csproj`)

| Параметр | Значение |
|---|---|
| `OutputType` | `Library` (DLL) |
| `TargetFrameworkVersion` | `v4.8` |
| `RootNamespace` / `AssemblyName` | `AcSapfir` / `AcSapfir` |
| `Platform` | `AnyCPU` |
| `Deterministic` | `true` |
| `OutputPath` | `bin\Debug\`, `bin\Release\` |

### Ссылки на API AutoCAD

| Сборка | `HintPath` | `Private` |
|---|---|---|
| `accoremgd` | `..\..\..\..\..\Program Files\Autodesk\AutoCAD 2020\accoremgd.dll` | `False` |
| `acdbmgd` | тот же путь | `False` |
| `acmgd` | тот же путь | `False` |

`HintPath` указан относительно каталога проекта, поэтому рабочим является полный путь `C:\Program Files\Autodesk\AutoCAD 2020\…`.

### COM-ссылки

```xml
<COMReference Include="SapfirLib">
  <Guid>{94863C7A-F77E-418F-9F50-3C0AD68A509A}</Guid>
  <VersionMajor>19</VersionMajor>
  <VersionMinor>0</VersionMinor>
  <WrapperTool>tlbimp</WrapperTool>
  <EmbedInteropTypes>True</EmbedInteropTypes>
</COMReference>
```

## Метаданные сборки (`AssemblyInfo.cs`)

| Атрибут | Значение |
|---|---|
| `AssemblyTitle` / `AssemblyProduct` | `AcSapfir` |
| `AssemblyVersion` / `AssemblyFileVersion` | `1.0.0.0` |
| `ComVisible` | `false` |
| `Guid` | `ffa1adba-6c12-4883-9338-436751316614` |

## Перечисления типов

### `ModelsTypes` (типы параметрических объектов)

```
TM_NONE = 0        TM_SITE = 1        TM_PROJECT = 3     TM_STOREY = 4
TM_BLOCK = 5       TM_FEAPROJECT = 6  TM_NODE = 7

TM_WALL = 10       TM_WALL_1 = 65546  TM_SLAB = 11       TM_SLAB_1 = 65547
TM_COLUMN = 12     TM_BEAM = 13       TM_BAR = 14        TM_ROOF = 15
TM_STAIRS = 16     TM_AXES = 17       TM_LINE = 18       TM_ZONE = 19
TM_DIMENSION = 20  TM_LOAD = 21       TM_MOMENT = 22     TM_CAPITAL = 23
TM_REF = 24        TM_UNDERCAP = 25   TM_TRUSS = 26

TM_LOAD_1 = 65557  TM_LOAD_2 = 131093 TM_LOAD_4 = 262165

TM_RECESS = 29     TM_HOLE = 30       TM_WINDOW = 31     TM_DOOR = 32
TM_MODVISION = 33  TM_MODVISIONSECTION = 34  TM_MODVISIONPLAN = 35
TM_MODVISION3D = 36  TM_MODVISIONCONSTRUCT = 37

TM_SPHERE = 60     TM_PIVOTAL = 61    TM_PRISM = 62      TM_PROFILE = 63
TM_CONE = 64       TM_HIPPAR = 65     TM_LIGHT = 66      TM_POLY = 67
TM_TEXT = 68       TM_CYLINDER = 69   TM_SURFACE = 70

TM_WIND = 71       TM_ARM_AREA = 72   TM_ARM_SLAB = 73   TM_KARKAS = 74
TM_PUNCH = 75      TM_ARM_WALL = 76   TM_ARM_BAR = 77    TM_ARM_ZONE = 78
TM_ARM_COLUMN = 79 TM_FEASCHEMA = 80  TM_FEABAR = 81     TM_FEASHELL = 82
TM_MULTICONT = 83  TM_OTHER = 128
```

### `Models3dTypes` (3D-геометрия и размеры)

```
TM3_LINE = 1   TM3_ARC = 2      TM3_BEZIER = 3  TM3_POLYLINE = 4  TM3_MESH3D = 5
TM3_MODEL3D = 6  TM3_TEXT3D = 7  TM3_FACE = 8   TM3_DIM3D = 9     TM3_TRIANGLE = 10
TM3_DIMENSION = 11  TM3_POLY = 12  TM3_AXIS = 13  TM3_FE_BAR = 14  TM3_FE_TRIANGLE = 15
TM3_FE_QUAD = 16  TM3_FE_TETRAEDR = 17  TM3_FE_PRYSM_3 = 18  TM3_FE_PRYSM_4 = 19
TM3_FE_SUPER = 20  TM3_ELLIPSE = 21  TM3_POINT = 22  TM_DRAFT = 100

DIM_ELEV = 1  DIM_LINEAR = 2  DIM_CHAIN = 3  DIM_RADIAL = 4  DIM_ANGULAR = 5
DIM_NOTE = 6  DIM_DIAMETR = 7  DIM_AXIS = 8  DIM_POINT = 9  DIM_ARC = 10
DIM_MARKER_CIRCLE = 11

DIM_DIR_X = 1  DIM_DIR_Y = 2  DIM_DIR_D = 3  DIM_DIR_XY = 3  DIM_DIR_Z = 4
```

### `ParameterOptionsTypes`

```
NPA_GROUP_OPEN = 65536  NPA_GROUP_CLOSE = 131072  NPA_READONLY = 262144
OBP_NAMES = 4  OBP_COMMENTS = 8  OBP_VALUES = 16
```

В текущем коде перечисление не используется — зарезервировано под работу с именованными параметрами.

## Сборка

```powershell
msbuild AcSapfir.csproj /t:Build /p:Configuration=Debug
```

Результат — `bin\Debug\AcSapfir.dll`. Подробности: [Установка и сборка](../getting-started/install.md).
