# Архитектура

## Состав сборки

Сборка `AcSapfir.dll` (namespace `AcSapfir`) содержит четыре класса и три перечисления:

| Файл | Тип | Назначение |
|---|---|---|
| `SapfirDrafter.cs` | класс `SapfirDrafter` | все команды AutoCAD и методы создания объектов САПФИР |
| `AcUtilites.cs` | класс `AcUtilites` (static) | выбор примитивов, итерация по ним, работа с XData |
| `AxesWizardForm.cs` | класс `AxesWizardForm` | мастер координационных осей |
| `StoreyGeneratorForm.cs` | класс `StoreyGeneratorForm` | генератор этажей |
| `MaterialSelectorForm.cs` | класс `MaterialSelectorForm` | выбор материала |
| `SapfirDrafter.cs` | enum `ModelsTypes`, `Models3dTypes` | числовые типы объектов и 3D-моделей САПФИР |
| `SapfirDrafter.cs` | enum `ParameterOptionsTypes` | константы управления именованными параметрами (в текущем коде не используется) |

Регистрация команд — атрибутом уровня сборки:

```csharp
[assembly: CommandClass(typeof(AcSapfir.SapfirDrafter))]
```

## Поток выполнения команды

```
AutoCAD: пользователь вызывает Sapfir_XXX
    │
    ├─ создаётся новый экземпляр SapfirDrafter  (конструктор)
    │     ├─ new SapfirLib.Application(); Visible = 1
    │     ├─ GetActiveSapfirView().GetDocument()  →  fallback GetActiveDoc()  →  fallback NewDocument()
    │     └─ активный проект / активный этаж (создаются при отсутствии)
    │
    ├─ ResolveActiveContext()   ← повторное определение активного контекста
    │
    ├─ AcUtilites.Selection()   ← PickFirst или интерактивный выбор
    │
    ├─ AcUtilites.ActionOn*(...)  ← транзакция AutoCAD + вызов делегата на каждый примитив
    │     └─ Create*(...)         ← создание модели САПФИР (COM) + запись XData-ID
    │
    └─ Editor.WriteMessage(...)   ← отчёт в командную строку
```

### Жизненный цикл экземпляра

AutoCAD создаёт **новый экземпляр** `SapfirDrafter` на каждый вызов команды. Поэтому:

- подключение к САПФИР и определение контекста происходят при каждом запуске команды;
- состояние между вызовами не сохраняется (кроме `storeySpf`/`docSpf` внутри одного вызова);
- поля класса (`appSpf`, `docSpf`, `projSpf`, `storeySpf`, `paramModelSpf`, `polylineSpf`, `buf1`, `selectedMaterialGuid`) — по сути рабочие переменные одного запуска.

## Работа с COM САПФИР

- `SapfirLib.Application` — корневой COM-объект; в конструкторе выставляется `Visible = 1`, чтобы окно САПФИР отображалось.
- Определение документа: активное окно (`GetActiveSapfirView()` → `GetDocument()`), затем активный документ (`GetActiveDoc()`), затем создание нового (`NewDocument()`).
- Передача координат в COM выполняется через `object[]` (поле `buf1`): плоский массив `x, y, z, x, y, z, …`; значения — `double`, поэтому используется явное приведение и распаковка (`(double)(object)...`).
- Модели создаются через `storeySpf.NewModel((int)ModelsTypes.TM_*)`, отверстия — `model.NewHole((int)ModelsTypes.TM_*)`.
- Параметры задаются индексатором: `model.Parameter["M_THICKNESS"] = t;`
- После заполнения параметров обязательно вызывается `RegenModel()`.
- Часть данных читается обратно: `polylineSpf.GetLine(0)`, `spfLine.GetPoints(ref buf1)`, `multiContour.GetContour()`, `model.TypeModel`, `model.Parameter["M_GUID"]`.

## Транзакции AutoCAD

Все операции с примитивами выполняются внутри транзакции, которую открывает `AcUtilites`:

```csharp
using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
{
    foreach (ObjectId acObjId in objIds)
    {
        Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
        if (acEnt is Line line)
            action(line, parameter);
    }
    acTrans.Commit();
}
```

Важные следствия:

- примитивы открываются **на запись** (`ForWrite`) — это нужно, чтобы записать XData с ID;
- `SetSapfirId()` опирается на `TopTransaction`, поэтому вызывать его разрешено только внутри активной транзакции;
- команды, создающие объекты за пределами транзакции (`Sapfir_AXES_WIZ`, `Sapfir_WINDOWS`, `Sapfir_COLUMNS`, `Sapfir_BEAMS`), открывают транзакцию самостоятельно.

## Модальные диалоги

Формы Windows Forms открываются через API AutoCAD:

```csharp
Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form);
```

Такой вызов корректно блокирует интерфейс AutoCAD на время работы диалога. Команда `Sapfir_STOREYS_WIZ` помечена `CommandFlags.Session`, чтобы диалог работал вне контекста документа.

## Обработка ошибок

| Место | Поведение |
|---|---|
| Конструктор `SapfirDrafter` | ловит исключение COM, печатает «Ошибка подключения к САПФИР: …» и **пробрасывает** исключение дальше |
| `Sapfir_WALL_EDIT` | тело цикла в `try { … } catch { }` — ошибки скрываются |
| `MaterialSelectorForm.LoadMaterials` | все ошибки чтения параметров скрываются (`catch { }`), производится переход на встроенный список |
| `StoreyGeneratorForm.UpdateStatus` | при недоступности САПФИР показывает «Sapfir: недоступен» |

## Схема взаимодействия компонентов

```
┌────────────────────────────┐        ┌──────────────────────────────┐
│ SapfirDrafter              │        │ Формы (WinForms)             │
│  • [CommandMethod] команды │◀──────▶│  • AxesWizardForm            │
│  • Create* методы          │        │  • StoreyGeneratorForm       │
│  • ResolveActiveContext    │        │  • MaterialSelectorForm      │
└───────────┬────────────────┘        └──────────────────────────────┘
            │
            ├──▶ AcUtilites: Selection / ActionOnLines / ActionOnPolylines /
            │                ActionOnPoints / GetMTextsData / XData
            │
            └──▶ SapfirLib (COM): Application → SapfirDoc → AutoProject → AutoStorey → AutoModel
```

## Смотрите также

- [XData-связь](xdata.md) — как примитивы «помнят» объекты САПФИР.
- [Масштабирование координат](scaling.md) — правила перевода мм → м.
- [SapfirLib COM](sapfir-lib.md) — перечень используемых методов.
