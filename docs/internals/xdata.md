# XData-связь

Плагин связывает примитивы AutoCAD с объектами САПФИР так: ID созданного объекта записывается в **расширенные данные (XData)** примитива под именем приложения `ACSAPFIR_ID`.

```csharp
internal const string XDataAppName = "ACSAPFIR_ID";
```

## Как читается и записывается

### Запись (`AcUtilites.SetSapfirId`)

```csharp
internal static void SetSapfirId(Entity ent, int sapfirId)
{
    Database db = ent.Database;
    Transaction tr = db.TransactionManager.TopTransaction;          // нужна активная транзакция

    RegAppTable regAppTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForWrite);
    if (!regAppTable.Has(XDataAppName))                             // регистрируем приложение один раз
    {
        RegAppTableRecord appRecord = new RegAppTableRecord();
        appRecord.Name = XDataAppName;
        regAppTable.Add(appRecord);
        tr.AddNewlyCreatedDBObject(appRecord, true);
    }

    ResultBuffer rb = new ResultBuffer(
        new TypedValue((int)DxfCode.ExtendedDataRegAppName, XDataAppName),
        new TypedValue((int)DxfCode.ExtendedDataInteger32, sapfirId)
    );
    ent.XData = rb;
}
```

### Чтение (`AcUtilites.GetSapfirId`)

```csharp
internal static int GetSapfirId(Entity ent)
{
    ResultBuffer rb = ent.GetXDataForApplication(XDataAppName);
    if (rb == null) return 0;
    foreach (TypedValue tv in rb.AsArray())
        if (tv.TypeCode == (int)DxfCode.ExtendedDataInteger32)
            return (int)tv.Value;
    return 0;
}
```

Возврат `0` означает «связи нет» — именно так команды отсеивают неподходящие примитивы.

## Структура данных

| Компонент | Значение |
|---|---|
| Имя приложения XData | `ACSAPFIR_ID` |
| Код значения | `ExtendedDataRegAppName` (1001) + `ExtendedDataInteger32` (1071) |
| Данные | ID модели САПФИР (`AutoModel.ID`) |

## Кто пишет и кто читает

| Команда | XData |
|---|---|
| `Sapfir_WALLS` | **пишет** ID стены |
| `Sapfir_SLABS` | **пишет** ID плиты |
| `Sapfir_Found_SLABS` | **пишет** ID фундаментной плиты |
| `Sapfir_COLUMNS` | **пишет** ID колонны |
| `Sapfir_BEAMS` | **пишет** ID балки |
| `Sapfir_LINES` | **пишет** ID линии построения |
| `Sapfir_WALL_EDIT` | **читает** ID стены и обновляет её геометрию |
| `Sapfir_SLAB_Holes` | **читает** ID плиты, выбранной отдельно |
| `Sapfir_AXES`, `Sapfir_AXES_WIZ` | не использует |
| `Sapfir_Point_LOADS`, `Sapfir_Line_LOADS`, `Sapfir_Line_LOADS2` | не использует |
| `Sapfir_DOORS`, `Sapfir_WINDOWS` | не использует (ищет стены по типу модели в активном этаже) |
| `Sapfir_STOREYS`, `Sapfir_STOREYS_WIZ` | не использует |

## Что даёт связь

- `Sapfir_WALL_EDIT` — обновление геометрии стен после правки чертежа.
- `Sapfir_SLAB_Holes` — понимание, в какую плиту добавить отверстие.
- Возможность внешней автоматизации: скрипты и надстройки могут читать ID и работать с объектами САПФИР напрямую.

## Как посмотреть XData вручную

AutoLISP (командная строка AutoCAD):

```lisp
(setq e (car (entsel)))
(entget e '("ACSAPFIR_ID"))
```

В результате будет список вида:

```
(-3 ("ACSAPFIR_ID" (1071 . 12)))
```

где `12` — ID модели в САПФИР.

Через .NET (внешняя надстройка):

```csharp
var rb = ent.GetXDataForApplication("ACSAPFIR_ID");
```

## Важные ограничения

!!! warning "XData перезаписывается целиком"
    `ent.XData = rb;` заменяет **всю** секцию XData примитива для этого приложения. Если ваша внешняя надстройка пишет что-то в `ACSAPFIR_ID`, данные плагина будут потеряны (и наоборот).

!!! note "Транзакция обязательна"
    `SetSapfirId` использует `TopTransaction`. Вызов вне активной транзакции приведёт к исключению — соблюдайте этот контракт при доработке кода.

!!! note "Копирование примитивов сохраняет XData"
    `COPY`/`MIRROR`/`ARRAY` копируют XData вместе с геометрией, поэтому копии будут «указывать» на тот же объект САПФИР. Учитывайте это при `Sapfir_WALL_EDIT`: две копии могут править одну и ту же стену.
