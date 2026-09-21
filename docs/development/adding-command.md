# Adding a command / Добавление новой команды

Пошаговый рецепт: как добавить в плагин новую команду экспорта.

## 1. Определите контракт

Опишите в двух строках:

- какой примитив AutoCAD служит входом (`Line`, `Polyline`, `DBPoint`, `MText`);
- какой тип модели САПФИР создаётся (`ModelsTypes.TM_*`);
- какие данные вводит пользователь (толщина, высота, значение нагрузки);
- нужен ли материал и нужна ли XData-связь (почти всегда — да).

## 2. Добавьте метод создания модели

В `SapfirDrafter.cs`, в секции «Создание конструктивных элементов Sapfir»:

```csharp
/// <summary>
/// Создаёт <что-то> в Sapfir на основе <примитива> AutoCAD.
/// Координаты масштабируются из миллиметров в метры (×0.001).
/// </summary>
/// <param name="line">Примитив AutoCAD.</param>
/// <param name="t">Параметр в метрах.</param>
void CreateMyElement(Line line, double t)
{
    paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_WALL);   // выберите нужный тип

    polylineSpf = paramModelSpf.GetAxisLine();
    buf1 = new object[]
    {
        line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
        line.EndPoint.X   * 0.001, line.EndPoint.Y   * 0.001, 0
    };
    polylineSpf.SetPoints(buf1);

    paramModelSpf.Parameter["M_THICKNESS"] = t;
    if (selectedMaterialGuid != null)
        paramModelSpf.Parameter["M_MATERIAL"] = selectedMaterialGuid;

    paramModelSpf.RegenModel();

    AcUtilites.SetSapfirId(line, paramModelSpf.ID);      // XData-связь (внутри транзакции!)
    AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
        "\n  Мой элемент ID={0}", paramModelSpf.ID);
}
```

Конвенции этого слоя:

- имя метода — `Create<Element>`;
- первым аргументом всегда примитив AutoCAD, затем параметры;
- координаты — только через `× 0.001` ([подробнее](../internals/scaling.md));
- после заполнения параметров — `RegenModel()`;
- затем `SetSapfirId` и сообщение с ID.

## 3. Добавьте саму команду

В секции «Команды AutoCAD»:

```csharp
[CommandMethod("Sapfir_MY_CMD", CommandFlags.UsePickSet)]
public void Sapfir_MY_CMD()
{
    ResolveActiveContext();
    selectedMaterialGuid = SelectMaterial("Мои элементы");   // если нужен материал

    Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
    PromptDoubleResult result = acDocEd.GetDouble("\nВведите толщину в метрах: ");
    if (result.Status != PromptStatus.OK) return;

    AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateMyElement, result.Value);
}
```

| Требование | Зачем |
|---|---|
| `CommandFlags.UsePickSet` | поддержка PickFirst (объекты, выбранные до вызова команды) |
| `ResolveActiveContext()` | актуализация активного документа/проекта/этажа САПФИР |
| Проверка `PromptStatus.OK` | корректная отмена по `Esc` |
| Никаких исключений наружу | ошибки COM печатаются через `Editor.WriteMessage` |

## 4. Если нужен новый способ итерации по примитивам

`AcUtilites` уже умеет: `ActionOnPoints`, `ActionOnLines` (3 перегрузки), `ActionOnPolylines` (2 перегрузки). Если нужен другой тип — добавьте перегрузку по образцу:

```csharp
public static void ActionOnCircles(ObjectId[] objIds, Action<Circle, double> action, double parameter)
{
    if (objIds == null) return;
    Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
    {
        foreach (ObjectId acObjId in objIds)
        {
            Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
            if (acEnt is Circle circle)
                action(circle, parameter);
        }
        acTrans.Commit();
    }
}
```

!!! warning "Открывайте примитивы на запись"
    `OpenMode.ForWrite` обязателен, если внутри действия вызывается `SetSapfirId` (запись XData). Использование `ForRead` приведёт к исключению.

## 5. Если добавлен новый файл — включите его в проект

```xml
<ItemGroup>
  <Compile Include="MyNewForm.cs" />
</ItemGroup>
```

## 6. Соберите и проверьте

```powershell
msbuild AcSapfir.csproj /t:Build /p:Configuration=Debug
```

Проверка в AutoCAD: `NETLOAD` → вызвать команду → убедиться, что:

- объект появился в **активном** этаже САПФИР;
- размеры соответствуют ожидаемым в метрах;
- в примитиве появился XData `ACSAPFIR_ID` (проверка через `(entget e '("ACSAPFIR_ID"))`);
- команда переживает отмену (`Esc`) на любом этапе.

## 7. Опишите новую команду в документации

1. Добавьте строку в таблицу [Команды](../commands/index.md).
2. Создайте/дополните страницу в `docs/commands/`.
3. Если команда попадает в группу существующего раздела — допишите `##`-секцию и якорь для ссылок.

## Чек-лист ревью

- [ ] команда помечена `CommandFlags.UsePickSet`;
- [ ] вызывается `ResolveActiveContext()`;
- [ ] все координаты умножены на `0.001`;
- [ ] параметры заданы до `RegenModel()`;
- [ ] XData-ID записан, если элемент должен участвовать в обновлениях;
- [ ] сообщения в командную строку понятны и содержат ID;
- [ ] нет «тихих» `catch { }` без комментария, поясняющего причину;
- [ ] документация обновлена.
