# Sapfir_STOREYS — команда создания этажей из MText

## Цель
Новая команда AutoCAD `Sapfir_STOREYS`, которая создаёт набор этажей в Sapfir на основе выбранных многострочных текстов (MText). Y-координата точки вставки MText в пространстве модели является абсолютной отметкой верха этажа в миллиметрах, а текст MText — именем этажа.

## Контекст
- Проект: AcSapfir — AutoCAD 2020 plugin для экспорта геометрии в Sapfir
- Язык: C# (.NET Framework 4.7)
- Зависимости: AutoCAD API, SapfirLib COM

## Решённые проектные решения
1. **Координата высоты:** Sapfir использует Z как высотную координату. Y-координата MText в AutoCAD переводится в Z-координату Sapfir с масштабом 0.001 (мм → м).
2. **Удаление существующих этажей:** Перед созданием новых все существующие этажи удаляются.
3. **Конструктор SapfirDrafter:** остаётся без изменений (создаёт один этаж по умолчанию).
4. **Активный этаж после команды:** storeySpf не обновляется. Пользователь выбирает нужный этаж вручную в Sapfir.

## Изменения в коде

### AcUtilites.cs — новый метод
```csharp
public static List<Tuple<string, double>> GetMTextsData(ObjectId[] objIds)
```
Открывает транзакцию, фильтрует объекты типа `MText`, возвращает список кортежей (текст, Y-координата). Паттерн — как у существующих `ActionOn*`.

### SapfirDrafter.cs — новый метод
```csharp
void CreateStoreys(List<Tuple<string, double>> storeyData)
```
- Удаляет все этажи: цикл `projSpf.DelStoreyByIndex(0)` пока `projSpf.CountStorey > 0`
- Сортирует `storeyData` по Y по возрастанию
- Для каждой пары: `storeySpf = projSpf.NewStorey(name)` → `storeySpf.SetPosition(0, 0, y * 0.001)`

### SapfirDrafter.cs — новая команда
```csharp
[CommandMethod("Sapfir_STOREYS", CommandFlags.UsePickSet)]
public void Sapfir_STOREYS()
```
- Вызывает `AcUtilites.Selection()` для получения ObjectId[]
- Вызывает `AcUtilites.GetMTextsData(objIds)` для получения данных
- Вызывает `CreateStoreys(data)`

### SapfirDrafter.cs — конструктор
Без изменений. Остаётся создание одного этажа по умолчанию.

## Обработка текста MText
Используется свойство `MText.PlainText()` (доступно в AutoCAD .NET API) — возвращает текст без форматирования MText (убирает коды типа `\P`, `{\...}` и т.д.). Если `PlainText()` недоступен, используется `MText.Text` как fallback.

## Сортировка
Этажи сортируются по Y (отрицательная Y = нижний этаж, положительная = верхний). Порядок: по возрастанию Y.

## Масштабирование
Y-координата MText (мм) → Z-координата SetPosition (м): `z = Y * 0.001`

## Существующие команды — влияние
Нет. Конструктор не меняется, другие команды не затрагиваются. Команда `Sapfir_STOREYS` полностью автономна.