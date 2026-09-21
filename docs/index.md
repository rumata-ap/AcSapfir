# AcSapfir

**AcSapfir** — плагин для AutoCAD 2020, который переносит геометрию из чертежа AutoCAD в расчётную схему **САПФИР 2018+**: координационные оси, стены, колонны, балки, плиты перекрытия и фундаментные плиты, проёмы, нагрузки и этажи.

Плагин собирается в .NET-сборку `AcSapfir.dll`, загружается в AutoCAD командой `NETLOAD` и общается с запущенным САПФИР через COM-библиотеку `SapfirLib`.

## Как это устроено

```
AutoCAD 2020 ──[команда Sapfir_*]──▶ AcSapfir.dll ──[COM: SapfirLib]──▶ САПФИР 2018+
 примитивы: отрезки, полилинии,       масштаб мм → м (×0.001)            активный этаж:
 точки, MText                        XData-трекинг ID (ACSAPFIR_ID)     модели элементов
```

1. Команда забирает выбранные примитивы AutoCAD (или использует текущий набор выбора — PickFirst).
2. Координаты пересчитываются из миллиметров в метры (`×0.001`) и передаются в САПФИР.
3. Созданный объект САПФИР получает **ID**, и этот ID записывается обратно в примитив AutoCAD через **XData** (приложение `ACSAPFIR_ID`).
4. Благодаря XData примитив «помнит» связанный объект — это используется командами `Sapfir_WALL_EDIT` и `Sapfir_SLAB_Holes`.

Все объекты создаются **в активном этаже активного проекта** САПФИР, который плагин определяет на момент вызова команды.

## Быстрые ссылки

| Раздел | Содержание |
|---|---|
| [Требования](getting-started/requirements.md) | AutoCAD 2020, САПФИР 2018+, .NET Framework 4.8 |
| [Установка и сборка](getting-started/install.md) | сборка `AcSapfir.dll` и загрузка через `NETLOAD` |
| [Быстрый старт](getting-started/quickstart.md) | первый экспорт за 5 минут |
| [Команды](commands/index.md) | сводная таблица всех команд |
| [Мастер осей](guides/axes-wizard.md) | 5 правил автоматической маркировки осей |
| [Генератор этажей](guides/storeys-generator.md) | этажи из таблицы, шаблоны «Подвал / Типовые / Чердак» |
| [Решение проблем](guides/troubleshooting.md) | типичные ошибки и их причины |
| [Внутреннее устройство](internals/architecture.md) | архитектура, XData, масштабирование, SapfirLib COM |

## Команды по группам

| Группа | Команды |
|---|---|
| Координационные оси | [`Sapfir_AXES`](commands/axes.md#sapfir_axes), [`Sapfir_AXES_WIZ`](commands/axes.md#sapfir_axes_wiz) |
| Нагрузки | [`Sapfir_Point_LOADS`](commands/loads.md#sapfir_point_loads), [`Sapfir_Line_LOADS`](commands/loads.md#sapfir_line_loads), [`Sapfir_Line_LOADS2`](commands/loads.md#sapfir_line_loads2) |
| Плиты | [`Sapfir_SLABS`](commands/slabs.md#sapfir_slabs), [`Sapfir_SLAB_Holes`](commands/slabs.md#sapfir_slab_holes), [`Sapfir_Found_SLABS`](commands/slabs.md#sapfir_found_slabs) |
| Стены | [`Sapfir_WALLS`](commands/walls.md#sapfir_walls), [`Sapfir_WALL_EDIT`](commands/walls.md#sapfir_wall_edit) |
| Проёмы | [`Sapfir_DOORS`](commands/openings.md#sapfir_doors), [`Sapfir_WINDOWS`](commands/openings.md#sapfir_windows) |
| Колонны | [`Sapfir_COLUMNS`](commands/columns.md) |
| Балки | [`Sapfir_BEAMS`](commands/beams.md) |
| Линии построения | [`Sapfir_LINES`](commands/lines.md) |
| Этажи | [`Sapfir_STOREYS`](commands/storeys.md#sapfir_storeys), [`Sapfir_STOREYS_WIZ`](commands/storeys.md#sapfir_storeys_wiz) |

!!! tip "Перед первым запуском"
    САПФИР должен быть **запущен**, а в нём открыт документ. Плагин при подключении подхватывает активное окно САПФИР (`GetActiveSapfirView()`), а если документа/проекта/этажа нет — создаёт их с именем этажа «1-й этаж».

## Что плагин не делает

- Не выполняет статический расчёт — только создаёт геометрию и нагрузки в САПФИР.
- Не импортирует обратно объекты из САПФИР в AutoCAD (кроме чтения ID и обновления геометрии по XData).
- Не работает с 3D-геометрией AutoCAD: плагин читает плоские координаты XY, высота задаётся параметрами (толщины, отметки этажей).
