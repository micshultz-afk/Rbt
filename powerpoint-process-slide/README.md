# ProcessSlide — обход слайда с поддержкой групп

Исправленная процедура `ProcessSlide` для PowerPoint 2013 VBA.

## Почему форматирование не работало на группу

`For Each shp In pptSlide.Shapes` обходит **только верхний уровень**.  
Сгруппированные объекты имеют тип `msoGroup`; текст, таблицы и диаграммы лежат в `shp.GroupItems`, а не на самом контейнере (`HasTextFrame` / `HasTable` / `HasChart` у группы обычно ложны).

Поэтому дочерние фигуры никогда не попадали в `ProcessTextInShape` / `ProcessTable` / `ProcessChart`.

**Исправление:** рекурсия `ProcessShapeTree` — при `msoGroup` обходятся все `GroupItems`, включая вложенные группы.

## Что ещё исправлено

| Было | Стало |
|---|---|
| `TraceElementEnd` вызывался только для «необработанных» фигур (все рабочие ветки делали `GoTo NextShape` мимо трейса) | `TraceElementEnd` всегда в конце `ProcessShapeTree` |
| `RemoveForcedLineBreaksInChartLabels` вызывался после `ProcessTable` | перенесён к ветке диаграммы |
| Имя файла через `ActivePresentation` (может быть не та презентация) | сначала `pptSlide.Parent.Parent.FullName` |
| Прямой доступ к `HasTextFrame` / `HasTable` / `HasChart` / `Type` | безопасные обёртки с `On Error Resume Next` |
| `IsFullyInsideSlide` на рамке группы отбрасывал всю группу | для группы проверка не делается; проверяется каждый ребёнок |
| Длинная лестница `GoTo NextShape` | короче: дерево + `FormatShapeByKind` |

## Установка

1. Скопируйте тело `ProcessSlide` / добавьте `ProcessShapeTree`, `FormatShapeByKind` и обёртки в ваш модуль форматирования (или импортируйте `ProcessSlide.bas` и перенесите `Private` → туда, где лежат зависимости).
2. Убедитесь, что зависимости из шапки `.bas` уже есть в проекте.

Макрос сам по себе не запускается — это замена внутренней процедуры обхода.
