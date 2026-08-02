# ProcessSlide — обход слайда с поддержкой групп

Минимальная замена `ProcessSlide` для PowerPoint 2013 VBA.

## Суть фикса

`pptSlide.Shapes` — только верхний уровень. Внутри `msoGroup` объекты в `GroupItems`.  
`ProcessShapeTree` рекурсивно обходит группы (включая вложенные).

## Прочее

- `TraceElementEnd` всегда в конце
- `RemoveForcedLineBreaksInChartLabels` только у диаграмм
- имя файла из `pptSlide.Parent.Parent`
- `IsFullyInsideSlide` не на контейнере группы

Скопируйте обе процедуры в модуль с зависимостями или импортируйте `.bas`.
