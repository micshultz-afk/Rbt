# ReplaceMultipleSpaces

PowerPoint VBA: любое `"  "+` → `" "` в `TextRange`, с сохранением формата между Runs.

```vba
n = CollapseMultiSpaces(shp.TextFrame.TextRange)   ' удалено; 0; -1 = ошибка
ReplaceMultipleSpacesDashes rng, slideIndex, shp   ' обёртка с логом проекта
```

## Почему старый код сбоил

| Баг | Эффект |
|---|---|
| `text` не обновлялся в цикле | до 100 пустых `Replace` |
| `cnt` общий, лимит 100 | хвост Runs не обрабатывался |
| правка только внутри Run | двойной пробел на стыке Runs оставался |
| обход `Runs` вперёд при мутации | пропуски из‑за перестройки коллекции |
| `g_ChangedElements` всегда +1 | ложная статистика |

## Как сделано

1. Нет `"  "` → сразу выход.
2. Один Run → строковый `Replace$` (быстро, формат однороден).
3. Несколько Runs → схлопывание внутри Run **с конца**, затем `TextRange.Replace` на стыках.
4. Выход из COM-цикла по `Len` (в PPT `Replace` может вернуть пустой диапазон, не `Nothing`).
5. При ошибке `Runs.Count` — `Fail` (−1), а не присвоение `rng.Text` (оно убило бы формат).

Обёртка требует символы проекта: `bDebugMode`, `TraceStart`, `TraceEnd`, `LogLine`, `ElapsedSeconds`, `g_ChangedElements`.

Проверка: `python3 text-cleanup/test_collapse_spaces.py`
