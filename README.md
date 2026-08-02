# Rbt — PowerPoint VBA utilities

Офлайн VBA-модули для PowerPoint.

## Модули

### [`powerpoint-shape-signature/`](powerpoint-shape-signature/README.md)

Устойчивые сигнатуры Shape и проверка совпадения.

```vba
sig = BuildSignatureByShapeName("MyTarget", ActivePresentation)
ok  = ShapeMatchesSignature(shp, sig)   ' True / False
```

### [`text-cleanup/`](text-cleanup/README.md)

Надёжная замена множественных пробелов на один в любом `TextRange`.

```vba
n = CollapseMultiSpaces(shp.TextFrame.TextRange)   ' лишних удалено; -1 = ошибка
```
