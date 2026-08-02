# Rbt — VBA-утилиты для PowerPoint 2013

## Shape Signature

Офлайн модуль: устойчивые сигнатуры Shape и проверка совпадения.

См. [`powerpoint-shape-signature/README.md`](powerpoint-shape-signature/README.md).

```vba
sig = BuildSignatureByShapeName("MyTarget", ActivePresentation)
ok  = ShapeMatchesSignature(shp, sig)   ' True / False
```

## ProcessSlide

Обход объектов слайда с рекурсией по группам (`msoGroup` / `GroupItems`).

См. [`powerpoint-process-slide/README.md`](powerpoint-process-slide/README.md).
