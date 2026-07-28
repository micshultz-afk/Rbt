# PowerPoint Shape Signature

Офлайн VBA-модуль для PowerPoint 2013: устойчивые сигнатуры Shape и проверка совпадения.

См. [`powerpoint-shape-signature/README.md`](powerpoint-shape-signature/README.md).

Ключевой API:

```vba
sig = BuildSignatureByShapeName("MyTarget", ActivePresentation)
ok  = ShapeMatchesSignature(shp, sig)   ' True / False
```
