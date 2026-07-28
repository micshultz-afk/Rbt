# ShapeSignature — PowerPoint 2013 (offline / Secret Net)

Устойчивые компактные сигнатуры Shape + проверка `True`/`False`.  
Работает **без Интернета**, **без CreateObject / Scripting / MSForms**, только объектная модель PowerPoint и чистый VBA.

## Новая функция: сигнатура по имени

```vba
Public Function BuildSignatureByShapeName( _
    ByVal shapeName As String, _
    ByRef pres As Presentation) As String
```

1. Находит **все** фигуры с именем `shapeName` во всей презентации (включая группы и повторы на других слайдах).
2. Берёт только **стабильные** между этими экземплярами признаки.
3. Отделяет их от остальных объектов презентации.
4. Возвращает сжатую сигнатуру `S2|...` (или `""` при ошибке/если не найдено).

Распознавание потом — без имени:

```vba
Dim ok As Boolean
Dim shp As Shape
Dim sig As String
Dim pres As Presentation

Set pres = ActivePresentation
sig = BuildSignatureByShapeName("MyTarget", pres)
Set shp = ActiveWindow.Selection.ShapeRange(1)
ok = ShapeMatchesSignature(shp, sig)   ' True / False
```

UI-обёртка: макрос `BuildSignatureByShapeNameUI`.

## Почему повторы на слайдах обрабатываются правильно

Имя одно, позиций может быть несколько:

| Признак | Если экземпляры на слайдах… | Поведение |
|---------|-----------------------------|-----------|
| Type / Auto / Size / Caps / Fill / Font / Text | совпадают | входят в сигнатуру |
| **Pos (Left/Top)** | различаются | **исключается** автоматически |
| Pos | совпадают (один шаблон) | может войти для большей уникальности |

Позиция и размер разделены (`F_POS` / `F_SIZE`): объект «тот же по размеру, но в другом месте» остаётся распознаваемым.

## Распознавание

```vba
Public Function ShapeMatchesSignature(ByRef shp As Shape, ByVal sig As String) As Boolean
```

Только 2 аргумента. Ошибки глотаются → `False` (без вылетов).

## Сжатие сигнатуры (S2)

Старый текстовый `S1` мог приближаться к ~90 символам на полном наборе.

Новый формат **`S2|<base64url>`**:

- бинарная упаковка полей + CRC24
- типичная длина **~25–45** символов
- полный набор признаков **&lt; 100** символов (обычно ~60–70)
- без потери точности признаков
- legacy `S1|...` ещё читается (best-effort)

Кодирование — чистый VBA Base64URL, без MSXML/интернета.

## Среда Secret Net / offline

| Используется | Не используется |
|--------------|-----------------|
| PowerPoint OM (`Shape`, `Presentation`, `FileDialog`) | `CreateObject` |
| VBA `Collection`, массивы `Byte` | `Scripting.Dictionary` |
| Собственный Base64/FNV | MSForms clipboard / MSXML / WinHTTP |

`Application.FileDialog` — штатный Office UI (не сеть). Если запрещён политикой, `BuildUniqueSignatureFromFiles` продолжит работу по активной презентации/выделению.

## Установка

1. `Alt+F11` → Import `ShapeSignature.bas` (+ опционально `ShapeSignatureExamples.bas`)
2. Сохранить как `.pptm`
3. Задать фигурам одно имя (область выделения) на нужных слайдах
4. `sig = BuildSignatureByShapeName("MyTarget", ActivePresentation)`

## Точность извлечения признаков

Берутся устойчивые свойства OM:

- тип / AutoShape / placeholder
- относительная геометрия (доли слайда, points → независимо от DPI/ОС/размера листа)
- заливка: RGB и/или scheme color (тема стабильнее «сырого» RGB)
- линия, caps (text/table/chart/group)
- текст: нормализация + FNV-1a; fallback `TextFrame2`
- шрифт: имя (hash) + относительный размер

**Не** используются: `Id`, экранные пиксели, интернет, хеш пикселей картинки (в VBA 2013 это хрупко и медленно).

Два визуально идентичных объекта без отличимых свойств OM физически неразделимы — алгоритм честно сообщит о коллизиях в отчёте UI.

## Файлы

- `ShapeSignature.bas` — ядро
- `ShapeSignatureExamples.bas` — примеры
- `test_signature_helpers.py` — unit-тесты S2 (на машине разработки)
