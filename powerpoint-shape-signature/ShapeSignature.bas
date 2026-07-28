Attribute VB_Name = "ShapeSignature"
'==============================================================================
' ShapeSignature — PowerPoint 2013 VBA
' Обучаемые сигнатуры текстовых блоков (колонтитулы, Label и любые другие).
'
' Среда: без Интернета, без CreateObject / Scripting / MSForms / внешних DLL.
' В PowerPoint нет ThisPresentation — все методы принимают путь или объект явно.
'
' ОБУЧЕНИЕ (опыт хранится внутри строки-сигнатуры):
'   sig = TrainSignatureFromFiles(badFiles, badName, goodFiles, goodName)
'     badFiles  — пути к файлам с НЕПРАВИЛЬНЫМ форматированием через ";"
'     goodFiles — пути к файлам с ПРАВИЛЬНЫМ форматированием через ";"
'     badName / goodName — имя фигуры (одинаковое на нужных слайдах)
'
'   Оба набора дают признаки РОЛИ объекта (что он есть).
'   Правильные файлы дополнительно дают ЭТАЛОН ФОРМАТА (как должно быть).
'   Признаки, которые «плавают» между примерами, автоматически отключаются.
'
' РАСПОЗНАВАНИЕ:
'   pct = ShapeSimilarityPercent(shp, sig)      ' 0..100
'   ok  = ShapeMatchesSignature(shp, sig)       ' True/False по порогу из сигнатуры
'   txt = ShapeSimilarityExplain(shp, sig)      ' понятное объяснение
'   dev = DescribeFormatDeviations(shp, sig)    ' что не так с форматированием
'==============================================================================
Option Explicit

'----- Формат сигнатуры -------------------------------------------------------
Private Const SIG_VER As String = "S3"
Private Const SIG_VER_NUM As Long = 3

' Геометрия в сигнатуре: 255 = 100% стороны слайда (шаг ~0.4%)
Private Const G8 As Long = 255

'----- Биты активных признаков роли -------------------------------------------
Private Const R_ZONE As Long = 1        ' зона центра фигуры
Private Const R_SIZE As Long = 2        ' полоса размеров
Private Const R_AR As Long = 4          ' пропорции
Private Const R_TYPE As Long = 8        ' множество типов фигур
Private Const R_PH As Long = 16         ' множество типов placeholder
Private Const R_FLAGS As Long = 32      ' обязательные признаки-флаги
Private Const R_TXT As Long = 64        ' структура текста
Private Const R_FONTREL As Long = 128   ' относительный кегль
Private Const R_ROT As Long = 256       ' поворот
Private Const R_ALL As Long = 511

'----- Биты флагов наблюдения -------------------------------------------------
Private Const FL_TEXTFRAME As Long = 1
Private Const FL_HASTEXT As Long = 2
Private Const FL_FILL As Long = 4
Private Const FL_BOLD As Long = 8
Private Const FL_COLORFONT As Long = 16
Private Const FL_PLACEHOLDER As Long = 32
Private Const FL_DIGITS As Long = 64
Private Const FL_MULTIPARA As Long = 128

'----- Базовые веса признаков -------------------------------------------------
Private Const W_ZONE As Double = 26
Private Const W_SIZE As Double = 18
Private Const W_TYPE As Double = 14
Private Const W_FLAGS As Double = 10
Private Const W_AR As Double = 8
Private Const W_PH As Double = 8
Private Const W_TXT As Double = 8
Private Const W_FONTREL As Double = 4
Private Const W_ROT As Double = 4

'----- Пределы, после которых признак считается нестабильным ------------------
Private Const LIM_ZONE_H As Long = 96
Private Const LIM_SIZE_H As Long = 64
Private Const LIM_AR_H As Long = 40
Private Const LIM_FONT_H As Long = 40
Private Const LIM_PARA_H As Long = 30
Private Const LIM_ROT_H As Long = 20
Private Const LIM_TYPE_COUNT As Long = 6

'----- Запасы мягкого затухания ----------------------------------------------
Private Const SLACK_ZONE As Long = 32
Private Const SLACK_SIZE As Long = 26
Private Const SLACK_AR As Long = 16
Private Const SLACK_FONT As Long = 24
Private Const SLACK_PARA As Long = 4
Private Const SLACK_ROT As Long = 10

Private Const MAX_OBS As Long = 500
Private Const TAG_PREFIX As String = "SHAPESIG.CLASS."

'==============================================================================
' ТИПЫ
'==============================================================================
Private Type ShapeFeat
    Valid As Boolean
    Deep As Boolean
    Blocked As Boolean          ' таблица/диаграмма/картинка/группа/OLE
    TypeId As Long
    PhType As Long              ' -1 если не placeholder
    Cx As Long                  ' 0..255 центр по X
    Cy As Long
    Wd As Long                  ' 0..255 ширина
    Ht As Long
    AR8 As Long                 ' (W/H)*8, максимум 255
    Rot2 As Long                ' градусы/2, 0..180
    Flags As Long
    Paras As Long               ' 1..255
    LenBucket As Long           ' 0..7
    FontRel As Long             ' кегль/высота слайда*1000, максимум 255
    FontFnv As Long
    FillBucket As Long          ' 0..7
    FontBucket As Long          ' 0..7
    Align As Long               ' 0..7
End Type

Private Type Rng
    HasVal As Boolean
    MinV As Long
    MaxV As Long
End Type

Private Type Acc
    n As Long
    Cx As Rng
    Cy As Rng
    Wd As Rng
    Ht As Rng
    AR As Rng
    Paras As Rng
    FontRel As Rng
    Rot As Rng
    TypeSet As Long
    PhSet As Long
    LenSet As Long
    FlagsSeen As Long
    FlagsAll As Long
    FlagsInit As Boolean
    FontFnv As Long
    FontFnvSame As Boolean
    FontFnvInit As Boolean
    FontBucketSet As Long
    FillBucketSet As Long
    AlignSet As Long
    StyleFlags As Long
End Type

Private Type RoleSig
    Valid As Boolean
    Mask As Long
    nBad As Long
    nGood As Long
    Thr As Long
    CxC As Long
    CxH As Long
    CyC As Long
    CyH As Long
    WC As Long
    WH As Long
    HC As Long
    HH As Long
    ARC As Long
    ARH As Long
    TypeSet As Long
    PhSet As Long
    FlagsSeen As Long
    FlagsAll As Long
    ParaC As Long
    ParaH As Long
    LenSet As Long
    FontRelC As Long
    FontRelH As Long
    RotC As Long
    RotH As Long
    HasFormat As Boolean
    FmtFontFnv As Long
    FmtFontRelC As Long
    FmtFontRelH As Long
    FmtFontBucketSet As Long
    FmtFillBucketSet As Long
    FmtStyleFlags As Long
    FmtAlignSet As Long
    FmtCxC As Long
    FmtCxH As Long
    FmtCyC As Long
    FmtCyH As Long
End Type

'----- Кеш разбора сигнатуры (ускоряет циклы поиска) -------------------------
Private m_cacheSig As String
Private m_cacheRole As RoleSig
Private m_lastReport As String

'==============================================================================
' ПУБЛИЧНОЕ API — ОБУЧЕНИЕ
'==============================================================================

'--- Основной метод обучения --------------------------------------------------
' badFiles  : "C:\a.pptx;C:\b.pptx"  — неправильное форматирование
' badName   : имя фигуры в этих файлах
' goodFiles : "C:\ok1.pptx;C:\ok2.pptx" — правильное форматирование (может быть "")
' goodName  : имя фигуры в правильных файлах (если "" — берётся badName)
' Возврат   : строка-сигнатура с накопленным опытом ("" при неудаче)
Public Function TrainSignatureFromFiles(ByVal badFiles As String, _
                                        ByVal badName As String, _
                                        Optional ByVal goodFiles As String = "", _
                                        Optional ByVal goodName As String = "") As String
    Dim obs() As ShapeFeat
    Dim nObs As Long
    Dim accAll As Acc
    Dim accGood As Acc
    Dim r As RoleSig
    Dim report As String
    Dim opened As Collection
    Dim nBad As Long
    Dim nGood As Long
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo Fail
    TrainSignatureFromFiles = ""
    m_lastReport = ""
    Set opened = New Collection
    ReDim obs(1 To MAX_OBS)
    nObs = 0

    badName = Trim$(badName)
    goodName = Trim$(goodName)
    If Len(goodName) = 0 Then goodName = badName
    If Len(badName) = 0 And Len(goodName) = 0 Then
        m_lastReport = "Не задано имя фигуры."
        Exit Function
    End If

    AccInit accAll
    AccInit accGood

    ' 1) Файлы с неправильным форматированием — только признаки роли
    nBad = LearnFromFileList(badFiles, badName, accAll, accGood, False, obs, nObs, opened, report)

    ' 2) Файлы с правильным форматированием — роль + эталон формата
    nGood = LearnFromFileList(goodFiles, goodName, accAll, accGood, True, obs, nObs, opened, report)

    CloseOpened opened

    If accAll.n = 0 Then
        m_lastReport = "Ни одной фигуры с указанным именем не найдено." & vbCrLf & report
        Exit Function
    End If

    BuildRoleFromAcc accAll, accGood, nBad, nGood, r, report
    If Not r.Valid Then
        m_lastReport = "Не удалось выделить устойчивые признаки." & vbCrLf & report
        Exit Function
    End If

    r.Thr = ComputeThreshold(obs, nObs, r)
    TrainSignatureFromFiles = EncodeRoleSig(r)

    m_lastReport = report & _
        "Примеров: неправильных=" & CStr(nBad) & ", правильных=" & CStr(nGood) & vbCrLf & _
        "Активные признаки: " & MaskDescription(r.Mask) & vbCrLf & _
        "Эталон формата: " & IIf(r.HasFormat, "есть", "нет") & vbCrLf & _
        "Порог совпадения: " & CStr(r.Thr) & "%" & vbCrLf & _
        "Длина сигнатуры: " & CStr(Len(TrainSignatureFromFiles)) & " символов"
    Exit Function
Fail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    CloseOpened opened
    On Error GoTo 0
    TrainSignatureFromFiles = ""
    m_lastReport = "Ошибка обучения (" & CStr(errNum) & "): " & errDesc
End Function

'--- Отчёт последнего обучения ------------------------------------------------
Public Function LastTrainReport() As String
    LastTrainReport = m_lastReport
End Function

'--- Обучение с диалогом ------------------------------------------------------
Public Sub TrainSignatureUI()
    Dim badFiles As String
    Dim goodFiles As String
    Dim nm As String
    Dim sig As String

    On Error GoTo Fail
    nm = InputBox("Имя фигуры (одинаковое на нужных слайдах):", "Обучение")
    If Len(Trim$(nm)) = 0 Then Exit Sub

    badFiles = InputBox("Пути к файлам с НЕПРАВИЛЬНЫМ форматированием (через ;):", "Обучение")
    goodFiles = InputBox("Пути к файлам с ПРАВИЛЬНЫМ форматированием (через ;), можно пусто:", "Обучение")

    sig = TrainSignatureFromFiles(badFiles, nm, goodFiles, nm)
    If Len(sig) = 0 Then
        MsgBox "Сигнатура не построена." & vbCrLf & vbCrLf & LastTrainReport(), vbExclamation
        Exit Sub
    End If
    InputBox LastTrainReport() & vbCrLf & vbCrLf & "Сигнатура:", "Обучение завершено", sig
    Exit Sub
Fail:
    MsgBox "Ошибка TrainSignatureUI: " & Err.Description, vbCritical
End Sub

'--- Обучение по уже открытой презентации (без путей) ------------------------
Public Function BuildSignatureByShapeName(ByVal shapeName As String, _
                                          ByRef pres As Presentation) As String
    Dim obs() As ShapeFeat
    Dim nObs As Long
    Dim accAll As Acc
    Dim accGood As Acc
    Dim r As RoleSig
    Dim report As String

    On Error GoTo SoftFail
    BuildSignatureByShapeName = ""
    m_lastReport = ""
    If pres Is Nothing Then Exit Function
    shapeName = Trim$(shapeName)
    If Len(shapeName) = 0 Then Exit Function

    ReDim obs(1 To MAX_OBS)
    nObs = 0
    AccInit accAll
    AccInit accGood

    If LearnFromPresentation(pres, shapeName, accAll, accGood, False, obs, nObs) = 0 Then
        m_lastReport = "Фигуры с именем «" & shapeName & "» не найдены."
        Exit Function
    End If

    BuildRoleFromAcc accAll, accGood, accAll.n, 0, r, report
    If Not r.Valid Then Exit Function
    r.Thr = ComputeThreshold(obs, nObs, r)
    BuildSignatureByShapeName = EncodeRoleSig(r)
    m_lastReport = report & "Примеров: " & CStr(accAll.n) & ", порог " & CStr(r.Thr) & "%"
    Exit Function
SoftFail:
    BuildSignatureByShapeName = ""
    m_lastReport = "Ошибка BuildSignatureByShapeName: " & Err.Description
End Function

'==============================================================================
' ПУБЛИЧНОЕ API — РАСПОЗНАВАНИЕ
'==============================================================================

'--- Похожесть в процентах: ByRef Shape + сигнатура → 0..100 -----------------
Public Function ShapeSimilarityPercent(ByRef shp As Shape, ByVal sig As String) As Long
    Dim r As RoleSig
    Dim f As ShapeFeat
    Dim sc As Double
    Dim expl As String
    Dim gate As String

    On Error GoTo SoftFail
    ShapeSimilarityPercent = 0
    If shp Is Nothing Then Exit Function
    If Not GetRole(sig, r) Then Exit Function

    If Not ExtractCheap(shp, f) Then Exit Function
    If Not PassGates(f, r, gate) Then Exit Function
    ExtractDeep shp, f

    sc = ScoreFeat(f, r, expl)
    If sc < 0# Then sc = 0#
    If sc > 100# Then sc = 100#
    ShapeSimilarityPercent = CLng(Round(sc, 0))
    Exit Function
SoftFail:
    ShapeSimilarityPercent = 0
End Function

'--- True/False по порогу, записанному в сигнатуре ---------------------------
Public Function ShapeMatchesSignature(ByRef shp As Shape, ByVal sig As String) As Boolean
    Dim r As RoleSig
    Dim pct As Long
    On Error GoTo SoftFail
    ShapeMatchesSignature = False
    If shp Is Nothing Then Exit Function
    If Not GetRole(sig, r) Then Exit Function
    pct = ShapeSimilarityPercent(shp, sig)
    ShapeMatchesSignature = (pct >= r.Thr)
    Exit Function
SoftFail:
    ShapeMatchesSignature = False
End Function

'--- Понятное объяснение результата ------------------------------------------
Public Function ShapeSimilarityExplain(ByRef shp As Shape, ByVal sig As String) As String
    Dim r As RoleSig
    Dim f As ShapeFeat
    Dim sc As Double
    Dim expl As String
    Dim gate As String
    Dim pct As Long
    Dim level As String

    On Error GoTo SoftFail
    ShapeSimilarityExplain = "0% — сравнение невозможно."

    If shp Is Nothing Then
        ShapeSimilarityExplain = "0% — фигура не задана."
        Exit Function
    End If
    If Not GetRole(sig, r) Then
        ShapeSimilarityExplain = "0% — сигнатура пуста, повреждена или неизвестного формата."
        Exit Function
    End If
    If Not ExtractCheap(shp, f) Then
        ShapeSimilarityExplain = "0% — не удалось прочитать свойства фигуры."
        Exit Function
    End If
    If Not PassGates(f, r, gate) Then
        ShapeSimilarityExplain = "0% — не подходит принципиально: " & gate & "."
        Exit Function
    End If
    ExtractDeep shp, f

    sc = ScoreFeat(f, r, expl)
    If sc < 0# Then sc = 0#
    If sc > 100# Then sc = 100#
    pct = CLng(Round(sc, 0))

    If pct >= r.Thr Then
        level = "совпадение (порог " & CStr(r.Thr) & "%)"
    ElseIf pct >= r.Thr - 15 Then
        level = "пограничный случай, требует проверки"
    ElseIf pct >= 40 Then
        level = "слабая похожесть"
    Else
        level = "не похож"
    End If

    ShapeSimilarityExplain = CStr(pct) & "% — " & level & ". " & expl
    Exit Function
SoftFail:
    ShapeSimilarityExplain = "0% — сбой сравнения: " & Err.Description
End Function

'--- Отклонения формата от эталона (для вашего исправителя) ------------------
Public Function DescribeFormatDeviations(ByRef shp As Shape, ByVal sig As String) As String
    Dim r As RoleSig
    Dim f As ShapeFeat
    Dim s As String

    On Error GoTo SoftFail
    DescribeFormatDeviations = ""
    If shp Is Nothing Then Exit Function
    If Not GetRole(sig, r) Then Exit Function
    If Not r.HasFormat Then
        DescribeFormatDeviations = "Эталон формата не обучен."
        Exit Function
    End If
    If Not ExtractCheap(shp, f) Then Exit Function
    ExtractDeep shp, f

    s = ""
    If r.FmtFontFnv <> 0 Then
        If f.FontFnv <> r.FmtFontFnv Then AppendPart s, "шрифт отличается от эталона"
    End If
    If r.FmtFontRelH < 200 Then
        If Abs(f.FontRel - r.FmtFontRelC) > r.FmtFontRelH + 2 Then
            If f.FontRel < r.FmtFontRelC Then
                AppendPart s, "кегль меньше эталона"
            Else
                AppendPart s, "кегль больше эталона"
            End If
        End If
    End If
    If r.FmtFontBucketSet <> 0 Then
        If Not BitSet(r.FmtFontBucketSet, f.FontBucket) Then AppendPart s, "цвет текста вне эталона"
    End If
    If r.FmtFillBucketSet <> 0 Then
        If Not BitSet(r.FmtFillBucketSet, f.FillBucket) Then AppendPart s, "цвет заливки вне эталона"
    End If
    If r.FmtAlignSet <> 0 Then
        If Not BitSet(r.FmtAlignSet, f.Align) Then AppendPart s, "выравнивание вне эталона"
    End If
    If (r.FmtStyleFlags And FL_BOLD) <> 0 Then
        If (f.Flags And FL_BOLD) = 0 Then AppendPart s, "ожидается полужирный"
    Else
        If (f.Flags And FL_BOLD) <> 0 Then AppendPart s, "лишний полужирный"
    End If
    If Abs(f.Cx - r.FmtCxC) > r.FmtCxH + 4 Then AppendPart s, "смещение по горизонтали"
    If Abs(f.Cy - r.FmtCyC) > r.FmtCyH + 4 Then AppendPart s, "смещение по вертикали"

    If Len(s) = 0 Then
        DescribeFormatDeviations = "Формат соответствует эталону."
    Else
        DescribeFormatDeviations = s & "."
    End If
    Exit Function
SoftFail:
    DescribeFormatDeviations = "Сбой анализа формата: " & Err.Description
End Function

'==============================================================================
' ПУБЛИЧНОЕ API — ПОИСК
'==============================================================================

'--- Лучший кандидат на слайде (для роли «один объект на слайд») -------------
Public Function FindBestMatchOnSlide(ByRef sld As Slide, ByVal sig As String, _
                                     Optional ByRef outPercent As Long, _
                                     Optional ByRef outAmbiguous As Boolean) As Shape
    Dim best As Shape
    Dim bestPct As Long
    Dim secondPct As Long
    Dim shp As Shape

    On Error GoTo SoftFail
    Set FindBestMatchOnSlide = Nothing
    outPercent = 0
    outAmbiguous = False
    If sld Is Nothing Then Exit Function
    If Len(Trim$(sig)) = 0 Then Exit Function

    bestPct = -1
    secondPct = -1
    For Each shp In sld.Shapes
        ScanBestInTree shp, sig, best, bestPct, secondPct
    Next shp

    If bestPct <= 0 Then Exit Function
    outPercent = bestPct
    If secondPct > 0 Then outAmbiguous = ((bestPct - secondPct) < 10)
    Set FindBestMatchOnSlide = best
    Exit Function
SoftFail:
    Set FindBestMatchOnSlide = Nothing
    outPercent = 0
End Function

'--- Все похожие в презентации ------------------------------------------------
Public Function FindSimilarShapes(ByRef pres As Presentation, ByVal sig As String, _
                                  Optional ByVal minPercent As Long = -1) As Collection
    Dim result As Collection
    Dim sld As Slide
    Dim shp As Shape
    Dim r As RoleSig

    On Error GoTo SoftFail
    Set result = New Collection
    Set FindSimilarShapes = result
    If pres Is Nothing Then Exit Function
    If Not GetRole(sig, r) Then Exit Function
    If minPercent < 0 Then minPercent = r.Thr
    If minPercent > 100 Then minPercent = 100

    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            CollectSimilarInTree shp, sig, minPercent, result
        Next shp
    Next sld
    Set FindSimilarShapes = result
    Exit Function
SoftFail:
    On Error Resume Next
    Set FindSimilarShapes = New Collection
End Function

'--- Текстовый отчёт поиска ---------------------------------------------------
Public Function FindSimilarShapesReport(ByRef pres As Presentation, ByVal sig As String, _
                                        Optional ByVal minPercent As Long = -1) As String
    Dim sld As Slide
    Dim best As Shape
    Dim pct As Long
    Dim amb As Boolean
    Dim lines As String
    Dim n As Long
    Dim r As RoleSig
    Dim nm As String

    On Error GoTo SoftFail
    FindSimilarShapesReport = ""
    If pres Is Nothing Then
        FindSimilarShapesReport = "Ошибка: презентация не задана."
        Exit Function
    End If
    If Not GetRole(sig, r) Then
        FindSimilarShapesReport = "Ошибка: сигнатура пуста или повреждена."
        Exit Function
    End If
    If minPercent < 0 Then minPercent = r.Thr

    lines = ""
    n = 0
    For Each sld In pres.Slides
        Set best = FindBestMatchOnSlide(sld, sig, pct, amb)
        If Not best Is Nothing Then
            If pct >= minPercent Then
                n = n + 1
                nm = SafeShapeName(best)
                lines = lines & "• Слайд " & CStr(SafeSlideIndex(sld)) & ": «" & nm & "» — " & _
                        CStr(pct) & "%"
                If amb Then lines = lines & " (неоднозначно)"
                lines = lines & vbCrLf
            End If
        End If
    Next sld

    If n = 0 Then
        FindSimilarShapesReport = "Совпадений >= " & CStr(minPercent) & "% не найдено."
    Else
        FindSimilarShapesReport = "Найдено на слайдах: " & CStr(n) & _
                                  " (порог " & CStr(minPercent) & "%)" & vbCrLf & lines
    End If
    Exit Function
SoftFail:
    FindSimilarShapesReport = "Сбой поиска: " & Err.Description
End Function

'==============================================================================
' ПУБЛИЧНОЕ API — КАТАЛОГ СИГНАТУР В ТЕГАХ ПРЕЗЕНТАЦИИ
'==============================================================================
Public Function SaveSignatureToPresentation(ByRef pres As Presentation, _
                                            ByVal key As String, _
                                            ByVal sig As String) As Boolean
    On Error GoTo SoftFail
    SaveSignatureToPresentation = False
    If pres Is Nothing Then Exit Function
    key = Trim$(key)
    If Len(key) = 0 Or Len(sig) = 0 Then Exit Function
    pres.Tags.Delete TAG_PREFIX & key
    pres.Tags.Add TAG_PREFIX & key, sig
    SaveSignatureToPresentation = True
    Exit Function
SoftFail:
    SaveSignatureToPresentation = False
End Function

Public Function LoadSignatureFromPresentation(ByRef pres As Presentation, _
                                              ByVal key As String) As String
    On Error GoTo SoftFail
    LoadSignatureFromPresentation = ""
    If pres Is Nothing Then Exit Function
    LoadSignatureFromPresentation = pres.Tags(TAG_PREFIX & Trim$(key))
    Exit Function
SoftFail:
    LoadSignatureFromPresentation = ""
End Function

'==============================================================================
' ТЕСТОВЫЕ МАКРОСЫ
'==============================================================================
Public Sub TestSimilaritySelected()
    Dim shp As Shape
    Dim sig As String
    On Error GoTo Fail
    If Not SelectionHasShapes() Then
        MsgBox "Выделите фигуру.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    sig = InputBox("Сигнатура:", "Похожесть")
    If Len(sig) = 0 Then Exit Sub
    MsgBox ShapeSimilarityExplain(shp, sig) & vbCrLf & vbCrLf & _
           "Формат: " & DescribeFormatDeviations(shp, sig), vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка TestSimilaritySelected: " & Err.Description, vbCritical
End Sub

Public Sub TestFindInActivePresentation()
    Dim sig As String
    On Error GoTo Fail
    If ActivePresentation Is Nothing Then
        MsgBox "Нет активной презентации.", vbExclamation
        Exit Sub
    End If
    sig = InputBox("Сигнатура:", "Поиск")
    If Len(sig) = 0 Then Exit Sub
    MsgBox FindSimilarShapesReport(ActivePresentation, sig), vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка TestFindInActivePresentation: " & Err.Description, vbCritical
End Sub

'==============================================================================
' ОБУЧЕНИЕ — ВНУТРЕННЕЕ
'==============================================================================
Private Function LearnFromFileList(ByVal fileList As String, ByVal shapeName As String, _
                                   ByRef accAll As Acc, ByRef accGood As Acc, _
                                   ByVal isGood As Boolean, _
                                   ByRef obs() As ShapeFeat, ByRef nObs As Long, _
                                   ByVal opened As Collection, _
                                   ByRef report As String) As Long
    Dim paths As Collection
    Dim i As Long
    Dim p As String
    Dim pres As Presentation
    Dim wasOpen As Boolean
    Dim cnt As Long
    Dim total As Long

    On Error GoTo SoftFail
    LearnFromFileList = 0
    If Len(Trim$(fileList)) = 0 Then Exit Function
    If Len(Trim$(shapeName)) = 0 Then Exit Function

    Set paths = SplitPaths(fileList)
    For i = 1 To paths.Count
        p = CStr(paths(i))
        Set pres = Nothing
        wasOpen = False

        On Error Resume Next
        Set pres = FindOpenPresentation(p)
        If Not pres Is Nothing Then
            wasOpen = True
        Else
            Set pres = Presentations.Open(FileName:=p, ReadOnly:=msoTrue, _
                                          Untitled:=msoFalse, WithWindow:=msoFalse)
            If Err.Number <> 0 Then
                Err.Clear
                Set pres = Nothing
            End If
            If Not pres Is Nothing Then opened.Add pres
        End If
        Err.Clear
        On Error GoTo SoftFail

        If pres Is Nothing Then
            report = report & "Не удалось открыть: " & p & vbCrLf
        Else
            cnt = LearnFromPresentation(pres, shapeName, accAll, accGood, isGood, obs, nObs)
            If cnt = 0 Then
                report = report & "Фигура «" & shapeName & "» не найдена в: " & p & vbCrLf
            End If
            total = total + cnt
        End If
    Next i

    LearnFromFileList = total
    Exit Function
SoftFail:
    LearnFromFileList = total
    report = report & "Сбой при обработке списка файлов: " & Err.Description & vbCrLf
End Function

Private Function LearnFromPresentation(ByVal pres As Presentation, ByVal shapeName As String, _
                                       ByRef accAll As Acc, ByRef accGood As Acc, _
                                       ByVal isGood As Boolean, _
                                       ByRef obs() As ShapeFeat, ByRef nObs As Long) As Long
    Dim sld As Slide
    Dim shp As Shape
    Dim cnt As Long
    On Error GoTo SoftFail
    LearnFromPresentation = 0
    If pres Is Nothing Then Exit Function
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            LearnTree shp, shapeName, accAll, accGood, isGood, obs, nObs, cnt
        Next shp
    Next sld
    LearnFromPresentation = cnt
    Exit Function
SoftFail:
    LearnFromPresentation = cnt
End Function

Private Sub LearnTree(ByVal shp As Shape, ByVal shapeName As String, _
                      ByRef accAll As Acc, ByRef accGood As Acc, _
                      ByVal isGood As Boolean, _
                      ByRef obs() As ShapeFeat, ByRef nObs As Long, _
                      ByRef cnt As Long)
    Dim f As ShapeFeat
    Dim i As Long
    Dim nm As String

    On Error Resume Next
    If shp Is Nothing Then Exit Sub

    nm = SafeShapeName(shp)
    If StrComp(nm, shapeName, vbTextCompare) = 0 Then
        If ExtractCheap(shp, f) Then
            ExtractDeep shp, f
            AccAdd accAll, f
            If isGood Then AccAdd accGood, f
            If nObs < MAX_OBS Then
                nObs = nObs + 1
                obs(nObs) = f
            End If
            cnt = cnt + 1
        End If
    End If

    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            LearnTree shp.GroupItems(i), shapeName, accAll, accGood, isGood, obs, nObs, cnt
        Next i
    End If
End Sub

Private Sub AccInit(ByRef a As Acc)
    a.n = 0
    RngInit a.Cx
    RngInit a.Cy
    RngInit a.Wd
    RngInit a.Ht
    RngInit a.AR
    RngInit a.Paras
    RngInit a.FontRel
    RngInit a.Rot
    a.TypeSet = 0
    a.PhSet = 0
    a.LenSet = 0
    a.FlagsSeen = 0
    a.FlagsAll = 0
    a.FlagsInit = False
    a.FontFnv = 0
    a.FontFnvSame = True
    a.FontFnvInit = False
    a.FontBucketSet = 0
    a.FillBucketSet = 0
    a.AlignSet = 0
    a.StyleFlags = 0
End Sub

Private Sub RngInit(ByRef r As Rng)
    r.HasVal = False
    r.MinV = 0
    r.MaxV = 0
End Sub

Private Sub RngAdd(ByRef r As Rng, ByVal v As Long)
    If Not r.HasVal Then
        r.HasVal = True
        r.MinV = v
        r.MaxV = v
    Else
        If v < r.MinV Then r.MinV = v
        If v > r.MaxV Then r.MaxV = v
    End If
End Sub

Private Sub AccAdd(ByRef a As Acc, ByRef f As ShapeFeat)
    On Error Resume Next
    a.n = a.n + 1

    RngAdd a.Cx, f.Cx
    RngAdd a.Cy, f.Cy
    RngAdd a.Wd, f.Wd
    RngAdd a.Ht, f.Ht
    RngAdd a.AR, f.AR8
    RngAdd a.Rot, f.Rot2
    If (f.Flags And FL_TEXTFRAME) <> 0 Then
        RngAdd a.Paras, f.Paras
        RngAdd a.FontRel, f.FontRel
        a.LenSet = SetBit(a.LenSet, f.LenBucket)
        a.FontBucketSet = SetBit(a.FontBucketSet, f.FontBucket)
        a.AlignSet = SetBit(a.AlignSet, f.Align)
        If (f.Flags And FL_BOLD) <> 0 Then a.StyleFlags = a.StyleFlags Or FL_BOLD
        If Not a.FontFnvInit Then
            a.FontFnvInit = True
            a.FontFnv = f.FontFnv
        ElseIf a.FontFnv <> f.FontFnv Then
            a.FontFnvSame = False
        End If
    End If
    a.FillBucketSet = SetBit(a.FillBucketSet, f.FillBucket)

    a.TypeSet = SetBit(a.TypeSet, TypeBitIndex(f.TypeId))
    a.PhSet = SetBit(a.PhSet, PhBitIndex(f.PhType))

    a.FlagsSeen = a.FlagsSeen Or f.Flags
    If Not a.FlagsInit Then
        a.FlagsInit = True
        a.FlagsAll = f.Flags
    Else
        a.FlagsAll = a.FlagsAll And f.Flags
    End If
End Sub

Private Sub BuildRoleFromAcc(ByRef accAll As Acc, ByRef accGood As Acc, _
                             ByVal nBad As Long, ByVal nGood As Long, _
                             ByRef r As RoleSig, ByRef report As String)
    Dim dropped As String

    On Error GoTo SoftFail
    r.Valid = False
    r.Mask = 0
    r.nBad = ClampLng(nBad, 0, 255)
    r.nGood = ClampLng(nGood, 0, 255)
    r.Thr = 60
    dropped = ""

    ' Зона центра
    If accAll.Cx.HasVal And accAll.Cy.HasVal Then
        CenterHalf accAll.Cx, 6, r.CxC, r.CxH
        CenterHalf accAll.Cy, 6, r.CyC, r.CyH
        If r.CxH <= LIM_ZONE_H And r.CyH <= LIM_ZONE_H Then
            r.Mask = r.Mask Or R_ZONE
        Else
            AppendPart dropped, "зона"
        End If
    End If

    ' Размеры
    If accAll.Wd.HasVal And accAll.Ht.HasVal Then
        CenterHalf accAll.Wd, 5, r.WC, r.WH
        CenterHalf accAll.Ht, 5, r.HC, r.HH
        If r.WH <= LIM_SIZE_H And r.HH <= LIM_SIZE_H Then
            r.Mask = r.Mask Or R_SIZE
        Else
            AppendPart dropped, "размер"
        End If
    End If

    ' Пропорции
    If accAll.AR.HasVal Then
        CenterHalf accAll.AR, 4, r.ARC, r.ARH
        If r.ARH <= LIM_AR_H Then
            r.Mask = r.Mask Or R_AR
        Else
            AppendPart dropped, "пропорции"
        End If
    End If

    ' Типы фигур
    If accAll.TypeSet <> 0 Then
        r.TypeSet = accAll.TypeSet
        If PopCount(r.TypeSet) <= LIM_TYPE_COUNT Then
            r.Mask = r.Mask Or R_TYPE
        Else
            AppendPart dropped, "тип"
        End If
    End If

    ' Placeholder
    If accAll.PhSet <> 0 Then
        r.PhSet = accAll.PhSet
        If PopCount(r.PhSet) <= 5 Then r.Mask = r.Mask Or R_PH
    End If

    ' Флаги
    r.FlagsSeen = accAll.FlagsSeen
    r.FlagsAll = accAll.FlagsAll
    If r.FlagsAll <> 0 Then r.Mask = r.Mask Or R_FLAGS

    ' Структура текста
    If accAll.Paras.HasVal Then
        CenterHalf accAll.Paras, 1, r.ParaC, r.ParaH
        r.LenSet = accAll.LenSet
        If r.ParaH <= LIM_PARA_H Then
            r.Mask = r.Mask Or R_TXT
        Else
            AppendPart dropped, "структура текста"
        End If
    End If

    ' Относительный кегль (часто это и есть дефект — легко отключается)
    If accAll.FontRel.HasVal Then
        CenterHalf accAll.FontRel, 2, r.FontRelC, r.FontRelH
        If r.FontRelH <= LIM_FONT_H Then
            r.Mask = r.Mask Or R_FONTREL
        Else
            AppendPart dropped, "кегль"
        End If
    End If

    ' Поворот
    If accAll.Rot.HasVal Then
        CenterHalf accAll.Rot, 1, r.RotC, r.RotH
        If r.RotH <= LIM_ROT_H Then r.Mask = r.Mask Or R_ROT
    End If

    ' Эталон формата — только из правильных файлов
    r.HasFormat = False
    If accGood.n > 0 Then
        r.HasFormat = True
        If accGood.FontFnvSame Then r.FmtFontFnv = accGood.FontFnv Else r.FmtFontFnv = 0
        If accGood.FontRel.HasVal Then
            CenterHalf accGood.FontRel, 1, r.FmtFontRelC, r.FmtFontRelH
        Else
            r.FmtFontRelC = 0
            r.FmtFontRelH = 255
        End If
        r.FmtFontBucketSet = accGood.FontBucketSet
        r.FmtFillBucketSet = accGood.FillBucketSet
        r.FmtStyleFlags = accGood.StyleFlags
        r.FmtAlignSet = accGood.AlignSet
        If accGood.Cx.HasVal Then
            CenterHalf accGood.Cx, 3, r.FmtCxC, r.FmtCxH
            CenterHalf accGood.Cy, 3, r.FmtCyC, r.FmtCyH
        End If
    End If

    If Len(dropped) > 0 Then
        report = report & "Отключены нестабильные признаки: " & dropped & vbCrLf
    End If

    r.Valid = (r.Mask <> 0)
    Exit Sub
SoftFail:
    r.Valid = False
End Sub

Private Sub CenterHalf(ByRef rg As Rng, ByVal pad As Long, _
                       ByRef c As Long, ByRef h As Long)
    Dim lo As Long, hi As Long
    lo = rg.MinV
    hi = rg.MaxV
    c = (lo + hi) \ 2
    h = ((hi - lo) \ 2) + pad
    If c < 0 Then c = 0
    If c > 255 Then c = 255
    If h < 0 Then h = 0
    If h > 255 Then h = 255
End Sub

Private Function ComputeThreshold(ByRef obs() As ShapeFeat, ByVal nObs As Long, _
                                  ByRef r As RoleSig) As Long
    Dim i As Long
    Dim sc As Double
    Dim mn As Double
    Dim expl As String

    On Error GoTo SoftFail
    ComputeThreshold = 60
    If nObs <= 0 Then Exit Function
    mn = 101#
    For i = 1 To nObs
        sc = ScoreFeat(obs(i), r, expl)
        If sc < mn Then mn = sc
    Next i
    If mn > 100# Then mn = 100#
    ComputeThreshold = ClampLng(CLng(Round(mn, 0)) - 10, 45, 85)
    Exit Function
SoftFail:
    ComputeThreshold = 60
End Function

'==============================================================================
' ИЗВЛЕЧЕНИЕ ПРИЗНАКОВ (двухфазное: дёшево → дорого)
'==============================================================================
Private Function ExtractCheap(ByVal shp As Shape, ByRef f As ShapeFeat) As Boolean
    Dim sld As Slide
    Dim pres As Presentation
    Dim sw As Double, sh As Double
    Dim lft As Double, tp As Double, wd As Double, ht As Double
    Dim t As Long

    On Error GoTo SoftFail
    FeatInit f
    ExtractCheap = False
    If shp Is Nothing Then Exit Function

    Set sld = ParentSlide(shp)
    If sld Is Nothing Then Exit Function
    On Error Resume Next
    Set pres = sld.Parent
    If pres Is Nothing Then GoTo SoftFail
    sw = pres.PageSetup.SlideWidth
    sh = pres.PageSetup.SlideHeight
    Err.Clear
    On Error GoTo SoftFail
    If sw <= 0# Or sh <= 0# Then Exit Function

    On Error Resume Next
    t = CLng(shp.Type)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    f.TypeId = t

    f.PhType = -1
    If t = msoPlaceholder Then
        f.PhType = CLng(shp.PlaceholderFormat.Type)
        If Err.Number <> 0 Then
            Err.Clear
            f.PhType = -1
        End If
        f.Flags = f.Flags Or FL_PLACEHOLDER
    End If

    lft = shp.Left
    tp = shp.Top
    wd = shp.Width
    ht = shp.Height
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    f.Cx = ClampLng(CLng(Round(((lft + wd / 2#) / sw) * G8, 0)), 0, 255)
    f.Cy = ClampLng(CLng(Round(((tp + ht / 2#) / sh) * G8, 0)), 0, 255)
    f.Wd = ClampLng(CLng(Round((wd / sw) * G8, 0)), 0, 255)
    f.Ht = ClampLng(CLng(Round((ht / sh) * G8, 0)), 0, 255)

    If ht > 0.0001 Then
        f.AR8 = ClampLng(CLng(Round((wd / ht) * 8#, 0)), 0, 255)
    Else
        f.AR8 = 0
    End If

    f.Rot2 = ClampLng(CLng(Round(NormalizeAngle(shp.Rotation) / 2#, 0)), 0, 180)
    If Err.Number <> 0 Then
        Err.Clear
        f.Rot2 = 0
    End If

    f.Blocked = False
    If t = msoTable Or t = msoChart Or t = msoPicture Or t = msoLinkedPicture _
       Or t = msoEmbeddedOLEObject Or t = msoLinkedOLEObject Or t = msoMedia _
       Or t = msoGroup Or t = msoDiagram Then
        f.Blocked = True
    End If

    If shp.HasTextFrame = msoTrue Then f.Flags = f.Flags Or FL_TEXTFRAME
    Err.Clear
    On Error GoTo SoftFail

    f.Valid = True
    ExtractCheap = True
    Exit Function
SoftFail:
    f.Valid = False
    ExtractCheap = False
End Function

Private Sub ExtractDeep(ByVal shp As Shape, ByRef f As ShapeFeat)
    Dim txt As String
    Dim fnt As String
    Dim sz As Single
    Dim got As Boolean
    Dim sld As Slide
    Dim pres As Presentation
    Dim sh As Double
    Dim ct As Long
    Dim rgbv As Long

    On Error Resume Next
    If f.Deep Then Exit Sub
    f.Deep = True

    ' Заливка: важен факт выделения, а не конкретный цвет
    If shp.Fill.Visible = msoTrue Then
        f.Flags = f.Flags Or FL_FILL
        ct = CLng(shp.Fill.Type)
        If Err.Number = 0 Then
            If ct = msoFillSolid Then
                rgbv = CLng(shp.Fill.ForeColor.RGB)
                If Err.Number = 0 Then f.FillBucket = ColorBucket(rgbv)
            End If
        End If
    End If
    Err.Clear

    ' Текст
    got = False
    If shp.HasTextFrame = msoTrue Then
        If shp.TextFrame.HasText Then
            txt = shp.TextFrame.TextRange.Text
            fnt = shp.TextFrame.TextRange.Font.Name
            sz = shp.TextFrame.TextRange.Font.Size
            f.Paras = ClampLng(CLng(shp.TextFrame.TextRange.Paragraphs.Count), 1, 255)
            f.Align = AlignBucket(CLng(shp.TextFrame.TextRange.ParagraphFormat.Alignment))
            If shp.TextFrame.TextRange.Font.Bold = msoTrue Then f.Flags = f.Flags Or FL_BOLD
            rgbv = CLng(shp.TextFrame.TextRange.Font.Color.RGB)
            If Err.Number = 0 Then f.FontBucket = ColorBucket(rgbv)
            got = True
        End If
    End If
    Err.Clear

    If Not got Then
        If shp.TextFrame2.HasText Then
            txt = shp.TextFrame2.TextRange.Text
            fnt = CStr(shp.TextFrame2.TextRange.Font.Name)
            sz = shp.TextFrame2.TextRange.Font.Size
            f.Paras = ClampLng(CLng(shp.TextFrame2.TextRange.Paragraphs.Count), 1, 255)
            got = True
        End If
    End If
    Err.Clear

    If Not got Then Exit Sub

    f.Flags = f.Flags Or FL_HASTEXT
    If f.Paras > 1 Then f.Flags = f.Flags Or FL_MULTIPARA
    If HasDigits(txt) Then f.Flags = f.Flags Or FL_DIGITS
    f.LenBucket = LenBucket(txt)
    If Len(fnt) > 0 Then f.FontFnv = Fnv1a32(LCase$(Trim$(fnt)))

    If f.FontBucket <> 0 Then
        If f.FontBucket <> 2 And f.FontBucket <> 3 Then f.Flags = f.Flags Or FL_COLORFONT
    End If

    Set sld = ParentSlide(shp)
    If Not sld Is Nothing Then
        Set pres = sld.Parent
        If Not pres Is Nothing Then
            sh = pres.PageSetup.SlideHeight
            If sh > 0# And sz > 0 Then
                f.FontRel = ClampLng(CLng(Round((CDbl(sz) / sh) * 1000#, 0)), 0, 255)
            End If
        End If
    End If
    Err.Clear
End Sub

Private Sub FeatInit(ByRef f As ShapeFeat)
    f.Valid = False
    f.Deep = False
    f.Blocked = False
    f.TypeId = 0
    f.PhType = -1
    f.Cx = 0
    f.Cy = 0
    f.Wd = 0
    f.Ht = 0
    f.AR8 = 0
    f.Rot2 = 0
    f.Flags = 0
    f.Paras = 0
    f.LenBucket = 0
    f.FontRel = 0
    f.FontFnv = 0
    f.FillBucket = 0
    f.FontBucket = 0
    f.Align = 0
End Sub

Private Function NormalizeAngle(ByVal deg As Single) As Double
    ' Без циклов: защита от зацикливания на аномальных значениях
    Dim d As Double
    On Error GoTo SoftFail
    d = CDbl(deg)
    If d > 100000# Or d < -100000# Then
        NormalizeAngle = 0#
        Exit Function
    End If
    d = d - Int(d / 360#) * 360#
    If d < 0# Then d = d + 360#
    If d >= 360# Then d = 0#
    NormalizeAngle = d
    Exit Function
SoftFail:
    NormalizeAngle = 0#
End Function

Private Function HasDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim n As Long
    Dim c As Long
    HasDigits = False
    n = Len(s)
    If n > 200 Then n = 200
    For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        If c >= 48 And c <= 57 Then
            HasDigits = True
            Exit Function
        End If
    Next i
End Function

Private Function LenBucket(ByVal s As String) As Long
    Dim n As Long
    n = Len(Trim$(s))
    If n = 0 Then
        LenBucket = 0
    ElseIf n <= 10 Then
        LenBucket = 1
    ElseIf n <= 25 Then
        LenBucket = 2
    ElseIf n <= 50 Then
        LenBucket = 3
    ElseIf n <= 100 Then
        LenBucket = 4
    ElseIf n <= 250 Then
        LenBucket = 5
    ElseIf n <= 1000 Then
        LenBucket = 6
    Else
        LenBucket = 7
    End If
End Function

' Грубая корзина цвета: устойчива к смене оттенка, различает «цветной / нет»
Private Function ColorBucket(ByVal rgbv As Long) As Long
    Dim r As Long, g As Long, b As Long
    Dim mx As Long, mn As Long
    On Error GoTo SoftFail
    r = rgbv And &HFF&
    g = (rgbv \ &H100&) And &HFF&
    b = (rgbv \ &H10000) And &HFF&
    mx = r
    If g > mx Then mx = g
    If b > mx Then mx = b
    mn = r
    If g < mn Then mn = g
    If b < mn Then mn = b

    If mx < 40 Then
        ColorBucket = 2                 ' почти чёрный
    ElseIf mn > 215 Then
        ColorBucket = 1                 ' почти белый
    ElseIf (mx - mn) < 30 Then
        ColorBucket = 3                 ' серый
    ElseIf r >= g And r >= b Then
        If g > 128 Then
            ColorBucket = 7             ' жёлтый / оранжевый
        Else
            ColorBucket = 4             ' красный
        End If
    ElseIf g >= r And g >= b Then
        ColorBucket = 5                 ' зелёный
    Else
        ColorBucket = 6                 ' синий
    End If
    Exit Function
SoftFail:
    ColorBucket = 0
End Function

Private Function AlignBucket(ByVal a As Long) As Long
    If a >= 1 And a <= 6 Then
        AlignBucket = a
    Else
        AlignBucket = 7
    End If
End Function

Private Function TypeBitIndex(ByVal t As Long) As Long
    If t >= 1 And t <= 29 Then
        TypeBitIndex = t - 1
    Else
        TypeBitIndex = 29
    End If
End Function

Private Function PhBitIndex(ByVal ph As Long) As Long
    If ph < 0 Then
        PhBitIndex = 15                 ' не placeholder
    ElseIf ph >= 1 And ph <= 15 Then
        PhBitIndex = ph - 1
    Else
        PhBitIndex = 14
    End If
End Function

'==============================================================================
' ОЦЕНКА
'==============================================================================
Private Function PassGates(ByRef f As ShapeFeat, ByRef r As RoleSig, _
                           ByRef reason As String) As Boolean
    On Error GoTo SoftFail
    PassGates = False
    reason = ""
    If Not f.Valid Then
        reason = "свойства не прочитаны"
        Exit Function
    End If

    ' Требование текстового фрейма, если он был у всех обучающих примеров
    If (r.FlagsAll And FL_TEXTFRAME) <> 0 Then
        If (f.Flags And FL_TEXTFRAME) = 0 Then
            reason = "нет текстового блока"
            Exit Function
        End If
        If f.Blocked Then
            reason = "это таблица/диаграмма/рисунок/группа"
            Exit Function
        End If
    End If

    ' Грубый отсев по зоне (дёшево, до чтения текста и шрифта)
    If (r.Mask And R_ZONE) <> 0 Then
        If Abs(f.Cx - r.CxC) > r.CxH + SLACK_ZONE * 3 Then
            reason = "далеко по горизонтали от обученной зоны"
            Exit Function
        End If
        If Abs(f.Cy - r.CyC) > r.CyH + SLACK_ZONE * 3 Then
            reason = "далеко по вертикали от обученной зоны"
            Exit Function
        End If
    End If

    PassGates = True
    Exit Function
SoftFail:
    PassGates = False
    reason = "сбой проверки"
End Function

Private Function ScoreFeat(ByRef f As ShapeFeat, ByRef r As RoleSig, _
                           ByRef explain As String) As Double
    Dim wSum As Double
    Dim sSum As Double
    Dim okp As String
    Dim bad As String
    Dim w As Double
    Dim s As Double

    On Error GoTo SoftFail
    ScoreFeat = 0#
    explain = ""
    wSum = 0#
    sSum = 0#
    okp = ""
    bad = ""

    If (r.Mask And R_ZONE) <> 0 Then
        s = 0.5 * BandScore(f.Cx, r.CxC, r.CxH, SLACK_ZONE) + _
            0.5 * BandScore(f.Cy, r.CyC, r.CyH, SLACK_ZONE)
        w = W_ZONE * Stability(r.CxH + r.CyH, LIM_ZONE_H * 2)
        Accumulate w, s, "зона", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_SIZE) <> 0 Then
        s = 0.5 * BandScore(f.Wd, r.WC, r.WH, SLACK_SIZE) + _
            0.5 * BandScore(f.Ht, r.HC, r.HH, SLACK_SIZE)
        w = W_SIZE * Stability(r.WH + r.HH, LIM_SIZE_H * 2)
        Accumulate w, s, "размер", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_AR) <> 0 Then
        s = BandScore(f.AR8, r.ARC, r.ARH, SLACK_AR)
        w = W_AR * Stability(r.ARH, LIM_AR_H)
        Accumulate w, s, "пропорции", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_TYPE) <> 0 Then
        If BitSet(r.TypeSet, TypeBitIndex(f.TypeId)) Then s = 1# Else s = 0#
        Accumulate W_TYPE, s, "тип", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_PH) <> 0 Then
        If BitSet(r.PhSet, PhBitIndex(f.PhType)) Then s = 1# Else s = 0.25
        Accumulate W_PH, s, "placeholder", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_FLAGS) <> 0 Then
        s = FlagScore(f.Flags, r.FlagsAll)
        Accumulate W_FLAGS, s, "признаки блока", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_TXT) <> 0 Then
        s = 0.6 * BandScore(f.Paras, r.ParaC, r.ParaH, SLACK_PARA) + _
            0.4 * BucketScore(r.LenSet, f.LenBucket)
        w = W_TXT * Stability(r.ParaH, LIM_PARA_H)
        Accumulate w, s, "структура текста", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_FONTREL) <> 0 Then
        s = BandScore(f.FontRel, r.FontRelC, r.FontRelH, SLACK_FONT)
        w = W_FONTREL * Stability(r.FontRelH, LIM_FONT_H)
        Accumulate w, s, "кегль", wSum, sSum, okp, bad
    End If

    If (r.Mask And R_ROT) <> 0 Then
        s = BandScore(f.Rot2, r.RotC, r.RotH, SLACK_ROT)
        Accumulate W_ROT, s, "поворот", wSum, sSum, okp, bad
    End If

    If wSum <= 0# Then
        explain = "Нет сопоставимых признаков."
        Exit Function
    End If

    ScoreFeat = 100# * (sSum / wSum)

    If Len(okp) > 0 Then explain = "Совпало: " & okp & "."
    If Len(bad) > 0 Then
        If Len(explain) > 0 Then explain = explain & " "
        explain = explain & "Отличается: " & bad & "."
    End If
    If Len(explain) = 0 Then explain = "Сравнение выполнено."
    Exit Function
SoftFail:
    ScoreFeat = 0#
    explain = "ошибка расчёта"
End Function

Private Sub Accumulate(ByVal w As Double, ByVal s As Double, ByVal label As String, _
                       ByRef wSum As Double, ByRef sSum As Double, _
                       ByRef okp As String, ByRef bad As String)
    If w <= 0# Then Exit Sub
    If s < 0# Then s = 0#
    If s > 1# Then s = 1#
    wSum = wSum + w
    sSum = sSum + w * s
    If s >= 0.75 Then
        AppendPart okp, label
    ElseIf s < 0.45 Then
        AppendPart bad, label
    End If
End Sub

Private Function BandScore(ByVal v As Long, ByVal c As Long, ByVal h As Long, _
                           ByVal slack As Long) As Double
    Dim d As Long
    Dim over As Long
    If slack <= 0 Then slack = 1
    d = Abs(v - c)
    If d <= h Then
        BandScore = 1#
        Exit Function
    End If
    over = d - h
    If over >= slack Then
        BandScore = 0#
    Else
        BandScore = 1# - (CDbl(over) / CDbl(slack))
    End If
End Function

' Широкий диапазон → признак менее надёжен → меньший вес (но не ноль)
Private Function Stability(ByVal h As Long, ByVal maxH As Long) As Double
    Dim k As Double
    If maxH <= 0 Then
        Stability = 1#
        Exit Function
    End If
    k = CDbl(h) / CDbl(maxH)
    If k > 1# Then k = 1#
    Stability = 1# - 0.7 * k
End Function

Private Function FlagScore(ByVal actual As Long, ByVal required As Long) As Double
    Dim i As Long
    Dim need As Long
    Dim hit As Long
    Dim bitv As Long
    If required = 0 Then
        FlagScore = 1#
        Exit Function
    End If
    For i = 0 To 7
        bitv = PowL(i)
        If (required And bitv) <> 0 Then
            need = need + 1
            If (actual And bitv) <> 0 Then hit = hit + 1
        End If
    Next i
    If need = 0 Then
        FlagScore = 1#
    Else
        FlagScore = CDbl(hit) / CDbl(need)
    End If
End Function

' Корзина совпала → 1; соседняя → 0.6; иначе → 0.25 (текст может меняться)
Private Function BucketScore(ByVal setMask As Long, ByVal b As Long) As Double
    If setMask = 0 Then
        BucketScore = 1#
        Exit Function
    End If
    If BitSet(setMask, b) Then
        BucketScore = 1#
    ElseIf BitSet(setMask, b - 1) Or BitSet(setMask, b + 1) Then
        BucketScore = 0.6
    Else
        BucketScore = 0.25
    End If
End Function

'==============================================================================
' ПОИСК — ВНУТРЕННЕЕ
'==============================================================================
Private Sub CollectSimilarInTree(ByVal shp As Shape, ByVal sig As String, _
                                 ByVal minPercent As Long, ByVal result As Collection)
    Dim i As Long
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    If ShapeSimilarityPercent(shp, sig) >= minPercent Then result.Add shp
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            CollectSimilarInTree shp.GroupItems(i), sig, minPercent, result
        Next i
    End If
End Sub

Private Sub ScanBestInTree(ByVal shp As Shape, ByVal sig As String, _
                           ByRef best As Shape, ByRef bestPct As Long, _
                           ByRef secondPct As Long)
    Dim pct As Long
    Dim i As Long
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    pct = ShapeSimilarityPercent(shp, sig)
    If pct > bestPct Then
        secondPct = bestPct
        bestPct = pct
        Set best = shp
    ElseIf pct > secondPct Then
        secondPct = pct
    End If
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            ScanBestInTree shp.GroupItems(i), sig, best, bestPct, secondPct
        Next i
    End If
End Sub

'==============================================================================
' КОДИРОВАНИЕ S3
'==============================================================================
Private Function EncodeRoleSig(ByRef r As RoleSig) As String
    Dim buf() As Byte
    Dim n As Long
    Dim crc As Long
    Dim flags As Long

    On Error GoTo SoftFail
    EncodeRoleSig = ""
    If Not r.Valid Or r.Mask = 0 Then Exit Function

    ReDim buf(0 To 127)
    n = 0
    flags = 1
    If r.HasFormat Then flags = flags Or 2

    AppendU8 buf, n, SIG_VER_NUM
    AppendU8 buf, n, flags
    AppendU16 buf, n, r.Mask
    AppendU8 buf, n, r.nBad
    AppendU8 buf, n, r.nGood
    AppendU8 buf, n, r.Thr

    If (r.Mask And R_ZONE) <> 0 Then
        AppendU8 buf, n, r.CxC
        AppendU8 buf, n, r.CxH
        AppendU8 buf, n, r.CyC
        AppendU8 buf, n, r.CyH
    End If
    If (r.Mask And R_SIZE) <> 0 Then
        AppendU8 buf, n, r.WC
        AppendU8 buf, n, r.WH
        AppendU8 buf, n, r.HC
        AppendU8 buf, n, r.HH
    End If
    If (r.Mask And R_AR) <> 0 Then
        AppendU8 buf, n, r.ARC
        AppendU8 buf, n, r.ARH
    End If
    If (r.Mask And R_TYPE) <> 0 Then AppendU32 buf, n, r.TypeSet
    If (r.Mask And R_PH) <> 0 Then AppendU16 buf, n, r.PhSet
    If (r.Mask And R_FLAGS) <> 0 Then
        AppendU8 buf, n, r.FlagsSeen
        AppendU8 buf, n, r.FlagsAll
    End If
    If (r.Mask And R_TXT) <> 0 Then
        AppendU8 buf, n, r.ParaC
        AppendU8 buf, n, r.ParaH
        AppendU8 buf, n, r.LenSet
    End If
    If (r.Mask And R_FONTREL) <> 0 Then
        AppendU8 buf, n, r.FontRelC
        AppendU8 buf, n, r.FontRelH
    End If
    If (r.Mask And R_ROT) <> 0 Then
        AppendU8 buf, n, r.RotC
        AppendU8 buf, n, r.RotH
    End If

    If r.HasFormat Then
        AppendU32 buf, n, r.FmtFontFnv
        AppendU8 buf, n, r.FmtFontRelC
        AppendU8 buf, n, r.FmtFontRelH
        AppendU8 buf, n, r.FmtFontBucketSet
        AppendU8 buf, n, r.FmtFillBucketSet
        AppendU8 buf, n, r.FmtStyleFlags
        AppendU8 buf, n, r.FmtAlignSet
        AppendU8 buf, n, r.FmtCxC
        AppendU8 buf, n, r.FmtCxH
        AppendU8 buf, n, r.FmtCyC
        AppendU8 buf, n, r.FmtCyH
    End If

    crc = Fnv1aBytes(buf, n) And &HFFFFFF
    AppendU8 buf, n, crc And &HFF&
    AppendU8 buf, n, (crc \ &H100&) And &HFF&
    AppendU8 buf, n, (crc \ &H10000) And &HFF&

    ReDim Preserve buf(0 To n - 1)
    EncodeRoleSig = SIG_VER & "|" & Base64UrlEncode(buf)
    Exit Function
SoftFail:
    EncodeRoleSig = ""
End Function

Private Function DecodeRoleSig(ByVal sig As String, ByRef r As RoleSig) As Boolean
    Dim p() As String
    Dim buf() As Byte
    Dim n As Long
    Dim pos As Long
    Dim bodyLen As Long
    Dim crc As Long
    Dim got As Long
    Dim flags As Long

    On Error GoTo SoftFail
    DecodeRoleSig = False
    r.Valid = False
    sig = Trim$(sig)
    If Len(sig) = 0 Then Exit Function

    p = Split(sig, "|")
    If UBound(p) <> 1 Then Exit Function
    If p(0) <> SIG_VER Then Exit Function
    If Not Base64UrlDecode(p(1), buf) Then Exit Function

    n = UBound(buf) + 1
    If n < 10 Then Exit Function
    bodyLen = n - 3
    crc = CLng(buf(bodyLen)) + CLng(buf(bodyLen + 1)) * &H100& + CLng(buf(bodyLen + 2)) * &H10000
    got = Fnv1aBytes(buf, bodyLen) And &HFFFFFF
    If crc <> got Then Exit Function

    pos = 0
    If ReadU8(buf, pos) <> SIG_VER_NUM Then Exit Function
    flags = ReadU8(buf, pos)
    r.Mask = ReadU16(buf, pos)
    If r.Mask = 0 Or (r.Mask And Not R_ALL) <> 0 Then Exit Function
    r.nBad = ReadU8(buf, pos)
    r.nGood = ReadU8(buf, pos)
    r.Thr = ReadU8(buf, pos)
    If r.Thr < 1 Or r.Thr > 100 Then r.Thr = 60

    If (r.Mask And R_ZONE) <> 0 Then
        r.CxC = ReadU8(buf, pos)
        r.CxH = ReadU8(buf, pos)
        r.CyC = ReadU8(buf, pos)
        r.CyH = ReadU8(buf, pos)
    End If
    If (r.Mask And R_SIZE) <> 0 Then
        r.WC = ReadU8(buf, pos)
        r.WH = ReadU8(buf, pos)
        r.HC = ReadU8(buf, pos)
        r.HH = ReadU8(buf, pos)
    End If
    If (r.Mask And R_AR) <> 0 Then
        r.ARC = ReadU8(buf, pos)
        r.ARH = ReadU8(buf, pos)
    End If
    If (r.Mask And R_TYPE) <> 0 Then r.TypeSet = ReadU32(buf, pos)
    If (r.Mask And R_PH) <> 0 Then r.PhSet = ReadU16(buf, pos)
    If (r.Mask And R_FLAGS) <> 0 Then
        r.FlagsSeen = ReadU8(buf, pos)
        r.FlagsAll = ReadU8(buf, pos)
    End If
    If (r.Mask And R_TXT) <> 0 Then
        r.ParaC = ReadU8(buf, pos)
        r.ParaH = ReadU8(buf, pos)
        r.LenSet = ReadU8(buf, pos)
    End If
    If (r.Mask And R_FONTREL) <> 0 Then
        r.FontRelC = ReadU8(buf, pos)
        r.FontRelH = ReadU8(buf, pos)
    End If
    If (r.Mask And R_ROT) <> 0 Then
        r.RotC = ReadU8(buf, pos)
        r.RotH = ReadU8(buf, pos)
    End If

    r.HasFormat = ((flags And 2) <> 0)
    If r.HasFormat Then
        r.FmtFontFnv = ReadU32(buf, pos)
        r.FmtFontRelC = ReadU8(buf, pos)
        r.FmtFontRelH = ReadU8(buf, pos)
        r.FmtFontBucketSet = ReadU8(buf, pos)
        r.FmtFillBucketSet = ReadU8(buf, pos)
        r.FmtStyleFlags = ReadU8(buf, pos)
        r.FmtAlignSet = ReadU8(buf, pos)
        r.FmtCxC = ReadU8(buf, pos)
        r.FmtCxH = ReadU8(buf, pos)
        r.FmtCyC = ReadU8(buf, pos)
        r.FmtCyH = ReadU8(buf, pos)
    End If

    If pos <> bodyLen Then Exit Function
    r.Valid = True
    DecodeRoleSig = True
    Exit Function
SoftFail:
    r.Valid = False
    DecodeRoleSig = False
End Function

' Разбор с кешем: в циклах поиска сигнатура парсится один раз
Private Function GetRole(ByVal sig As String, ByRef r As RoleSig) As Boolean
    On Error GoTo SoftFail
    GetRole = False
    sig = Trim$(sig)
    If Len(sig) = 0 Then Exit Function

    If sig = m_cacheSig Then
        If m_cacheRole.Valid Then
            r = m_cacheRole
            GetRole = True
            Exit Function
        End If
        Exit Function
    End If

    If DecodeRoleSig(sig, r) Then
        m_cacheSig = sig
        m_cacheRole = r
        GetRole = True
    Else
        m_cacheSig = sig
        m_cacheRole.Valid = False
    End If
    Exit Function
SoftFail:
    GetRole = False
End Function

Private Function MaskDescription(ByVal mask As Long) As String
    Dim s As String
    s = ""
    If (mask And R_ZONE) <> 0 Then AppendPart s, "зона"
    If (mask And R_SIZE) <> 0 Then AppendPart s, "размер"
    If (mask And R_AR) <> 0 Then AppendPart s, "пропорции"
    If (mask And R_TYPE) <> 0 Then AppendPart s, "тип"
    If (mask And R_PH) <> 0 Then AppendPart s, "placeholder"
    If (mask And R_FLAGS) <> 0 Then AppendPart s, "признаки блока"
    If (mask And R_TXT) <> 0 Then AppendPart s, "структура текста"
    If (mask And R_FONTREL) <> 0 Then AppendPart s, "кегль"
    If (mask And R_ROT) <> 0 Then AppendPart s, "поворот"
    If Len(s) = 0 Then MaskDescription = "(нет)" Else MaskDescription = s
End Function

'==============================================================================
' УТИЛИТЫ
'==============================================================================
Private Function SplitPaths(ByVal s As String) As Collection
    Dim c As Collection
    Dim parts() As String
    Dim i As Long
    Dim t As String
    Set c = New Collection
    On Error GoTo SoftFail
    parts = Split(s, ";")
    For i = LBound(parts) To UBound(parts)
        t = Trim$(parts(i))
        If Len(t) >= 2 Then
            If Left$(t, 1) = """" And Right$(t, 1) = """" Then t = Mid$(t, 2, Len(t) - 2)
        End If
        t = Trim$(t)
        If Len(t) > 0 Then c.Add t
    Next i
SoftFail:
    Set SplitPaths = c
End Function

Private Function FindOpenPresentation(ByVal fullPath As String) As Presentation
    Dim p As Presentation
    Dim target As String
    On Error Resume Next
    Set FindOpenPresentation = Nothing
    target = LCase$(Trim$(fullPath))
    For Each p In Presentations
        If LCase$(p.FullName) = target Then
            Set FindOpenPresentation = p
            Exit Function
        End If
    Next p
End Function

Private Sub CloseOpened(ByVal opened As Collection)
    Dim i As Long
    Dim p As Presentation
    On Error Resume Next
    If opened Is Nothing Then Exit Sub
    For i = opened.Count To 1 Step -1
        Set p = opened(i)
        p.Close
    Next i
End Sub

Private Function ParentSlide(ByVal shp As Shape) As Slide
    Dim p As Object
    On Error Resume Next
    Set ParentSlide = Nothing
    Set p = shp.Parent
    Do While Not p Is Nothing
        If TypeName(p) = "Slide" Then
            Set ParentSlide = p
            Exit Function
        End If
        Set p = p.Parent
    Loop
End Function

Private Function SafeShapeName(ByVal shp As Shape) As String
    On Error Resume Next
    SafeShapeName = ""
    SafeShapeName = shp.Name
    If Err.Number <> 0 Then
        Err.Clear
        SafeShapeName = "?"
    End If
End Function

Private Function SafeSlideIndex(ByVal sld As Slide) As Long
    On Error Resume Next
    SafeSlideIndex = 0
    SafeSlideIndex = sld.SlideIndex
    If Err.Number <> 0 Then
        Err.Clear
        SafeSlideIndex = 0
    End If
End Function

Private Function SelectionHasShapes() As Boolean
    On Error Resume Next
    SelectionHasShapes = False
    If ActiveWindow Is Nothing Then Exit Function
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        SelectionHasShapes = (ActiveWindow.Selection.ShapeRange.Count > 0)
    End If
End Function

Private Sub AppendPart(ByRef bag As String, ByVal part As String)
    If Len(bag) = 0 Then
        bag = part
    Else
        bag = bag & ", " & part
    End If
End Sub

Private Function ClampLng(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If v < lo Then
        ClampLng = lo
    ElseIf v > hi Then
        ClampLng = hi
    Else
        ClampLng = v
    End If
End Function

Private Function PowL(ByVal bit As Long) As Long
    Dim i As Long
    Dim v As Long
    If bit < 0 Or bit > 30 Then
        PowL = 0
        Exit Function
    End If
    v = 1
    For i = 1 To bit
        v = v * 2
    Next i
    PowL = v
End Function

Private Function SetBit(ByVal mask As Long, ByVal bit As Long) As Long
    Dim b As Long
    b = PowL(bit)
    If b = 0 Then
        SetBit = mask
    Else
        SetBit = mask Or b
    End If
End Function

Private Function BitSet(ByVal mask As Long, ByVal bit As Long) As Boolean
    Dim b As Long
    b = PowL(bit)
    If b = 0 Then
        BitSet = False
    Else
        BitSet = ((mask And b) <> 0)
    End If
End Function

Private Function PopCount(ByVal mask As Long) As Long
    Dim i As Long
    Dim n As Long
    For i = 0 To 30
        If (mask And PowL(i)) <> 0 Then n = n + 1
    Next i
    PopCount = n
End Function

'----- Буфер байт -------------------------------------------------------------
Private Sub EnsureBuf(ByRef buf() As Byte, ByVal need As Long)
    If need > UBound(buf) Then ReDim Preserve buf(0 To need + 64)
End Sub

Private Sub AppendU8(ByRef buf() As Byte, ByRef n As Long, ByVal v As Long)
    EnsureBuf buf, n
    buf(n) = CByte(v And &HFF&)
    n = n + 1
End Sub

Private Sub AppendU16(ByRef buf() As Byte, ByRef n As Long, ByVal v As Long)
    AppendU8 buf, n, v And &HFF&
    AppendU8 buf, n, (v \ &H100&) And &HFF&
End Sub

Private Sub AppendU32(ByRef buf() As Byte, ByRef n As Long, ByVal v As Long)
    Dim uh As Double
    Dim b As Long
    Dim i As Long
    If v < 0 Then
        uh = CDbl(v And &H7FFFFFFF) + 2147483648#
    Else
        uh = CDbl(v)
    End If
    For i = 1 To 4
        b = CLng(uh - Int(uh / 256#) * 256#)
        AppendU8 buf, n, b
        uh = Int(uh / 256#)
    Next i
End Sub

Private Function ReadU8(ByRef buf() As Byte, ByRef pos As Long) As Long
    ReadU8 = buf(pos)
    pos = pos + 1
End Function

Private Function ReadU16(ByRef buf() As Byte, ByRef pos As Long) As Long
    ReadU16 = CLng(buf(pos)) + CLng(buf(pos + 1)) * &H100&
    pos = pos + 2
End Function

Private Function ReadU32(ByRef buf() As Byte, ByRef pos As Long) As Long
    Dim uh As Double
    uh = CDbl(buf(pos)) + CDbl(buf(pos + 1)) * 256# + _
         CDbl(buf(pos + 2)) * 65536# + CDbl(buf(pos + 3)) * 16777216#
    pos = pos + 4
    If uh >= 2147483648# Then
        ReadU32 = CLng(uh - 4294967296#)
    Else
        ReadU32 = CLng(uh)
    End If
End Function

'----- Base64URL (чистый VBA) -------------------------------------------------
Private Function Base64UrlEncode(ByRef buf() As Byte) As String
    Dim alphabet As String
    Dim out As String
    Dim i As Long, n As Long, remn As Long
    Dim a As Long, b As Long, c As Long

    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    n = UBound(buf) + 1
    out = ""
    i = 0
    Do While i + 2 < n
        a = buf(i)
        b = buf(i + 1)
        c = buf(i + 2)
        out = out & Mid$(alphabet, ((a \ 4) And 63) + 1, 1)
        out = out & Mid$(alphabet, (((a And 3) * 16) Or ((b \ 16) And 15)) + 1, 1)
        out = out & Mid$(alphabet, (((b And 15) * 4) Or ((c \ 64) And 3)) + 1, 1)
        out = out & Mid$(alphabet, (c And 63) + 1, 1)
        i = i + 3
    Loop
    remn = n - i
    If remn = 1 Then
        a = buf(i)
        out = out & Mid$(alphabet, ((a \ 4) And 63) + 1, 1)
        out = out & Mid$(alphabet, ((a And 3) * 16) + 1, 1)
    ElseIf remn = 2 Then
        a = buf(i)
        b = buf(i + 1)
        out = out & Mid$(alphabet, ((a \ 4) And 63) + 1, 1)
        out = out & Mid$(alphabet, (((a And 3) * 16) Or ((b \ 16) And 15)) + 1, 1)
        out = out & Mid$(alphabet, ((b And 15) * 4) + 1, 1)
    End If
    Base64UrlEncode = out
End Function

Private Function Base64UrlDecode(ByVal s As String, ByRef buf() As Byte) As Boolean
    Dim alphabet As String
    Dim vals() As Long
    Dim i As Long, n As Long, o As Long, p As Long
    Dim a As Long, b As Long, c As Long, d As Long

    On Error GoTo SoftFail
    Base64UrlDecode = False
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    s = Replace(Replace(Trim$(s), "+", "-"), "/", "_")
    s = Replace(s, "=", "")
    n = Len(s)
    If n < 2 Then Exit Function

    ReDim vals(0 To n - 1)
    For i = 1 To n
        p = InStr(1, alphabet, Mid$(s, i, 1), vbBinaryCompare)
        If p <= 0 Then Exit Function
        vals(i - 1) = p - 1
    Next i

    ReDim buf(0 To ((n * 3) \ 4) + 3)
    o = 0
    i = 0
    Do While i + 3 < n
        a = vals(i)
        b = vals(i + 1)
        c = vals(i + 2)
        d = vals(i + 3)
        buf(o) = CByte(((a * 4) Or (b \ 16)) And &HFF&)
        buf(o + 1) = CByte((((b And 15) * 16) Or (c \ 4)) And &HFF&)
        buf(o + 2) = CByte((((c And 3) * 64) Or d) And &HFF&)
        o = o + 3
        i = i + 4
    Loop
    If n - i = 2 Then
        a = vals(i)
        b = vals(i + 1)
        buf(o) = CByte(((a * 4) Or (b \ 16)) And &HFF&)
        o = o + 1
    ElseIf n - i = 3 Then
        a = vals(i)
        b = vals(i + 1)
        c = vals(i + 2)
        buf(o) = CByte(((a * 4) Or (b \ 16)) And &HFF&)
        buf(o + 1) = CByte((((b And 15) * 16) Or (c \ 4)) And &HFF&)
        o = o + 2
    ElseIf n - i <> 0 Then
        Exit Function
    End If
    If o = 0 Then Exit Function
    ReDim Preserve buf(0 To o - 1)
    Base64UrlDecode = True
    Exit Function
SoftFail:
    Base64UrlDecode = False
End Function

'----- FNV-1a 32 --------------------------------------------------------------
Private Function Fnv1a32(ByVal s As String) As Long
    Dim h As Long
    Dim i As Long
    h = -2128831035
    For i = 1 To Len(s)
        h = h Xor (AscW(Mid$(s, i, 1)) And &HFF&)
        h = FnvMul(h)
    Next i
    Fnv1a32 = h
End Function

Private Function Fnv1aBytes(ByRef buf() As Byte, ByVal n As Long) As Long
    Dim h As Long
    Dim i As Long
    h = -2128831035
    For i = 0 To n - 1
        h = h Xor (CLng(buf(i)) And &HFF&)
        h = FnvMul(h)
    Next i
    Fnv1aBytes = h
End Function

Private Function FnvMul(ByVal h As Long) As Long
    ' (h * 16777619) mod 2^32 через 16-битные половины:
    ' все промежуточные значения остаются точными в Double
    Dim uh As Double
    Dim lo As Double
    Dim hi As Double
    Dim r As Double
    Dim t As Double

    If h < 0 Then
        uh = CDbl(h And &H7FFFFFFF) + 2147483648#
    Else
        uh = CDbl(h)
    End If

    hi = Int(uh / 65536#)
    lo = uh - hi * 65536#

    r = lo * 16777619#
    r = r - Int(r / 4294967296#) * 4294967296#

    t = hi * 16777619#
    t = t - Int(t / 65536#) * 65536#

    r = r + t * 65536#
    r = r - Int(r / 4294967296#) * 4294967296#

    If r >= 2147483648# Then
        FnvMul = CLng(r - 4294967296#)
    Else
        FnvMul = CLng(r)
    End If
End Function
