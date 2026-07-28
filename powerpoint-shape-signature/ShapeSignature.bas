Attribute VB_Name = "ShapeSignature"
'==============================================================================
' ShapeSignature — PowerPoint 2013 VBA
' Устойчивые сигнатуры Shape для распознавания между презентациями.
'
' Не использует Name / Id (ненадёжны между файлами и при копировании).
' Геометрия нормализуется к размеру слайда (доли), единицы — Points (1/72"),
' поэтому результат не зависит от DPI монитора, ОС и размера листа.
'
' Публичный API:
'   MarkSelectedAsTargets
'   ClearTargetMarks
'   BuildUniqueSignatureFromFiles   — мультивыбор файлов; поиск аналогов по семени/тегам
'   BuildSignatureFromSelection     — быстрый анализ выделения vs соседей
'   ShapeMatchesSignature(shp, sig) As Boolean
'   MakeSignatureFromShape(shp) As String
'   BuildSignatureFromCollections(targets, others) As String
'   GetShapeFeatureString(shp) As String
'   TestMatchSelected
'==============================================================================
Option Explicit

'----- Константы --------------------------------------------------------------
Private Const SIG_VER As String = "S1"
Private Const TAG_TARGET As String = "SHAPESIG_TARGET"
Private Const TAG_TARGET_VAL As String = "1"

' Квантование геометрии: 1 единица = 0.01% стороны слайда (0..10000)
Private Const GEO_SCALE As Long = 10000
' Допуск при сопоставлении геометрии (в тех же единицах). 50 = 0.5% слайда
Private Const GEO_TOL As Long = 50
' Допуск угла (десятые доли градуса): 5 = 0.5°
Private Const ROT_TOL As Long = 5
' Допуск относительного размера шрифта (промилле высоты слайда)
Private Const FONT_TOL As Long = 2

' Битовая маска признаков (порядок фиксирован — не менять!)
Private Const F_TYPE As Long = 1          ' bit 0  — Shape.Type
Private Const F_AUTO As Long = 2          ' bit 1  — AutoShape / ContainedType
Private Const F_PH As Long = 4            ' bit 2  — PlaceholderType
Private Const F_GEO As Long = 8           ' bit 3  — относит. L,T,W,H
Private Const F_AR As Long = 16           ' bit 4  — aspect ratio
Private Const F_ROT As Long = 32          ' bit 5  — rotation
Private Const F_FLIP As Long = 64         ' bit 6  — flip H/V
Private Const F_FILL As Long = 128        ' bit 7  — fill type+RGB
Private Const F_LINE As Long = 256        ' bit 8  — line
Private Const F_CAPS As Long = 512        ' bit 9  — hasText/Table/Chart/Group
Private Const F_TXH As Long = 1024        ' bit 10 — hash текста
Private Const F_FONT As Long = 2048       ' bit 11 — font name hash + rel size
Private Const F_ALL As Long = 4095        ' все биты 0..11
' Маска поиска «похожих» целей в других файлах (без Fill/Text — они могут плавать)
Private Const SEED_FIND_MASK As Long = F_TYPE + F_AUTO + F_PH + F_CAPS + F_AR + F_GEO + F_ROT + F_FLIP

' Порядок добавления признаков при поиске минимальной уникальной маски
Private Const DISC_ORDER As String = "1,2,4,512,16,8,32,64,128,256,2048,1024"

'==============================================================================
' ПУБЛИЧНЫЙ API
'==============================================================================

'--- Пометить выделенные фигуры как целевые (для пакетного анализа) ----------
Public Sub MarkSelectedAsTargets()
    Dim i As Long
    Dim n As Long
    On Error GoTo Fail
    If Not SelectionHasShapes() Then
        MsgBox "Выделите одну или несколько фигур (Shape).", vbExclamation
        Exit Sub
    End If
    With ActiveWindow.Selection.ShapeRange
        For i = 1 To .Count
            SafeTagAdd .Item(i), TAG_TARGET, TAG_TARGET_VAL
            n = n + 1
        Next i
    End With
    MsgBox "Помечено целевых фигур: " & CStr(n) & vbCrLf & _
           "Тег: " & TAG_TARGET & "=" & TAG_TARGET_VAL, vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка MarkSelectedAsTargets: " & Err.Description, vbCritical
End Sub

'--- Снять метки с выделенных (если ничего не выделено — со всего слайда) ----
Public Sub ClearTargetMarks()
    Dim sld As Slide
    Dim shp As Shape
    Dim n As Long
    On Error GoTo Fail

    If SelectionHasShapes() Then
        Dim i As Long
        With ActiveWindow.Selection.ShapeRange
            For i = 1 To .Count
                If SafeTagRemove(.Item(i), TAG_TARGET) Then n = n + 1
            Next i
        End With
    ElseIf ActiveWindow.ViewType = ppViewNormal Or _
           ActiveWindow.ViewType = ppViewSlide Then
        Set sld = ActiveWindow.View.Slide
        For Each shp In sld.Shapes
            If SafeTagRemove(shp, TAG_TARGET) Then n = n + 1
            n = n + ClearTargetMarksInGroup(shp)
        Next shp
    End If

    MsgBox "Снято меток: " & CStr(n), vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка ClearTargetMarks: " & Err.Description, vbCritical
End Sub

'--- Анализ выделенных фигур vs остальные на тех же слайдах ------------------
Public Sub BuildSignatureFromSelection()
    Dim targets As Collection
    Dim others As Collection
    Dim i As Long
    Dim shp As Shape
    Dim sld As Slide
    Dim sig As String
    Dim mask As Long
    Dim report As String

    On Error GoTo Fail
    If Not SelectionHasShapes() Then
        MsgBox "Выделите целевые фигуры.", vbExclamation
        Exit Sub
    End If

    Set targets = New Collection
    Set others = New Collection

    With ActiveWindow.Selection.ShapeRange
        For i = 1 To .Count
            targets.Add .Item(i)
        Next i
    End With

    ' Соседи: все фигуры на слайдах выделенных объектов, кроме самих целей
    Dim seenSlides As Object
    Set seenSlides = CreateObject("Scripting.Dictionary")
    For i = 1 To targets.Count
        Set shp = targets(i)
        Set sld = ParentSlide(shp)
        If Not sld Is Nothing Then
            If Not seenSlides.Exists(CStr(sld.SlideID)) Then
                seenSlides.Add CStr(sld.SlideID), True
                CollectOthersOnSlide sld, targets, others
            End If
        End If
    Next i

    sig = BuildDiscriminatingSignature(targets, others, mask, report)
    PresentSignatureResult sig, mask, targets.Count, others.Count, report
    Exit Sub
Fail:
    MsgBox "Ошибка BuildSignatureFromSelection: " & Err.Description, vbCritical
End Sub

'--- Мультивыбор презентаций → анализ помеченных фигур → сигнатура -----------
Public Sub BuildUniqueSignatureFromFiles()
    Dim fd As FileDialog
    Dim paths As Collection
    Dim v As Variant
    Dim targets As Collection
    Dim others As Collection
    Dim pres As Presentation
    Dim wasOpen As Boolean
    Dim p As String
    Dim sig As String
    Dim mask As Long
    Dim report As String
    Dim openedHere As Collection
    Dim i As Long
    Dim si As Long

    On Error GoTo Fail
    Set paths = New Collection
    Set targets = New Collection
    Set others = New Collection
    Set openedHere = New Collection

    ' 1) Текущее выделение — семена / цели
    If SelectionHasShapes() Then
        With ActiveWindow.Selection.ShapeRange
            For si = 1 To .Count
                targets.Add .Item(si)
            Next si
        End With
    End If

    ' 2) Диалог выбора файлов
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = True
        .Title = "Выберите презентации для анализа сигнатур"
        .Filters.Clear
        .Filters.Add "PowerPoint", "*.pptx;*.ppt;*.pptm;*.ppsx;*.pps", 1
        .Filters.Add "Все файлы", "*.*", 2
        If .Show <> -1 Then
            If targets.Count = 0 Then
                MsgBox "Отменено. Нет целевых фигур." & vbCrLf & _
                       "Выделите фигуры или выполните MarkSelectedAsTargets.", _
                       vbExclamation
                Exit Sub
            End If
        Else
            For Each v In .SelectedItems
                paths.Add CStr(v)
            Next v
        End If
    End With

    ' 3) Семена для поиска аналогов
    Dim seedFeats() As ShapeFeat
    Dim hasSeeds As Boolean
    hasSeeds = False
    If targets.Count > 0 Then
        ReDim seedFeats(1 To targets.Count)
        For i = 1 To targets.Count
            seedFeats(i) = ExtractFeatures(targets(i))
            If seedFeats(i).Valid Then hasSeeds = True
        Next i
    End If

    ' 4) Активная презентация целиком
    If Not ActivePresentation Is Nothing Then
        If hasSeeds Then
            CollectBySeedInPresentation ActivePresentation, seedFeats, targets, others
        Else
            CollectMarkedInPresentation ActivePresentation, targets, others
        End If
    End If

    ' 5) Обход выбранных файлов: теги ИЛИ совпадение с семенем
    Dim activePath As String
    activePath = ""
    On Error Resume Next
    If Not ActivePresentation Is Nothing Then activePath = LCase$(ActivePresentation.FullName)
    On Error GoTo Fail

    For Each v In paths
        p = CStr(v)
        If Len(activePath) > 0 Then
            If LCase$(p) = activePath Then GoTo ContLoop
        End If
        wasOpen = IsPresentationOpen(p)
        If wasOpen Then
            Set pres = FindOpenPresentation(p)
        Else
            Set pres = Presentations.Open(FileName:=p, ReadOnly:=msoTrue, _
                                          Untitled:=msoFalse, WithWindow:=msoFalse)
            openedHere.Add pres
        End If
        If Not pres Is Nothing Then
            If hasSeeds Then
                CollectBySeedInPresentation pres, seedFeats, targets, others
            Else
                CollectMarkedInPresentation pres, targets, others
            End If
        End If
ContLoop:
    Next v

    If targets.Count = 0 Then
        CleanupOpened openedHere
        MsgBox "Целевые фигуры не найдены." & vbCrLf & _
               "Выделите фигуры и/или пометьте их через MarkSelectedAsTargets," & vbCrLf & _
               "затем запустите снова.", vbExclamation
        Exit Sub
    End If

    ' Убрать цели из антипримеров и устранить дубликаты ссылок
    DeduplicateShapes targets
    RemoveTargetsFromOthers targets, others

    sig = BuildDiscriminatingSignature(targets, others, mask, report)
    PresentSignatureResult sig, mask, targets.Count, others.Count, report

    CleanupOpened openedHere
    Exit Sub
Fail:
    On Error Resume Next
    CleanupOpened openedHere
    On Error GoTo 0
    MsgBox "Ошибка BuildUniqueSignatureFromFiles: " & Err.Description, vbCritical
End Sub

'--- Проверка: соответствует ли Shape сигнатуре?  True / False ---------------
' Сигнатура: ByRef Shape + ByVal String → Boolean. Других аргументов нет.
Public Function ShapeMatchesSignature(ByRef shp As Shape, ByVal sig As String) As Boolean
    Dim mask As Long
    Dim feat() As String
    Dim f As ShapeFeat

    ShapeMatchesSignature = False
    If shp Is Nothing Then Exit Function
    If Len(sig) = 0 Then Exit Function

    If Not ParseSignature(sig, mask, feat) Then Exit Function
    f = ExtractFeatures(shp)
    ShapeMatchesSignature = FeaturesMatchMask(f, mask, feat)
End Function

'--- Полный вектор признаков одной фигуры (отладка / ручная сборка) ----------
Public Function GetShapeFeatureString(ByVal shp As Shape) As String
    Dim f As ShapeFeat
    f = ExtractFeatures(shp)
    GetShapeFeatureString = EncodeFeatures(F_ALL, f)
End Function

'--- Тест: выделить фигуру, вставить сигнатуру → True/False ------------------
Public Sub TestMatchSelected()
    Dim sig As String
    Dim shp As Shape
    Dim r As Boolean
    On Error GoTo Fail
    If Not SelectionHasShapes() Then
        MsgBox "Выделите одну фигуру.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    sig = InputBox("Вставьте строку сигнатуры:", "ShapeMatchesSignature")
    If Len(sig) = 0 Then Exit Sub
    r = ShapeMatchesSignature(shp, sig)
    MsgBox "Результат: " & IIf(r, "True", "False"), vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка TestMatchSelected: " & Err.Description, vbCritical
End Sub

'--- Программная сборка сигнатуры одной фигуры (маска = все поля) ---
Public Function MakeSignatureFromShape(ByVal shp As Shape) As String
    Dim f As ShapeFeat
    MakeSignatureFromShape = ""
    If shp Is Nothing Then Exit Function
    f = ExtractFeatures(shp)
    If Not f.Valid Then Exit Function
    MakeSignatureFromShape = BuildSignatureString(F_ALL, f)
End Function

'--- Программная проверка без UI: коллекции целей и антипримеров -------------
Public Function BuildSignatureFromCollections(ByVal targets As Collection, _
                                              ByVal others As Collection) As String
    Dim mask As Long
    Dim report As String
    BuildSignatureFromCollections = BuildDiscriminatingSignature(targets, others, mask, report)
End Function

'==============================================================================
' СТРУКТУРА ПРИЗНАКОВ
'==============================================================================
Private Type ShapeFeat
    TypeId As Long
    AutoId As Long
    PhType As Long
    L As Long          ' geo 0..GEO_SCALE
    T As Long
    W As Long
    H As Long
    AR As Long         ' aspect*1000
    Rot As Long        ' degrees*10
    Flip As Long       ' bit0=H, bit1=V
    FillKey As String  ' "t:rgb" or "t"
    LineKey As String
    Caps As Long       ' bit0=text,1=table,2=chart,3=group
    TxHash As String   ' base36 FNV
    FontHash As String
    FontRel As Long    ' font size / slideH * 1000
    Valid As Boolean
End Type

'==============================================================================
' ИЗВЛЕЧЕНИЕ ПРИЗНАКОВ (нормализованных)
'==============================================================================
Private Function ExtractFeatures(ByVal shp As Shape) As ShapeFeat
    Dim f As ShapeFeat
    Dim sld As Slide
    Dim sw As Double, sh As Double
    Dim effType As Long

    On Error GoTo SoftFail
    f.Valid = False
    If shp Is Nothing Then Exit Function

    Set sld = ParentSlide(shp)
    If sld Is Nothing Then
        ' fallback: берём PageSetup активной презентации
        sw = ActivePresentation.PageSetup.SlideWidth
        sh = ActivePresentation.PageSetup.SlideHeight
    Else
        sw = sld.Parent.PageSetup.SlideWidth
        sh = sld.Parent.PageSetup.SlideHeight
    End If
    If sw <= 0 Or sh <= 0 Then Exit Function

    f.TypeId = CLng(shp.Type)

    ' Эффективный «автотип»: для placeholder — ContainedType
    f.AutoId = -1
    f.PhType = -1
    On Error Resume Next
    If shp.Type = msoPlaceholder Then
        f.PhType = CLng(shp.PlaceholderFormat.Type)
        f.AutoId = CLng(shp.PlaceholderFormat.ContainedType)
    ElseIf shp.Type = msoAutoShape Or shp.Type = msoFreeform Then
        f.AutoId = CLng(shp.AutoShapeType)
    Else
        f.AutoId = CLng(shp.Type)
    End If
    On Error GoTo SoftFail

    f.L = ClampLng(CLng(Round((shp.Left / sw) * GEO_SCALE, 0)), 0, GEO_SCALE * 2)
    f.T = ClampLng(CLng(Round((shp.Top / sh) * GEO_SCALE, 0)), 0, GEO_SCALE * 2)
    f.W = ClampLng(CLng(Round((shp.Width / sw) * GEO_SCALE, 0)), 0, GEO_SCALE * 2)
    f.H = ClampLng(CLng(Round((shp.Height / sh) * GEO_SCALE, 0)), 0, GEO_SCALE * 2)

    If shp.Height > 0.0001 Then
        f.AR = CLng(Round((shp.Width / shp.Height) * 1000#, 0))
    Else
        f.AR = 0
    End If

    f.Rot = CLng(Round(shp.Rotation * 10#, 0))
    f.Flip = 0
    If shp.HorizontalFlip Then f.Flip = f.Flip Or 1
    If shp.VerticalFlip Then f.Flip = f.Flip Or 2

    f.FillKey = ReadFillKey(shp)
    f.LineKey = ReadLineKey(shp)
    f.Caps = ReadCaps(shp)

    f.TxHash = "0"
    f.FontHash = "0"
    f.FontRel = 0
    ReadTextFeatures shp, sh, f

    f.Valid = True
    ExtractFeatures = f
    Exit Function
SoftFail:
    f.Valid = False
    ExtractFeatures = f
End Function

Private Function ReadFillKey(ByVal shp As Shape) As String
    Dim t As Long
    Dim rgbv As Long
    On Error Resume Next
    t = CLng(shp.Fill.Type)
    If Err.Number <> 0 Then
        ReadFillKey = "x"
        Exit Function
    End If
    Err.Clear
    If t = msoFillSolid Then
        rgbv = CLng(shp.Fill.ForeColor.RGB)
        ReadFillKey = CStr(t) & ":" & Hex$(rgbv And &HFFFFFF)
    Else
        ReadFillKey = CStr(t)
    End If
End Function

Private Function ReadLineKey(ByVal shp As Shape) As String
    Dim vis As Long
    Dim rgbv As Long
    Dim w As Long
    On Error Resume Next
    If shp.Line.Visible = msoFalse Then
        ReadLineKey = "0"
        Exit Function
    End If
    vis = 1
    rgbv = CLng(shp.Line.ForeColor.RGB)
    w = CLng(Round(shp.Line.Weight * 100#, 0))
    ReadLineKey = "1:" & Hex$(rgbv And &HFFFFFF) & ":" & ToB36(w)
End Function

Private Function ReadCaps(ByVal shp As Shape) As Long
    Dim c As Long
    On Error Resume Next
    If shp.HasTextFrame = msoTrue Then
        If shp.TextFrame.HasText Then c = c Or 1
    End If
    If shp.HasTable Then c = c Or 2
    If shp.HasChart Then c = c Or 4
    If shp.Type = msoGroup Then c = c Or 8
    ReadCaps = c
End Function

Private Sub ReadTextFeatures(ByVal shp As Shape, ByVal slideH As Double, ByRef f As ShapeFeat)
    Dim txt As String
    Dim fnt As String
    Dim sz As Single
    On Error Resume Next
    If shp.HasTextFrame <> msoTrue Then Exit Sub
    If shp.TextFrame.HasText <> msoTrue Then Exit Sub

    txt = shp.TextFrame.TextRange.Text
    txt = NormalizeText(txt)
    f.TxHash = Fnv1aB36(txt)

    fnt = ""
    sz = 0
    ' Берём шрифт первого символа — быстрее и стабильнее для PPT 2013
    fnt = shp.TextFrame.TextRange.Font.Name
    sz = shp.TextFrame.TextRange.Font.Size
    If Len(fnt) > 0 Then f.FontHash = Fnv1aB36(LCase$(fnt))
    If slideH > 0 And sz > 0 Then
        f.FontRel = CLng(Round((CDbl(sz) / slideH) * 1000#, 0))
    End If
End Sub

Private Function NormalizeText(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, ChrW$(160), " ")
    t = Trim$(t)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeText = LCase$(t)
End Function

'==============================================================================
' КОДИРОВАНИЕ / ДЕКОДИРОВАНИЕ СИГНАТУРЫ
'==============================================================================
' Формат: S1|<maskB36>|<fields...>|<crcB36>
' Поля (в фиксированном порядке битов маски), разделитель «;»
'   TYPE, AUTO, PH, L.T.W.H, AR, ROT, FLIP, FILL, LINE, CAPS, TXH, FONT(hash.rel)

Private Function EncodeFeatures(ByVal mask As Long, ByRef f As ShapeFeat) As String
    Dim parts As String
    parts = ""
    If mask And F_TYPE Then AppendPart parts, ToB36(f.TypeId)
    If mask And F_AUTO Then AppendPart parts, ToB36(f.AutoId)
    If mask And F_PH Then AppendPart parts, ToB36(f.PhType)
    If mask And F_GEO Then AppendPart parts, ToB36(f.L) & "." & ToB36(f.T) & "." & _
                                            ToB36(f.W) & "." & ToB36(f.H)
    If mask And F_AR Then AppendPart parts, ToB36(f.AR)
    If mask And F_ROT Then AppendPart parts, ToB36(f.Rot)
    If mask And F_FLIP Then AppendPart parts, ToB36(f.Flip)
    If mask And F_FILL Then AppendPart parts, f.FillKey
    If mask And F_LINE Then AppendPart parts, f.LineKey
    If mask And F_CAPS Then AppendPart parts, ToB36(f.Caps)
    If mask And F_TXH Then AppendPart parts, f.TxHash
    If mask And F_FONT Then AppendPart parts, f.FontHash & "." & ToB36(f.FontRel)
    EncodeFeatures = parts
End Function

Private Sub AppendPart(ByRef bag As String, ByVal part As String)
    If Len(bag) = 0 Then
        bag = part
    Else
        bag = bag & ";" & part
    End If
End Sub

Private Function BuildSignatureString(ByVal mask As Long, ByRef f As ShapeFeat) As String
    Dim body As String
    Dim crc As String
    body = SIG_VER & "|" & ToB36(mask) & "|" & EncodeFeatures(mask, f)
    crc = Fnv1aB36(body)
    BuildSignatureString = body & "|" & crc
End Function

Private Function ParseSignature(ByVal sig As String, ByRef mask As Long, ByRef feat() As String) As Boolean
    Dim p() As String
    Dim body As String
    Dim crc As String
    Dim fields As String

    ParseSignature = False
    sig = Trim$(sig)
    If Len(sig) = 0 Then Exit Function

    ' Строгий формат: S1|mask|fields|crc  (ровно 4 сегмента)
    p = Split(sig, "|")
    If UBound(p) <> 3 Then Exit Function
    If p(0) <> SIG_VER Then Exit Function

    mask = FromB36(p(1))
    fields = p(2)
    crc = p(3)
    body = p(0) & "|" & p(1) & "|" & p(2)
    If StrComp(Fnv1aB36(body), crc, vbTextCompare) <> 0 Then Exit Function

    If Len(fields) = 0 Then
        ReDim feat(0 To 0)
        feat(0) = ""
    Else
        feat = Split(fields, ";")
    End If
    ParseSignature = True
End Function

Private Function FeaturesMatchMask(ByRef f As ShapeFeat, ByVal mask As Long, ByRef feat() As String) As Boolean
    Dim idx As Long
    Dim geo() As String
    Dim font() As String
    Dim v As Long

    FeaturesMatchMask = False
    If Not f.Valid Then Exit Function
    If mask = 0 Then Exit Function
    idx = 0

    If mask And F_TYPE Then
        If Not Need(feat, idx) Then Exit Function
        If FromB36(feat(idx)) <> f.TypeId Then Exit Function
        idx = idx + 1
    End If
    If mask And F_AUTO Then
        If Not Need(feat, idx) Then Exit Function
        If FromB36(feat(idx)) <> f.AutoId Then Exit Function
        idx = idx + 1
    End If
    If mask And F_PH Then
        If Not Need(feat, idx) Then Exit Function
        If FromB36(feat(idx)) <> f.PhType Then Exit Function
        idx = idx + 1
    End If
    If mask And F_GEO Then
        If Not Need(feat, idx) Then Exit Function
        geo = Split(feat(idx), ".")
        If UBound(geo) <> 3 Then Exit Function
        If Abs(FromB36(geo(0)) - f.L) > GEO_TOL Then Exit Function
        If Abs(FromB36(geo(1)) - f.T) > GEO_TOL Then Exit Function
        If Abs(FromB36(geo(2)) - f.W) > GEO_TOL Then Exit Function
        If Abs(FromB36(geo(3)) - f.H) > GEO_TOL Then Exit Function
        idx = idx + 1
    End If
    If mask And F_AR Then
        If Not Need(feat, idx) Then Exit Function
        v = FromB36(feat(idx))
        ' допуск ~2% к aspect
        If Abs(v - f.AR) > MaxLng(20, CLng(v * 0.02)) Then Exit Function
        idx = idx + 1
    End If
    If mask And F_ROT Then
        If Not Need(feat, idx) Then Exit Function
        If Abs(FromB36(feat(idx)) - f.Rot) > ROT_TOL Then Exit Function
        idx = idx + 1
    End If
    If mask And F_FLIP Then
        If Not Need(feat, idx) Then Exit Function
        If FromB36(feat(idx)) <> f.Flip Then Exit Function
        idx = idx + 1
    End If
    If mask And F_FILL Then
        If Not Need(feat, idx) Then Exit Function
        If StrComp(feat(idx), f.FillKey, vbTextCompare) <> 0 Then Exit Function
        idx = idx + 1
    End If
    If mask And F_LINE Then
        If Not Need(feat, idx) Then Exit Function
        If StrComp(feat(idx), f.LineKey, vbTextCompare) <> 0 Then Exit Function
        idx = idx + 1
    End If
    If mask And F_CAPS Then
        If Not Need(feat, idx) Then Exit Function
        If FromB36(feat(idx)) <> f.Caps Then Exit Function
        idx = idx + 1
    End If
    If mask And F_TXH Then
        If Not Need(feat, idx) Then Exit Function
        If StrComp(feat(idx), f.TxHash, vbTextCompare) <> 0 Then Exit Function
        idx = idx + 1
    End If
    If mask And F_FONT Then
        If Not Need(feat, idx) Then Exit Function
        font = Split(feat(idx), ".")
        If UBound(font) <> 1 Then Exit Function
        If StrComp(font(0), f.FontHash, vbTextCompare) <> 0 Then Exit Function
        If Abs(FromB36(font(1)) - f.FontRel) > FONT_TOL Then Exit Function
        idx = idx + 1
    End If

    FeaturesMatchMask = True
End Function

Private Function Need(ByRef feat() As String, ByVal idx As Long) As Boolean
    On Error Resume Next
    Need = (idx <= UBound(feat))
End Function

'==============================================================================
' ПОИСК МИНИМАЛЬНОЙ УНИКАЛЬНОЙ МАСКИ
'==============================================================================
Private Function BuildDiscriminatingSignature(ByVal targets As Collection, _
                                              ByVal others As Collection, _
                                              ByRef outMask As Long, _
                                              ByRef report As String) As String
    Dim tFeats() As ShapeFeat
    Dim oFeats() As ShapeFeat
    Dim i As Long, j As Long
    Dim bits() As String
    Dim mask As Long
    Dim bit As Long
    Dim okTargets As Boolean
    Dim unique As Boolean
    Dim prototype As ShapeFeat
    Dim sig As String

    ReDim tFeats(1 To targets.Count)
    For i = 1 To targets.Count
        tFeats(i) = ExtractFeatures(targets(i))
        If Not tFeats(i).Valid Then
            report = "Не удалось извлечь признаки у целевой фигуры #" & CStr(i)
            outMask = 0
            BuildDiscriminatingSignature = ""
            Exit Function
        End If
    Next i

    If others.Count > 0 Then
        ReDim oFeats(1 To others.Count)
        For i = 1 To others.Count
            oFeats(i) = ExtractFeatures(others(i))
        Next i
    End If

    ' Проверяем согласованность целей на полном наборе признаков
    If Not AllTargetsAgree(tFeats, F_ALL) Then
        ' Цели различаются — берём пересечение стабильных признаков
        mask = StableMaskAcrossTargets(tFeats)
        If mask = 0 Then
            report = "Целевые фигуры слишком различаются: общий стабильный набор признаков пуст."
            outMask = 0
            BuildDiscriminatingSignature = ""
            Exit Function
        End If
    Else
        mask = 0
    End If

    bits = Split(DISC_ORDER, ",")
    ' Жадное наращивание маски до уникальности относительно others
    unique = False
    For j = 0 To UBound(bits)
        bit = CLng(bits(j))
        If (mask And bit) = 0 Then
            ' добавляем бит только если он стабилен на всех targets
            If AllTargetsAgree(tFeats, mask Or bit) Then
                mask = mask Or bit
            End If
        End If
        If mask <> 0 Then
            If others.Count = 0 Then
                unique = True
                Exit For
            End If
            unique = True
            For i = 1 To others.Count
                If oFeats(i).Valid Then
                    If FeaturesEqualMask(tFeats(1), oFeats(i), mask) Then
                        unique = False
                        Exit For
                    End If
                End If
            Next i
            If unique Then Exit For
        End If
    Next j

    ' Прототип = первая цель (все согласованы по mask)
    prototype = tFeats(1)
    outMask = mask
    sig = BuildSignatureString(mask, prototype)

    report = "Маска=0x" & Hex$(mask) & " (" & MaskDescription(mask) & ")" & vbCrLf & _
             "Уникальна среди соседей: " & IIf(unique, "ДА", "НЕТ — добавьте больше примеров/свойств") & vbCrLf & _
             "Признаки: " & EncodeFeatures(mask, prototype)

    If Not unique Then
        report = report & vbCrLf & _
                 "ВНИМАНИЕ: полной уникальности на обучающей выборке не достигнуто." & vbCrLf & _
                 "Сигнатура всё же построена по стабильным признакам целей."
    End If

    BuildDiscriminatingSignature = sig
End Function

Private Function AllTargetsAgree(ByRef tFeats() As ShapeFeat, ByVal mask As Long) As Boolean
    Dim i As Long
    AllTargetsAgree = True
    For i = 2 To UBound(tFeats)
        If Not FeaturesEqualMask(tFeats(1), tFeats(i), mask) Then
            AllTargetsAgree = False
            Exit Function
        End If
    Next i
End Function

Private Function StableMaskAcrossTargets(ByRef tFeats() As ShapeFeat) As Long
    Dim candidates As Variant
    Dim i As Long
    Dim bit As Long
    Dim m As Long
    candidates = Array(F_TYPE, F_AUTO, F_PH, F_CAPS, F_AR, F_GEO, F_ROT, F_FLIP, F_FILL, F_LINE, F_FONT, F_TXH)
    m = 0
    For i = LBound(candidates) To UBound(candidates)
        bit = CLng(candidates(i))
        If AllTargetsAgree(tFeats, bit) Then m = m Or bit
    Next i
    StableMaskAcrossTargets = m
End Function

Private Function FeaturesEqualMask(ByRef a As ShapeFeat, ByRef b As ShapeFeat, ByVal mask As Long) As Boolean
    ' Используем ту же семантику допуска, что и матчер
    Dim enc As String
    Dim feat() As String
    enc = EncodeFeatures(mask, a)
    If Len(enc) = 0 Then
        FeaturesEqualMask = (mask = 0)
        Exit Function
    End If
    feat = Split(enc, ";")
    FeaturesEqualMask = FeaturesMatchMask(b, mask, feat)
End Function

Private Function MaskDescription(ByVal mask As Long) As String
    Dim s As String
    s = ""
    If mask And F_TYPE Then s = s & "+Type"
    If mask And F_AUTO Then s = s & "+Auto"
    If mask And F_PH Then s = s & "+Ph"
    If mask And F_GEO Then s = s & "+Geo%"
    If mask And F_AR Then s = s & "+AR"
    If mask And F_ROT Then s = s & "+Rot"
    If mask And F_FLIP Then s = s & "+Flip"
    If mask And F_FILL Then s = s & "+Fill"
    If mask And F_LINE Then s = s & "+Line"
    If mask And F_CAPS Then s = s & "+Caps"
    If mask And F_TXH Then s = s & "+TextHash"
    If mask And F_FONT Then s = s & "+Font"
    If Len(s) = 0 Then
        MaskDescription = "(empty)"
    Else
        MaskDescription = Mid$(s, 2)
    End If
End Function

'==============================================================================
' СБОР ФИГУР ИЗ ПРЕЗЕНТАЦИЙ
'==============================================================================
Private Sub CollectMarkedInPresentation(ByVal pres As Presentation, _
                                        ByVal targets As Collection, _
                                        ByVal others As Collection)
    Dim sld As Slide
    Dim shp As Shape
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            CollectShapeTree shp, targets, others, False
        Next shp
    Next sld
End Sub

' Поиск целей по семени (признаки выделенных фигур) + тегам
Private Sub CollectBySeedInPresentation(ByVal pres As Presentation, _
                                        ByRef seedFeats() As ShapeFeat, _
                                        ByVal targets As Collection, _
                                        ByVal others As Collection)
    Dim sld As Slide
    Dim shp As Shape
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            ClassifyShapeBySeed shp, seedFeats, targets, others
            If shp.Type = msoGroup Then
                CollectBySeedGroupItems shp, seedFeats, targets, others
            End If
        Next shp
    Next sld
End Sub

Private Sub ClassifyShapeBySeed(ByVal shp As Shape, _
                                ByRef seedFeats() As ShapeFeat, _
                                ByVal targets As Collection, _
                                ByVal others As Collection)
    Dim f As ShapeFeat
    Dim isTarget As Boolean
    Dim k As Long
    isTarget = IsMarkedTarget(shp)
    If Not isTarget Then
        f = ExtractFeatures(shp)
        If f.Valid Then
            For k = LBound(seedFeats) To UBound(seedFeats)
                If seedFeats(k).Valid Then
                    If FeaturesEqualMask(seedFeats(k), f, SEED_FIND_MASK) Then
                        isTarget = True
                        Exit For
                    End If
                End If
            Next k
        End If
    End If
    If isTarget Then
        targets.Add shp
    Else
        others.Add shp
    End If
End Sub

Private Sub CollectBySeedGroupItems(ByVal grp As Shape, _
                                    ByRef seedFeats() As ShapeFeat, _
                                    ByVal targets As Collection, _
                                    ByVal others As Collection)
    Dim i As Long
    Dim shp As Shape
    On Error Resume Next
    For i = 1 To grp.GroupItems.Count
        Set shp = grp.GroupItems(i)
        ClassifyShapeBySeed shp, seedFeats, targets, others
        If shp.Type = msoGroup Then
            CollectBySeedGroupItems shp, seedFeats, targets, others
        End If
    Next i
End Sub

Private Sub CollectShapeTree(ByVal shp As Shape, _
                             ByVal targets As Collection, _
                             ByVal others As Collection, _
                             ByVal ancestorMarked As Boolean)
    Dim i As Long
    ' Метка не наследуется на детей: помечается только явно выбранный объект
    If IsMarkedTarget(shp) Then
        targets.Add shp
    Else
        others.Add shp
    End If

    If shp.Type = msoGroup Then
        On Error Resume Next
        For i = 1 To shp.GroupItems.Count
            CollectShapeTree shp.GroupItems(i), targets, others, False
        Next i
        On Error GoTo 0
    End If
End Sub

Private Sub CollectOthersOnSlide(ByVal sld As Slide, ByVal targets As Collection, ByVal others As Collection)
    Dim shp As Shape
    Dim i As Long
    Dim isTarget As Boolean
    Dim t As Shape
    For Each shp In sld.Shapes
        isTarget = False
        For i = 1 To targets.Count
            Set t = targets(i)
            If IsSameShapeRef(t, shp) Then
                isTarget = True
                Exit For
            End If
        Next i
        If Not isTarget Then others.Add shp
        ' Разбор группы: вложенные элементы как соседи, если не в targets
        If shp.Type = msoGroup Then
            CollectGroupOthers shp, targets, others
        End If
    Next shp
End Sub

Private Sub CollectGroupOthers(ByVal grp As Shape, ByVal targets As Collection, ByVal others As Collection)
    Dim i As Long, j As Long
    Dim gi As Shape
    Dim isTarget As Boolean
    On Error Resume Next
    For i = 1 To grp.GroupItems.Count
        Set gi = grp.GroupItems(i)
        isTarget = False
        For j = 1 To targets.Count
            If IsSameShapeRef(targets(j), gi) Then
                isTarget = True
                Exit For
            End If
        Next j
        If Not isTarget Then others.Add gi
        If gi.Type = msoGroup Then CollectGroupOthers gi, targets, others
    Next i
End Sub


Private Sub DeduplicateShapes(ByVal col As Collection)
    Dim out As Collection
    Dim i As Long, j As Long
    Dim shp As Shape
    Dim dup As Boolean
    Set out = New Collection
    For i = 1 To col.Count
        Set shp = col(i)
        dup = False
        For j = 1 To out.Count
            If IsSameShapeRef(out(j), shp) Then
                dup = True
                Exit For
            End If
        Next j
        If Not dup Then out.Add shp
    Next i
    ' переписать исходную коллекцию
    Do While col.Count > 0
        col.Remove 1
    Loop
    For i = 1 To out.Count
        col.Add out(i)
    Next i
End Sub

Private Sub RemoveTargetsFromOthers(ByVal targets As Collection, ByVal others As Collection)
    Dim i As Long, j As Long
    Dim keep As Collection
    Dim o As Shape
    Dim isT As Boolean
    Set keep = New Collection
    For i = 1 To others.Count
        Set o = others(i)
        isT = False
        For j = 1 To targets.Count
            If IsSameShapeRef(targets(j), o) Then
                isT = True
                Exit For
            End If
        Next j
        If Not isT Then keep.Add o
    Next i
    Do While others.Count > 0
        others.Remove 1
    Loop
    For i = 1 To keep.Count
        others.Add keep(i)
    Next i
End Sub

Private Function IsSameShapeRef(ByVal a As Shape, ByVal b As Shape) As Boolean
    Dim sa As Slide
    Dim sb As Slide
    On Error Resume Next
    IsSameShapeRef = False
    If a Is Nothing Or b Is Nothing Then Exit Function
    If a.Id <> b.Id Then Exit Function
    Set sa = ParentSlide(a)
    Set sb = ParentSlide(b)
    If sa Is Nothing Or sb Is Nothing Then Exit Function
    IsSameShapeRef = (sa.SlideID = sb.SlideID)
End Function

Private Function IsMarkedTarget(ByVal shp As Shape) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 1 To shp.Tags.Count
        If StrComp(shp.Tags.Name(i), TAG_TARGET, vbTextCompare) = 0 Then
            If StrComp(shp.Tags.Value(i), TAG_TARGET_VAL, vbTextCompare) = 0 Then
                IsMarkedTarget = True
                Exit Function
            End If
        End If
    Next i
    IsMarkedTarget = False
End Function

Private Function ClearTargetMarksInGroup(ByVal shp As Shape) As Long
    Dim i As Long
    Dim n As Long
    On Error Resume Next
    If shp.Type <> msoGroup Then Exit Function
    For i = 1 To shp.GroupItems.Count
        If SafeTagRemove(shp.GroupItems(i), TAG_TARGET) Then n = n + 1
        n = n + ClearTargetMarksInGroup(shp.GroupItems(i))
    Next i
    ClearTargetMarksInGroup = n
End Function

'==============================================================================
' ВСПОМОГАТЕЛЬНЫЕ
'==============================================================================
Private Sub PresentSignatureResult(ByVal sig As String, ByVal mask As Long, _
                                   ByVal nTargets As Long, ByVal nOthers As Long, _
                                   ByVal report As String)
    Dim msg As String
    If Len(sig) = 0 Then
        MsgBox report, vbExclamation, "Сигнатура не построена"
        Exit Sub
    End If
    On Error Resume Next
    CopyToClipboard sig
    On Error GoTo 0

    msg = "Целей: " & CStr(nTargets) & " | Соседей (антипримеры): " & CStr(nOthers) & vbCrLf & _
          report & vbCrLf & vbCrLf & _
          "Сигнатура (скопирована в буфер, если доступно):" & vbCrLf & sig

    ' InputBox удобен для копирования длинной строки
    InputBox msg, "ShapeSignature", sig
End Sub

Private Sub CopyToClipboard(ByVal s As String)
    Dim clip As Object
    ' MSForms.DataObject через ProgID-класс (без обязательной ссылки)
    Set clip = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clip.SetText s
    clip.PutInClipboard
End Sub

Private Function SelectionHasShapes() As Boolean
    On Error Resume Next
    SelectionHasShapes = False
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        SelectionHasShapes = (ActiveWindow.Selection.ShapeRange.Count > 0)
    End If
End Function

Private Function ParentSlide(ByVal shp As Shape) As Slide
    Dim p As Object
    On Error Resume Next
    Set p = shp.Parent
    ' Parent может быть Slide или GroupShape — поднимаемся вверх
    Do While Not p Is Nothing
        If TypeName(p) = "Slide" Then
            Set ParentSlide = p
            Exit Function
        End If
        Set p = p.Parent
    Loop
    Set ParentSlide = Nothing
End Function

Private Sub SafeTagAdd(ByVal shp As Shape, ByVal nm As String, ByVal val As String)
    On Error Resume Next
    shp.Tags.Delete nm
    shp.Tags.Add nm, val
End Sub

Private Function SafeTagRemove(ByVal shp As Shape, ByVal nm As String) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 1 To shp.Tags.Count
        If StrComp(shp.Tags.Name(i), nm, vbTextCompare) = 0 Then
            shp.Tags.Delete nm
            SafeTagRemove = True
            Exit Function
        End If
    Next i
    SafeTagRemove = False
End Function

Private Function IsPresentationOpen(ByVal fullPath As String) As Boolean
    IsPresentationOpen = Not (FindOpenPresentation(fullPath) Is Nothing)
End Function

Private Function FindOpenPresentation(ByVal fullPath As String) As Presentation
    Dim p As Presentation
    Dim target As String
    target = LCase$(fullPath)
    For Each p In Presentations
        On Error Resume Next
        If LCase$(p.FullName) = target Then
            Set FindOpenPresentation = p
            Exit Function
        End If
        On Error GoTo 0
    Next p
    Set FindOpenPresentation = Nothing
End Function

Private Function GetFileName(ByVal fullPath As String) As String
    Dim i As Long
    i = InStrRev(fullPath, "\")
    If i = 0 Then i = InStrRev(fullPath, "/")
    If i > 0 Then
        GetFileName = Mid$(fullPath, i + 1)
    Else
        GetFileName = fullPath
    End If
End Function

Private Sub CleanupOpened(ByVal openedHere As Collection)
    Dim i As Long
    Dim p As Presentation
    On Error Resume Next
    If openedHere Is Nothing Then Exit Sub
    For i = openedHere.Count To 1 Step -1
        Set p = openedHere(i)
        p.Close
    Next i
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

Private Function MaxLng(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLng = a Else MaxLng = b
End Function

'----- Base36 / FNV-1a (32-bit) ----------------------------------------------
Private Function ToB36(ByVal n As Long) As String
    Dim digits As String
    Dim neg As Boolean
    Dim r As Long
    Dim v As Long
    digits = "0123456789abcdefghijklmnopqrstuvwxyz"
    If n = 0 Then
        ToB36 = "0"
        Exit Function
    End If
    neg = (n < 0)
    ' Работаем через абсолютное, избегая Overflow на Long Min
    If neg Then
        v = -n
    Else
        v = n
    End If
    ToB36 = ""
    Do While v > 0
        r = v Mod 36
        ToB36 = Mid$(digits, r + 1, 1) & ToB36
        v = v \ 36
    Loop
    If neg Then ToB36 = "-" & ToB36
End Function

Private Function FromB36(ByVal s As String) As Long
    Dim digits As String
    Dim i As Long
    Dim c As String
    Dim v As Long
    Dim neg As Boolean
    Dim p As Long
    digits = "0123456789abcdefghijklmnopqrstuvwxyz"
    s = LCase$(Trim$(s))
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "-" Then
        neg = True
        s = Mid$(s, 2)
    End If
    v = 0
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        p = InStr(1, digits, c, vbBinaryCompare) - 1
        If p < 0 Then
            FromB36 = 0
            Exit Function
        End If
        v = v * 36 + p
    Next i
    If neg Then v = -v
    FromB36 = v
End Function

Private Function Fnv1aB36(ByVal s As String) As String
    ' FNV-1a 32-bit → unsigned → Base36 (компактно, быстро)
    Dim h As Long
    Dim i As Long
    Dim b As Long
    Dim uh As Double
    h = -2128831035 ' &H811C9DC5 as signed Long
    For i = 1 To Len(s)
        b = AscW(Mid$(s, i, 1)) And &HFF&
        h = (h Xor b)
        h = FnvMul(h)
    Next i
    ' интерпретация как unsigned
    If h < 0 Then
        uh = CDbl(h And &H7FFFFFFF) + 2147483648#
    Else
        uh = CDbl(h)
    End If
    Fnv1aB36 = ToB36Unsigned(uh)
End Function

Private Function FnvMul(ByVal h As Long) As Long
    ' h * FNV_prime(16777619) mod 2^32 через пошаговое умножение
    Dim r As Long
    r = h
    r = LngMulAdd(r, 16777619, 0)
    FnvMul = r
End Function

Private Function LngMulAdd(ByVal a As Long, ByVal m As Long, ByVal addv As Long) As Long
    ' (a * m + addv) mod 2^32, результат как signed Long
    Dim da As Double, dm As Double, dr As Double
    Dim hi As Double
    If a < 0 Then
        da = CDbl(a And &H7FFFFFFF) + 2147483648#
    Else
        da = CDbl(a)
    End If
    dm = CDbl(m)
    dr = da * dm + CDbl(addv)
    hi = Int(dr / 4294967296#)
    dr = dr - hi * 4294967296#
    If dr >= 2147483648# Then
        LngMulAdd = CLng(dr - 4294967296#)
    Else
        LngMulAdd = CLng(dr)
    End If
End Function

Private Function ToB36Unsigned(ByVal uh As Double) As String
    Dim digits As String
    Dim r As Long
    Dim s As String
    digits = "0123456789abcdefghijklmnopqrstuvwxyz"
    If uh <= 0 Then
        ToB36Unsigned = "0"
        Exit Function
    End If
    s = ""
    Do While uh >= 1
        r = CLng(uh - Int(uh / 36#) * 36#)
        s = Mid$(digits, r + 1, 1) & s
        uh = Int(uh / 36#)
    Loop
    ToB36Unsigned = s
End Function
