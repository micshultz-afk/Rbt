Attribute VB_Name = "ShapeSignature"
'==============================================================================
' ShapeSignature — PowerPoint 2013 VBA (offline / Secret Net friendly)
'
' Ограничения среды:
'   - без Интернета
'   - без CreateObject / Scripting / MSForms / внешних COM
'   - только объектная модель PowerPoint + чистый VBA
'
' Публичный API:
'   ShapeMatchesSignature(ByRef shp, ByVal sig) As Boolean
'   BuildSignatureByShapeName(ByVal shapeName, ByRef pres) As String
'   BuildSignatureFromSelection
'   BuildUniqueSignatureFromFiles
'   MakeSignatureFromShape / GetShapeFeatureString / TestMatchSelected
'   MarkSelectedAsTargets / ClearTargetMarks
'==============================================================================
Option Explicit

Private Const SIG_VER As String = "S2"
Private Const SIG_VER_LEGACY As String = "S1"
Private Const TAG_TARGET As String = "SHAPESIG_TARGET"
Private Const TAG_TARGET_VAL As String = "1"

Private Const GEO_SCALE As Long = 10000
Private Const GEO_TOL As Long = 50
Private Const ROT_TOL As Long = 5
Private Const FONT_TOL As Long = 2

' Биты маски S2 (порядок фиксирован)
Private Const F_TYPE As Long = 1
Private Const F_AUTO As Long = 2
Private Const F_PH As Long = 4
Private Const F_POS As Long = 8          ' относит. Left, Top
Private Const F_SIZE As Long = 16       ' относит. Width, Height
Private Const F_AR As Long = 32
Private Const F_ROT As Long = 64
Private Const F_FLIP As Long = 128
Private Const F_FILL As Long = 256
Private Const F_LINE As Long = 512
Private Const F_CAPS As Long = 1024
Private Const F_TXH As Long = 2048
Private Const F_FONT As Long = 4096
Private Const F_ALL As Long = 8191

' Поиск аналогов между файлами (без Fill/Text/POS — они чаще плавают)
Private Const SEED_FIND_MASK As Long = F_TYPE + F_AUTO + F_PH + F_CAPS + F_SIZE + F_AR + F_ROT + F_FLIP

' Для объектов с одним именем на многих слайдах: сначала «кто», потом «где»
Private Const DISC_ORDER As String = "1,2,4,1024,16,32,64,128,256,512,4096,2048,8"

Private Type ShapeFeat
    TypeId As Long
    AutoId As Long
    PhType As Long
    L As Long
    T As Long
    W As Long
    H As Long
    AR As Long
    Rot As Long
    Flip As Long
    FillType As Long
    FillRgb As Long       ' -1 если не RGB
    FillScheme As Long    ' -1 если не scheme/theme
    LineVis As Long       ' 0/1
    LineRgb As Long
    LineWeight As Long    ' *100
    Caps As Long
    TxFnv As Long
    FontFnv As Long
    FontRel As Long
    Valid As Boolean
End Type

'==============================================================================
' ПУБЛИЧНЫЙ API
'==============================================================================

'--- Главная функция распознавания: ByRef Shape + ByVal sig → True/False ------
Public Function ShapeMatchesSignature(ByRef shp As Shape, ByVal sig As String) As Boolean
    Dim mask As Long
    Dim proto As ShapeFeat
    Dim f As ShapeFeat

    On Error GoTo SoftFail
    ShapeMatchesSignature = False
    If shp Is Nothing Then Exit Function
    If Len(sig) = 0 Then Exit Function
    If Not ParseSignatureToFeat(sig, mask, proto) Then Exit Function
    If mask = 0 Then Exit Function

    f = ExtractFeatures(shp)
    If Not f.Valid Then Exit Function
    ShapeMatchesSignature = FeaturesMatchFeat(f, mask, proto)
    Exit Function
SoftFail:
    ShapeMatchesSignature = False
End Function

'--- Сигнатура по уникальному имени Shape внутри презентации -----------------
' Находит ВСЕ фигуры с данным Name (в т.ч. на разных слайдах и в группах),
' берёт только стабильные между ними признаки и отсекает остальные объекты.
' Позиция (POS) включается только если она совпадает у всех экземпляров.
Public Function BuildSignatureByShapeName(ByVal shapeName As String, _
                                          ByRef pres As Presentation) As String
    Dim targets As Collection
    Dim others As Collection
    Dim mask As Long
    Dim report As String
    Dim sig As String

    On Error GoTo SoftFail
    BuildSignatureByShapeName = ""

    shapeName = Trim$(shapeName)
    If Len(shapeName) = 0 Then Exit Function
    If pres Is Nothing Then Exit Function

    Set targets = New Collection
    Set others = New Collection

    If Not CollectByNameInPresentation(pres, shapeName, targets, others) Then Exit Function
    If targets.Count = 0 Then Exit Function

    DeduplicateShapes targets
    DeduplicateShapes others
    RemoveTargetsFromOthers targets, others

    sig = BuildDiscriminatingSignature(targets, others, mask, report)
    BuildSignatureByShapeName = sig
    Exit Function
SoftFail:
    BuildSignatureByShapeName = ""
End Function

'--- То же + отчёт (для отладки из макроса) ----------------------------------
Public Sub BuildSignatureByShapeNameUI()
    Dim nm As String
    Dim sig As String
    Dim pres As Presentation
    Dim targets As Collection
    Dim others As Collection
    Dim mask As Long
    Dim report As String

    On Error GoTo Fail
    Set pres = ActivePresentation
    If pres Is Nothing Then
        MsgBox "Нет активной презентации.", vbExclamation
        Exit Sub
    End If

    nm = InputBox("Имя Shape (как в области выделения):", "BuildSignatureByShapeName")
    If Len(Trim$(nm)) = 0 Then Exit Sub

    Set targets = New Collection
    Set others = New Collection
    If Not CollectByNameInPresentation(pres, Trim$(nm), targets, others) Then
        MsgBox "Ошибка обхода презентации.", vbCritical
        Exit Sub
    End If
    If targets.Count = 0 Then
        MsgBox "Фигуры с именем «" & nm & "» не найдены.", vbExclamation
        Exit Sub
    End If

    DeduplicateShapes targets
    DeduplicateShapes others
    RemoveTargetsFromOthers targets, others

    sig = BuildDiscriminatingSignature(targets, others, mask, report)
    PresentSignatureResult sig, mask, targets.Count, others.Count, report
    Exit Sub
Fail:
    MsgBox "Ошибка BuildSignatureByShapeNameUI: " & Err.Description, vbCritical
End Sub

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
    MsgBox "Помечено целевых фигур: " & CStr(n), vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка MarkSelectedAsTargets: " & Err.Description, vbCritical
End Sub

Public Sub ClearTargetMarks()
    Dim sld As Slide
    Dim shp As Shape
    Dim n As Long
    Dim i As Long
    On Error GoTo Fail

    If SelectionHasShapes() Then
        With ActiveWindow.Selection.ShapeRange
            For i = 1 To .Count
                If SafeTagRemove(.Item(i), TAG_TARGET) Then n = n + 1
            Next i
        End With
    Else
        On Error Resume Next
        Set sld = ActiveWindow.View.Slide
        On Error GoTo Fail
        If sld Is Nothing Then
            MsgBox "Нет активного слайда и нет выделения.", vbExclamation
            Exit Sub
        End If
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

Public Sub BuildSignatureFromSelection()
    Dim targets As Collection
    Dim others As Collection
    Dim i As Long
    Dim shp As Shape
    Dim sld As Slide
    Dim sig As String
    Dim mask As Long
    Dim report As String
    Dim seen As Collection

    On Error GoTo Fail
    If Not SelectionHasShapes() Then
        MsgBox "Выделите целевые фигуры.", vbExclamation
        Exit Sub
    End If

    Set targets = New Collection
    Set others = New Collection
    Set seen = New Collection

    With ActiveWindow.Selection.ShapeRange
        For i = 1 To .Count
            targets.Add .Item(i)
        Next i
    End With

    For i = 1 To targets.Count
        Set shp = targets(i)
        Set sld = ParentSlide(shp)
        If Not sld Is Nothing Then
            If Not CollectionHasKey(seen, CStr(sld.SlideID)) Then
                CollectionAddKey seen, CStr(sld.SlideID)
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

Public Sub BuildUniqueSignatureFromFiles()
    Dim fd As FileDialog
    Dim paths As Collection
    Dim v As Variant
    Dim targets As Collection
    Dim others As Collection
    Dim seedTargets As Collection
    Dim pres As Presentation
    Dim wasOpen As Boolean
    Dim p As String
    Dim sig As String
    Dim mask As Long
    Dim report As String
    Dim openedHere As Collection
    Dim i As Long
    Dim si As Long
    Dim errNum As Long
    Dim errDesc As String
    Dim seedFeats() As ShapeFeat
    Dim hasSeeds As Boolean
    Dim activePath As String

    On Error GoTo Fail
    Set paths = New Collection
    Set targets = New Collection
    Set others = New Collection
    Set seedTargets = New Collection
    Set openedHere = New Collection

    If SelectionHasShapes() Then
        With ActiveWindow.Selection.ShapeRange
            For si = 1 To .Count
                seedTargets.Add .Item(si)
                targets.Add .Item(si)
            Next si
        End With
    End If

    If Not ActivePresentation Is Nothing Then
        CollectMarkedTargetsOnly ActivePresentation, seedTargets
        CollectMarkedTargetsOnly ActivePresentation, targets
    End If
    DeduplicateShapes seedTargets
    DeduplicateShapes targets

    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If fd Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        On Error GoTo Fail
        If targets.Count = 0 Then
            MsgBox "Диалог файлов недоступен и нет целевых фигур.", vbExclamation
            Exit Sub
        End If
        GoTo AfterDialog
    End If
    Err.Clear
    On Error GoTo Fail

    With fd
        .AllowMultiSelect = True
        .Title = "Выберите презентации для анализа сигнатур"
        On Error Resume Next
        .Filters.Clear
        .Filters.Add "PowerPoint", "*.pptx;*.ppt;*.pptm;*.ppsx;*.pps", 1
        .Filters.Add "Все файлы", "*.*", 2
        Err.Clear
        On Error GoTo Fail
        If .Show <> -1 Then
            If targets.Count = 0 Then
                MsgBox "Отменено. Нет целевых фигур.", vbExclamation
                Exit Sub
            End If
        Else
            For Each v In .SelectedItems
                paths.Add CStr(v)
            Next v
        End If
    End With

AfterDialog:
    hasSeeds = False
    If seedTargets.Count > 0 Then
        ReDim seedFeats(1 To seedTargets.Count)
        For i = 1 To seedTargets.Count
            seedFeats(i) = ExtractFeatures(seedTargets(i))
            If seedFeats(i).Valid Then hasSeeds = True
        Next i
    End If

    If Not ActivePresentation Is Nothing Then
        If hasSeeds Then
            CollectBySeedInPresentation ActivePresentation, seedFeats, targets, others
        Else
            CollectMarkedInPresentation ActivePresentation, targets, others
        End If
    End If

    activePath = ""
    On Error Resume Next
    If Not ActivePresentation Is Nothing Then activePath = LCase$(ActivePresentation.FullName)
    Err.Clear
    On Error GoTo Fail

    For Each v In paths
        p = CStr(v)
        If Len(activePath) > 0 Then
            If LCase$(p) = activePath Then GoTo ContLoop
        End If

        wasOpen = IsPresentationOpen(p)
        Set pres = Nothing
        On Error Resume Next
        If wasOpen Then
            Set pres = FindOpenPresentation(p)
        Else
            Set pres = Presentations.Open(FileName:=p, ReadOnly:=msoTrue, _
                                          Untitled:=msoFalse, WithWindow:=msoFalse)
            If Not pres Is Nothing Then openedHere.Add pres
        End If
        Err.Clear
        On Error GoTo Fail

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
        MsgBox "Целевые фигуры не найдены.", vbExclamation
        Exit Sub
    End If

    DeduplicateShapes targets
    RemoveTargetsFromOthers targets, others

    sig = BuildDiscriminatingSignature(targets, others, mask, report)
    PresentSignatureResult sig, mask, targets.Count, others.Count, report
    CleanupOpened openedHere
    Exit Sub
Fail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    CleanupOpened openedHere
    On Error GoTo 0
    MsgBox "Ошибка BuildUniqueSignatureFromFiles (" & CStr(errNum) & "): " & errDesc, vbCritical
End Sub

Public Function MakeSignatureFromShape(ByRef shp As Shape) As String
    Dim f As ShapeFeat
    On Error GoTo SoftFail
    MakeSignatureFromShape = ""
    If shp Is Nothing Then Exit Function
    f = ExtractFeatures(shp)
    If Not f.Valid Then Exit Function
    MakeSignatureFromShape = BuildSignatureString(F_ALL, f)
    Exit Function
SoftFail:
    MakeSignatureFromShape = ""
End Function

Public Function BuildSignatureFromCollections(ByVal targets As Collection, _
                                              ByVal others As Collection) As String
    Dim mask As Long
    Dim report As String
    On Error GoTo SoftFail
    BuildSignatureFromCollections = BuildDiscriminatingSignature(targets, others, mask, report)
    Exit Function
SoftFail:
    BuildSignatureFromCollections = ""
End Function

Public Function GetShapeFeatureString(ByRef shp As Shape) As String
    Dim f As ShapeFeat
    On Error GoTo SoftFail
    GetShapeFeatureString = ""
    If shp Is Nothing Then Exit Function
    f = ExtractFeatures(shp)
    If Not f.Valid Then Exit Function
    GetShapeFeatureString = "T=" & CStr(f.TypeId) & ";A=" & CStr(f.AutoId) & _
        ";Ph=" & CStr(f.PhType) & ";Pos=" & CStr(f.L) & "," & CStr(f.T) & _
        ";Size=" & CStr(f.W) & "," & CStr(f.H) & ";AR=" & CStr(f.AR) & _
        ";Rot=" & CStr(f.Rot) & ";Flip=" & CStr(f.Flip) & _
        ";Fill=" & CStr(f.FillType) & "/" & CStr(f.FillRgb) & "/" & CStr(f.FillScheme) & _
        ";Line=" & CStr(f.LineVis) & "/" & CStr(f.LineRgb) & "/" & CStr(f.LineWeight) & _
        ";Caps=" & CStr(f.Caps) & ";Tx=" & Hex$(f.TxFnv) & _
        ";Font=" & Hex$(f.FontFnv) & "/" & CStr(f.FontRel)
    Exit Function
SoftFail:
    GetShapeFeatureString = ""
End Function

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

'==============================================================================
' ИЗВЛЕЧЕНИЕ ПРИЗНАКОВ
'==============================================================================
Private Function ExtractFeatures(ByVal shp As Shape) As ShapeFeat
    Dim f As ShapeFeat
    Dim sld As Slide
    Dim sw As Double, sh As Double
    Dim pres As Presentation

    On Error GoTo SoftFail
    f.Valid = False
    InitFeat f
    If shp Is Nothing Then Exit Function

    Set sld = ParentSlide(shp)
    If sld Is Nothing Then
        On Error Resume Next
        Set pres = ActivePresentation
        If pres Is Nothing Then GoTo SoftFail
        sw = pres.PageSetup.SlideWidth
        sh = pres.PageSetup.SlideHeight
        Err.Clear
        On Error GoTo SoftFail
    Else
        On Error Resume Next
        Set pres = sld.Parent
        If pres Is Nothing Then GoTo SoftFail
        sw = pres.PageSetup.SlideWidth
        sh = pres.PageSetup.SlideHeight
        Err.Clear
        On Error GoTo SoftFail
    End If
    If sw <= 0# Or sh <= 0# Then Exit Function

    f.TypeId = CLng(shp.Type)

    f.AutoId = -1
    f.PhType = -1
    On Error Resume Next
    If shp.Type = msoPlaceholder Then
        f.PhType = CLng(shp.PlaceholderFormat.Type)
        If Err.Number <> 0 Then
            Err.Clear
            f.PhType = -1
        End If
        f.AutoId = CLng(shp.PlaceholderFormat.ContainedType)
        If Err.Number <> 0 Then
            Err.Clear
            f.AutoId = -1
        End If
    ElseIf shp.Type = msoAutoShape Or shp.Type = msoFreeform Then
        f.AutoId = CLng(shp.AutoShapeType)
        If Err.Number <> 0 Then
            Err.Clear
            f.AutoId = -1
        End If
    Else
        f.AutoId = CLng(shp.Type)
    End If
    Err.Clear
    On Error GoTo SoftFail

    f.L = ClampLng(CLng(Round((shp.Left / sw) * GEO_SCALE, 0)), -GEO_SCALE, GEO_SCALE * 2)
    f.T = ClampLng(CLng(Round((shp.Top / sh) * GEO_SCALE, 0)), -GEO_SCALE, GEO_SCALE * 2)
    f.W = ClampLng(CLng(Round((shp.Width / sw) * GEO_SCALE, 0)), 0, GEO_SCALE * 2)
    f.H = ClampLng(CLng(Round((shp.Height / sh) * GEO_SCALE, 0)), 0, GEO_SCALE * 2)

    If shp.Height > 0.0001 Then
        f.AR = CLng(Round((shp.Width / shp.Height) * 1000#, 0))
    Else
        f.AR = 0
    End If

    On Error Resume Next
    f.Rot = CLng(Round(shp.Rotation * 10#, 0))
    If Err.Number <> 0 Then
        Err.Clear
        f.Rot = 0
    End If
    f.Flip = 0
    If shp.HorizontalFlip Then f.Flip = f.Flip Or 1
    If shp.VerticalFlip Then f.Flip = f.Flip Or 2
    Err.Clear
    On Error GoTo SoftFail

    ReadFillFeat shp, f
    ReadLineFeat shp, f
    f.Caps = ReadCaps(shp)
    ReadTextFeatures shp, sh, f

    f.Valid = True
    ExtractFeatures = f
    Exit Function
SoftFail:
    f.Valid = False
    ExtractFeatures = f
End Function

Private Sub InitFeat(ByRef f As ShapeFeat)
    f.TypeId = 0: f.AutoId = -1: f.PhType = -1
    f.L = 0: f.T = 0: f.W = 0: f.H = 0: f.AR = 0
    f.Rot = 0: f.Flip = 0
    f.FillType = -1: f.FillRgb = -1: f.FillScheme = -1
    f.LineVis = 0: f.LineRgb = -1: f.LineWeight = 0
    f.Caps = 0: f.TxFnv = 0: f.FontFnv = 0: f.FontRel = 0
    f.Valid = False
End Sub

Private Sub ReadFillFeat(ByVal shp As Shape, ByRef f As ShapeFeat)
    Dim ct As Long
    On Error Resume Next
    f.FillType = CLng(shp.Fill.Type)
    If Err.Number <> 0 Then
        Err.Clear
        f.FillType = -1
        Exit Sub
    End If

    f.FillRgb = -1
    f.FillScheme = -1
    ct = CLng(shp.Fill.ForeColor.Type)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    If ct = msoColorTypeRGB Then
        f.FillRgb = CLng(shp.Fill.ForeColor.RGB) And &HFFFFFF
    ElseIf ct = msoColorTypeScheme Then
        f.FillScheme = CLng(shp.Fill.ForeColor.SchemeColor)
        ' дополнительно фиксируем разрешённый RGB, если доступен
        f.FillRgb = CLng(shp.Fill.ForeColor.RGB) And &HFFFFFF
        If Err.Number <> 0 Then
            Err.Clear
            f.FillRgb = -1
        End If
    Else
        f.FillRgb = CLng(shp.Fill.ForeColor.RGB) And &HFFFFFF
        If Err.Number <> 0 Then
            Err.Clear
            f.FillRgb = -1
        End If
    End If
    Err.Clear
End Sub

Private Sub ReadLineFeat(ByVal shp As Shape, ByRef f As ShapeFeat)
    Dim ct As Long
    On Error Resume Next
    f.LineVis = 0
    f.LineRgb = -1
    f.LineWeight = 0
    If shp.Line.Visible <> msoTrue Then
        Err.Clear
        Exit Sub
    End If
    f.LineVis = 1
    f.LineWeight = CLng(Round(shp.Line.Weight * 100#, 0))
    If Err.Number <> 0 Then
        Err.Clear
        f.LineWeight = 0
    End If

    ct = CLng(shp.Line.ForeColor.Type)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    If ct = msoColorTypeRGB Or ct = msoColorTypeScheme Then
        f.LineRgb = CLng(shp.Line.ForeColor.RGB) And &HFFFFFF
        If Err.Number <> 0 Then
            Err.Clear
            f.LineRgb = -1
        End If
    End If
    Err.Clear
End Sub

Private Function ReadCaps(ByVal shp As Shape) As Long
    Dim c As Long
    On Error Resume Next
    If shp.HasTextFrame = msoTrue Then
        If shp.TextFrame.HasText Then c = c Or 1
    End If
    If (c And 1) = 0 Then
        If shp.TextFrame2.HasText Then c = c Or 1
    End If
    If shp.HasTable Then c = c Or 2
    If shp.HasChart Then c = c Or 4
    If shp.Type = msoGroup Then c = c Or 8
    ReadCaps = c
    Err.Clear
End Function

Private Sub ReadTextFeatures(ByVal shp As Shape, ByVal slideH As Double, ByRef f As ShapeFeat)
    Dim txt As String
    Dim fnt As String
    Dim sz As Single
    Dim got As Boolean

    On Error Resume Next
    got = False
    txt = ""
    fnt = ""
    sz = 0

    If shp.HasTextFrame = msoTrue Then
        If shp.TextFrame.HasText Then
            txt = shp.TextFrame.TextRange.Text
            fnt = shp.TextFrame.TextRange.Font.Name
            sz = shp.TextFrame.TextRange.Font.Size
            got = True
        End If
    End If

    If Not got Then
        If shp.TextFrame2.HasText Then
            txt = shp.TextFrame2.TextRange.Text
            fnt = CStr(shp.TextFrame2.TextRange.Font.Name)
            sz = shp.TextFrame2.TextRange.Font.Size
            got = True
        End If
    End If
    Err.Clear

    If Not got Then Exit Sub
    f.TxFnv = Fnv1a32(NormalizeText(txt))
    If Len(fnt) > 0 Then f.FontFnv = Fnv1a32(LCase$(fnt))
    If slideH > 0# And sz > 0 Then
        f.FontRel = CLng(Round((CDbl(sz) / slideH) * 1000#, 0))
    End If
End Sub

Private Function NormalizeText(ByVal s As String) As String
    Dim t As String
    On Error Resume Next
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
' СРАВНЕНИЕ / УНИКАЛЬНОСТЬ
'==============================================================================
Private Function FeaturesMatchFeat(ByRef f As ShapeFeat, ByVal mask As Long, _
                                   ByRef proto As ShapeFeat) As Boolean
    On Error GoTo SoftFail
    FeaturesMatchFeat = False
    If Not f.Valid Or Not proto.Valid Then Exit Function
    If mask = 0 Then Exit Function

    If mask And F_TYPE Then If f.TypeId <> proto.TypeId Then Exit Function
    If mask And F_AUTO Then If f.AutoId <> proto.AutoId Then Exit Function
    If mask And F_PH Then If f.PhType <> proto.PhType Then Exit Function

    If mask And F_POS Then
        If Abs(f.L - proto.L) > GEO_TOL Then Exit Function
        If Abs(f.T - proto.T) > GEO_TOL Then Exit Function
    End If
    If mask And F_SIZE Then
        If Abs(f.W - proto.W) > GEO_TOL Then Exit Function
        If Abs(f.H - proto.H) > GEO_TOL Then Exit Function
    End If
    If mask And F_AR Then
        If Abs(f.AR - proto.AR) > MaxLng(20, CLng(Abs(proto.AR) * 0.02)) Then Exit Function
    End If
    If mask And F_ROT Then If Abs(f.Rot - proto.Rot) > ROT_TOL Then Exit Function
    If mask And F_FLIP Then If f.Flip <> proto.Flip Then Exit Function

    If mask And F_FILL Then
        If f.FillType <> proto.FillType Then Exit Function
        If proto.FillScheme >= 0 Then
            If f.FillScheme <> proto.FillScheme Then Exit Function
        ElseIf proto.FillRgb >= 0 Then
            If f.FillRgb <> proto.FillRgb Then Exit Function
        End If
    End If

    If mask And F_LINE Then
        If f.LineVis <> proto.LineVis Then Exit Function
        If proto.LineVis <> 0 Then
            If proto.LineRgb >= 0 Then
                If f.LineRgb <> proto.LineRgb Then Exit Function
            End If
            If Abs(f.LineWeight - proto.LineWeight) > 25 Then Exit Function
        End If
    End If

    If mask And F_CAPS Then If f.Caps <> proto.Caps Then Exit Function
    If mask And F_TXH Then If f.TxFnv <> proto.TxFnv Then Exit Function
    If mask And F_FONT Then
        If f.FontFnv <> proto.FontFnv Then Exit Function
        If Abs(f.FontRel - proto.FontRel) > FONT_TOL Then Exit Function
    End If

    FeaturesMatchFeat = True
    Exit Function
SoftFail:
    FeaturesMatchFeat = False
End Function

Private Function FeaturesEqualMask(ByRef a As ShapeFeat, ByRef b As ShapeFeat, ByVal mask As Long) As Boolean
    FeaturesEqualMask = FeaturesMatchFeat(b, mask, a)
End Function

Private Function AllTargetsAgree(ByRef tFeats() As ShapeFeat, ByVal mask As Long) As Boolean
    Dim i As Long
    AllTargetsAgree = False
    If mask = 0 Then Exit Function
    AllTargetsAgree = True
    For i = 2 To UBound(tFeats)
        If Not FeaturesEqualMask(tFeats(1), tFeats(i), mask) Then
            AllTargetsAgree = False
            Exit Function
        End If
    Next i
End Function

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
    Dim unique As Boolean
    Dim prototype As ShapeFeat
    Dim hasOthers As Boolean
    Dim collisions As Long

    On Error GoTo SoftFail
    outMask = 0
    BuildDiscriminatingSignature = ""
    report = ""

    If targets Is Nothing Then
        report = "Коллекция целей пуста (Nothing)."
        Exit Function
    End If
    If targets.Count = 0 Then
        report = "Коллекция целей пуста."
        Exit Function
    End If
    If others Is Nothing Then Set others = New Collection

    ReDim tFeats(1 To targets.Count)
    For i = 1 To targets.Count
        tFeats(i) = ExtractFeatures(targets(i))
        If Not tFeats(i).Valid Then
            report = "Не удалось извлечь признаки у цели #" & CStr(i)
            Exit Function
        End If
    Next i

    hasOthers = (others.Count > 0)
    If hasOthers Then
        ReDim oFeats(1 To others.Count)
        For i = 1 To others.Count
            oFeats(i) = ExtractFeatures(others(i))
        Next i
    End If

    bits = Split(DISC_ORDER, ",")
    mask = 0
    unique = False

    For j = 0 To UBound(bits)
        bit = CLng(bits(j))
        If (mask And bit) = 0 Then
            If AllTargetsAgree(tFeats, bit) Then
                ' добавляем бит только если он стабилен сам по себе на целях
                If AllTargetsAgree(tFeats, mask Or bit) Then mask = mask Or bit
            End If
        End If

        If mask = 0 Then GoTo ContBits

        If Not hasOthers Then
            If (mask And SEED_FIND_MASK) = SEED_FIND_MASK Then
                unique = True
                Exit For
            End If
        Else
            collisions = 0
            For i = 1 To others.Count
                If oFeats(i).Valid Then
                    If FeaturesEqualMask(tFeats(1), oFeats(i), mask) Then collisions = collisions + 1
                End If
            Next i
            unique = (collisions = 0)
            If unique Then Exit For
        End If
ContBits:
    Next j

    If Not hasOthers And mask <> 0 Then unique = True

    If mask = 0 Then
        report = "Нет стабильных признаков у экземпляров с этим именем/выборкой."
        Exit Function
    End If

    prototype = tFeats(1)
    outMask = mask
    BuildDiscriminatingSignature = BuildSignatureString(mask, prototype)

    report = "Маска=0x" & Hex$(mask) & " (" & MaskDescription(mask) & ")" & vbCrLf & _
             "Экземпляров-целей: " & CStr(targets.Count) & vbCrLf & _
             "Уникальна среди соседей: " & IIf(unique, "ДА", "НЕТ") & vbCrLf & _
             "Длина сигнатуры: " & CStr(Len(BuildDiscriminatingSignature)) & " символов"

    If Not unique Then
        report = report & vbCrLf & _
                 "ВНИМАНИЕ: на выборке остались коллизии. Добавьте отличия объектам или больше антипримеров."
    End If
    Exit Function
SoftFail:
    outMask = 0
    BuildDiscriminatingSignature = ""
    report = "Сбой BuildDiscriminatingSignature: " & Err.Description
End Function

Private Function MaskDescription(ByVal mask As Long) As String
    Dim s As String
    s = ""
    If mask And F_TYPE Then s = s & "+Type"
    If mask And F_AUTO Then s = s & "+Auto"
    If mask And F_PH Then s = s & "+Ph"
    If mask And F_POS Then s = s & "+Pos%"
    If mask And F_SIZE Then s = s & "+Size%"
    If mask And F_AR Then s = s & "+AR"
    If mask And F_ROT Then s = s & "+Rot"
    If mask And F_FLIP Then s = s & "+Flip"
    If mask And F_FILL Then s = s & "+Fill"
    If mask And F_LINE Then s = s & "+Line"
    If mask And F_CAPS Then s = s & "+Caps"
    If mask And F_TXH Then s = s & "+TextHash"
    If mask And F_FONT Then s = s & "+Font"
    If Len(s) = 0 Then MaskDescription = "(empty)" Else MaskDescription = Mid$(s, 2)
End Function

'==============================================================================
' УПАКОВКА СИГНАТУРЫ S2 (бинарно + Base64URL, без внешних COM)
' Формат: S2|<base64url payload>
' payload = mask(u16) + fields + crc24(fnv of body)
'==============================================================================
Private Function BuildSignatureString(ByVal mask As Long, ByRef f As ShapeFeat) As String
    Dim buf() As Byte
    Dim n As Long
    Dim crc As Long
    Dim bodyLen As Long
    Dim i As Long
    Dim b64 As String

    On Error GoTo SoftFail
    BuildSignatureString = ""
    If mask = 0 Or Not f.Valid Then Exit Function

    ReDim buf(0 To 127)
    n = 0
    AppendU16 buf, n, mask And &HFFFF&

    If mask And F_TYPE Then AppendU8 buf, n, f.TypeId And &HFF&
    If mask And F_AUTO Then AppendI16 buf, n, f.AutoId
    If mask And F_PH Then AppendI16 buf, n, f.PhType
    If mask And F_POS Then
        AppendI16 buf, n, f.L
        AppendI16 buf, n, f.T
    End If
    If mask And F_SIZE Then
        AppendI16 buf, n, f.W
        AppendI16 buf, n, f.H
    End If
    If mask And F_AR Then AppendI32 buf, n, f.AR
    If mask And F_ROT Then AppendI16 buf, n, f.Rot
    If mask And F_FLIP Then AppendU8 buf, n, f.Flip And &HFF&
    If mask And F_FILL Then
        AppendI16 buf, n, f.FillType
        AppendI32 buf, n, f.FillRgb
        AppendI16 buf, n, f.FillScheme
    End If
    If mask And F_LINE Then
        AppendU8 buf, n, f.LineVis And &HFF&
        AppendI32 buf, n, f.LineRgb
        AppendI16 buf, n, f.LineWeight
    End If
    If mask And F_CAPS Then AppendU8 buf, n, f.Caps And &HFF&
    If mask And F_TXH Then AppendI32 buf, n, f.TxFnv
    If mask And F_FONT Then
        AppendI32 buf, n, f.FontFnv
        AppendI16 buf, n, f.FontRel
    End If

    bodyLen = n
    crc = Fnv1aBytes(buf, bodyLen) And &HFFFFFF
    AppendU8 buf, n, crc And &HFF&
    AppendU8 buf, n, (crc \ &H100&) And &HFF&
    AppendU8 buf, n, (crc \ &H10000) And &HFF&

    ReDim Preserve buf(0 To n - 1)
    b64 = Base64UrlEncode(buf)
    BuildSignatureString = SIG_VER & "|" & b64
    Exit Function
SoftFail:
    BuildSignatureString = ""
End Function

Private Function ParseSignatureToFeat(ByVal sig As String, ByRef mask As Long, _
                                      ByRef proto As ShapeFeat) As Boolean
    Dim p() As String
    On Error GoTo SoftFail
    ParseSignatureToFeat = False
    InitFeat proto
    mask = 0
    sig = Trim$(sig)
    If Len(sig) = 0 Then Exit Function

    p = Split(sig, "|")
    If UBound(p) < 1 Then Exit Function

    If p(0) = SIG_VER Then
        If UBound(p) <> 1 Then Exit Function
        ParseSignatureToFeat = ParseS2Payload(p(1), mask, proto)
        Exit Function
    End If

    ' Legacy S1: S1|mask|fields|crc  — только чтение для совместимости
    If p(0) = SIG_VER_LEGACY Then
        ParseSignatureToFeat = ParseLegacyS1(sig, mask, proto)
        Exit Function
    End If
    Exit Function
SoftFail:
    ParseSignatureToFeat = False
End Function

Private Function ParseS2Payload(ByVal b64 As String, ByRef mask As Long, _
                                ByRef proto As ShapeFeat) As Boolean
    Dim buf() As Byte
    Dim n As Long
    Dim pos As Long
    Dim bodyLen As Long
    Dim crc As Long
    Dim got As Long

    On Error GoTo SoftFail
    ParseS2Payload = False
    InitFeat proto
    If Len(b64) = 0 Then Exit Function
    If Not Base64UrlDecode(b64, buf) Then Exit Function
    n = UBound(buf) + 1
    If n < 5 Then Exit Function

    bodyLen = n - 3
    crc = CLng(buf(bodyLen)) + CLng(buf(bodyLen + 1)) * &H100& + CLng(buf(bodyLen + 2)) * &H10000
    got = Fnv1aBytes(buf, bodyLen) And &HFFFFFF
    If crc <> got Then Exit Function

    pos = 0
    mask = ReadU16(buf, pos)
    If mask = 0 Or (mask And Not F_ALL) <> 0 Then Exit Function

    If mask And F_TYPE Then proto.TypeId = ReadU8(buf, pos)
    If mask And F_AUTO Then proto.AutoId = ReadI16(buf, pos)
    If mask And F_PH Then proto.PhType = ReadI16(buf, pos)
    If mask And F_POS Then
        proto.L = ReadI16(buf, pos)
        proto.T = ReadI16(buf, pos)
    End If
    If mask And F_SIZE Then
        proto.W = ReadI16(buf, pos)
        proto.H = ReadI16(buf, pos)
    End If
    If mask And F_AR Then proto.AR = ReadI32(buf, pos)
    If mask And F_ROT Then proto.Rot = ReadI16(buf, pos)
    If mask And F_FLIP Then proto.Flip = ReadU8(buf, pos)
    If mask And F_FILL Then
        proto.FillType = ReadI16(buf, pos)
        proto.FillRgb = ReadI32(buf, pos)
        proto.FillScheme = ReadI16(buf, pos)
    End If
    If mask And F_LINE Then
        proto.LineVis = ReadU8(buf, pos)
        proto.LineRgb = ReadI32(buf, pos)
        proto.LineWeight = ReadI16(buf, pos)
    End If
    If mask And F_CAPS Then proto.Caps = ReadU8(buf, pos)
    If mask And F_TXH Then proto.TxFnv = ReadI32(buf, pos)
    If mask And F_FONT Then
        proto.FontFnv = ReadI32(buf, pos)
        proto.FontRel = ReadI16(buf, pos)
    End If

    If pos <> bodyLen Then Exit Function
    proto.Valid = True
    ParseS2Payload = True
    Exit Function
SoftFail:
    ParseS2Payload = False
End Function

' Упрощённый разбор legacy S1 → ShapeFeat (best-effort)
Private Function ParseLegacyS1(ByVal sig As String, ByRef mask As Long, ByRef proto As ShapeFeat) As Boolean
    Dim p() As String
    Dim fields As String
    Dim feat() As String
    Dim idx As Long
    Dim geo() As String
    Dim font() As String
    Dim fillParts() As String
    Dim lineParts() As String

    On Error GoTo SoftFail
    ParseLegacyS1 = False
    InitFeat proto
    p = Split(sig, "|")
    If UBound(p) <> 3 Then Exit Function
    If p(0) <> SIG_VER_LEGACY Then Exit Function
    If StrComp(Fnv1aB36(p(0) & "|" & p(1) & "|" & p(2)), p(3), vbTextCompare) <> 0 Then Exit Function

    ' Старый S1 использовал другую раскладку битов (GEO=L.T.W.H одним полем).
    ' Поддерживаем только через строковые поля старого порядка:
    ' TYPE,AUTO,PH,GEO,AR,ROT,FLIP,FILL,LINE,CAPS,TXH,FONT
    mask = FromB36(p(1))
    fields = p(2)
    If Len(fields) = 0 Then Exit Function
    feat = Split(fields, ";")
    idx = 0

    ' Маппинг legacy bits → новые (приблизительно)
    ' legacy: 1 type,2 auto,4 ph,8 geo,16 ar,32 rot,64 flip,128 fill,256 line,512 caps,1024 tx,2048 font
    Dim legMask As Long
    Dim newMask As Long
    legMask = mask
    newMask = 0

    If legMask And 1 Then
        If idx > UBound(feat) Then Exit Function
        proto.TypeId = FromB36(feat(idx)): idx = idx + 1: newMask = newMask Or F_TYPE
    End If
    If legMask And 2 Then
        If idx > UBound(feat) Then Exit Function
        proto.AutoId = FromB36(feat(idx)): idx = idx + 1: newMask = newMask Or F_AUTO
    End If
    If legMask And 4 Then
        If idx > UBound(feat) Then Exit Function
        proto.PhType = FromB36(feat(idx)): idx = idx + 1: newMask = newMask Or F_PH
    End If
    If legMask And 8 Then
        If idx > UBound(feat) Then Exit Function
        geo = Split(feat(idx), ".")
        If UBound(geo) <> 3 Then Exit Function
        proto.L = FromB36(geo(0)): proto.T = FromB36(geo(1))
        proto.W = FromB36(geo(2)): proto.H = FromB36(geo(3))
        idx = idx + 1
        newMask = newMask Or F_POS Or F_SIZE
    End If
    If legMask And 16 Then
        If idx > UBound(feat) Then Exit Function
        proto.AR = FromB36(feat(idx)): idx = idx + 1: newMask = newMask Or F_AR
    End If
    If legMask And 32 Then
        If idx > UBound(feat) Then Exit Function
        proto.Rot = FromB36(feat(idx)): idx = idx + 1: newMask = newMask Or F_ROT
    End If
    If legMask And 64 Then
        If idx > UBound(feat) Then Exit Function
        proto.Flip = FromB36(feat(idx)): idx = idx + 1: newMask = newMask Or F_FLIP
    End If
    If legMask And 128 Then
        If idx > UBound(feat) Then Exit Function
        fillParts = Split(feat(idx), ":")
        proto.FillType = CLng(Val(fillParts(0)))
        proto.FillScheme = -1
        proto.FillRgb = -1
        If UBound(fillParts) >= 1 Then
            If Left$(fillParts(1), 1) = "s" Or Left$(fillParts(1), 1) = "S" Then
                proto.FillScheme = CLng(Val(Mid$(fillParts(1), 2)))
            Else
                proto.FillRgb = CLng("&H" & fillParts(1)) And &HFFFFFF
            End If
        End If
        idx = idx + 1: newMask = newMask Or F_FILL
    End If
    If legMask And 256 Then
        If idx > UBound(feat) Then Exit Function
        If feat(idx) = "0" Then
            proto.LineVis = 0
        Else
            lineParts = Split(feat(idx), ":")
            proto.LineVis = 1
            If UBound(lineParts) >= 1 Then proto.LineRgb = CLng("&H" & lineParts(1)) And &HFFFFFF
            If UBound(lineParts) >= 2 Then proto.LineWeight = FromB36(lineParts(2))
        End If
        idx = idx + 1: newMask = newMask Or F_LINE
    End If
    If legMask And 512 Then
        If idx > UBound(feat) Then Exit Function
        proto.Caps = FromB36(feat(idx)): idx = idx + 1: newMask = newMask Or F_CAPS
    End If
    If legMask And 1024 Then
        If idx > UBound(feat) Then Exit Function
        proto.TxFnv = FromB36UnsignedToLong(feat(idx)): idx = idx + 1: newMask = newMask Or F_TXH
    End If
    If legMask And 2048 Then
        If idx > UBound(feat) Then Exit Function
        font = Split(feat(idx), ".")
        If UBound(font) <> 1 Then Exit Function
        proto.FontFnv = FromB36UnsignedToLong(font(0))
        proto.FontRel = FromB36(font(1))
        idx = idx + 1: newMask = newMask Or F_FONT
    End If

    mask = newMask
    proto.Valid = True
    ParseLegacyS1 = True
    Exit Function
SoftFail:
    ParseLegacyS1 = False
End Function

'==============================================================================
' СБОР ПО ИМЕНИ / ТЕГАМ / СЕМЕНИ
'==============================================================================
Private Function CollectByNameInPresentation(ByVal pres As Presentation, _
                                             ByVal shapeName As String, _
                                             ByVal targets As Collection, _
                                             ByVal others As Collection) As Boolean
    Dim sld As Slide
    Dim shp As Shape
    On Error GoTo SoftFail
    CollectByNameInPresentation = False
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            ClassifyByName shp, shapeName, targets, others
            If shp.Type = msoGroup Then CollectByNameGroupItems shp, shapeName, targets, others
        Next shp
    Next sld
    CollectByNameInPresentation = True
    Exit Function
SoftFail:
    CollectByNameInPresentation = False
End Function

Private Sub ClassifyByName(ByVal shp As Shape, ByVal shapeName As String, _
                           ByVal targets As Collection, ByVal others As Collection)
    Dim nm As String
    On Error Resume Next
    nm = shp.Name
    If Err.Number <> 0 Then
        Err.Clear
        others.Add shp
        Exit Sub
    End If
    If StrComp(nm, shapeName, vbTextCompare) = 0 Then
        targets.Add shp
    Else
        others.Add shp
    End If
End Sub

Private Sub CollectByNameGroupItems(ByVal grp As Shape, ByVal shapeName As String, _
                                    ByVal targets As Collection, ByVal others As Collection)
    Dim i As Long
    Dim shp As Shape
    On Error Resume Next
    For i = 1 To grp.GroupItems.Count
        Set shp = grp.GroupItems(i)
        If Not shp Is Nothing Then
            ClassifyByName shp, shapeName, targets, others
            If shp.Type = msoGroup Then CollectByNameGroupItems shp, shapeName, targets, others
        End If
    Next i
End Sub

Private Sub CollectMarkedTargetsOnly(ByVal pres As Presentation, ByVal targets As Collection)
    Dim sld As Slide
    Dim shp As Shape
    On Error Resume Next
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            CollectMarkedTargetsTree shp, targets
        Next shp
    Next sld
End Sub

Private Sub CollectMarkedTargetsTree(ByVal shp As Shape, ByVal targets As Collection)
    Dim i As Long
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    If IsMarkedTarget(shp) Then targets.Add shp
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            CollectMarkedTargetsTree shp.GroupItems(i), targets
        Next i
    End If
End Sub

Private Sub CollectMarkedInPresentation(ByVal pres As Presentation, _
                                        ByVal targets As Collection, _
                                        ByVal others As Collection)
    Dim sld As Slide
    Dim shp As Shape
    On Error Resume Next
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            CollectShapeTree shp, targets, others
        Next shp
    Next sld
End Sub

Private Sub CollectShapeTree(ByVal shp As Shape, ByVal targets As Collection, ByVal others As Collection)
    Dim i As Long
    On Error Resume Next
    If IsMarkedTarget(shp) Then targets.Add shp Else others.Add shp
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            CollectShapeTree shp.GroupItems(i), targets, others
        Next i
    End If
End Sub

Private Sub CollectBySeedInPresentation(ByVal pres As Presentation, _
                                        ByRef seedFeats() As ShapeFeat, _
                                        ByVal targets As Collection, _
                                        ByVal others As Collection)
    Dim sld As Slide
    Dim shp As Shape
    On Error Resume Next
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            ClassifyShapeBySeed shp, seedFeats, targets, others
            If shp.Type = msoGroup Then CollectBySeedGroupItems shp, seedFeats, targets, others
        Next shp
    Next sld
End Sub

Private Sub ClassifyShapeBySeed(ByVal shp As Shape, ByRef seedFeats() As ShapeFeat, _
                                ByVal targets As Collection, ByVal others As Collection)
    Dim f As ShapeFeat
    Dim isTarget As Boolean
    Dim k As Long
    On Error Resume Next
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
    If isTarget Then targets.Add shp Else others.Add shp
End Sub

Private Sub CollectBySeedGroupItems(ByVal grp As Shape, ByRef seedFeats() As ShapeFeat, _
                                    ByVal targets As Collection, ByVal others As Collection)
    Dim i As Long
    Dim shp As Shape
    On Error Resume Next
    For i = 1 To grp.GroupItems.Count
        Set shp = grp.GroupItems(i)
        ClassifyShapeBySeed shp, seedFeats, targets, others
        If shp.Type = msoGroup Then CollectBySeedGroupItems shp, seedFeats, targets, others
    Next i
End Sub

Private Sub CollectOthersOnSlide(ByVal sld As Slide, ByVal targets As Collection, ByVal others As Collection)
    Dim shp As Shape
    Dim i As Long
    Dim isTarget As Boolean
    On Error Resume Next
    For Each shp In sld.Shapes
        isTarget = False
        For i = 1 To targets.Count
            If IsSameShapeRef(targets(i), shp) Then
                isTarget = True
                Exit For
            End If
        Next i
        If Not isTarget Then others.Add shp
        If shp.Type = msoGroup Then CollectGroupOthers shp, targets, others
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
    On Error Resume Next
    If col Is Nothing Then Exit Sub
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
    On Error Resume Next
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
    On Error Resume Next
    IsSameShapeRef = False
    If a Is Nothing Or b Is Nothing Then Exit Function
    IsSameShapeRef = (a Is b)
End Function

Private Function IsMarkedTarget(ByVal shp As Shape) As Boolean
    Dim i As Long
    On Error Resume Next
    IsMarkedTarget = False
    For i = 1 To shp.Tags.Count
        If StrComp(shp.Tags.Name(i), TAG_TARGET, vbTextCompare) = 0 Then
            If StrComp(shp.Tags.Value(i), TAG_TARGET_VAL, vbTextCompare) = 0 Then
                IsMarkedTarget = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ClearTargetMarksInGroup(ByVal shp As Shape) As Long
    Dim i As Long
    Dim n As Long
    On Error Resume Next
    ClearTargetMarksInGroup = 0
    If shp.Type <> msoGroup Then Exit Function
    For i = 1 To shp.GroupItems.Count
        If SafeTagRemove(shp.GroupItems(i), TAG_TARGET) Then n = n + 1
        n = n + ClearTargetMarksInGroup(shp.GroupItems(i))
    Next i
    ClearTargetMarksInGroup = n
End Function

'==============================================================================
' ВСПОМОГАТЕЛЬНЫЕ (без CreateObject)
'==============================================================================
Private Sub PresentSignatureResult(ByVal sig As String, ByVal mask As Long, _
                                   ByVal nTargets As Long, ByVal nOthers As Long, _
                                   ByVal report As String)
    Dim msg As String
    On Error Resume Next
    If Len(sig) = 0 Then
        MsgBox report, vbExclamation, "Сигнатура не построена"
        Exit Sub
    End If
    msg = "Целей: " & CStr(nTargets) & " | Соседей: " & CStr(nOthers) & vbCrLf & _
          report & vbCrLf & vbCrLf & _
          "Сигнатура (скопируйте из поля ниже):" & vbCrLf & sig
    InputBox msg, "ShapeSignature", sig
End Sub

Private Function SelectionHasShapes() As Boolean
    On Error Resume Next
    SelectionHasShapes = False
    If ActiveWindow Is Nothing Then Exit Function
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        SelectionHasShapes = (ActiveWindow.Selection.ShapeRange.Count > 0)
    End If
End Function

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

Private Sub SafeTagAdd(ByVal shp As Shape, ByVal nm As String, ByVal val As String)
    On Error Resume Next
    shp.Tags.Delete nm
    shp.Tags.Add nm, val
End Sub

Private Function SafeTagRemove(ByVal shp As Shape, ByVal nm As String) As Boolean
    Dim i As Long
    On Error Resume Next
    SafeTagRemove = False
    For i = 1 To shp.Tags.Count
        If StrComp(shp.Tags.Name(i), nm, vbTextCompare) = 0 Then
            shp.Tags.Delete nm
            SafeTagRemove = True
            Exit Function
        End If
    Next i
End Function

Private Function IsPresentationOpen(ByVal fullPath As String) As Boolean
    IsPresentationOpen = Not (FindOpenPresentation(fullPath) Is Nothing)
End Function

Private Function FindOpenPresentation(ByVal fullPath As String) As Presentation
    Dim p As Presentation
    Dim target As String
    On Error Resume Next
    Set FindOpenPresentation = Nothing
    target = LCase$(fullPath)
    For Each p In Presentations
        If LCase$(p.FullName) = target Then
            Set FindOpenPresentation = p
            Exit Function
        End If
    Next p
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

Private Function CollectionHasKey(ByVal col As Collection, ByVal key As String) As Boolean
    Dim t As Variant
    On Error Resume Next
    t = col.Item(key)
    CollectionHasKey = (Err.Number = 0)
    Err.Clear
End Function

Private Sub CollectionAddKey(ByVal col As Collection, ByVal key As String)
    On Error Resume Next
    col.Add True, key
    Err.Clear
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

'----- Byte buffer / Base64URL / FNV -----------------------------------------
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

Private Sub AppendI16(ByRef buf() As Byte, ByRef n As Long, ByVal v As Long)
    Dim u As Long
    If v < 0 Then
        u = v + 65536
    Else
        u = v And &HFFFF&
    End If
    AppendU16 buf, n, u
End Sub

Private Sub AppendI32(ByRef buf() As Byte, ByRef n As Long, ByVal v As Long)
    Dim uh As Double
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    If v < 0 Then
        uh = CDbl(v And &H7FFFFFFF) + 2147483648#
    Else
        uh = CDbl(v)
    End If
    b0 = CLng(uh - Int(uh / 256#) * 256#)
    uh = Int(uh / 256#)
    b1 = CLng(uh - Int(uh / 256#) * 256#)
    uh = Int(uh / 256#)
    b2 = CLng(uh - Int(uh / 256#) * 256#)
    uh = Int(uh / 256#)
    b3 = CLng(uh - Int(uh / 256#) * 256#)
    AppendU8 buf, n, b0
    AppendU8 buf, n, b1
    AppendU8 buf, n, b2
    AppendU8 buf, n, b3
End Sub

Private Function ReadU8(ByRef buf() As Byte, ByRef pos As Long) As Long
    ReadU8 = buf(pos)
    pos = pos + 1
End Function

Private Function ReadU16(ByRef buf() As Byte, ByRef pos As Long) As Long
    ReadU16 = CLng(buf(pos)) + CLng(buf(pos + 1)) * &H100&
    pos = pos + 2
End Function

Private Function ReadI16(ByRef buf() As Byte, ByRef pos As Long) As Long
    Dim u As Long
    u = ReadU16(buf, pos)
    If u >= 32768 Then ReadI16 = u - 65536 Else ReadI16 = u
End Function

Private Function ReadI32(ByRef buf() As Byte, ByRef pos As Long) As Long
    Dim uh As Double
    uh = CDbl(buf(pos)) + CDbl(buf(pos + 1)) * 256# + CDbl(buf(pos + 2)) * 65536# + CDbl(buf(pos + 3)) * 16777216#
    pos = pos + 4
    If uh >= 2147483648# Then
        ReadI32 = CLng(uh - 4294967296#)
    Else
        ReadI32 = CLng(uh)
    End If
End Function

Private Function Base64UrlEncode(ByRef buf() As Byte) As String
    Dim enc As String
    Dim i As Long
    Dim n As Long
    Dim a As Long, b As Long, c As Long
    Dim alphabet As String
    Dim out As String
    Dim remn As Long

    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    n = UBound(buf) + 1
    out = ""
    i = 0
    Do While i + 2 < n
        a = buf(i): b = buf(i + 1): c = buf(i + 2)
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
        a = buf(i): b = buf(i + 1)
        out = out & Mid$(alphabet, ((a \ 4) And 63) + 1, 1)
        out = out & Mid$(alphabet, (((a And 3) * 16) Or ((b \ 16) And 15)) + 1, 1)
        out = out & Mid$(alphabet, ((b And 15) * 4) + 1, 1)
    End If
    Base64UrlEncode = out
End Function

Private Function Base64UrlDecode(ByVal s As String, ByRef buf() As Byte) As Boolean
    Dim alphabet As String
    Dim i As Long
    Dim v As Long
    Dim vals() As Long
    Dim n As Long
    Dim o As Long
    Dim a As Long, b As Long, c As Long, d As Long
    Dim ch As String
    Dim p As Long

    On Error GoTo SoftFail
    Base64UrlDecode = False
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    s = Replace(Replace(Trim$(s), "+", "-"), "/", "_")
    s = Replace(s, "=", "")
    If Len(s) = 0 Then Exit Function

    n = Len(s)
    ReDim vals(0 To n - 1)
    For i = 1 To n
        ch = Mid$(s, i, 1)
        p = InStr(1, alphabet, ch, vbBinaryCompare)
        If p <= 0 Then Exit Function
        vals(i - 1) = p - 1
    Next i

    ReDim buf(0 To ((n * 3) \ 4) + 3)
    o = 0
    i = 0
    Do While i + 3 < n
        a = vals(i): b = vals(i + 1): c = vals(i + 2): d = vals(i + 3)
        buf(o) = CByte(((a * 4) Or (b \ 16)) And &HFF&)
        buf(o + 1) = CByte((((b And 15) * 16) Or (c \ 4)) And &HFF&)
        buf(o + 2) = CByte((((c And 3) * 64) Or d) And &HFF&)
        o = o + 3
        i = i + 4
    Loop
    If n - i = 2 Then
        a = vals(i): b = vals(i + 1)
        buf(o) = CByte(((a * 4) Or (b \ 16)) And &HFF&)
        o = o + 1
    ElseIf n - i = 3 Then
        a = vals(i): b = vals(i + 1): c = vals(i + 2)
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

Private Function Fnv1a32(ByVal s As String) As Long
    Dim h As Long
    Dim i As Long
    Dim b As Long
    h = -2128831035
    For i = 1 To Len(s)
        b = AscW(Mid$(s, i, 1)) And &HFF&
        h = (h Xor b)
        h = FnvMul(h)
    Next i
    Fnv1a32 = h
End Function

Private Function Fnv1aBytes(ByRef buf() As Byte, ByVal n As Long) As Long
    Dim h As Long
    Dim i As Long
    h = -2128831035
    For i = 0 To n - 1
        h = (h Xor (buf(i) And &HFF&))
        h = FnvMul(h)
    Next i
    Fnv1aBytes = h
End Function

Private Function FnvMul(ByVal h As Long) As Long
    FnvMul = LngMulAdd(h, 16777619, 0)
End Function

Private Function LngMulAdd(ByVal a As Long, ByVal m As Long, ByVal addv As Long) As Long
    Dim da As Double, dm As Double, dr As Double, hi As Double
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

' Legacy helpers for S1 read
Private Function ToB36(ByVal n As Long) As String
    Dim digits As String
    Dim neg As Boolean
    Dim r As Long
    Dim v As Long
    digits = "0123456789abcdefghijklmnopqrstuvwxyz"
    If n = 0 Then ToB36 = "0": Exit Function
    neg = (n < 0)
    If neg Then v = -n Else v = n
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
    If Left$(s, 1) = "-" Then neg = True: s = Mid$(s, 2)
    v = 0
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        p = InStr(1, digits, c, vbBinaryCompare) - 1
        If p < 0 Then FromB36 = 0: Exit Function
        v = v * 36 + p
    Next i
    If neg Then v = -v
    FromB36 = v
End Function

Private Function ToB36Unsigned(ByVal uh As Double) As String
    Dim digits As String
    Dim r As Long
    Dim s As String
    digits = "0123456789abcdefghijklmnopqrstuvwxyz"
    If uh <= 0 Then ToB36Unsigned = "0": Exit Function
    s = ""
    Do While uh >= 1
        r = CLng(uh - Int(uh / 36#) * 36#)
        s = Mid$(digits, r + 1, 1) & s
        uh = Int(uh / 36#)
    Loop
    ToB36Unsigned = s
End Function

Private Function Fnv1aB36(ByVal s As String) As String
    Dim h As Long
    Dim uh As Double
    h = Fnv1a32(s)
    If h < 0 Then
        uh = CDbl(h And &H7FFFFFFF) + 2147483648#
    Else
        uh = CDbl(h)
    End If
    Fnv1aB36 = ToB36Unsigned(uh)
End Function

Private Function FromB36UnsignedToLong(ByVal s As String) As Long
    Dim uh As Double
    Dim digits As String
    Dim i As Long
    Dim p As Long
    digits = "0123456789abcdefghijklmnopqrstuvwxyz"
    s = LCase$(Trim$(s))
    uh = 0
    For i = 1 To Len(s)
        p = InStr(1, digits, Mid$(s, i, 1), vbBinaryCompare) - 1
        If p < 0 Then Exit Function
        uh = uh * 36# + p
    Next i
    If uh >= 2147483648# Then
        FromB36UnsignedToLong = CLng(uh - 4294967296#)
    Else
        FromB36UnsignedToLong = CLng(uh)
    End If
End Function
