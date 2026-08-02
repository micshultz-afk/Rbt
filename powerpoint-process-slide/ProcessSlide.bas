Attribute VB_Name = "ProcessSlideMod"
'==============================================================================
' ProcessSlide — обход и форматирование объектов слайда (PowerPoint 2013 VBA)
'
' Ключевое исправление: фигуры внутри msoGroup обрабатываются рекурсивно
' через GroupItems (в т.ч. вложенные группы). Раньше For Each pptSlide.Shapes
' видел только верхний уровень — текст/таблицы/диаграммы в группах пропускались.
'
' Зависит от уже существующих в проекте:
'   bDebugMode, g_FooterBottomBorder, SettingsForm, InfoWindow,
'   TraceStart/TraceEnd/TraceElementStart/TraceElementEnd/LogLine,
'   IsFullyInsideSlide, GlobalIgnoreShapeSettings, IsProjectLabelShape,
'   FormatProjectLabel, ProcessTextInShape, ProcessTable, ProcessChart,
'   ProcessLine, RemoveForcedLineBreaksInChartLabels
'==============================================================================
Option Explicit

' === Обработка одного слайда по объектам ===
' Public — чтобы можно было вызвать из другого модуля после Import File.
Public Sub ProcessSlide(ByRef pptSlide As Slide, ByVal slideIndex As Long)
    Dim t As Double
    Dim shp As Shape
    Dim presName As String

    If Not bDebugMode Then On Error Resume Next

    t = TraceStart("ProcessSlide", slideIndex)

    presName = vbNullString
    On Error Resume Next
    presName = pptSlide.Parent.Parent.FullName
    If Len(presName) = 0 Then presName = pptSlide.Application.ActivePresentation.FullName
    On Error GoTo 0
    If Not bDebugMode Then On Error Resume Next

    LogLine "ProcessSlide:", presName

    For Each shp In pptSlide.Shapes
        ProcessShapeTree shp, slideIndex
    Next shp

    TraceEnd "ProcessSlide", t, slideIndex
End Sub

' Рекурсивный обход: группа -> GroupItems; иначе — форматирование листа.
Private Sub ProcessShapeTree(ByRef shp As Shape, ByVal slideIndex As Long)
    Dim te As Double
    Dim i As Long
    Dim n As Long

    te = TraceElementStart(slideIndex, shp)
    ShowElementProgress shp

    ' Группа: спускаемся к детям (вложенные группы тоже).
    ' Не проверяем IsFullyInsideSlide на контейнере — рамка группы часто
    ' шире слайда, а дочерние объекты при этом валидны.
    If SafeShapeType(shp) = msoGroup Then
        On Error Resume Next
        n = shp.GroupItems.Count
        On Error GoTo 0
        If Not bDebugMode Then On Error Resume Next

        For i = 1 To n
            ProcessShapeTree shp.GroupItems(i), slideIndex
        Next i
        GoTo Done
    End If

    If Not IsFullyInsideSlide(shp) Then GoTo Done

    If FormatShapeByKind(shp, slideIndex) Then GoTo Done

Done:
    TraceElementEnd slideIndex, shp, te
End Sub

' Возвращает True, если фигура обработана (или намеренно пропущена по правилам).
Private Function FormatShapeByKind(ByRef shp As Shape, ByVal slideIndex As Long) As Boolean
    FormatShapeByKind = True

    ' --- Текст ---
    If SafeHasTextFrame(shp) Then
        If SettingsForm.cbFormatText.Value Then
            If shp.Top > g_FooterBottomBorder Then
                If GlobalIgnoreShapeSettings(shp) Then Exit Function
                If IsProjectLabelShape(shp) Then
                    FormatProjectLabel shp
                    Exit Function
                End If
                ProcessTextInShape shp, slideIndex
                Exit Function
            End If
        End If
    End If

    ' --- Таблица ---
    If SafeHasTable(shp) Then
        If SettingsForm.cbFormatTableGroup.Value Then
            ProcessTable shp, slideIndex
            Exit Function
        End If
    End If

    ' --- Диаграмма ---
    If SafeHasChart(shp) Then
        If SettingsForm.cbFormatCharts.Value Then
            ProcessChart shp, slideIndex
            RemoveForcedLineBreaksInChartLabels shp
            Exit Function
        End If
    End If

    ' --- Линия / стрелка ---
    If SafeShapeType(shp) = msoLine Then
        If SettingsForm.cbArrowhead.Value Then
            ProcessLine shp, slideIndex
            Exit Function
        End If
    End If

    FormatShapeByKind = False
End Function

Private Sub ShowElementProgress(ByRef shp As Shape)
    Dim nm As String
    Dim idVal As Long

    On Error Resume Next
    nm = shp.Name
    idVal = shp.Id
    InfoWindow.lblHeader2 = "Форматирование элемента " & nm & " (id:" & idVal & ")"
    InfoWindow.Repaint
End Sub

'----- Безопасные обёртки (PPT 2013: часть свойств падает на «чужих» типах) ---

Private Function SafeShapeType(ByRef shp As Shape) As MsoShapeType
    On Error Resume Next
    SafeShapeType = shp.Type
End Function

Private Function SafeHasTextFrame(ByRef shp As Shape) As Boolean
    On Error Resume Next
    SafeHasTextFrame = (shp.HasTextFrame = msoTrue)
End Function

Private Function SafeHasTable(ByRef shp As Shape) As Boolean
    On Error Resume Next
    SafeHasTable = CBool(shp.HasTable)
End Function

Private Function SafeHasChart(ByRef shp As Shape) As Boolean
    On Error Resume Next
    SafeHasChart = CBool(shp.HasChart)
End Function
