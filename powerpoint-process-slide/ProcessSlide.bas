Attribute VB_Name = "ProcessSlideMod"
' ProcessSlide: обход слайда + рекурсия msoGroup/GroupItems (PPT 2013).
' Зависит от: bDebugMode, g_FooterBottomBorder, SettingsForm, InfoWindow,
' Trace*, LogLine, IsFullyInsideSlide, GlobalIgnoreShapeSettings,
' IsProjectLabelShape, FormatProjectLabel, ProcessTextInShape, ProcessTable,
' ProcessChart, ProcessLine, RemoveForcedLineBreaksInChartLabels
Option Explicit

Public Sub ProcessSlide(ByRef pptSlide As Slide, ByVal slideIndex As Long)
    Dim t As Double, shp As Shape, nm As String
    If Not bDebugMode Then On Error Resume Next
    t = TraceStart("ProcessSlide", slideIndex)
    On Error Resume Next
    nm = pptSlide.Parent.Parent.FullName
    If Len(nm) = 0 Then nm = pptSlide.Application.ActivePresentation.FullName
    On Error GoTo 0
    If Not bDebugMode Then On Error Resume Next
    LogLine "ProcessSlide:", nm
    For Each shp In pptSlide.Shapes
        ProcessShapeTree shp, slideIndex
    Next
    TraceEnd "ProcessSlide", t, slideIndex
End Sub

Private Sub ProcessShapeTree(ByRef shp As Shape, ByVal slideIndex As Long)
    Dim te As Double, i As Long
    If Not bDebugMode Then On Error Resume Next
    te = TraceElementStart(slideIndex, shp)
    On Error Resume Next
    InfoWindow.lblHeader2 = "Форматирование элемента " & shp.Name & " (id:" & shp.Id & ")"
    InfoWindow.Repaint
    On Error GoTo 0
    If Not bDebugMode Then On Error Resume Next

    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            ProcessShapeTree shp.GroupItems(i), slideIndex
        Next
    ElseIf IsFullyInsideSlide(shp) Then
        If shp.HasTextFrame = msoTrue And SettingsForm.cbFormatText.Value And shp.Top > g_FooterBottomBorder Then
            If Not GlobalIgnoreShapeSettings(shp) Then
                If IsProjectLabelShape(shp) Then FormatProjectLabel shp Else ProcessTextInShape shp, slideIndex
            End If
        ElseIf shp.HasTable And SettingsForm.cbFormatTableGroup.Value Then
            ProcessTable shp, slideIndex
        ElseIf shp.HasChart And SettingsForm.cbFormatCharts.Value Then
            ProcessChart shp, slideIndex
            RemoveForcedLineBreaksInChartLabels shp
        ElseIf shp.Type = msoLine And SettingsForm.cbArrowhead.Value Then
            ProcessLine shp, slideIndex
        End If
    End If

    TraceElementEnd slideIndex, shp, te
End Sub
