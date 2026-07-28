Attribute VB_Name = "ShapeSignatureExamples"
'==============================================================================
' Примеры (offline). Импортируйте вместе с ShapeSignature.bas
'==============================================================================
Option Explicit

Public Sub Example_BuildByName()
    Dim sig As String
    Dim pres As Presentation
    Dim shp As Shape
    Dim ok As Boolean

    On Error GoTo Fail
    Set pres = ActivePresentation
    If pres Is Nothing Then
        MsgBox "Нет активной презентации", vbExclamation
        Exit Sub
    End If

    ' Задайте одно и то же Name фигурам-ролям на разных слайдах
    sig = BuildSignatureByShapeName("MyTarget", pres)
    If Len(sig) = 0 Then
        MsgBox "Не удалось построить сигнатуру для имени MyTarget", vbExclamation
        Exit Sub
    End If

    InputBox "Сигнатура:", "BuildSignatureByShapeName", sig

    For Each shp In ActiveWindow.View.Slide.Shapes
        ok = ShapeMatchesSignature(shp, sig)
        If ok Then Debug.Print "match: " & shp.Name
    Next shp
    Exit Sub
Fail:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

Public Sub Example_CheckSelected()
    Dim sig As String
    Dim ok As Boolean
    Dim shp As Shape

    On Error GoTo Fail
    sig = InputBox("Сигнатура:", "Проверка")
    If Len(sig) = 0 Then Exit Sub
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Выделите фигуру", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    ok = ShapeMatchesSignature(shp, sig)
    MsgBox "Совпадает: " & ok
    Exit Sub
Fail:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

Public Function CountMatchesInPresentation(ByRef pres As Presentation, _
                                           ByVal sig As String) As Long
    Dim sld As Slide
    Dim shp As Shape
    Dim n As Long
    On Error GoTo SoftFail
    CountMatchesInPresentation = 0
    If pres Is Nothing Then Exit Function
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If ShapeMatchesSignature(shp, sig) Then n = n + 1
        Next shp
    Next sld
    CountMatchesInPresentation = n
    Exit Function
SoftFail:
    CountMatchesInPresentation = 0
End Function

Public Sub Example_SimilarityPercent()
    Dim shp As Shape
    Dim sig As String
    Dim pct As Long
    Dim explain As String

    On Error GoTo Fail
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Выделите фигуру", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    sig = InputBox("Сигнатура:", "Похожесть %")
    If Len(sig) = 0 Then Exit Sub

    pct = ShapeSimilarityPercent(shp, sig)
    explain = ShapeSimilarityExplain(shp, sig)
    MsgBox "Процент: " & CStr(pct) & vbCrLf & explain, vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

Public Sub Example_FindSimilar()
    Dim sig As String
    Dim col As Collection
    Dim i As Long
    Dim shp As Shape
    Dim msg As String

    On Error GoTo Fail
    sig = InputBox("Сигнатура:", "Поиск похожих")
    If Len(sig) = 0 Then Exit Sub

    Set col = FindSimilarShapes(ActivePresentation, sig, 70)
    msg = "Найдено объектов: " & CStr(col.Count) & vbCrLf
    For i = 1 To col.Count
        Set shp = col(i)
        msg = msg & "• " & shp.Name & " — " & _
              CStr(ShapeSimilarityPercent(shp, sig)) & "%" & vbCrLf
        If i >= 30 Then
            msg = msg & "…"
            Exit For
        End If
    Next i
    MsgBox msg, vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub
