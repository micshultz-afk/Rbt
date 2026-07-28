Attribute VB_Name = "ShapeSignatureExamples"
'==============================================================================
' Примеры использования ShapeSignature (PowerPoint 2013+)
' Импортируйте вместе с ShapeSignature.bas
'==============================================================================
Option Explicit

'--- Пример 1: быстрая проверка выделенной фигуры ----------------------------
Public Sub Example_CheckSelected()
    Dim sig As String
    Dim ok As Boolean
    Dim shp As Shape

    ' Вставьте сюда сигнатуру, полученную из BuildSignatureFromSelection
    sig = "S1|..." ' замените на реальную строку

    If Not (ActiveWindow.Selection.Type = ppSelectionShapes) Then
        MsgBox "Выделите фигуру", vbExclamation
        Exit Sub
    End If

    ' ByRef Shape — передаём переменную, не выражение ShapeRange(1)
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    ok = ShapeMatchesSignature(shp, sig)
    MsgBox "Совпадает: " & ok
End Sub

'--- Пример 2: найти все совпадения на текущем слайде ------------------------
Public Sub Example_FindMatchesOnSlide()
    Dim sig As String
    Dim shp As Shape
    Dim sld As Slide
    Dim n As Long
    Dim names As String

    sig = InputBox("Сигнатура:", "Поиск на слайде")
    If Len(sig) = 0 Then Exit Sub

    Set sld = ActiveWindow.View.Slide
    For Each shp In sld.Shapes
        If ShapeMatchesSignature(shp, sig) Then
            n = n + 1
            names = names & shp.Name & " (Id=" & shp.Id & ")" & vbCrLf
        End If
    Next shp

    MsgBox "Найдено: " & n & vbCrLf & names, vbInformation
End Sub

'--- Пример 3: обход всей презентации ----------------------------------------
Public Function CountMatchesInPresentation(ByVal pres As Presentation, _
                                           ByVal sig As String) As Long
    Dim sld As Slide
    Dim shp As Shape
    Dim n As Long
    If pres Is Nothing Then Exit Function
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If ShapeMatchesSignature(shp, sig) Then n = n + 1
        Next shp
    Next sld
    CountMatchesInPresentation = n
End Function

'--- Пример 4: использовать в своём коде как предикат ------------------------
Public Sub Example_ProcessIfMatch()
    Dim sig As String
    Dim shp As Shape
    sig = "S1|..." ' ваша сигнатура

    For Each shp In ActiveWindow.View.Slide.Shapes
        If ShapeMatchesSignature(shp, sig) Then
            ' ваша бизнес-логика
        End If
    Next shp
End Sub
