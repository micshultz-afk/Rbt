Attribute VB_Name = "ShapeSignatureExamples"
'==============================================================================
' Примеры использования ShapeSignature (PowerPoint 2013, offline)
' Импортируйте вместе с ShapeSignature.bas
'==============================================================================
Option Explicit

'--- 1. Обучение: файлы с ошибками + эталонные файлы --------------------------
Public Sub Example_Train()
    Dim badFiles As String
    Dim goodFiles As String
    Dim nm As String
    Dim sig As String

    On Error GoTo Fail

    ' На нужных слайдах во всех этих файлах фигура названа одним именем
    nm = "FooterLabel"
    badFiles = "C:\Decks\bad1.pptx;C:\Decks\bad2.pptx;C:\Decks\bad3.pptx"
    goodFiles = "C:\Decks\reference.pptx"

    sig = TrainSignatureFromFiles(badFiles, nm, goodFiles, nm)

    If Len(sig) = 0 Then
        MsgBox "Обучение не удалось:" & vbCrLf & LastTrainReport(), vbExclamation
        Exit Sub
    End If

    Debug.Print "Сигнатура: " & sig
    Debug.Print LastTrainReport()

    ' Сохранить в каталоге тегов презентации (переносится вместе с файлом)
    SaveSignatureToPresentation ActivePresentation, "FOOTER", sig

    InputBox LastTrainReport(), "Обучение завершено", sig
    Exit Sub
Fail:
    MsgBox "Ошибка обучения: " & Err.Description, vbCritical
End Sub

'--- 2. Похожесть выделенной фигуры в процентах -------------------------------
Public Sub Example_Percent()
    Dim shp As Shape
    Dim sig As String
    Dim pct As Long

    On Error GoTo Fail
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Выделите фигуру", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    sig = LoadSignatureFromPresentation(ActivePresentation, "FOOTER")
    If Len(sig) = 0 Then sig = InputBox("Сигнатура:", "Похожесть")
    If Len(sig) = 0 Then Exit Sub

    pct = ShapeSimilarityPercent(shp, sig)

    MsgBox "Похожесть: " & CStr(pct) & "%" & vbCrLf & vbCrLf & _
           ShapeSimilarityExplain(shp, sig) & vbCrLf & vbCrLf & _
           "Формат: " & DescribeFormatDeviations(shp, sig), vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

'--- 3. Поиск и обработка: по одному объекту на слайд -------------------------
Public Sub Example_ProcessEachSlide()
    Dim sig As String
    Dim sld As Slide
    Dim shp As Shape
    Dim pct As Long
    Dim amb As Boolean
    Dim log As String

    On Error GoTo Fail
    sig = LoadSignatureFromPresentation(ActivePresentation, "FOOTER")
    If Len(sig) = 0 Then sig = InputBox("Сигнатура:", "Поиск")
    If Len(sig) = 0 Then Exit Sub

    For Each sld In ActivePresentation.Slides
        Set shp = FindBestMatchOnSlide(sld, sig, pct, amb)
        If shp Is Nothing Then
            log = log & "Слайд " & sld.SlideIndex & ": не найдено" & vbCrLf
        ElseIf amb Then
            log = log & "Слайд " & sld.SlideIndex & ": неоднозначно (" & pct & "%) — пропуск" & vbCrLf
        ElseIf ShapeMatchesSignature(shp, sig) Then
            log = log & "Слайд " & sld.SlideIndex & ": «" & shp.Name & "» " & pct & "% — " & _
                  DescribeFormatDeviations(shp, sig) & vbCrLf
            ' Здесь вызывайте свой код исправления форматирования для shp
        Else
            log = log & "Слайд " & sld.SlideIndex & ": ниже порога (" & pct & "%)" & vbCrLf
        End If
    Next sld

    Debug.Print log
    MsgBox log, vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

'--- 4. Все совпадения в презентации -----------------------------------------
Public Sub Example_FindAll()
    Dim sig As String
    On Error GoTo Fail
    sig = InputBox("Сигнатура:", "Поиск всех")
    If Len(sig) = 0 Then Exit Sub
    MsgBox FindSimilarShapesReport(ActivePresentation, sig), vbInformation
    Exit Sub
Fail:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

'--- 5. Обход открытых презентаций без ThisPresentation ----------------------
Public Function CountMatchesInPresentation(ByRef pres As Presentation, _
                                           ByVal sig As String) As Long
    Dim col As Collection
    On Error GoTo SoftFail
    CountMatchesInPresentation = 0
    If pres Is Nothing Then Exit Function
    Set col = FindSimilarShapes(pres, sig)
    CountMatchesInPresentation = col.Count
    Exit Function
SoftFail:
    CountMatchesInPresentation = 0
End Function
