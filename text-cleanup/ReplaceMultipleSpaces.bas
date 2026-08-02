Attribute VB_Name = "ReplaceMultipleSpaces"
'==============================================================================
' ReplaceMultipleSpaces — PowerPoint VBA
' Любое "  "+ → " " в TextRange. Формат между Runs сохраняется.
'
'   n = CollapseMultiSpaces(rng)                 ' удалено лишних; 0; -1 = ошибка
'   ReplaceMultipleSpacesDashes rng, slideIndex, shp
'==============================================================================
Option Explicit

' Схлопывает множественные пробелы. Возврат: удалено / 0 / -1.
Public Function CollapseMultiSpaces(ByRef rng As TextRange) As Long
    Dim s As String, s2 As String
    Dim run As TextRange
    Dim i As Long, nRuns As Long, before As Long, prev As Long, guard As Long

    On Error GoTo Fail
    CollapseMultiSpaces = 0
    If rng Is Nothing Then Exit Function

    s = rng.Text
    before = Len(s)
    If before < 2 Then Exit Function
    If InStr(s, "  ") = 0 Then Exit Function

    nRuns = rng.Runs.Count
    If nRuns <= 1 Then
        ' Один Run — формат однороден, строковая правка безопасна и быстра.
        s2 = CollapseSpacesString(s)
        If Len(s2) <> before Then
            rng.Text = s2
            CollapseMultiSpaces = before - Len(s2)
        End If
        Exit Function
    End If

    ' Несколько Runs: сначала внутри каждого (с конца — индексы не плывут).
    For i = nRuns To 1 Step -1
        Set run = rng.Runs(i)
        s = run.Text
        If InStr(s, "  ") > 0 Then
            s2 = CollapseSpacesString(s)
            If StrComp(s2, s, vbBinaryCompare) <> 0 Then run.Text = s2
        End If
    Next i

    ' Стыки Runs: " " | " " не видны внутри Run.
    ' Выход по факту сжатия — Replace в PPT иногда возвращает пустой
    ' TextRange вместо Nothing, поэтому на found полагаться нельзя.
    guard = 0
    Do While InStr(rng.Text, "  ") > 0
        prev = Len(rng.Text)
        rng.Replace "  ", " "
        If Len(rng.Text) >= prev Then Exit Do
        guard = guard + 1
        If guard > before Then Exit Do
    Loop

    CollapseMultiSpaces = before - Len(rng.Text)
    Exit Function
Fail:
    CollapseMultiSpaces = -1
End Function

' Серии пробелов → один. Каждый Replace$ сжимает все пары за проход.
Public Function CollapseSpacesString(ByVal s As String) As String
    Dim prev As Long
    Do While InStr(s, "  ") > 0
        prev = Len(s)
        s = Replace$(s, "  ", " ")
        If Len(s) >= prev Then Exit Do
    Loop
    CollapseSpacesString = s
End Function

' Обёртка под лог проекта (bDebugMode, TraceStart/End, LogLine, g_ChangedElements).
Public Sub ReplaceMultipleSpacesDashes(ByRef Rng As TextRange, ByVal slideIndex As Long, ByRef shp As Shape)
    Dim t As Double, removed As Long, msg As String
    Dim shpName As String, shpId As Long

    On Error Resume Next
    shpName = shp.Name
    shpId = shp.Id
    On Error GoTo 0

    If Not bDebugMode Then On Error GoTo EH
    t = TraceStart("ReplaceMultipleSpaces", slideIndex, shpName, shpId)
    removed = CollapseMultiSpaces(Rng)

    If removed > 0 Then
        g_ChangedElements = g_ChangedElements + 1
        msg = "Заменены множественные пробелы на одиночные (" & removed & ")"
    ElseIf removed = 0 Then
        msg = "Множественных пробелов не найдено"
    Else
        msg = "Сбой при замене множественных пробелов"
    End If

    LogLine "ReplaceMultipleSpaces", msg, slideIndex, shpName, shpId, ElapsedSeconds(t), removed < 0
    TraceEnd "ReplaceMultipleSpaces", t, slideIndex, shpName, shpId
    Exit Sub
EH:
    On Error Resume Next
    LogLine "ReplaceMultipleSpaces", "Ошибка: " & Err.Description, slideIndex, shpName, shpId, 0, True
    TraceEnd "ReplaceMultipleSpaces", t, slideIndex, shpName, shpId
End Sub
