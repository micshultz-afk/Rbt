Attribute VB_Name = "ReplaceMultipleSpaces"
'==============================================================================
' ReplaceMultipleSpaces — PowerPoint VBA
' Схлопывает любое число пробелов (>1) в один во всём TextRange.
'
' Публичный API:
'   n = CollapseMultiSpaces(rng)                 ' ядро: число удалённых лишних пробелов, -1 при сбое
'   ok = TryCollapseMultiSpaces(rng, removed)    ' то же с Boolean-результатом
'   ReplaceMultipleSpacesDashes rng, slideIndex, shp   ' обёртка под проектный лог/трассировку
'
' Сохраняет форматирование между Runs. Внутри одного Run правка через строку
' (быстро); пробелы на стыке Runs добиваются через TextRange.Replace.
'==============================================================================
Option Explicit

Private Const MAX_BOUNDARY_PASSES As Long = 10000

'==============================================================================
' ПУБЛИЧНОЕ ЯДРО
'==============================================================================

' Схлопывает "  "+ → " " во всём rng.
' Возвращает: число удалённых лишних пробелов; 0 если менять нечего; -1 при ошибке.
Public Function CollapseMultiSpaces(ByRef rng As TextRange) As Long
    Dim s As String
    Dim s2 As String
    Dim run As TextRange
    Dim found As TextRange
    Dim i As Long
    Dim nRuns As Long
    Dim removed As Long
    Dim guard As Long
    Dim beforeLen As Long
    Dim afterLen As Long
    Dim stillHas As Boolean

    On Error GoTo Fail
    CollapseMultiSpaces = 0

    If rng Is Nothing Then Exit Function

    beforeLen = 0
    On Error Resume Next
    beforeLen = Len(rng.Text)
    On Error GoTo Fail
    If beforeLen < 2 Then Exit Function

    s = vbNullString
    On Error Resume Next
    s = rng.Text
    On Error GoTo Fail
    If Len(s) < 2 Then Exit Function
    If InStr(s, "  ") = 0 Then Exit Function

    nRuns = 0
    On Error Resume Next
    nRuns = rng.Runs.Count
    On Error GoTo Fail

    ' Один Run (или Runs недоступны): правка в памяти — O(log n) проходов Replace$,
    ' без COM на каждое вхождение. Форматирование Run однородно → Text безопасен.
    If nRuns <= 1 Then
        s2 = CollapseSpacesString(s)
        If StrComp(s2, s, vbBinaryCompare) <> 0 Then
            rng.Text = s2
            CollapseMultiSpaces = Len(s) - Len(s2)
        End If
        Exit Function
    End If

    ' Несколько Runs: сначала схлопываем внутри каждого Run (с конца — индексы
    ' не «плывут» при укорочении текста). Формат между Runs сохраняется.
    removed = 0
    For i = nRuns To 1 Step -1
        Set run = Nothing
        On Error Resume Next
        Set run = rng.Runs(i)
        On Error GoTo Fail
        If Not run Is Nothing Then
            s = vbNullString
            On Error Resume Next
            s = run.Text
            On Error GoTo Fail
            If Len(s) >= 2 Then
                If InStr(s, "  ") > 0 Then
                    s2 = CollapseSpacesString(s)
                    If StrComp(s2, s, vbBinaryCompare) <> 0 Then
                        run.Text = s2
                        removed = removed + (Len(s) - Len(s2))
                    End If
                End If
            End If
        End If
    Next i

    ' Стыки Runs: " " | " " → визуально два пробела, внутри Run не видны.
    ' TextRange.Replace сохраняет формат и берёт первое вхождение за вызов.
    guard = 0
    Do While guard < MAX_BOUNDARY_PASSES
        s = vbNullString
        On Error Resume Next
        s = rng.Text
        On Error GoTo Fail
        If InStr(s, "  ") = 0 Then Exit Do

        Set found = Nothing
        On Error Resume Next
        Set found = rng.Replace("  ", " ")
        On Error GoTo Fail
        If found Is Nothing Then Exit Do

        removed = removed + 1
        guard = guard + 1
    Loop

    ' Страховка: если после лимита ещё остались двойные пробелы — бинарное
    ' схлопывание по всему диапазону (редко нужно, но не оставляет мусор).
    stillHas = False
    On Error Resume Next
    stillHas = (InStr(rng.Text, "  ") > 0)
    On Error GoTo Fail
    If stillHas Then
        removed = removed + CollapseSpacesBinary(rng)
    End If

    afterLen = beforeLen
    On Error Resume Next
    afterLen = Len(rng.Text)
    On Error GoTo Fail
    If removed <= 0 And afterLen < beforeLen Then
        removed = beforeLen - afterLen
    End If

    CollapseMultiSpaces = removed
    Exit Function

Fail:
    CollapseMultiSpaces = -1
End Function

' Обёртка: True если ошибок не было (в т.ч. когда менять было нечего).
Public Function TryCollapseMultiSpaces(ByRef rng As TextRange, ByRef removed As Long) As Boolean
    On Error GoTo Fail
    removed = CollapseMultiSpaces(rng)
    TryCollapseMultiSpaces = (removed >= 0)
    Exit Function
Fail:
    removed = -1
    TryCollapseMultiSpaces = False
End Function

'==============================================================================
' ОБЁРТКА ПОД существующий проект (TraceStart / LogLine / g_ChangedElements)
' Подключите модуль в ту же VBA-проекцию, где объявлены эти символы.
'==============================================================================

Public Sub ReplaceMultipleSpacesDashes(ByRef Rng As TextRange, ByVal slideIndex As Long, ByRef shp As Shape)
    Dim t As Double
    Dim removed As Long
    Dim shpName As String
    Dim shpId As Long
    Dim errDesc As String

    shpName = vbNullString
    shpId = 0
    On Error Resume Next
    shpName = shp.Name
    shpId = shp.Id
    On Error GoTo 0

    If Not bDebugMode Then On Error GoTo EH

    t = TraceStart("ReplaceMultipleSpaces", slideIndex, shpName, shpId)
    removed = CollapseMultiSpaces(Rng)

    If removed > 0 Then
        g_ChangedElements = g_ChangedElements + 1
        LogLine "ReplaceMultipleSpaces", _
                "Заменены множественные пробелы на одиночные (" & CStr(removed) & ")", _
                slideIndex, shpName, shpId, ElapsedSeconds(t), False
    ElseIf removed = 0 Then
        LogLine "ReplaceMultipleSpaces", _
                "Множественных пробелов не найдено", _
                slideIndex, shpName, shpId, ElapsedSeconds(t), False
    Else
        LogLine "ReplaceMultipleSpaces", _
                "Сбой при замене множественных пробелов", _
                slideIndex, shpName, shpId, ElapsedSeconds(t), True
    End If

    TraceEnd "ReplaceMultipleSpaces", t, slideIndex, shpName, shpId
    Exit Sub

EH:
    errDesc = Err.Description
    On Error Resume Next
    LogLine "ReplaceMultipleSpaces", "Ошибка: " & errDesc, slideIndex, shpName, shpId, ElapsedSeconds(t), True
    TraceEnd "ReplaceMultipleSpaces", t, slideIndex, shpName, shpId
End Sub

'==============================================================================
' ВНУТРЕННИЕ ХЕЛПЕРЫ
'==============================================================================

' Чистая строка: все серии пробелов → один. O(len * log maxRun).
Public Function CollapseSpacesString(ByVal s As String) As String
    Dim guard As Long
    Dim prevLen As Long

    If Len(s) < 2 Then
        CollapseSpacesString = s
        Exit Function
    End If
    If InStr(s, "  ") = 0 Then
        CollapseSpacesString = s
        Exit Function
    End If

    guard = 0
    Do While InStr(s, "  ") > 0
        prevLen = Len(s)
        s = Replace$(s, "  ", " ")
        guard = guard + 1
        ' Защита от теоретического зацикливания (Replace$ не сжал строку).
        If Len(s) >= prevLen Then Exit Do
        If guard > Len(s) + 64 Then Exit Do
    Loop

    CollapseSpacesString = s
End Function

' Бинарное схлопывание через TextRange.Replace: 32→1, 16→1, …, 2→1.
' Меньше COM-вызовов, чем по одному "  " за раз, на длинных сериях пробелов.
Private Function CollapseSpacesBinary(ByRef rng As TextRange) As Long
    Dim chunk As Long
    Dim guard As Long
    Dim total As Long
    Dim found As TextRange
    Dim needle As String
    Dim maxGuard As Long

    On Error GoTo Fail
    CollapseSpacesBinary = 0
    If rng Is Nothing Then Exit Function

    maxGuard = Len(rng.Text) + 32
    If maxGuard < 32 Then maxGuard = 32

    chunk = 32
    Do While chunk >= 2
        needle = String$(chunk, " ")
        guard = 0
        Do While InStr(rng.Text, needle) > 0
            Set found = Nothing
            On Error Resume Next
            Set found = rng.Replace(needle, " ")
            On Error GoTo Fail
            If found Is Nothing Then Exit Do
            total = total + (chunk - 1)
            guard = guard + 1
            If guard > maxGuard Then Exit Do
        Loop
        chunk = chunk \ 2
    Loop

    CollapseSpacesBinary = total
    Exit Function
Fail:
    CollapseSpacesBinary = total
End Function
