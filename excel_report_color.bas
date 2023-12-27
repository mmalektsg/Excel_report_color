Option Explicit

'a sub to color the font of the characters in a string that match the given pattern
'pattern is a string of characters that will be colored
'color is the color to use

Public Sub colorString( _
    ByRef str As String, _
    ByRef pattern As String, _
    ByRef color As Long, _
    ByRef target As Range _
)

Dim i As Long
Dim j As Long
Dim pos As Long
Dim patternLen As Long
Dim nextChar As String

patternLen = Len(pattern)
If Right(pattern, 1) = "*" Then
    patternLen = patternLen - 1
End If

'make it case insensitive
str = LCase(str)
pattern = LCase(pattern)

'reset indexes
i = 1
j = 1

Do While i <= Len(str)
    pos = 0
    For j = 1 To patternLen
        If Mid(str, i + j - 1, 1) <> Mid(pattern, j, 1) Then
            pos = 1
            Exit For
        End If
    Next j
    If pos = 0 Then
        If Right(pattern, 1) = "*" Then
            While i + patternLen <= Len(str) And Not (Mid(str, i + patternLen, 1) Like "[ ,.-;:(){}/\|!@#$%^&*~`<>?""']")
                target.Characters(i + patternLen, 1).Font.color = color
                patternLen = patternLen + 1
            Wend
        End If
        For j = 1 To patternLen
            target.Characters(i + j - 1, 1).Font.color = color
        Next j
        i = i + patternLen
    Else
        i = i + 1
    End If
Loop

End Sub

'a sub to create a dictionary of patterns based on the cell values in column A of the sheet named "Patterns"

Public Sub createPatternDictionary(ByRef rng As Range, ByRef dict As Object)
    Dim cell As Range
    Dim pattern As String
    Dim color As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    dict.comparemode = vbTextCompare
    For Each cell In rng
        pattern = cell.Value
        color = cell.Offset(0, 1).Value
        If IsEmpty(color) Then
            color = vbRed ' domyślna wartość, jeśli komórka jest pusta
        End If
        If dict.exists(pattern) = False Then
            dict.Add pattern, color
        End If
    Next cell
End Sub

'a sub to loop through all the cells in a range and color the font of the characters in a string that match the given pattern

Public Sub colorStringRange()
    Dim cell As Range
    Dim dict As Object
    Dim key As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim tempStr As String

    'disable all interactive features
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Ustaw ws na arkusz 'Color' w ThisWorkbook
    Set ws = ThisWorkbook.Sheets("Color")

    ' Znajdź ostatni wiersz z danymi w kolumnie A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Ustaw rng na wszystkie niepuste wartości z kolumny A
    Set rng = ws.Range("A1:A" & lastRow)

    ' Utwórz słownik wzorców
    createPatternDictionary rng, dict

    'initialise progress bar
    Application.DisplayStatusBar = True
    Application.StatusBar = "Coloring strings..."

    ' Przejdź przez wszystkie komórki w zaznaczonym zakresie
    For Each cell In Selection

        'update progress bar
        Application.StatusBar = "Coloring strings... " & cell.Address & " of " & Selection.Rows.Count

        tempStr = cell.Value

        ' Przejdź przez wszystkie klucze w słowniku
        For Each key In dict.Keys
            ' Koloruj ciąg zgodnie z wzorcem i kolorem
            colorString tempStr, CStr(key), dict(key), cell
        Next key

    Next cell

    're-enable interactive features
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
End Sub