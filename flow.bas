Public Sub ShiftUp()
    If Selection.Row < 3 Then Exit Sub
    Selection.Rows(Selection.Rows.Count + 1).Insert Shift:=xlDown
    Selection.Rows(1).Offset(-1).Cut Selection.Rows(Selection.Rows.Count + 1)
    Selection.Rows(1).Offset(-1).Delete Shift:=xlUp
    Selection.Offset(-1).Select
End Sub

Public Sub ShiftDown()
    Selection.Rows(1).Insert Shift:=xlDown
    Selection.Rows(Selection.Rows.Count).Offset(2).Cut Selection.Rows(1)
    Selection.Rows(Selection.Rows.Count).Offset(2).Delete Shift:=xlUp
    Selection.Offset(1).Select
End Sub

Public Sub InsertRowAbove()
    Dim r As Range
    Set r = ActiveCell.EntireRow
    r.Insert Shift:=xlDown
    r.Offset(-1).ClearContents
    Set r = Nothing
End Sub

Public Sub InsertCellAbove()
    Dim r As Range
    Set r = ActiveCell
    r.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    r.Offset(-1, 0).ClearContents
    Set r = Nothing
End Sub

Public Sub ToggleEvidence()
    Dim targetRange As Range
    Dim origColor As Long

    origColor = Cells(2, ActiveCell.Column).Font.Color
    
    If ActiveCell.Font.ColorIndex <> 1 Then
        ActiveCell.Font.ColorIndex = 1
        ActiveCell.Font.Bold = True
    Else
        ActiveCell.Font.Color = origColor
        ActiveCell.Font.Bold = False
    End If
End Sub

Public Sub ToggleHighlighting()
    If ActiveCell.Interior.ColorIndex = 6 Then
        ActiveCell.Interior.ColorIndex = xlNone
    Else
        ActiveCell.Interior.ColorIndex = 6
    End If
End Sub

Public Sub SwitchSpeech()
    Dim speech As String
    Dim i As Integer
    Dim ws As Worksheet
    Dim wsSide As String
    Dim col As Integer

    speech = SpeechName()

    For i = 6 To ActiveWorkbook.Worksheets.Count
        Set ws = ActiveWorkbook.Worksheets(i)

        If ws.Range("A2").Value = "1AC" Then
            wsSide = "AFF"
        Else
            wsSide = "NEG"
        End If

        col = ColumnFromSpeech(speech, wsSide)


        If col > 0 Then
            ws.Activate
            ws.Cells(3, col).Select
        End If
    Next i


    ActiveWorkbook.Worksheets(6).Activate
End Sub

Private Function ColumnFromSpeech(SpeechName As String, Side As String) As Integer
    Dim baseColumn As Integer

    Select Case SpeechName
        Case "1AC"
            baseColumn = 1
        Case "1NC"
            baseColumn = 2
        Case "2AC"
            baseColumn = 3
        Case "Block"
            baseColumn = 4
        Case "1AR"
            baseColumn = 5
        Case "2NR"
            baseColumn = 6
        Case "2AR"
            baseColumn = 7
        Case Else
            ColumnFromSpeech = 0
            Exit Function
    End Select

    If Side = "AFF" Then
        ColumnFromSpeech = baseColumn
    ElseIf Side = "NEG" Then
        ColumnFromSpeech = baseColumn - 1
    Else
        ColumnFromSpeech = 0
    End If
End Function

Private Function SpeechName As String
    Dim i As Integer

    For i = 37 To 32 Step -1
        If Range("C" & i).Value = True Then
            SpeechName = Range("B" & i).Value
            Exit Function
        End If
    Next i

    SpeechName = Range("B31").Value
End Function

Public Sub CreateArgumentSection()
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    startRow = Selection.Row
    endRow = Selection.Row + Selection.Rows.Count - 1
    
    ws.Rows(startRow).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Rows(startRow).Borders(xlEdgeTop).Weight = xlMedium
    ws.Rows(startRow).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
    
    ws.Rows(endRow).Borders(xlEdgeBottom).LineStyle = xlDot
    ws.Rows(endRow).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Rows(endRow).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
End Sub

Public Sub ExtendSectionDown()
    Dim ws As Worksheet
    Dim currentRow As Long
    
    Set ws = ActiveSheet
    currentRow = ActiveCell.Row
    
    ws.Rows(currentRow).Borders(xlEdgeBottom).LineStyle = xlNone

    For i = currentRow + 1 To currentRow + 3
        ws.Rows(i).Borders(xlEdgeBottom).LineStyle = xlNone
    Next i
    
    ws.Rows(currentRow + 3).Borders(xlEdgeBottom).LineStyle = xlDot
    ws.Rows(currentRow + 3).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Rows(currentRow + 3).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
End Sub

Public Sub ExtendSectionUp()
    Dim ws As Worksheet
    Dim currentRow As Long
    
    Set ws = ActiveSheet
    currentRow = ActiveCell.Row
    
    ws.Rows(currentRow).Borders(xlEdgeTop).LineStyle = xlNone

    For i = currentRow - 3 To currentRow - 1
        ws.Rows(i).Borders(xlEdgeTop).LineStyle = xlNone
    Next i

    ws.Rows(currentRow - 3).Borders(xlEdgeTop).LineStyle = xlDot
    ws.Rows(currentRow - 3).Borders(xlEdgeTop).Weight = xlMedium
    ws.Rows(currentRow - 3).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
End Sub

Public Sub RemoveArgumentSection()
    Dim ws As Worksheet
    Dim topRow As Long, bottomRow As Long
    
    Set ws = ActiveSheet
    
    topRow = ActiveCell.Row
    Do While topRow > 1
        If ws.Rows(topRow).Borders(xlEdgeTop).LineStyle <> xlNone Then
            Exit Do
        End If
        topRow = topRow - 1
    Loop
    
    bottomRow = ActiveCell.Row
    Do While bottomRow <= ws.Rows.Count
        If ws.Rows(bottomRow).Borders(xlEdgeBottom).LineStyle <> xlNone Then
            Exit Do
        End If
        bottomRow = bottomRow + 1
    Loop
    
    If ws.Rows(topRow).Borders(xlEdgeTop).LineStyle = xlNone And _
       ws.Rows(bottomRow).Borders(xlEdgeBottom).LineStyle = xlNone Then
        Exit Sub
    End If
    
    ws.Rows(topRow).Borders(xlEdgeTop).LineStyle = xlNone
    ws.Rows(bottomRow).Borders(xlEdgeBottom).LineStyle = xlNone
End Sub

Public Sub ShiftSectionUp()
    Dim topRow As Long, bottomRow As Long
    Dim ws As Worksheet
    Dim currentCol As Long
    Dim offsetFromTop As Long
    Dim newCursorRow As Long
    Dim foundTopBorder As Boolean
    Dim foundBottomBorder As Boolean
    
    Set ws = ActiveSheet
    currentCol = ActiveCell.Column
    foundTopBorder = False
    foundBottomBorder = False
    
    topRow = ActiveCell.Row
    Do While topRow > 1
        If ws.Rows(topRow).Borders(xlEdgeTop).LineStyle <> xlNone Then
            foundTopBorder = True
            Exit Do
        End If
        topRow = topRow - 1
    Loop
    
    bottomRow = ActiveCell.Row
    Do While bottomRow <= ws.Rows.Count
        If ws.Rows(bottomRow).Borders(xlEdgeBottom).LineStyle <> xlNone Then
            foundBottomBorder = True
            Exit Do
        End If
        bottomRow = bottomRow + 1
    Loop
    
    If Not foundTopBorder Or Not foundBottomBorder Then
        MsgBox "No bordered section found. Make sure cursor is within a section with top and bottom borders.", vbExclamation
        Exit Sub
    End If
    
    If topRow <= 3 Then
        MsgBox "Cannot move section above row 3.", vbExclamation
        Exit Sub
    End If
    
    offsetFromTop = ActiveCell.Row - topRow
    
    On Error Resume Next
    ws.Rows(topRow - 1).Cut
    ws.Rows(bottomRow + 1).Insert Shift:=xlDown
    Application.CutCopyMode = False
    On Error GoTo 0
    
    newCursorRow = (topRow - 1) + offsetFromTop
    ws.Cells(newCursorRow, currentCol).Select
End Sub

Public Sub ShiftSectionDown()
    Dim topRow As Long, bottomRow As Long
    Dim ws As Worksheet
    Dim currentCol As Long
    Dim offsetFromTop As Long
    Dim newCursorRow As Long
    Dim foundTopBorder As Boolean
    Dim foundBottomBorder As Boolean
    
    Set ws = ActiveSheet
    currentCol = ActiveCell.Column
    foundTopBorder = False
    foundBottomBorder = False
    
    topRow = ActiveCell.Row
    Do While topRow > 1
        If ws.Rows(topRow).Borders(xlEdgeTop).LineStyle <> xlNone Then
            foundTopBorder = True
            Exit Do
        End If
        topRow = topRow - 1
    Loop
    
    bottomRow = ActiveCell.Row
    Do While bottomRow <= ws.Rows.Count
        If ws.Rows(bottomRow).Borders(xlEdgeBottom).LineStyle <> xlNone Then
            foundBottomBorder = True
            Exit Do
        End If
        bottomRow = bottomRow + 1
    Loop
    
    If Not foundTopBorder Or Not foundBottomBorder Then
        Exit Sub
    End If
    
    offsetFromTop = ActiveCell.Row - topRow
    
    On Error Resume Next
    ws.Rows(bottomRow + 1).Cut
    ws.Rows(topRow).Insert Shift:=xlDown
    Application.CutCopyMode = False
    On Error GoTo 0
    
    newCursorRow = (topRow + 1) + offsetFromTop
    ws.Cells(newCursorRow, currentCol).Select
End Sub
