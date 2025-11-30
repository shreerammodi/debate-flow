Function GetMaxNum(Side As String)
    Dim maxNum As Integer
    Dim ws As Worksheet

    maxNum = 0

    For Each ws in ActiveWorkbook.Worksheets
        If Left(ws.Name, 3) = Side Then
            Dim numStr As String
            numStr = Mid(ws.Name, 5)

            If IsNumeric(numStr) Then
                If CInt(NumStr) > maxNum Then
                    maxNum = CInt(numStr)
                End If
            End If
        End If
    Next ws

    GetMaxNum = maxNum
End Function

Sub CreateNEGSheets()
    Dim sheetCount As Integer
    Dim i As Integer
    Dim newSheet As Worksheet
    Dim maxNum As Integer

    With ActiveWorkbook
        sheetCount = Sheets("Info").Range("I20").Value

        For i = 1 To sheetCount Step 1
            maxNum = GetMaxNum("NEG")

            Sheets("NEG").Visible = True
            Worksheets("NEG").Copy After:=Sheets(Sheets.Count)

            Set newSheet = Sheets(Sheets.Count)
            newSheet.Name = "NEG-" & maxNum + 1

            Sheets("NEG").Visible = False
        Next i

    End With
    Sheets("Info").Select
End Sub

Sub CreateAFFSheets()
    Dim sheetCount As Integer
    Dim i As Integer
    Dim newSheet As Worksheet
    Dim maxNum As Integer

    With ActiveWorkbook
        sheetCount = Sheets("Info").Range("I20").Value

        For i = 1 To sheetCount Step 1
            maxNum = GetMaxNum("AFF")

            Sheets("AFF").Visible = True
            Worksheets("AFF").Copy After:=Sheets(Sheets.Count)

            Set newSheet = Sheets(Sheets.Count)
            newSHeet.Name = "AFF-" & maxNum + 1

            Sheets("AFF").Visible = False
        Next i

    End With
    Sheets("Info").Select
End Sub
