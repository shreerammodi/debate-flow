Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
' Sets sheet name automatically based on value in A1
    If sh.Name = "Info" Or sh.Name = "CX" Or sh.Name = "Decisions" Then Exit Sub
    With sh
        If Not Intersect(Target, .Range("A1")) Is Nothing Then
            On Error GoTo Handler
            Application.EnableEvents = False
            If .Range("A1").Value <> "" Then
                .Name = .Range("A1").Value
            End If
        End If
    End With
Handler:
    Application.EnableEvents = True
End Sub

Private Sub Workbook_Open()
    Dim i As Integer
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set ws = ActiveWorkbook.Worksheets(i)

        ws.Activate
        ActiveWindow.DisplayHeadings = False
    Next i

    Application.ScreenUpdating = True
    ActiveWorkbook.Worksheets(1).Activate

    Startup.Start
End Sub

Private Sub Workbook_NewWindow()
    ActiveWindow.DisplayHeadings = False
End Sub
