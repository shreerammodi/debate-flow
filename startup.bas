Public Function GetModifierKey() As String
    ' Mac Excel versions after 2011 can't assign shortcuts to Command (character code = *), so we have to use Ctrl for now
    #If Mac Then
        GetModifierKey = "^" ' Ctrl Key
    #Else
        GetModifierKey = "^" ' Ctrl Key
    #End If
End Function


Public Sub Start()
    Startup.SetKeyboardShortcuts
End Sub

Public Sub SetKeyboardShortcuts()
' + = Shift, ^ = Ctrl, % = Alt, * = Command
    Dim Modifier As String
    Modifier = Startup.GetModifierKey

    Application.OnKey Modifier & "+i", "Flow.InsertCellAbove"
    Application.OnKey Modifier & "+o", "Flow.InsertRowAbove"

    Application.OnKey Modifier & "+k", "Flow.ShiftUp"
    Application.OnKey Modifier & "+j", "Flow.ShiftDown"

    Application.OnKey Modifier & "+a", "Flow.ToggleEvidence"
    Application.OnKey Modifier & "+h", "Flow.ToggleHighlighting"

    Application.OnKey Modifier & "+n", "Flow.CreateArgumentSection"
    Application.OnKey Modifier & "+x", "Flow.RemoveArgumentSection"

    Application.OnKey Modifier & "+y", "Flow.ShiftSectionUp"
    Application.OnKey Modifier & "+e", "Flow.ShiftSectionDown"

    Application.OnKey Modifier & "+u", "Flow.ExtendSectionUp"
    Application.OnKey Modifier & "+m", "Flow.ExtendSectionDown"
End Sub
