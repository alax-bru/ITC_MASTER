Attribute VB_Name = "Module3"
Sub MacroARenommer()

    Application.ScreenUpdating = False
    With ActiveWorkbook.Sheets("01.3-ITC MASTER WBS")
        .Activate
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        Application.DisplayFullScreen = True
        
        ' Réinitialiser
        .Columns("A:DZ").Hidden = False
        .Rows("1:1000").Hidden = False
        
        ' Masquer tout sauf B,C,D,I,J,L:R,AD:AF
        .Columns("A:A").Hidden = True
        .Columns("E:H").Hidden = True
        .Columns("K:K").Hidden = True
        .Columns("S:AC").Hidden = True
        .Columns("AG:DZ").Hidden = True
        
        ' Masquer lignes hors zone
        .Rows("1:165").Hidden = True
    End With

    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("I9").Select   ' F9 ? I9 (décalage +2 après E)
    ActiveWindow.Zoom = 50
    Application.ScreenUpdating = True

End Sub

