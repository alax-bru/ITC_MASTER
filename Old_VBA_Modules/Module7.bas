Attribute VB_Name = "Module7"
Sub InputClient()

    StatusBar = True
    Application.ScreenUpdating = False
    With ActiveWorkbook.Sheets("01.3-ITC MASTER WBS")
        .Activate
        Application.DisplayFullScreen = True
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        
        .Columns("A:DZ").Hidden = True
        .Rows("1:1000").Hidden = True
        
        .Columns("B:Q").Hidden = False
        .Columns("W:AB").Hidden = False
        
        .Rows("168:674").Hidden = False
        .Rows("7:54").Hidden = False
        .Rows("70:80").Hidden = False
    End With
    
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    ActiveWindow.Zoom = 57
    Application.ScreenUpdating = True

End Sub

