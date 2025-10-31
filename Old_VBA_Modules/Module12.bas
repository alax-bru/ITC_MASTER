Attribute VB_Name = "Module12"
Sub PPBOOS()
    Application.ScreenUpdating = False
    With ActiveWorkbook.Sheets("01.3-ITC MASTER WBS")
        .Activate
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        
        Application.DisplayFullScreen = True
        .Columns("A:DZ").Hidden = True
        .Rows("1:1000").Hidden = True
        
        .Rows("1").Hidden = False
        .Rows("165:674").Hidden = False
        .Rows("694:702").Hidden = False
        
        .Columns("B:D").Hidden = False
        .Columns("K:L").Hidden = False
        .Columns("N:T").Hidden = False
        .Columns("AF:AH").Hidden = False
    End With
    
    ActiveWindow.Zoom = 58
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("L165").Select
    Application.ScreenUpdating = True
End Sub

