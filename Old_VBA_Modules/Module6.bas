Attribute VB_Name = "Module6"
Sub PriceEstimation()

    Application.ScreenUpdating = False
    With ActiveWorkbook.Sheets("01.3-ITC MASTER WBS")
    .Activate
        Application.DisplayFullScreen = True
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        
        .Columns("A:DZ").Hidden = True
        .Rows("2:157").Hidden = True
        
        .Columns("A").Hidden = False
        .Columns("D").Hidden = False
        .Columns("T").Hidden = False
        .Columns("U").Hidden = False
        .Columns("V").Hidden = False
        .Columns("AF").Hidden = False
        .Columns("AG").Hidden = False
        .Columns("AH").Hidden = False
        .Columns("AK").Hidden = False
        .Columns("AN").Hidden = False
        .Columns("AR").Hidden = False
        
        .Rows("1").Hidden = False
        .Rows("158:664").Hidden = False
        
        .Rows("168").Hidden = True
        .Rows("197:209").Hidden = True
        .Rows("239:242").Hidden = True
        .Rows("244:247").Hidden = True
        .Rows("325:328").Hidden = True
        .Rows("480:484").Hidden = True
        .Rows("581").Hidden = True
        .Rows("666:674").Hidden = True
        .Rows("691").Hidden = True
    End With
    
    ActiveWindow.Zoom = 75
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("AR166").Select
    Application.ScreenUpdating = True

End Sub

