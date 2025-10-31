Attribute VB_Name = "Module1"
Sub BilanNumerique()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True

    Sheets("01.3-ITC MASTER WBS").Columns("A").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("B").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("D").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("H").Hidden = False   ' F devient H
    Sheets("01.3-ITC MASTER WBS").Columns("J").Hidden = False   ' H devient J
    Sheets("01.3-ITC MASTER WBS").Columns("K").Hidden = False   ' I devient K
    Sheets("01.3-ITC MASTER WBS").Columns("BW:CA").Hidden = False   ' BU:BY devient BW:CA
    Sheets("01.3-ITC MASTER WBS").Columns("CC:CK").Hidden = False   ' CA:CI devient CC:CK

    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("7:8").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("13").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("55:57").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("59").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("60").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("63").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("65").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("69:96").Hidden = False

    Sheets("01.3-ITC MASTER WBS").Rows("97").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("100").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("103").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("106").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("109").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("112").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("690:707").Hidden = False

    ActiveWindow.Zoom = 57

    Sheets("01.3-ITC MASTER WBS").Rows("1").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("7:8").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("13").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("55:57").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("59").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("60").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("63").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("65").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("69:96").AutoFit

    Sheets("01.3-ITC MASTER WBS").Rows("97").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("100").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("103").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("106").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("109").AutoFit
    Sheets("01.3-ITC MASTER WBS").Rows("112").AutoFit

    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True
End Sub

