Attribute VB_Name = "Module8"
Sub PSBOOS()
Attribute PSBOOS.VB_ProcData.VB_Invoke_Func = "S\n14"

Application.ScreenUpdating = False
ActiveWorkbook.Sheets("2.7-PS ITC Global").Activate
Application.DisplayFullScreen = True
Sheets("2.7-PS ITC Global").Columns("A:DZ").Hidden = False
Sheets("2.7-PS ITC Global").Rows("1:1000").Hidden = False
Sheets("2.7-PS ITC Global").Columns("B").Hidden = True
Sheets("2.7-PS ITC Global").Columns("D:E").Hidden = True

ActiveWindow.Zoom = 84
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("D9").Select
Application.ScreenUpdating = True

End Sub
Sub WBSGraphic()

Application.ScreenUpdating = False
ActiveWorkbook.Sheets("01.2-WBS & PIC").Activate
Application.DisplayFullScreen = True
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("01.2-WBS & PIC").Columns("A:DZ").Hidden = False
Sheets("01.2-WBS & PIC").Rows("1:1000").Hidden = False
Sheets("01.2-WBS & PIC").Rows("1").Hidden = True
Sheets("01.2-WBS & PIC").Columns("V:AA").Hidden = True
ActiveWindow.Zoom = 97
Range("A2").Select
Application.GoTo Reference:=Range("A2"), Scroll:=True
Range("G13").Select
Application.ScreenUpdating = True

End Sub

Sub AnnexeA()
ActiveWorkbook.Sheets("2.3-PP & SOW Annexe 1").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
ActiveWindow.Zoom = 56
Sheets("2.3-PP & SOW Annexe 1").Columns("A:DZ").Hidden = False
Sheets("2.3-PP & SOW Annexe 1").Rows("1:10000").Hidden = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("F4").Select
End Sub
Sub UndoAnnexeA()
ActiveWorkbook.Sheets("2.3-PP & SOW Annexe 1").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.3-PP & SOW Annexe 1").Columns("A:DZ").Hidden = False
Sheets("2.3-PP & SOW Annexe 1").Rows("1:10000").Hidden = False
Application.DisplayFullScreen = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("F4").Select
ActiveWindow.Zoom = 56
End Sub
Sub AnnexeBInput()
ActiveWorkbook.Sheets("1.5-Office Layout (INPUT Anx 3)").Activate
ActiveWindow.Zoom = 59
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("1.5-Office Layout (INPUT Anx 3)").Columns("A:DZ").Hidden = False
Sheets("1.5-Office Layout (INPUT Anx 3)").Rows("1:10000").Hidden = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("F5").Select
End Sub
Sub UndoAnnexeBInput()
ActiveWorkbook.Sheets("1.5-Office Layout (INPUT Anx 3)").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("1.5-Office Layout (INPUT Anx 3)").Columns("A:DZ").Hidden = False
Sheets("1.5-Office Layout (INPUT Anx 3)").Rows("1:10000").Hidden = False
Application.DisplayFullScreen = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("F5").Select
ActiveWindow.Zoom = 59
End Sub
Sub AnnexeBOutput()
ActiveWorkbook.Sheets("2.4-PP & SOW Annexe 2").Activate
ActiveWindow.Zoom = 100
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.4-PP & SOW Annexe 2").Columns("A:DZ").Hidden = False
Sheets("2.4-PP & SOW Annexe 2").Rows("1:10000").Hidden = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("E3").Select
End Sub
Sub UndoAnnexeBOutput()
ActiveWorkbook.Sheets("2.4-PP & SOW Annexe 2").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.4-PP & SOW Annexe 2").Columns("A:DZ").Hidden = False
Sheets("2.4-PP & SOW Annexe 2").Rows("1:10000").Hidden = False
Application.DisplayFullScreen = False
ActiveWindow.Zoom = 100
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("E3").Select
End Sub
Sub AnnexeC()
ActiveWorkbook.Sheets("2.5-PP & SOW Annexe 3").Activate
ActiveWindow.Zoom = 110
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.5-PP & SOW Annexe 3").Columns("A:DZ").Hidden = False
Sheets("2.5-PP & SOW Annexe 3").Rows("1:10000").Hidden = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("C3").Select
End Sub
Sub UndoAnnexeC()
ActiveWorkbook.Sheets("2.5-PP & SOW Annexe 3").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.5-PP & SOW Annexe 3").Columns("A:DZ").Hidden = False
Sheets("2.5-PP & SOW Annexe 3").Rows("1:10000").Hidden = False
Application.DisplayFullScreen = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("C3").Select
ActiveWindow.Zoom = 110
End Sub
Sub AnnexeD()
ActiveWorkbook.Sheets("SOW Annexe 4").Activate
ActiveWindow.Zoom = 110
Sheets("SOW Annexe 4").Columns("A:DZ").Hidden = False
Sheets("SOW Annexe 4").Rows("1:10000").Hidden = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("D3").Select
End Sub
Sub UndoAnnexeD()
ActiveWorkbook.Sheets("SOW Annexe 4").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("SOW Annexe 4").Columns("A:DZ").Hidden = False
Sheets("SOW Annexe 4").Rows("1:10000").Hidden = False
Application.DisplayFullScreen = False
ActiveWindow.Zoom = 110
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("D3").Select
End Sub

Sub SiteVisitDoc()
ActiveWorkbook.Sheets("01.1-Site Visit Doc").Activate
ActiveWindow.Zoom = 60
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("01.1-Site Visit Doc").Columns("A:DZ").Hidden = False
Sheets("01.1-Site Visit Doc").Rows("1:10000").Hidden = False
Sheets("01.1-Site Visit Doc").Columns("B:C").Hidden = False
Application.DisplayFullScreen = True
Range("D4").Select
Application.GoTo Reference:=Range("D4"), Scroll:=True
End Sub
Sub UndoSiteVisitDoc()
ActiveWorkbook.Sheets("01.1-Site Visit Doc").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("01.1-Site Visit Doc").Columns("A:DZ").Hidden = False
Sheets("01.1-Site Visit Doc").Rows("1:10000").Hidden = False
Application.DisplayFullScreen = False
ActiveWindow.Zoom = 60
Range("E6").Select
End Sub
Sub QTYDurMhrs()
    Application.ScreenUpdating = False
    With ActiveWorkbook.Sheets("01.3-ITC MASTER WBS")
        .Activate
        ActiveWindow.Zoom = 46
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        
        .Columns("A:DZ").Hidden = True
        .Rows("1:10000").Hidden = True
        
        .Rows("1").Hidden = False
        .Rows("7:54").Hidden = False
        .Rows("694:701").Hidden = False
        
        ' A:H ? devient A:J après insertion de 2 colonnes
        .Columns("A:J").Hidden = False
    End With
    
    Application.ScreenUpdating = True
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("D8").Select
End Sub
Sub BilanInput()
Application.ScreenUpdating = False
ActiveWorkbook.Sheets("2.1-Bilan ITC MASTER by familly").Activate
ActiveWindow.Zoom = 57
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.1-Bilan ITC MASTER by familly").Columns("A:DZ").Hidden = False
Sheets("2.1-Bilan ITC MASTER by familly").Rows("1:10000").Hidden = False
Application.ScreenUpdating = True
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("F6").Select
End Sub
