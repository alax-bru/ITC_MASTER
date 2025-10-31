Attribute VB_Name = "Module4"
Sub ReportingGraphique()
Attribute ReportingGraphique.VB_ProcData.VB_Invoke_Func = "G\n14"

Application.ScreenUpdating = False
ActiveWorkbook.Sheets("1.4-Bilan Graphique").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("1.4-Bilan Graphique").Rows("2:6").Hidden = True
Application.DisplayFullScreen = True
ActiveWindow.Zoom = 50
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("H8").Select
Application.ScreenUpdating = True


End Sub

Sub UnhideAllGraphique()

Application.ScreenUpdating = False
ActiveWorkbook.Sheets("1.4-Bilan Graphique").Activate
Sheets("1.4-Bilan Graphique").Rows("1:5000").Hidden = False
Sheets("1.4-Bilan Graphique").Columns("A:ZZ").Hidden = False
On Error Resume Next
Sheets("1.4-Bilan Graphique").ShowAllData
On Error GoTo 0
Application.DisplayFullScreen = False
Application.ScreenUpdating = True
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("H8").Select
ActiveWindow.Zoom = 17

End Sub


