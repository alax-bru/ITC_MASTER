Attribute VB_Name = "Module2"
Sub UnhideAllITCMaster()
Attribute UnhideAllITCMaster.VB_ProcData.VB_Invoke_Func = "U\n14"
ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
Application.ScreenUpdating = False
Sheets("01.3-ITC MASTER WBS").Rows("1:5000").Hidden = False
Sheets("01.3-ITC MASTER WBS").Columns("A:ZZ").Hidden = False
Sheets("01.3-ITC MASTER WBS").Columns("F:G").Hidden = True
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Application.DisplayFullScreen = False
Application.ScreenUpdating = True
ActiveWindow.Zoom = 60
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("D8").Select
' Activer la ligne au dessous si on veut que le bouton réinitialisation ouvre la sheet de guide
' ActiveWorkbook.Sheets("06.0-Guide").Activate
End Sub

