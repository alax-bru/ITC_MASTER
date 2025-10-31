Attribute VB_Name = "Module9"
Sub UndoWBS()
ActiveWorkbook.Sheets("01.2-WBS & PIC").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("01.2-WBS & PIC").Columns("A:DZ").Hidden = False
Sheets("01.2-WBS & PIC").Rows("1:10000").Hidden = False
Application.DisplayFullScreen = False
ActiveWindow.Zoom = 40
Range("A1").Select
Application.GoTo Reference:=Range("A2"), Scroll:=True
End Sub

Sub UndoPSBOOS()
ActiveWorkbook.Sheets("2.7-PS ITC Global").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.7-PS ITC Global").Rows("1:5000").Hidden = False
Sheets("2.7-PS ITC Global").Columns("A:ZZ").Hidden = False
Application.DisplayFullScreen = False
Range("D9").Select
Application.GoTo Reference:=Range("D9"), Scroll:=True
End Sub

