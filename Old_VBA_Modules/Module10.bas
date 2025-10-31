Attribute VB_Name = "Module10"
Sub BilanManpower()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("56:68").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:D").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("J").Hidden = False   ' H ? J
    ActiveWindow.Zoom = 75
    Range("B8").Select
    Application.GoTo Reference:=Range("B8"), Scroll:=True
    Application.ScreenUpdating = True
End Sub

Sub SurfaceDispo()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("70:80").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:J").Hidden = False   ' A:H ? A:J
    
    Sheets("01.3-ITC MASTER WBS").Columns("H").Hidden = True      ' F ? H
    Sheets("01.3-ITC MASTER WBS").Columns("K").Hidden = False     ' I ? K
    Sheets("01.3-ITC MASTER WBS").Columns("C").Hidden = True

    ActiveWindow.Zoom = 62
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True
End Sub

Sub SurfaceEstimee()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("82:93").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:J").Hidden = False   ' A:H ? A:J
    
    Sheets("01.3-ITC MASTER WBS").Columns("K").Hidden = False     ' I ? K

    ActiveWindow.Zoom = 50
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True
End Sub

Sub BilanSurfaces()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("95").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:J").Hidden = False   ' A:H ? A:J
    ActiveWindow.Zoom = 55
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True
End Sub

Sub BilanEnergies()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("97:102").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("103:114").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:D").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("J").Hidden = False   ' H ? J
    Sheets("01.3-ITC MASTER WBS").Columns("L").Hidden = False   ' J ? L
    ActiveWindow.Zoom = 75
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("B95").Select
    Application.ScreenUpdating = True
End Sub

Sub BilanEnergieEncore()
    Dim wsSource As Worksheet
    Dim wsTemp As Worksheet
    Dim lignesPrioritaires As Variant
    Dim ligne As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim NextRow As Long
    Dim ligneDejaCopiee As Object

    Application.ScreenUpdating = False

    Set wsSource = ThisWorkbook.Sheets("01.3-ITC MASTER WBS")

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("VueTemporaire").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = "VueTemporaire"

    lignesPrioritaires = Array(97, 103, 105, 98, 104, 106, 99, 126, 128, 100, 127, 129, 101, 149, 151, 103, 150, 152)

    Set ligneDejaCopiee = CreateObject("Scripting.Dictionary")
    NextRow = 1

    For Each ligne In lignesPrioritaires
        wsSource.Rows(ligne).Copy Destination:=wsTemp.Rows(NextRow)
        ligneDejaCopiee(ligne) = True
        NextRow = NextRow + 1
    Next ligne

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row

    For r = 1 To lastRow
        If wsSource.Rows(r).EntireRow.Hidden = False Then
            If Not ligneDejaCopiee.Exists(r) Then
                wsSource.Rows(r).Copy Destination:=wsTemp.Rows(NextRow)
                NextRow = NextRow + 1
            End If
        End If
    Next r

    wsTemp.Columns.AutoFit
    wsTemp.Activate
    wsTemp.Range("A1").Select
    
    wsTemp.Columns("E").Hidden = True
    wsTemp.Columns("H").Hidden = True   ' F ? H
    wsTemp.Columns("I").Hidden = True   ' G ? I
    wsTemp.Columns("M:R").Hidden = True ' K:P ? M:R
    wsTemp.Columns.AutoFit
    wsTemp.Rows.AutoFit

    Application.ScreenUpdating = True
    MsgBox "Vue temporaire créée avec succès dans l'ordre demandé.", vbInformation
End Sub

Sub SelectionMarchesTravaux()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("165:674").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:D").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("K:O").Hidden = False   ' I:M ? K:O
    ActiveWindow.Zoom = 68
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("L165").Select   ' J165 ? L165
    Application.ScreenUpdating = True
End Sub

Sub ServiesSpecifiquesPar()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("165:674").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:D").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("K").Hidden = False   ' I ? K
    Sheets("01.3-ITC MASTER WBS").Columns("P:R").Hidden = False ' N:P ? P:R
    ActiveWindow.Zoom = 71
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("P165").Select   ' N165 ? P165
    Application.ScreenUpdating = True
End Sub

Sub PhasagePartieB()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("165:674").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A:D").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("K").Hidden = False      ' I ? K
    Sheets("01.3-ITC MASTER WBS").Columns("V:AF").Hidden = False   ' T:AD ? V:AF
    ActiveWindow.Zoom = 65
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("W166").Select    ' U166 ? W166
    Application.ScreenUpdating = True
End Sub

Sub PhasagePartieA()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("10:17").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("L:R").Hidden = False   ' J:P ? L:R
    ActiveWindow.Zoom = 200
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("L10").Select     ' J10 ? L10
    Application.ScreenUpdating = True
End Sub

Sub PhasagePartA()
    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Application.DisplayFullScreen = True
    Sheets("01.3-ITC MASTER WBS").Columns("A:DZ").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1:1000").Hidden = True
    Sheets("01.3-ITC MASTER WBS").Rows("1").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("10:17").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Rows("694:702").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("A").Hidden = False
    Sheets("01.3-ITC MASTER WBS").Columns("L:R").Hidden = False   ' J:P ? L:R
    ActiveWindow.Zoom = 200
    Range("A1").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True
End Sub


