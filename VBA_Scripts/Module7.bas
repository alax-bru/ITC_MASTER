' ==========================================================================================================================================================
' ==========================================================================================================================================================
' ================================================================ MACROS GUIDE PARTIE 1 ===================================================================
' ==========================================================================================================================================================
' ==========================================================================================================================================================

Option Explicit

Sub UnhideAllITCMASTER()
ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("01.3-ITC MASTER WBS").Rows("1:5000").Hidden = False
Sheets("01.3-ITC MASTER WBS").Columns("A:ZZ").Hidden = False
Sheets("01.3-ITC MASTER WBS").Columns("F:G").Hidden = True
Application.DisplayFullScreen = False
Range("D9").Select
Application.GoTo Reference:=Range("D9"), Scroll:=True
End Sub


Sub UndoWBS()
ThisWorkbook.Sheets("01.2-WBS & PIC").Activate
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
ThisWorkbook.Sheets("2.7-PS ITC Global").Activate
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
Sheets("2.7-PS ITC Global").Rows("1:5000").Hidden = False
Sheets("2.7-PS ITC Global").Columns("A:ZZ").Hidden = False
Application.DisplayFullScreen = False
Range("D9").Select
Application.GoTo Reference:=Range("D9"), Scroll:=True
End Sub

Sub PSBOOS()

Application.ScreenUpdating = False
ThisWorkbook.Sheets("2.7-PS ITC Global").Activate
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
ThisWorkbook.Sheets("01.2-WBS & PIC").Activate
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
ThisWorkbook.Sheets("2.3-PP & SOW Annexe 1").Activate
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
ThisWorkbook.Sheets("2.3-PP & SOW Annexe 1").Activate
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
ThisWorkbook.Sheets("1.5-Office Layout (INPUT Anx 3)").Activate
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
ThisWorkbook.Sheets("1.5-Office Layout (INPUT Anx 3)").Activate
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
ThisWorkbook.Sheets("2.4-PP & SOW Annexe 2").Activate
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
ThisWorkbook.Sheets("2.4-PP & SOW Annexe 2").Activate
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
ThisWorkbook.Sheets("2.5-PP & SOW Annexe 3").Activate
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
ThisWorkbook.Sheets("2.5-PP & SOW Annexe 3").Activate
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
ThisWorkbook.Sheets("SOW Annexe 4").Activate
ActiveWindow.Zoom = 110
Sheets("SOW Annexe 4").Columns("A:DZ").Hidden = False
Sheets("SOW Annexe 4").Rows("1:10000").Hidden = False
Range("A1").Select
Application.GoTo Reference:=Range("A1"), Scroll:=True
Range("D3").Select
End Sub
Sub UndoAnnexeD()
ThisWorkbook.Sheets("SOW Annexe 4").Activate
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
ThisWorkbook.Sheets("01.1-Site Visit Doc").Activate
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
ThisWorkbook.Sheets("01.1-Site Visit Doc").Activate
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
    With ThisWorkbook.Sheets("01.3-ITC MASTER WBS")
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
ThisWorkbook.Sheets("2.1-Bilan ITC MASTER by familly").Activate
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

Sub InputClient()

    StatusBar = True
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("01.3-ITC MASTER WBS")
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

Sub PriceEstimation()

    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("01.3-ITC MASTER WBS")
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

Sub ReportingGraphique()

Application.ScreenUpdating = False
ThisWorkbook.Sheets("1.4-Bilan Graphique").Activate
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
ThisWorkbook.Sheets("1.4-Bilan Graphique").Activate
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


Sub MacroARenommer()

    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("01.3-ITC MASTER WBS")
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

Sub BilanNumerique()
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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

Sub BilanManpower()
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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
    ThisWorkbook.Sheets("01.3-ITC MASTER WBS").Activate
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


Sub PPBOOS()
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("01.3-ITC MASTER WBS")
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










Sub Rapports_Numer_Graph_Export()

    Application.ScreenUpdating = False
    
    Dim wbSource As Workbook
    Dim wbReport As Workbook
    Dim wsWBS As Worksheet
    Dim wsGraph As Worksheet
    Dim wsRapNum As Worksheet
    Dim wsRapGraph As Worksheet
    
    Dim rngSrcWBS As Range
    Dim rngSrc As Range
    Dim firstRow As Long, firstCol As Long
    Dim rowCount As Long, colCount As Long
    Dim i As Long, j As Long
    Dim graphLastRow As Long
    Dim savePath As String

    Set wbSource = ThisWorkbook
    Set wsWBS = wbSource.Sheets("01.3-ITC MASTER WBS")
    Set wsGraph = wbSource.Sheets("1.4-Bilan Graphique")   ' adapter le nom si besoin
    
    '--- 1. Préparation de la vue WBS ---
    wsWBS.Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    wsWBS.Columns("A:DZ").Hidden = True
    wsWBS.Rows("1:1000").Hidden = True

    wsWBS.Columns("A").Hidden = False
    wsWBS.Columns("B").Hidden = False
    wsWBS.Columns("D").Hidden = False
    wsWBS.Columns("H").Hidden = False   ' F devient H
    wsWBS.Columns("J").Hidden = False   ' H devient J
    wsWBS.Columns("K").Hidden = False   ' I devient K
    wsWBS.Columns("BW:CA").Hidden = False   ' BU:BY devient BW:CA
    wsWBS.Columns("CC:CK").Hidden = False   ' CA:CI devient CC:CK

    wsWBS.Rows("1").Hidden = False
    wsWBS.Rows("7:8").Hidden = False
    wsWBS.Rows("13").Hidden = False
    wsWBS.Rows("55:57").Hidden = False
    wsWBS.Rows("59").Hidden = False
    wsWBS.Rows("60").Hidden = False
    wsWBS.Rows("63").Hidden = False
    wsWBS.Rows("65").Hidden = False
    wsWBS.Rows("69:96").Hidden = False

    wsWBS.Rows("97").Hidden = False
    wsWBS.Rows("100").Hidden = False
    wsWBS.Rows("103").Hidden = False
    wsWBS.Rows("106").Hidden = False
    wsWBS.Rows("109").Hidden = False
    wsWBS.Rows("112").Hidden = False
    wsWBS.Rows("690:707").Hidden = False

    ActiveWindow.Zoom = 57

    wsWBS.Rows("1").AutoFit
    wsWBS.Rows("7:8").AutoFit
    wsWBS.Rows("13").AutoFit
    wsWBS.Rows("55:57").AutoFit
    wsWBS.Rows("59").AutoFit
    wsWBS.Rows("60").AutoFit
    wsWBS.Rows("63").AutoFit
    wsWBS.Rows("65").AutoFit
    wsWBS.Rows("69:96").AutoFit

    wsWBS.Rows("97").AutoFit
    wsWBS.Rows("100").AutoFit
    wsWBS.Rows("103").AutoFit
    wsWBS.Rows("106").AutoFit
    wsWBS.Rows("109").AutoFit
    wsWBS.Rows("112").AutoFit

    wsWBS.Range("A1").Select
    Application.GoTo Reference:=wsWBS.Range("A1"), Scroll:=True
    
    ' ======================================================================
    ' 2. CREATION DU NOUVEAU CLASSEUR "RAPPORT"
    ' ======================================================================
    Set wbReport = Workbooks.Add(xlWBATWorksheet)
    Set wsRapNum = wbReport.Sheets(1)
    On Error Resume Next
    wsRapNum.Name = "Rapport Numérique"
    On Error GoTo 0
    
    ' ---------- 2.1 Rapport Numérique ----------
    Set rngSrcWBS = wsWBS.UsedRange
    If Not rngSrcWBS Is Nothing Then
        
        firstRow = rngSrcWBS.row
        firstCol = rngSrcWBS.Column
        rowCount = rngSrcWBS.Rows.Count
        colCount = rngSrcWBS.Columns.Count
        
        ' Copie complète (valeurs + formats + merges)
        rngSrcWBS.Copy Destination:=wsRapNum.Range("A1")
        Application.CutCopyMode = False
        
        ' Casser les formules (on garde formats / merges)
        wsRapNum.UsedRange.Value = wsRapNum.UsedRange.Value
        
        ' Hauteurs de lignes + masquage identique
        For i = 1 To rowCount
            wsRapNum.Rows(i).RowHeight = wsWBS.Rows(firstRow + i - 1).RowHeight
            wsRapNum.Rows(i).Hidden = wsWBS.Rows(firstRow + i - 1).Hidden
        Next i
        
        ' Largeurs de colonnes + masquage identique
        For j = 1 To colCount
            wsRapNum.Columns(j).ColumnWidth = wsWBS.Columns(firstCol + j - 1).ColumnWidth
            wsRapNum.Columns(j).Hidden = wsWBS.Columns(firstCol + j - 1).Hidden
        Next j
        
    Else
        MsgBox "Aucune cellule à copier dans '01.3-ITC MASTER WBS'.", vbExclamation, "Rapport numérique"
    End If
    
    ' ---------- 2.2 Rapport graphique ----------
    ' Créer / récupérer la feuille "Rapport graphique"
    Set wsRapGraph = Nothing
    On Error Resume Next
    Set wsRapGraph = wbReport.Sheets("Rapport graphique")
    On Error GoTo 0
    
    If wsRapGraph Is Nothing Then
        Set wsRapGraph = wbReport.Sheets.Add(After:=wsRapNum)
        On Error Resume Next
        wsRapGraph.Name = "Rapport graphique"
        On Error GoTo 0
    End If
    
    ' Copier A1:Y100 tel quel, sans masquer de colonnes/lignes
    graphLastRow = 100   ' à adapter si besoin
    
    wsGraph.Range("A1:Y" & graphLastRow).Copy Destination:=wsRapGraph.Range("A1")
    Application.CutCopyMode = False
    
    ' Casser les formules
    wsRapGraph.UsedRange.Value = wsRapGraph.UsedRange.Value
    
    ' Reproduire hauteurs de lignes et largeurs de colonnes
    For i = 1 To graphLastRow
        wsRapGraph.Rows(i).RowHeight = wsGraph.Rows(i).RowHeight
    Next i
    
    For j = 1 To 25   ' A à Y = 25 colonnes
        wsRapGraph.Columns(j).ColumnWidth = wsGraph.Columns(j).ColumnWidth
    Next j
    
    wsRapNum.Activate
    wsRapNum.Range("A1").Select

    ' ======================================================================
    ' 3. SAUVEGARDE DU CLASSEUR RAPPORT
    ' ======================================================================
    If wbSource.Path <> "" Then
        savePath = wbSource.Path & Application.PathSeparator & _
                   "Rapports_Graphiques_et_Numérique_ITC_MASTER.xlsx"
    Else
        ' Si le classeur source n'est pas encore sauvegardé
        savePath = Application.DefaultFilePath & Application.PathSeparator & _
                   "Rapports_Graphiques_et_Numérique_ITC_MASTER.xlsx"
    End If
    
    wbReport.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook  ' .xlsx
    
    ' ======================================================================
    ' 4. RÉINITIALISATION DE LA SHEET SOURCE
    ' ======================================================================
    wbSource.Sheets("01.3-ITC MASTER WBS").Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    wbSource.Sheets("01.3-ITC MASTER WBS").Rows("1:5000").Hidden = False
    wbSource.Sheets("01.3-ITC MASTER WBS").Columns("A:ZZ").Hidden = False
    wbSource.Sheets("01.3-ITC MASTER WBS").Columns("F:G").Hidden = True
    Application.DisplayFullScreen = False
    Range("D9").Select
    Application.GoTo Reference:=Range("D9"), Scroll:=True
    
    Application.ScreenUpdating = True

End Sub



