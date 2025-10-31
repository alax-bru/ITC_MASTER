Attribute VB_Name = "Module17"

'============================================================================
'                                 Annexe 3c
'============================================================================


Public Sub PP_SOW_8002_FR_Annexe_3c()
    Dim ws As Worksheet
    Dim WordApp As Object, WordDoc As Object
    Dim anchor As Object
    Dim startCell As Range, endCell As Range
    Dim rStart As Long, rEnd As Long, cStart As Long, cEnd As Long
    Dim arr As Variant
    Dim i As Long, j As Long, n As Long, kept As Long
    Dim lines() As String, rowCells() As String
    Dim rngIns As Object, wordTable As Object
    Dim cheminDoc As String
    Dim w() As Double, perc() As Double, tot As Double
    Dim t0 As Single: t0 = Timer

    ' ---- Feuille source
    Set ws = ThisWorkbook.Worksheets("2.5-PP & SOW Annexe 3")

    ' ---- Repï¿½res
    With ws.Cells
        Set startCell = .Find("4 Lignes au dessus de debut Annexe 3c", LookAt:=xlWhole, LookIn:=xlValues)
        Set endCell = .Find("Cellule 4 Lignes Après Dernière Cellule Range Annexe 3c", LookAt:=xlWhole, LookIn:=xlValues)
    End With
    If startCell Is Nothing Or endCell Is Nothing Then
        MsgBox "Repï¿½res introuvables", vbCritical: Exit Sub
    End If

    ' ---- Dï¿½but/fin lignes avec +4 et -4
    rStart = startCell.row + 4
    rEnd = endCell.row - 4
    If rStart > rEnd Then MsgBox "Plage vide", vbExclamation: Exit Sub

    ' ---- Colonnes = de la colonne du 1er repï¿½re ï¿½ celle du 2ï¿½me repï¿½re
    cStart = startCell.Column
    cEnd = endCell.Column
    If cStart > cEnd Then
        Dim tmp As Long: tmp = cStart: cStart = cEnd: cEnd = tmp
    End If

    ' ---- Charger plage
    arr = ws.Range(ws.Cells(rStart, cStart), ws.Cells(rEnd, cEnd)).Value

    ' ---- Construire texte tabulï¿½
    n = UBound(arr, 1)
    ReDim lines(1 To n)
    ReDim rowCells(1 To (cEnd - cStart + 1))
    kept = 0
    For i = 1 To n
        Dim allEmpty As Boolean: allEmpty = True
        For j = 1 To (cEnd - cStart + 1)
            rowCells(j) = CleanCellText(CStr(arr(i, j)))
            If rowCells(j) <> "" Then allEmpty = False
        Next j
        If Not allEmpty Then
            kept = kept + 1
            lines(kept) = Join(rowCells, vbTab)
        End If
    Next i
    If kept = 0 Then MsgBox "Toutes les lignes vides", vbInformation: Exit Sub
    If kept < n Then ReDim Preserve lines(1 To kept)

    ' ---- Word : ouvrir et chercher repï¿½re
    cheminDoc = ThisWorkbook.Path & Application.PathSeparator & "PP_8002-FR.dotx"
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    WordApp.Visible = True
    Set WordDoc = WordApp.Documents.Open(cheminDoc)

    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3c)": .Forward = True: .Wrap = 1
        .Execute
    End With
    If Not anchor.Find.Found Then MsgBox "Repï¿½re '(Annexe 3c)' introuvable dans Word": Exit Sub
    anchor.Text = "": anchor.Collapse 0

    ' ---- Insertion
    Set rngIns = WordDoc.Range(anchor.Start, anchor.Start)
    rngIns.Text = Join(lines, vbCr) & vbCr
    Set wordTable = rngIns.ConvertToTable(Separator:=1, NumColumns:=(cEnd - cStart + 1))

    ' ---- Style + 1ï¿½re ligne
    On Error Resume Next
    wordTable.Range.Style = "Text in table"
    If Err.Number <> 0 Then
        Err.Clear: wordTable.Range.Style = "Texte dans le tableau"
    End If
    On Error GoTo 0
    wordTable.Rows(1).Range.Font.Bold = True
    wordTable.Rows(1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

    ' ---- Supprimer espace aprï¿½s paragraphe
    wordTable.Range.ParagraphFormat.SpaceAfter = 0

    ' ---- Largeurs depuis Excel
    Dim colCount As Long: colCount = cEnd - cStart + 1
    ReDim w(1 To colCount)
    ReDim perc(1 To colCount)
    For j = 1 To colCount
        w(j) = ws.Columns(cStart + j - 1).ColumnWidth
        If w(j) <= 0 Then w(j) = 1
        tot = tot + w(j)
    Next j
    For j = 1 To colCount
        perc(j) = (w(j) / tot) * 100#
        wordTable.Columns(j).PreferredWidthType = 2
        wordTable.Columns(j).PreferredWidth = perc(j)
    Next j
    wordTable.PreferredWidthType = 2
    wordTable.PreferredWidth = 100

    ' ---- Bordures
    Dim k As Long: For k = 1 To 6: wordTable.Borders(k).LineStyle = 1: Next k

    ' ---- Avancer ancre
    anchor.SetRange wordTable.Range.End, wordTable.Range.End
    anchor.InsertParagraphAfter
    anchor.Collapse 0

    ' ---- Timer affichï¿½
    MsgBox "Annexe 3c exportï¿½e : " & (cEnd - cStart + 1) & " colonnes, " & kept & " lignes." & vbCrLf & _
           "Temps ï¿½coulï¿½ : " & Format(Timer - t0, "0.00") & " s", vbInformation
End Sub

Private Function CleanCellText(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    CleanCellText = Trim$(s)
End Function









