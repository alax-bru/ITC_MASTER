Attribute VB_Name = "Module13"
Option Explicit

' =====================================================================
'  PP8002_AnnexeX — Export Excel -> Word (Annexe 3 Input)
' ---------------------------------------------------------------------
'  DESIGN / CONTRATS
'   • Parcours C10:L125 (paramétrable via R_START/R_END/C_START/C_END).
'   • Détecte les blocs de lignes non vides ? un tableau Word par bloc.
'   • Si l'ancre est DANS un tableau modèle Word : on conserve ses largeurs
'     et son style, puis on remplace à l’identique (pas de redistribution).
'   • Sinon : largeurs Word = % des largeurs Excel (col C..L), pas d’égalisation.
'   • Fusions Excel (ligne/col) répliquées côté Word (ordre bottom-right).
'   • Contenu cellule = texte AFFICHÉ (NumberFormatLocal) + style global (Bold/Italic/Underline/Color).
'   • Orientation par cellule (horizontal / vertical / upward / downward / angle num.).
'   • Repère "Annexe 3 Input" cherché **UNE FOIS** (sinon abort).
' =====================================================================

' === Constantes Word (late binding) ===
Private Const wdAutoFitFixed As Long = 0
Private Const wdLineStyleSingle As Long = 1
Private Const wdLineWidth050pt As Long = 2
Private Const wdLineWidth025pt As Long = 1
Private Const wdRowHeightAtLeast As Long = 1

' Orientation texte Word
Private Const wdTextOrientationHorizontal As Long = 0
Private Const wdTextOrientationVerticalFarEast As Long = 1
Private Const wdTextOrientationUpward As Long = 2
Private Const wdTextOrientationDownward As Long = 3

' Indices de bordures
Private Const B_TOP As Long = 1, B_LEFT As Long = 2, B_BOTTOM As Long = 3, B_RIGHT As Long = 4
Private Const B_HORZ As Long = 5, B_VERT As Long = 6

' === Plage source (C:10..L:125) ===
Private Const R_START As Long = 14     ' ligne 14
Private Const R_END   As Long = 125    ' ligne 125
Private Const C_START As Long = 4      ' Col D
Private Const C_END   As Long = 10     ' Col J
Private Const N_COLS  As Long = 6      ' 6 colonnes

Public Sub PP8002_Annexe3a()
    Dim ws As Worksheet
    Dim WordApp As Object, WordDoc As Object
    Dim cheminDoc As String
    Dim anchor As Object
    Dim t0 As Single
    Dim i As Long, blockStart As Long, blockEnd As Long

    t0 = Timer
    Set ws = ThisWorkbook.Worksheets("1.5-Office Layout (INPUT Anx 3)")

    ' --- Word (late binding)
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    If WordApp Is Nothing Then
        MsgBox "Impossible d'ouvrir Word.", vbCritical: Exit Sub
    End If
    WordApp.Visible = True

    cheminDoc = ThisWorkbook.Path & Application.PathSeparator & "PP_8002-FR.dotx"
    Set WordDoc = WordApp.Documents.Open(cheminDoc)

    ' === Chercher "Annexe 3 Input" **UNE SEULE FOIS** (sinon abort)
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3a)": .Forward = True: .Wrap = 1: .Execute
    End With
    If Not anchor.Find.Found Then
        MsgBox "Repère 'Annexe 3a' introuvable dans " & cheminDoc, vbExclamation
        Exit Sub
    End If
    anchor.Text = "": anchor.Collapse 0     ' 0 = wdCollapseEnd

    ' === Parcours des blocs de lignes non vides
    i = R_START
    Do While i <= R_END
        Do While i <= R_END And IsRowEmptyOnRange(ws, i, C_START, C_END): i = i + 1: Loop
        If i > R_END Then Exit Do
        blockStart = i

        Do While i <= R_END And Not IsRowEmptyOnRange(ws, i, C_START, C_END): i = i + 1: Loop
        blockEnd = i - 1

        CreateAndFillTableFromBlock ws, WordDoc, WordApp, anchor, blockStart, blockEnd
        DoEvents
    Loop

    MsgBox "Annexe 3 Input exportée (C10:L125) ? " & cheminDoc & vbCrLf & _
           "Temps : " & Format(Timer - t0, "0.00") & " s", vbInformation
End Sub

' ---------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------
Private Function IsRowEmptyOnRange(ws As Worksheet, ByVal rowNum As Long, ByVal cStart As Long, ByVal cEnd As Long) As Boolean
    Dim c As Long
    For c = cStart To cEnd
        If LenB(Trim$(CStr(ws.Cells(rowNum, c).Value))) <> 0 Then IsRowEmptyOnRange = False: Exit Function
    Next c
    IsRowEmptyOnRange = True
End Function

Private Function GetDisplayTextSafe(ByVal cel As Range) As String
    Dim s As String: s = CStr(cel.Text)
    If InStr(s, "#") > 0 Or Len(s) = 0 Then
        On Error Resume Next
        s = Format$(cel.Value, cel.NumberFormatLocal)
        If Err.Number <> 0 Or Len(s) = 0 Then s = CStr(cel.Value)
        On Error GoTo 0
    End If
    GetDisplayTextSafe = s
End Function

Private Function MapExcelOrientationToWord(xlCell As Range) As Long
    Dim o As Variant: o = xlCell.Orientation
    On Error Resume Next
    Select Case o
        Case xlVertical: MapExcelOrientationToWord = wdTextOrientationVerticalFarEast
        Case xlUpward:   MapExcelOrientationToWord = wdTextOrientationUpward
        Case xlDownward: MapExcelOrientationToWord = wdTextOrientationDownward
        Case xlHorizontal, 0: MapExcelOrientationToWord = wdTextOrientationHorizontal
        Case Else
            If IsNumeric(o) Then
                If o >= 45 Then
                    MapExcelOrientationToWord = wdTextOrientationUpward
                ElseIf o <= -45 Then
                    MapExcelOrientationToWord = wdTextOrientationDownward
                Else
                    MapExcelOrientationToWord = wdTextOrientationHorizontal
                End If
            Else
                MapExcelOrientationToWord = wdTextOrientationHorizontal
            End If
    End Select
    On Error GoTo 0
End Function

Private Sub ApplyCellTextDirection(wdCellRange As Object, ByVal wdOrientation As Long)
    On Error Resume Next
    wdCellRange.Orientation = wdOrientation
    On Error GoTo 0
End Sub

Private Sub ApplyExcelMergesToWord(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                   ByVal c1 As Long, ByVal c2 As Long, ByVal tbl As Object)
    Dim r As Long, c As Long, merges() As Variant, m As Long
    Dim xlCell As Range, area As Range
    Dim topRow As Long, leftCol As Long, bottomRow As Long, rightCol As Long
    Dim iTop As Long, iLeft As Long, iBottom As Long, iRight As Long
    Dim hRows As Long, hCols As Long
    Dim wr As Long, wc As Long, wr2 As Long, wc2 As Long
    Dim i As Long, j As Long, a As Variant, b As Variant, tmp As Variant

    m = 0
    For r = r1 To r2
        For c = c1 To c2
            Set xlCell = ws.Cells(r, c)
            If xlCell.MergeCells Then
                Set area = xlCell.MergeArea
                topRow = area.row: leftCol = area.Column
                bottomRow = topRow + area.Rows.Count - 1
                rightCol = leftCol + area.Columns.Count - 1

                iTop = IIf(topRow > r1, topRow, r1)
                iLeft = IIf(leftCol > c1, leftCol, c1)
                iBottom = IIf(bottomRow < r2, bottomRow, r2)
                iRight = IIf(rightCol < c2, rightCol, c2)

                If (iTop <= iBottom) And (iLeft <= iRight) Then
                    If r = iTop And c = iLeft Then
                        hRows = iBottom - iTop + 1: hCols = iRight - iLeft + 1
                        If hRows > 1 Or hCols > 1 Then
                            wr = iTop - r1 + 1: wc = iLeft - c1 + 1
                            wr2 = wr + hRows - 1: wc2 = wc + hCols - 1
                            m = m + 1
                            If m = 1 Then
                                ReDim merges(1 To 1)
                            Else
                                ReDim Preserve merges(1 To m)
                            End If
                            merges(m) = Array(wr, wc, wr2, wc2)
                        End If
                    End If
                End If
            End If
        Next c
    Next r
    If m = 0 Then Exit Sub

    For i = 1 To m - 1
        For j = i + 1 To m
            a = merges(i): b = merges(j)
            If (b(1) > a(1)) Or ((b(1) = a(1)) And (b(0) > a(0))) Then
                tmp = merges(i): merges(i) = merges(j): merges(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To m
        a = merges(i)
        On Error Resume Next
        tbl.cell(a(0), a(1)).Merge tbl.cell(a(2), a(3))
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------------
' Crée/remplit un tableau Word depuis le bloc r1..r2 / C_START..C_END
' ---------------------------------------------------------------------
Private Sub CreateAndFillTableFromBlock(ws As Worksheet, WordDoc As Object, WordApp As Object, _
                                        ByRef anchor As Object, ByVal r1 As Long, ByVal r2 As Long)

    Dim numRows As Long: numRows = r2 - r1 + 1
    Dim tbl As Object, tRng As Object
    Dim r As Long, c As Long
    Dim cellRng As Object
    Dim vU As Variant

    Dim useModel As Boolean
    Dim modelTbl As Object
    Dim savedW() As Single
    Dim i As Long
    Dim modelStyle As Variant

    Dim weights(1 To N_COLS) As Double
    Dim doApplyExcelWidths As Boolean

    useModel = False
    On Error Resume Next
    If anchor.Cells.Count > 0 Then
        Set modelTbl = anchor.Cells(1).Table
        If Not modelTbl Is Nothing Then useModel = True
    End If
    On Error GoTo 0

    If useModel Then
        ReDim savedW(1 To N_COLS)
        For i = 1 To N_COLS
            On Error Resume Next
            If i <= modelTbl.Columns.Count Then
                savedW(i) = modelTbl.Columns(i).Width
            Else
                savedW(i) = modelTbl.Columns(modelTbl.Columns.Count).Width
            End If
            On Error GoTo 0
        Next i
        On Error Resume Next: modelStyle = modelTbl.Style: On Error GoTo 0
        Set tRng = modelTbl.Range
        tRng.Delete
        doApplyExcelWidths = False
    Else
        Set tRng = WordDoc.Range(anchor.Start, anchor.Start)
        doApplyExcelWidths = True
        For c = 1 To N_COLS
            weights(c) = ws.Columns(C_START + c - 1).ColumnWidth
            If weights(c) <= 0 Then weights(c) = 1
        Next c
    End If

    Set tbl = WordDoc.Tables.Add(Range:=tRng, numRows:=numRows, NumColumns:=N_COLS)

    With tbl
        On Error Resume Next
        .AutoFitBehavior wdAutoFitFixed
        .TopPadding = 0: .BottomPadding = 0: .LeftPadding = 0: .RightPadding = 0
        .AllowAutoFit = False
        .Rows.AllowBreakAcrossPages = True
        .Rows.HeightRule = wdRowHeightAtLeast
        If useModel Then
            If Not IsEmpty(modelStyle) Then .Style = modelStyle
            For i = 1 To N_COLS: .Columns(i).Width = savedW(i): Next i
        Else
            .Style = "Grille du tableau": If Err.Number <> 0 Then Err.Clear: .Style = "Table Grid"
        End If
        Dim idx As Variant
        For Each idx In Array(B_TOP, B_LEFT, B_BOTTOM, B_RIGHT, B_HORZ, B_VERT)
            With .Borders(idx)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
            End With
        Next idx
        On Error GoTo 0
    End With

    If doApplyExcelWidths Then
        ApplyWidthsFromExcelWeights WordDoc, tbl, weights
    End If

    For r = 1 To numRows
        For c = 1 To N_COLS
            With tbl.cell(r, c)
                Set cellRng = .Range: cellRng.End = cellRng.End - 1
                cellRng.Text = GetDisplayTextSafe(ws.Cells(r1 + r - 1, C_START + c - 1))
                On Error Resume Next
                cellRng.Style = "Text in table"
                cellRng.Font.Size = 8   ' Forcer police 8
                ' Alignement conditionnel
                If c <= 2 Then
                    cellRng.ParagraphFormat.Alignment = 0   ' wdAlignParagraphLeft
                Else
                    cellRng.ParagraphFormat.Alignment = 1   ' wdAlignParagraphCenter
                End If
                On Error GoTo 0

                If Not IsNull(ws.Cells(r1 + r - 1, C_START + c - 1).Font.Bold) Then _
                    cellRng.Font.Bold = IIf(CBool(ws.Cells(r1 + r - 1, C_START + c - 1).Font.Bold), -1, 0)
                If Not IsNull(ws.Cells(r1 + r - 1, C_START + c - 1).Font.Italic) Then _
                    cellRng.Font.Italic = IIf(CBool(ws.Cells(r1 + r - 1, C_START + c - 1).Font.Italic), -1, 0)
                vU = ws.Cells(r1 + r - 1, C_START + c - 1).Font.Underline
                If Not IsNull(vU) Then cellRng.Font.Underline = IIf((vU = -4142 Or vU = 0), 0, 1)
                On Error Resume Next
                .Shading.BackgroundPatternColor = ws.Cells(r1 + r - 1, C_START + c - 1).Interior.Color
                On Error GoTo 0
                .WordWrap = True: .FitText = False

                ApplyCellTextDirection .Range, MapExcelOrientationToWord(ws.Cells(r1 + r - 1, C_START + c - 1))
            End With
        Next c

        With tbl.Rows(r)
            .HeightRule = wdRowHeightAtLeast
            .Height = WordApp.CentimetersToPoints(0.4)
        End With

        If r Mod 20 = 0 Then DoEvents
    Next r

    ApplyExcelMergesToWord ws, r1, r2, C_START, C_END, tbl

    anchor.SetRange tbl.Range.End, tbl.Range.End
    anchor.InsertParagraphAfter
    anchor.Collapse 0
End Sub

Private Sub ApplyWidthsFromExcelWeights(WordDoc As Object, tbl As Object, weights() As Double)
    Dim total As Double, perc(1 To N_COLS) As Double
    Dim c As Long, okPercent As Boolean

    For c = 1 To N_COLS
        If weights(c) <= 0 Then weights(c) = 1
        total = total + weights(c)
    Next c
    If total <= 0 Then total = N_COLS

    For c = 1 To N_COLS
        perc(c) = (weights(c) / total) * 100#
    Next c

    On Error Resume Next
    tbl.PreferredWidthType = 2
    tbl.PreferredWidth = 100
    okPercent = True
    For c = 1 To N_COLS
        Err.Clear
        tbl.Columns(c).PreferredWidth = perc(c)
        If Err.Number <> 0 Then okPercent = False
    Next c
    If Not okPercent Then
        Dim usable As Single
        usable = GetUsableWidthPts(WordDoc)
        For c = 1 To N_COLS
            tbl.Columns(c).Width = usable * (perc(c) / 100#)
        Next c
    End If
    On Error GoTo 0
End Sub

Private Function GetUsableWidthPts(WordDoc As Object) As Single
    On Error Resume Next
    Dim w As Single
    w = WordDoc.PageSetup.pageWidth - WordDoc.PageSetup.LeftMargin - WordDoc.PageSetup.RightMargin
    If w <= 0 Then w = 500
    GetUsableWidthPts = w
    On Error GoTo 0
End Function


