Attribute VB_Name = "Module14"
Option Explicit

' =====================================================================
'  PP8002_Annexe1_NonContigue — Export Excel -> Word (Annexe 1)
'  Feuille source : "01.3-ITC MASTER WBS"
'  Lignes (dynamiques) :
'    - AUTO   : de la ligne contenant "ICI LE DEBUT DE L'ANNEXE 1"
'               à la ligne contenant "ICI LA FIN DE L'ANNEXE 1" (inclus)
'    - FORCÉ  : décommentez la ligne prévue ci-dessous
'    - DEFAUT : 166..664 si repères absents/incomplets
'  Colonnes (fixes) : D, L, M, N, O, P, Q, R, AD, AE, AF  (11 colonnes)
'  Fidélité Word : ancre "(Annexe 1)" cherchée une seule fois, largeurs distribuées
'  selon Excel (en % à partir de ColumnWidth), fusions projetées, orientation cellule,
'  style "Text in table" + **police 6 pt**.
' =====================================================================

' === Constantes Word (late binding) ===
Private Const wdAutoFitFixed As Long = 0
Private Const wdLineStyleSingle As Long = 1
Private Const wdLineWidth050pt As Long = 2
Private Const wdLineWidth025pt As Long = 1
Private Const wdRowHeightAtLeast As Long = 1

' Orientation texte Word (cellule de tableau)
Private Const wdTextOrientationHorizontal As Long = 0
Private Const wdTextOrientationVerticalFarEast As Long = 1
Private Const wdTextOrientationUpward As Long = 2
Private Const wdTextOrientationDownward As Long = 3

' Indices de bordures
Private Const B_TOP As Long = 1, B_LEFT As Long = 2, B_BOTTOM As Long = 3, B_RIGHT As Long = 4
Private Const B_HORZ As Long = 5, B_VERT As Long = 6

' === Plage source (dynamiques) ===
Private R_START As Long
Private R_END   As Long

' --- Liste des colonnes non contiguës (numéros Excel) ---
Private Function SRC_COLS() As Variant
    ' D, L, M, N, O, P, Q, R, AD, AE, AF
    SRC_COLS = Array(4, 12, 13, 14, 15, 16, 17, 18, 30, 31, 32)
End Function

' === Helper: nombre de colonnes sélectionnées ===
Private Function COL_COUNT() As Long
    COL_COUNT = UBound(SRC_COLS) - LBound(SRC_COLS) + 1
End Function

' =====================================================================
'                             MACRO
' =====================================================================
Public Sub PP8002_Annexe1()
    Dim ws As Worksheet
    Dim WordApp As Object, WordDoc As Object
    Dim cheminDoc As String
    Dim anchor As Object
    Dim t0 As Single
    Dim i As Long, blockStart As Long, blockEnd As Long
    Dim cols As Variant
    Dim cDeb As Range, cFin As Range
    Dim N_COLS As Long

    t0 = Timer
    Set ws = ThisWorkbook.Worksheets("01.3-ITC MASTER WBS")
    cols = SRC_COLS
    N_COLS = COL_COUNT()

    ' ============================================================
    ' ???? RANGE FORCÉ (optionnel) ????
    ' ? Pour ignorer la détection auto, décommentez la ligne suivante :
    'R_START = 166: R_END = 664
    ' ============================================================

    ' === AUTO : chercher "ICI LE DEBUT DE L'ANNEXE 1" / "ICI LA FIN DE L'ANNEXE 1"
    If R_START = 0 And R_END = 0 Then
        With ws.Cells
            Set cDeb = .Find(What:="ICI LE DEBUT DE L'ANNEXE 1", LookAt:=xlWhole, LookIn:=xlValues)
            Set cFin = .Find(What:="ICI LA FIN DE L'ANNEXE 1", LookAt:=xlWhole, LookIn:=xlValues)
        End With
        If Not cDeb Is Nothing Then R_START = cDeb.row
        If Not cFin Is Nothing Then R_END = cFin.row
    End If

    ' === DEFAUT : si repères absents/incomplets, retomber sur l’ancien range
    If R_START = 0 Or R_END = 0 Then
        R_START = 166
        R_END = 664
    End If
    If R_START > R_END Then
        Dim tmp As Long: tmp = R_START: R_START = R_END: R_END = tmp
    End If

    ' --- Word ---
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

    ' === Repère "(Annexe 1)" — UNE SEULE FOIS (sinon abort)
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 1)": .Forward = True: .Wrap = 1 ' wdFindStop
        .Execute
    End With
    If Not anchor.Find.Found Then
        MsgBox "Repère '(Annexe 1)' introuvable dans " & cheminDoc, vbExclamation
        Exit Sub
    End If
    anchor.Text = "": anchor.Collapse 0    ' 0 = wdCollapseEnd
    ' >>> aucune nouvelle recherche d’ancre après ce point

    ' === Parcours des blocs de lignes non vides (sur nos colonnes fixes)
    i = R_START
    Do While i <= R_END
        Do While i <= R_END And IsRowEmptyOnCols(ws, i, cols): i = i + 1: Loop
        If i > R_END Then Exit Do
        blockStart = i

        Do While i <= R_END And Not IsRowEmptyOnCols(ws, i, cols): i = i + 1: Loop
        blockEnd = i - 1

        ' Table Word pour le bloc
        CreateAndFillTableFromBlock_NonContig ws, WordDoc, WordApp, anchor, blockStart, blockEnd, cols
        DoEvents
    Loop

    MsgBox "Annexe 1 exportée (" & R_START & ":" & R_END & ") ? " & cheminDoc & vbCrLf & _
           "Temps : " & Format(Timer - t0, "0.00") & " s", vbInformation
End Sub

' =====================================================================
'                             HELPERS
' =====================================================================

Private Function IsRowEmptyOnCols(ws As Worksheet, ByVal rowNum As Long, ByVal colList As Variant) As Boolean
    Dim k As Long, c As Long
    For k = LBound(colList) To UBound(colList)
        c = colList(k)
        If LenB(Trim$(CStr(ws.Cells(rowNum, c).Value))) <> 0 Then IsRowEmptyOnCols = False: Exit Function
    Next k
    IsRowEmptyOnCols = True
End Function

' Texte "affiché" (NumberFormatLocal) même si colonne étroite (évite ####)
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

' Map Excel -> Word pour l’orientation par cellule
Private Function MapExcelOrientationToWord(xlCell As Range) As Long
    Dim o As Variant: o = xlCell.Orientation
    On Error Resume Next
    Select Case o
        Case xlVertical
            MapExcelOrientationToWord = wdTextOrientationVerticalFarEast
        Case xlUpward
            MapExcelOrientationToWord = wdTextOrientationUpward
        Case xlDownward
            MapExcelOrientationToWord = wdTextOrientationDownward
        Case xlHorizontal, 0
            MapExcelOrientationToWord = wdTextOrientationHorizontal
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

' --- Fusions Excel sur colonnes NON contiguës (projection) ---
Private Sub ApplyExcelMergesToWord_NonContig(ws As Worksheet, _
        ByVal r1 As Long, ByVal r2 As Long, ByVal colList As Variant, ByVal tbl As Object)

    Dim merges() As Variant, m As Long
    Dim r As Long, k As Long, cXL As Long
    Dim xlCell As Range, area As Range
    Dim topRow As Long, leftCol As Long, bottomRow As Long, rightCol As Long
    Dim iTop As Long, iBottom As Long
    Dim leftIdx As Long, rightIdx As Long, haveCols As Boolean
    Dim wr As Long, wr2 As Long, wc As Long, wc2 As Long
    Dim i As Long, j As Long, a As Variant, b As Variant, tmp As Variant

    ' 1) Collecte des rectangles projetés
    m = 0
    For r = r1 To r2
        For k = LBound(colList) To UBound(colList)
            cXL = colList(k)
            Set xlCell = ws.Cells(r, cXL)
            If xlCell.MergeCells Then
                Set area = xlCell.MergeArea

                topRow = area.row
                leftCol = area.Column
                bottomRow = topRow + area.Rows.Count - 1
                rightCol = leftCol + area.Columns.Count - 1

                ' Intersection lignes
                iTop = Application.Max(topRow, r1)
                iBottom = Application.Min(bottomRow, r2)
                If iTop > iBottom Then GoTo NextK

                ' Intersection colonnes (projection sur nos colonnes)
                leftIdx = 0: rightIdx = 0: haveCols = False
                For j = LBound(colList) To UBound(colList)
                    If colList(j) >= leftCol And colList(j) <= rightCol Then
                        If leftIdx = 0 Then leftIdx = (j - LBound(colList) + 1)
                        rightIdx = (j - LBound(colList) + 1)
                        haveCols = True
                    End If
                Next j
                If Not haveCols Then GoTo NextK

                ' N'agir qu'au coin haut-gauche projeté
                If r = iTop And cXL = colList(LBound(colList) + leftIdx - 1) Then
                    wr = iTop - r1 + 1
                    wr2 = iBottom - r1 + 1
                    wc = leftIdx
                    wc2 = rightIdx
                    If (wr2 > wr) Or (wc2 > wc) Then
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
NextK:
        Next k
    Next r
    If m = 0 Then Exit Sub

    ' 2) Trier (bottom-right -> top-left) pour éviter le décalage d’index
    For i = 1 To m - 1
        For j = i + 1 To m
            a = merges(i): b = merges(j)
            If (b(1) > a(1)) Or ((b(1) = a(1)) And (b(0) > a(0))) Then
                tmp = merges(i): merges(i) = merges(j): merges(j) = tmp
            End If
        Next j
    Next i

    ' 3) Exécuter les merges
    For i = 1 To m
        a = merges(i)
        On Error Resume Next
        tbl.cell(a(0), a(1)).Merge tbl.cell(a(2), a(3))
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' --- Création + remplissage (colonnes non contiguës) ---
Private Sub CreateAndFillTableFromBlock_NonContig(ws As Worksheet, WordDoc As Object, WordApp As Object, _
                                        ByRef anchor As Object, ByVal r1 As Long, ByVal r2 As Long, _
                                        ByVal colList As Variant)

    Dim N_COLS As Long: N_COLS = UBound(colList) - LBound(colList) + 1
    Dim numRows As Long: numRows = r2 - r1 + 1
    Dim tbl As Object, tRng As Object
    Dim r As Long, k As Long, cXL As Long
    Dim cellRng As Object
    Dim vU As Variant

    Dim useModel As Boolean
    Dim modelTbl As Object
    Dim savedW() As Single
    Dim i As Long
    Dim modelStyle As Variant

    Dim weights() As Double
    Dim doApplyExcelWidths As Boolean

    ' 0) Sommes-nous dans un tableau modèle à l’ancre ?
    useModel = False
    On Error Resume Next
    If anchor.Cells.Count > 0 Then
        Set modelTbl = anchor.Cells(1).Table
        If Not modelTbl Is Nothing Then useModel = True
    End If
    On Error GoTo 0

    ' 1) Préparer insertion + mémoriser largeurs si modèle
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
        tRng.Delete                 ' supprime le tableau modèle (laisse un ¶)
        doApplyExcelWidths = False  ' on réapplique savedW
    Else
        Set tRng = WordDoc.Range(anchor.Start, anchor.Start)
        doApplyExcelWidths = True
        ' Poids depuis Excel pour nos colonnes non contiguës
        ReDim weights(1 To N_COLS)
        For k = LBound(colList) To UBound(colList)
            weights(k - LBound(colList) + 1) = ws.Columns(colList(k)).ColumnWidth
            If weights(k - LBound(colList) + 1) <= 0 Then weights(k - LBound(colList) + 1) = 1
        Next k
    End If

    ' 2) Créer le tableau Word
    Set tbl = WordDoc.Tables.Add(Range:=tRng, numRows:=numRows, NumColumns:=N_COLS)

    ' 3) Config de base (style, bordures, padding, etc.)
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

    ' 4) Si pas de modèle : appliquer la distribution % depuis Excel (colonne par colonne)
    If doApplyExcelWidths Then
        ApplyWidthsFromWeights WordDoc, tbl, weights
    End If

    ' 5) Remplir + orientation par cellule
    For r = 1 To numRows
        For k = LBound(colList) To UBound(colList)
            cXL = colList(k)
            With tbl.cell(r, k - LBound(colList) + 1)
                Set cellRng = .Range: cellRng.End = cellRng.End - 1
                cellRng.Text = GetDisplayTextSafe(ws.Cells(r1 + r - 1, cXL))
                On Error Resume Next
                cellRng.Style = "Text in table"
                On Error GoTo 0
                ' >>> Présentation : force la taille de police à 6 pt pour chaque cellule
                cellRng.Font.Size = 6

                ' Format global depuis Excel (macro "simple")
                If Not IsNull(ws.Cells(r1 + r - 1, cXL).Font.Bold) Then _
                    cellRng.Font.Bold = IIf(CBool(ws.Cells(r1 + r - 1, cXL).Font.Bold), -1, 0)
                If Not IsNull(ws.Cells(r1 + r - 1, cXL).Font.Italic) Then _
                    cellRng.Font.Italic = IIf(CBool(ws.Cells(r1 + r - 1, cXL).Font.Italic), -1, 0)
                vU = ws.Cells(r1 + r - 1, cXL).Font.Underline
                If Not IsNull(vU) Then cellRng.Font.Underline = IIf((vU = -4142 Or vU = 0), 0, 1)

                On Error Resume Next
                .Shading.BackgroundPatternColor = ws.Cells(r1 + r - 1, cXL).Interior.Color
                On Error GoTo 0

                .WordWrap = True: .FitText = False

                ' Orientation UNIQUEMENT pour cette cellule
                ApplyCellTextDirection .Range, MapExcelOrientationToWord(ws.Cells(r1 + r - 1, cXL))
            End With
        Next k

        With tbl.Rows(r)
            .HeightRule = wdRowHeightAtLeast
            .Height = WordApp.CentimetersToPoints(0.4)
        End With

        If r Mod 20 = 0 Then DoEvents
    Next r

    ' 6) Fusions Excel (projection non contiguë)
    ApplyExcelMergesToWord_NonContig ws, r1, r2, colList, tbl

    ' 7) Avancer l’ancre (¶ après le tableau)
    anchor.SetRange tbl.Range.End, tbl.Range.End
    anchor.InsertParagraphAfter
    anchor.Collapse 0
End Sub

' --- Largeurs en % depuis Excel (fallback en points si % impossible) ---
Private Sub ApplyWidthsFromWeights(WordDoc As Object, tbl As Object, ByRef weights() As Double)
    Dim total As Double, n As Long, perc() As Double
    Dim c As Long, okPercent As Boolean

    n = UBound(weights) - LBound(weights) + 1
    ReDim perc(1 To n)

    For c = 1 To n
        If weights(c) <= 0 Then weights(c) = 1
        total = total + weights(c)
    Next c
    If total <= 0 Then total = n

    For c = 1 To n
        perc(c) = (weights(c) / total) * 100#
    Next c

    On Error Resume Next
    tbl.PreferredWidthType = 2 ' wdPreferredWidthPercent
    tbl.PreferredWidth = 100

    okPercent = True
    For c = 1 To n
        Err.Clear
        tbl.Columns(c).PreferredWidthType = 2 ' wdPreferredWidthPercent
        tbl.Columns(c).PreferredWidth = perc(c)
        If Err.Number <> 0 Then okPercent = False
    Next c

    If Not okPercent Then
        Dim usable As Single: usable = GetUsableWidthPts(WordDoc)
        For c = 1 To n
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








