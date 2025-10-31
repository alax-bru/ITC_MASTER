Attribute VB_Name = "Module11"
Option Explicit

' =====================================================================
'  PP8002_Annexe3b — Export Excel -> Word (Annexe 3)
' ---------------------------------------------------------------------
'  SCOPE / DESIGN
'    - Annexe plus simple que "Annexe 3 Input" : Titres + blocs tableau 4 col (AC..AF).
'    - Fidélité :
'        • Repère "(Annexe 3)" cherché **UNE SEULE FOIS** ; si absent ? arrêt.
'        • Références de cellules **figées** par bloc (anti croisements).
'        • Rich-text au caractère (Characters) + texte affiché (NumberFormatLocal).
'        • Largeurs Word = % des largeurs Excel, sans toucher Excel.
'        • ? Orientation par cellule (horizontal / vertical / up / down / angle ? mapping).
' =====================================================================

' === Constantes Word (late binding) ===
Private Const wdPreferredWidthPercent As Long = 2
Private Const wdAutoFitFixed As Long = 0
Private Const wdLineStyleSingle As Long = 1
Private Const wdLineWidth050pt As Long = 2
Private Const wdLineWidth025pt As Long = 1
Private Const wdRowHeightAtLeast As Long = 1
Private Const wdCollapseEnd As Long = 0

' --- Orientation texte Word (cellule de tableau)
Private Const wdTextOrientationHorizontal As Long = 0
Private Const wdTextOrientationVerticalFarEast As Long = 1
Private Const wdTextOrientationUpward As Long = 2
Private Const wdTextOrientationDownward As Long = 3

' Indices de bordures
Private Const B_TOP As Long = 1, B_LEFT As Long = 2, B_BOTTOM As Long = 3, B_RIGHT As Long = 4
Private Const B_HORZ As Long = 5, B_VERT As Long = 6

' === Bornes dynamiques (calculées à l'exécution) ===
Private LIGNE_DEBUT As Long, LIGNE_FIN As Long, COLONNE_DEBUT As Long, COLONNE_FIN As Long

' === Cache de ligne ===
Private Type LigneInfo
    ValeurAA As String
    ValeurAB As String
    ValeursACtoAF(1 To 4) As String
    FormatsACtoAF(1 To 4) As Variant  ' Array(Bold, Italic, Underline, Color)
    EstVide As Boolean
    EstTitre As Boolean
    EstSousTitre As Boolean
    EstTableau As Boolean
End Type

' =====================================================================
'                           MACRO PRINCIPALE
' =====================================================================
Public Sub PP8002_Annexe3b()
    Dim ws As Worksheet
    Dim WordApp As Object, WordDoc As Object
    Dim anchor As Object
    Dim i As Long, t0 As Single
    Dim lignesData() As LigneInfo, dataRange As Variant
    Dim cellDebut As Range, cellFin As Range
    Dim cheminDoc As String

    t0 = Timer

    ' --- EXCEL perfs ---
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- Feuille source ---
    Set ws = ThisWorkbook.Worksheets("2.5-PP & SOW Annexe 3")

    ' ============================================================
    ' ?? Forçage manuel (optionnel) — 1 ligne à décommenter :
    'LIGNE_DEBUT = 139: LIGNE_FIN = 562: COLONNE_DEBUT = 27: COLONNE_FIN = 32
    ' ============================================================

    ' --- Repères auto si non forcé ---
    If LIGNE_DEBUT = 0 And LIGNE_FIN = 0 And COLONNE_DEBUT = 0 And COLONNE_FIN = 0 Then
        With ws.Cells
            Set cellDebut = .Find("Cellule 6 Lignes Avant Premiere Cellule Range Annexe 3b", LookAt:=xlWhole, LookIn:=xlValues)
            Set cellFin = .Find("Cellule 2 Lignes Après Dernière Cellule Range Annexe 3b", LookAt:=xlWhole, LookIn:=xlValues)
        End With
        If Not cellDebut Is Nothing Then
            LIGNE_DEBUT = cellDebut.row + 6
            COLONNE_DEBUT = cellDebut.Column
        End If
        If Not cellFin Is Nothing Then
            LIGNE_FIN = cellFin.row - 2
            COLONNE_FIN = cellFin.Column
        End If
    End If

    ' --- Fallback bornes historiques ---
    If LIGNE_DEBUT = 0 Or LIGNE_FIN = 0 Or COLONNE_DEBUT = 0 Or COLONNE_FIN = 0 Then
        LIGNE_DEBUT = 139: LIGNE_FIN = 562
        COLONNE_DEBUT = 27: COLONNE_FIN = 32
    End If

    ' --- Sécurité colonnes : exiger AA..AF (6 col mini) ---
    If COLONNE_FIN - COLONNE_DEBUT + 1 < 6 Then
        COLONNE_DEBUT = 27: COLONNE_FIN = 32
    End If

    ' --- Chargement mémoire (AA..AF dynamiques) ---
    ReDim lignesData(LIGNE_DEBUT To LIGNE_FIN)
    dataRange = ws.Range(ws.Cells(LIGNE_DEBUT, COLONNE_DEBUT), ws.Cells(LIGNE_FIN, COLONNE_FIN)).Value

    ' --- Prétraitement (classification + snapshot formats globaux AC..AF) ---
    For i = LIGNE_DEBUT To LIGNE_FIN
        PretraiterLigne ws, i, dataRange, lignesData(i)
    Next i

    ' --- Word : ouverture + repère unique ---
    cheminDoc = ThisWorkbook.Path & Application.PathSeparator & "PP_8002-FR.dotx"
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    If WordApp Is Nothing Then MsgBox "Impossible d'ouvrir Word.", vbCritical: GoTo CleanUp
    WordApp.Visible = True

    Set WordDoc = WordApp.Documents.Open(cheminDoc)

    ' >>> Chercher "(Annexe 3)" **UNE SEULE FOIS** ; si absent ? on stoppe.
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3)": .Forward = True: .Wrap = 1 ' wdFindStop
        .Execute
    End With
    If Not anchor.Find.Found Then
        MsgBox "Repère '(Annexe 3)' introuvable. Opération annulée.", vbCritical
        GoTo CleanUp
    End If
    anchor.Text = ""                ' consomme le repère
    anchor.Collapse wdCollapseEnd   ' positionne le curseur
    ' (Aucune autre recherche du repère après ce point)

    ' --- Parcours principal ---
    i = LIGNE_DEBUT
    Do While i <= LIGNE_FIN
        If lignesData(i).EstVide Then
            i = i + 1
        ElseIf lignesData(i).EstTitre Then
            AjouterTitreOptimise WordDoc, anchor, lignesData(i).ValeurAA, "Titre 2"
            i = i + 1
        ElseIf lignesData(i).EstSousTitre Then
            AjouterTitreOptimise WordDoc, anchor, lignesData(i).ValeurAB, "Titre 3"
            i = i + 1
        ElseIf lignesData(i).EstTableau Then
            i = TraiterTableauOptimise(ws, WordDoc, WordApp, anchor, i, LIGNE_FIN, lignesData)
        Else
            i = i + 1
        End If
        If i Mod 20 = 0 Then DoEvents
    Loop

CleanUp:
    ' --- Excel : remise en état ---
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Export Annexe 3 terminé !" & vbCrLf & _
           "Temps écoulé : " & Format(Timer - t0, "0.00") & " s", vbInformation
End Sub



' =====================================================================
'                             PRÉTRAITEMENT
' =====================================================================
Private Sub PretraiterLigne(ws As Worksheet, ByVal ligne As Long, dataRange As Variant, ByRef info As LigneInfo)
    Dim j As Long, idx As Long
    Dim hasAA As Boolean, hasAB As Boolean, hasACtoAF As Boolean

    idx = ligne - LIGNE_DEBUT + 1
    info.ValeurAA = CStr(dataRange(idx, 1)) ' AA
    info.ValeurAB = CStr(dataRange(idx, 2)) ' AB

    hasACtoAF = False
    For j = 1 To 4                             ' AC..AF
        info.ValeursACtoAF(j) = CStr(dataRange(idx, j + 2))
        If Trim$(info.ValeursACtoAF(j)) <> "" Then hasACtoAF = True
        With ws.Cells(ligne, COLONNE_DEBUT + 1 + j) ' AC..AF relatifs
            info.FormatsACtoAF(j) = Array(.Font.Bold, .Font.Italic, .Font.Underline, .Interior.Color)
        End With
    Next j

    hasAA = (Trim$(info.ValeurAA) <> "")
    hasAB = (Trim$(info.ValeurAB) <> "")

    info.EstVide = (Not hasAA And Not hasAB And Not hasACtoAF)
    info.EstTitre = (hasAA And Not hasAB And Not hasACtoAF)
    info.EstSousTitre = (Not hasAA And hasAB And Not hasACtoAF)
    info.EstTableau = hasACtoAF
End Sub

' =====================================================================
'                        TITRES (1 seul paragraphe)
' =====================================================================
Private Sub AjouterTitreOptimise(WordDoc As Object, ByRef anchor As Object, ByVal texte As String, ByVal styleNom As String)
    If Len(Trim$(texte)) = 0 Then Exit Sub
    Dim s As String, rng As Object
    s = CStr(texte)
    s = Replace(s, vbCrLf, Chr$(11))
    s = Replace(s, vbCr, Chr$(11))
    s = Replace(s, vbLf, Chr$(11))
    Set rng = WordDoc.Range(anchor.Start, anchor.Start)
    rng.Text = s & vbCr
    rng.Style = styleNom
    anchor.SetRange rng.End, rng.End
End Sub

' =====================================================================
'   TABLEAU (AC..AF) — rich text + widths Excel % + réf. figées + ORIENTATION
' =====================================================================
Private Function TraiterTableauOptimise(ws As Worksheet, WordDoc As Object, WordApp As Object, _
                                       ByRef anchor As Object, ByVal ligneDebut As Long, _
                                       ByVal derniereLigne As Long, lignesData() As LigneInfo) As Long
    Dim ligneFin As Long, nbLignes As Long
    Dim wordTable As Object, tblRng As Object
    Dim i As Long, j As Long, ligneIdx As Long
    Dim weights(1 To 4) As Double
    Dim colStart As Long
    Dim lb As Long, ub As Long

    ' --- Sécurités bornes ---
    lb = LBound(lignesData): ub = UBound(lignesData)
    If ligneDebut < lb Or ligneDebut > ub Then TraiterTableauOptimise = ligneDebut + 1: Exit Function
    If derniereLigne > ub Then derniereLigne = ub

    ' --- Délimitation du bloc consécutif ---
    ligneFin = ligneDebut
    Do While ligneFin <= derniereLigne
        If Not lignesData(ligneFin).EstTableau Then Exit Do
        ligneFin = ligneFin + 1
    Loop
    ligneFin = ligneFin - 1
    If ligneFin < ligneDebut Then TraiterTableauOptimise = ligneDebut + 1: Exit Function

    nbLignes = ligneFin - ligneDebut + 1
    colStart = COLONNE_DEBUT + 2 ' AC

    ' --- Poids des colonnes (largeur Excel) ---
    For j = 1 To 4
        weights(j) = ws.Columns(colStart + j - 1).ColumnWidth
        If weights(j) <= 0 Then weights(j) = 1
    Next j

    ' --- Figer les références des cellules du bloc (anti-glissement) ---
    Dim cellRef() As Range
    ReDim cellRef(1 To nbLignes, 1 To 4)
    For i = 1 To nbLignes
        ligneIdx = ligneDebut + i - 1
        For j = 1 To 4
            Set cellRef(i, j) = ws.Cells(ligneIdx, colStart + (j - 1))
        Next j
    Next i

    ' --- Créer le tableau Word + config standard ---
    Set tblRng = WordDoc.Range(anchor.Start, anchor.Start)
    Set wordTable = WordDoc.Tables.Add(Range:=tblRng, numRows:=nbLignes, NumColumns:=4)
    ConfigurerTableau wordTable, WordApp

    ' --- Largeurs % (depuis Excel) ---
    SetColumnWidthsFromWeightsPercent WordDoc, wordTable, weights

    ' --- Remplissage + ORIENTATION PAR CELLULE ---
    For i = 1 To nbLignes
        For j = 1 To 4
            ' 1) Contenu (rich-text ou texte affiché + format global)
            WriteCellWithDisplayAndRich cellRef(i, j), wordTable.cell(i, j).Range, WordDoc
            ' 2) Shading (fond de la cellule Excel)
            On Error Resume Next
            wordTable.cell(i, j).Shading.BackgroundPatternColor = cellRef(i, j).Interior.Color
            On Error GoTo 0
            ' 3) ORIENTATION (uniquement cette cellule)
            ApplyCellTextDirection wordTable.cell(i, j).Range, MapExcelOrientationToWord(cellRef(i, j))
        Next j
        ' Hauteur homogène
        With wordTable.Rows(i)
            .HeightRule = wdRowHeightAtLeast
            .Height = WordApp.CentimetersToPoints(0.4)
        End With
        If i Mod 20 = 0 Then DoEvents
    Next i

    ' --- Avancer l'ancre ---
    anchor.SetRange wordTable.Range.End, wordTable.Range.End
    anchor.InsertParagraphAfter
    anchor.Collapse wdCollapseEnd

    TraiterTableauOptimise = ligneFin + 1
End Function

' === Orientation : Excel ? Word (mapping + seuils pour angles numériques) ===
Private Function MapExcelOrientationToWord(xlCell As Range) As Long
    Dim o As Variant
    o = xlCell.Orientation
    On Error Resume Next
    Select Case o
        Case xlVertical                 ' texte empilé (Excel)
            MapExcelOrientationToWord = wdTextOrientationVerticalFarEast
        Case xlUpward                   ' 90° vers le haut
            MapExcelOrientationToWord = wdTextOrientationUpward
        Case xlDownward                 ' 90° vers le bas
            MapExcelOrientationToWord = wdTextOrientationDownward
        Case xlHorizontal, 0            ' horizontal (par défaut)
            MapExcelOrientationToWord = wdTextOrientationHorizontal
        Case Else
            ' Angle numérique (-90 .. +90). Word ne gère pas l'arbitraire,
            ' on approxime : >= +45° ? Upward ; <= -45° ? Downward ; sinon Horizontal.
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

' =====================================================================
'                  TEXTE AFFICHÉ & RICH-TEXT — cellule -> Word
' =====================================================================
Private Sub WriteCellWithDisplayAndRich(xlCell As Range, wdCellRange As Object, WordDoc As Object)
    Dim sDisp As String, sRaw As String, hasRich As Boolean
    ' (1) Texte affiché (unités, €, %, …)
    sDisp = GetDisplayTextSafe(xlCell)
    sDisp = Replace(Replace(sDisp, vbCrLf, vbLf), vbCr, vbLf)
    ' (2) Rich-text ?
    hasRich = HasMixedFormatting(xlCell)

    ' Nettoyage de la cellule Word (sans supprimer le marqueur de fin)
    wdCellRange.End = wdCellRange.End - 1
    wdCellRange.Text = ""
    On Error Resume Next
    wdCellRange.Style = "Text in table"   ' style paragraphe (n'écrase pas le format inline)
    On Error GoTo 0

    If hasRich Then
        ' Reconstruire les runs depuis la valeur texte (rich)
        sRaw = CStr(xlCell.Value)
        sRaw = Replace(Replace(sRaw, vbCrLf, vbLf), vbCr, vbLf)
        WriteRichRuns WordDoc, wdCellRange, xlCell, sRaw
    Else
        ' Pas de rich-text ? écrire le texte affiché + appliquer le format global
        Dim ins As Object, beforeEnd As Long, seg As Object, textOut As String
        textOut = Replace(sDisp, vbLf, vbCr) ' paragraphes Word
        Set ins = wdCellRange.Duplicate
        ins.Collapse wdCollapseEnd
        beforeEnd = ins.End
        ins.InsertAfter textOut
        Set seg = WordDoc.Range(beforeEnd, beforeEnd + Len(textOut))
        ApplyCellFontFormat seg, xlCell
    End If
End Sub

' === Helpers rich-text / display ===
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

Private Function HasMixedFormatting(xlCell As Range) As Boolean
    Dim s As String, n As Long, p As Long
    Dim b0 As Variant, i0 As Variant, u0 As Variant, c0 As Variant
    Dim b As Variant, i As Variant, u As Variant, c As Variant
    On Error Resume Next
    If VarType(xlCell.Value) <> vbString Then HasMixedFormatting = False: Exit Function
    s = CStr(xlCell.Value): n = Len(s)
    If n <= 1 Then HasMixedFormatting = False: Exit Function
    With xlCell.Characters(1, 1).Font
        b0 = .Bold: i0 = .Italic: u0 = .Underline: c0 = .Color
    End With
    For p = 2 To n
        With xlCell.Characters(p, 1).Font
            b = .Bold: i = .Italic: u = .Underline: c = .Color
        End With
        If (b <> b0) Or (i <> i0) Or (u <> u0) Or (c <> c0) Then HasMixedFormatting = True: Exit Function
    Next p
    HasMixedFormatting = False
End Function

Private Sub WriteRichRuns(WordDoc As Object, wdCellRange As Object, xlCell As Range, sRaw As String)
    Dim n As Long, p As Long, startSeg As Long
    Dim bPrev As Variant, iPrev As Variant, uPrev As Variant, colPrev As Variant
    Dim bCur As Variant, iCur As Variant, uCur As Variant, colCur As Variant
    Dim ins As Object, seg As Object, segText As String, beforeEnd As Long

    n = Len(sRaw)
    Set ins = wdCellRange.Duplicate
    ins.Collapse wdCollapseEnd
    If n = 0 Then Exit Sub

    GetCharProps xlCell, 1, bPrev, iPrev, uPrev, colPrev
    startSeg = 1

    For p = 2 To n + 1
        If p <= n Then
            GetCharProps xlCell, p, bCur, iCur, uCur, colCur
        Else
            bCur = Empty: iCur = Empty: uCur = Empty: colCur = Empty
        End If
        If Not SameProps(bPrev, iPrev, uPrev, colPrev, bCur, iCur, uCur, colCur) Then
            segText = Mid$(sRaw, startSeg, p - startSeg)
            segText = Replace(segText, vbLf, vbCr)
            beforeEnd = ins.End
            ins.InsertAfter segText
            Set seg = WordDoc.Range(beforeEnd, beforeEnd + Len(segText))
            ApplyRunFormat seg, bPrev, iPrev, uPrev, colPrev
            startSeg = p
            bPrev = bCur: iPrev = iCur: uPrev = uCur: colPrev = colCur
            ins.SetRange seg.End, seg.End
        End If
    Next p
End Sub

Private Sub ApplyCellFontFormat(wdRange As Object, xlCell As Range)
    On Error Resume Next
    With xlCell.Font
        If Not IsNull(.Bold) Then wdRange.Font.Bold = IIf(CBool(.Bold), -1, 0)
        If Not IsNull(.Italic) Then wdRange.Font.Italic = IIf(CBool(.Italic), -1, 0)
        If Not IsNull(.Underline) Then wdRange.Font.Underline = IIf((.Underline = -4142 Or .Underline = 0), 0, 1)
        If Not IsNull(.Color) Then wdRange.Font.Color = .Color
    End With
    On Error GoTo 0
End Sub

Private Sub GetCharProps(xlCell As Range, ByVal pos As Long, _
                         ByRef b As Variant, ByRef i As Variant, ByRef u As Variant, ByRef col As Variant)
    With xlCell.Characters(pos, 1).Font
        b = .Bold: i = .Italic: u = .Underline: col = .Color
    End With
End Sub

Private Function SameProps(b1 As Variant, i1 As Variant, u1 As Variant, c1 As Variant, _
                           b2 As Variant, i2 As Variant, u2 As Variant, c2 As Variant) As Boolean
    SameProps = (b1 = b2 And i1 = i2 And u1 = u2 And c1 = c2)
End Function

Private Sub ApplyRunFormat(wdRange As Object, b As Variant, i As Variant, u As Variant, col As Variant)
    On Error Resume Next
    If Not IsNull(b) Then wdRange.Font.Bold = IIf(CBool(b), -1, 0)
    If Not IsNull(i) Then wdRange.Font.Italic = IIf(CBool(i), -1, 0)
    If Not IsNull(u) Then wdRange.Font.Underline = IIf((u = -4142 Or u = 0), 0, 1)
    If Not IsNull(col) Then wdRange.Font.Color = col
    On Error GoTo 0
End Sub

' =====================================================================
'                      CONFIGURATION TABLEAU (Word)
' =====================================================================
Private Sub ConfigurerTableau(wordTable As Object, WordApp As Object)
    With wordTable
        .AutoFitBehavior wdAutoFitFixed
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        ' Défauts (si % impossible)
        .Columns(1).PreferredWidthType = wdPreferredWidthPercent: .Columns(1).PreferredWidth = 55
        .Columns(2).PreferredWidthType = wdPreferredWidthPercent: .Columns(2).PreferredWidth = 22
        .Columns(3).PreferredWidthType = wdPreferredWidthPercent: .Columns(3).PreferredWidth = 13
        .Columns(4).PreferredWidthType = wdPreferredWidthPercent: .Columns(4).PreferredWidth = 8
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
        .AllowAutoFit = False
        .Rows.AllowBreakAcrossPages = True
        On Error Resume Next
        .Style = "Grille du tableau": If Err.Number <> 0 Then Err.Clear: .Style = "Table Grid"
        On Error GoTo 0
        Dim idx As Variant
        For Each idx In Array(B_TOP, B_LEFT, B_BOTTOM, B_RIGHT, B_HORZ, B_VERT)
            With .Borders(idx)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
            End With
        Next idx
    End With
End Sub

' =====================================================================
'                  LARGEURS % (depuis POIDS capturés Excel)
' =====================================================================
Private Sub SetColumnWidthsFromWeightsPercent(WordDoc As Object, wordTable As Object, weights As Variant)
    Dim total As Double, perc() As Double
    Dim c As Long, nCols As Long, lb As Long, ub As Long
    Dim okPercent As Boolean

    lb = LBound(weights): ub = UBound(weights)
    nCols = ub - lb + 1
    ReDim perc(1 To nCols)

    For c = 1 To nCols
        If weights(lb + c - 1) <= 0 Then weights(lb + c - 1) = 1
        total = total + weights(lb + c - 1)
    Next c
    If total <= 0 Then total = nCols

    For c = 1 To nCols
        perc(c) = (weights(lb + c - 1) / total) * 100#
    Next c

    On Error Resume Next
    wordTable.PreferredWidthType = wdPreferredWidthPercent
    wordTable.PreferredWidth = 100
    okPercent = True
    For c = 1 To nCols
        Err.Clear
        wordTable.Columns(c).PreferredWidthType = wdPreferredWidthPercent
        wordTable.Columns(c).PreferredWidth = perc(c)
        If Err.Number <> 0 Then okPercent = False
    Next c
    If Not okPercent Then
        Dim usable As Single: usable = GetUsableWidthPts(WordDoc)
        For c = 1 To nCols
            wordTable.Columns(c).Width = usable * (perc(c) / 100#)
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


'================================================================================================
'                                             3c NIGGA
'================================================================================================




