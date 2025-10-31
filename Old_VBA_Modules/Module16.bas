Attribute VB_Name = "Module16"
'============================================================================
'                                 Annexe 3b - VERSION CORRIGÉE MULTI-LANGUE (+ décalage titres)
'============================================================================

' === Constantes Word (late binding) ===
Private Const wdCollapseEnd As Long = 0
Private Const wdPreferredWidthPercent As Long = 2
Private Const wdRowHeightAuto As Long = 1

' === Bornes dynamiques (calculées à l'exécution) ===
Private LIGNE_DEBUT As Long, LIGNE_FIN As Long
Private COLONNE_DEBUT As Long, COLONNE_FIN As Long

' === Structure de ligne ===
Private Type LigneInfo
    ValeurAA As String
    ValeurAB As String
    EstVide As Boolean
    EstTitre As Boolean
    EstSousTitre As Boolean
    EstTableau As Boolean
End Type

Public Sub PP_SOW_8002_FR_Annexe_3b()
    Dim ws As Worksheet
    Dim WordApp As Object, WordDoc As Object
    Dim cheminDoc As String, anchor As Object
    Dim i As Long, t0 As Single
    Dim lignesData() As LigneInfo, dataRange As Variant
    Dim cellDebut As Range, cellFin As Range

    On Error GoTo ErreurGlobale
    t0 = Timer
    Set ws = ThisWorkbook.Worksheets("2.5-PP & SOW Annexe 3")

    ' --- Validation worksheet ---
    If ws Is Nothing Then
        MsgBox "Feuille '2.5-PP & SOW Annexe 3' introuvable", vbCritical
        Exit Sub
    End If

    ' --- Nettoyage initial ---
    Application.CutCopyMode = False
    DoEvents

    ' --- Détection bornes auto avec NOUVELLES RÉFÉRENCES ---
    If LIGNE_DEBUT = 0 And LIGNE_FIN = 0 Then
        With ws.Cells
            Set cellDebut = .Find("Cellule 6 Lignes Avant Premiere Cellule Range Annexe 3b", LookAt:=xlWhole)
            Set cellFin = .Find("Cellule 2 Lignes Après Derniere Cellule Range Annexe 3b", LookAt:=xlWhole)
        End With
        
        If cellDebut Is Nothing Then
            MsgBox "REPÈRE DÉBUT NON TROUVÉ: 'Cellule 6 Lignes Avant Premiere Cellule Range Annexe 3b'" & vbCrLf & _
                   "Vérifiez l'orthographe exacte dans la feuille Excel", vbCritical
            Exit Sub
        Else
            Debug.Print "? Repère DÉBUT trouvé: " & cellDebut.Address
        End If
        
        If cellFin Is Nothing Then
            MsgBox "REPÈRE FIN NON TROUVÉ: 'Cellule 2 Lignes Après Derniere Cellule Range Annexe 3b'" & vbCrLf & _
                   "Vérifiez l'orthographe exacte dans la feuille Excel", vbCritical
            Exit Sub
        Else
            Debug.Print "? Repère FIN trouvé: " & cellFin.Address
        End If
        
        LIGNE_DEBUT = cellDebut.row + 6
        COLONNE_DEBUT = cellDebut.Column
        LIGNE_FIN = cellFin.row - 2
        COLONNE_FIN = cellFin.Column
        
        If COLONNE_DEBUT = COLONNE_FIN Then
            COLONNE_FIN = COLONNE_DEBUT + 4
            Debug.Print "CORRECTION: Repères dans même colonne -> Extension à " & COLONNE_FIN
        End If
    End If

    If LIGNE_DEBUT = 0 Or LIGNE_FIN = 0 Or COLONNE_DEBUT = 0 Or COLONNE_FIN = 0 Then
        MsgBox "Impossible de détecter les bornes automatiquement.", vbCritical
        Exit Sub
    End If
    
    If LIGNE_DEBUT >= LIGNE_FIN Or COLONNE_DEBUT >= COLONNE_FIN Then
        MsgBox "Bornes détectées invalides :" & vbCrLf & _
               "Ligne début: " & LIGNE_DEBUT & " | Ligne fin: " & LIGNE_FIN & vbCrLf & _
               "Colonne début: " & COLONNE_DEBUT & " | Colonne fin: " & COLONNE_FIN, vbCritical
        Exit Sub
    End If

    On Error Resume Next
    ReDim lignesData(LIGNE_DEBUT To LIGNE_FIN)
    dataRange = ws.Range(ws.Cells(LIGNE_DEBUT, COLONNE_DEBUT), ws.Cells(LIGNE_FIN, COLONNE_FIN)).Value
    
    If Err.Number <> 0 Then
        MsgBox "Erreur lors du chargement des données: " & Err.Description, vbCritical
        On Error GoTo ErreurGlobale
        Exit Sub
    End If
    On Error GoTo ErreurGlobale

    For i = LIGNE_DEBUT To LIGNE_FIN
        PretraiterLigne i, dataRange, lignesData(i)
    Next i

    cheminDoc = ThisWorkbook.Path & Application.PathSeparator & "PP_8002-FR.dotx"
    If Dir(cheminDoc) = "" Then
        MsgBox "Fichier template introuvable: " & cheminDoc, vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then Set WordApp = CreateObject("Word.Application")
    On Error GoTo ErreurGlobale

    If WordApp Is Nothing Then
        MsgBox "Impossible d'ouvrir Word.", vbCritical
        Exit Sub
    End If

    WordApp.Visible = True
    Set WordDoc = WordApp.Documents.Open(cheminDoc)

    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3b)"
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    
    If Not anchor.Find.Found Then
        MsgBox "Repère '(Annexe 3b)' introuvable dans le document Word.", vbCritical
        Exit Sub
    End If
    
    anchor.Text = ""
    anchor.Collapse wdCollapseEnd

    ' --- Parcours ligne par ligne avec compteurs ---
    Dim nbTitres As Long, nbSousTitres As Long, nbTableaux As Long, nbEchecs As Long
    Dim styleTitre2 As String, styleTitre3 As String

    ' Détection langue interface Excel
    Dim lngExcel As Long
    lngExcel = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If lngExcel = 1036 Then
        styleTitre2 = "Titre 3"   ' décalage
        styleTitre3 = "Titre 4"   ' décalage
    ElseIf lngExcel = 1033 Then
        styleTitre2 = "Heading 3" ' décalage
        styleTitre3 = "Heading 4" ' décalage
    Else
        styleTitre2 = "Heading 3"
        styleTitre3 = "Heading 4"
    End If

    i = LIGNE_DEBUT
    
    Do While i <= LIGNE_FIN
        DoEvents
        If i > UBound(lignesData) Then Exit Do
        
        If lignesData(i).EstVide Then
            i = i + 1
        ElseIf lignesData(i).EstTitre Then
            AjouterTitre WordDoc, anchor, lignesData(i).ValeurAA, styleTitre2
            nbTitres = nbTitres + 1
            i = i + 1
        ElseIf lignesData(i).EstSousTitre Then
            AjouterTitre WordDoc, anchor, lignesData(i).ValeurAB, styleTitre3
            nbSousTitres = nbSousTitres + 1
            i = i + 1
        ElseIf lignesData(i).EstTableau Then
            Dim resultat As Long
            resultat = ExporterBlocTableau(ws, WordDoc, WordApp, anchor, i, lignesData)
            If resultat = -1 Then
                nbEchecs = nbEchecs + 1
                Do While i <= LIGNE_FIN And i <= UBound(lignesData)
                    If Not lignesData(i).EstTableau Then Exit Do
                    i = i + 1
                Loop
            Else
                nbTableaux = nbTableaux + 1
                i = resultat
            End If
        Else
            i = i + 1
        End If
    Loop

    Application.CutCopyMode = False

    MsgBox "Annexe 3b exportée avec succès !" & vbCrLf & _
           "Titres: " & nbTitres & " | Sous-titres: " & nbSousTitres & vbCrLf & _
           "Tableaux: " & nbTableaux & " | Échecs: " & nbEchecs & vbCrLf & _
           "Temps : " & Format(Timer - t0, "0.00") & " s", vbInformation

    Exit Sub

ErreurGlobale:
    Application.CutCopyMode = False
    MsgBox "Erreur globale: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Ligne: " & Erl, vbCritical
End Sub

' =====================================================================
'  Prétraitement : classer la ligne
' =====================================================================
Private Sub PretraiterLigne(ByVal ligne As Long, dataRange As Variant, ByRef info As LigneInfo)
    On Error GoTo ErreurPretraitement
    
    Dim idx As Long, hasAA As Boolean, hasAB As Boolean, hasData As Boolean, j As Long
    idx = ligne - LIGNE_DEBUT + 1
    
    If idx < 1 Or idx > UBound(dataRange, 1) Then
        info.EstVide = True
        Exit Sub
    End If

    info.ValeurAA = CStr(dataRange(idx, 1))
    info.ValeurAB = CStr(dataRange(idx, 2))
    hasAA = (Trim$(info.ValeurAA) <> "")
    hasAB = (Trim$(info.ValeurAB) <> "")

    hasData = False
    If UBound(dataRange, 2) > 2 Then
        For j = 3 To UBound(dataRange, 2)
            If Trim$(CStr(dataRange(idx, j))) <> "" Then
                hasData = True
                Exit For
            End If
        Next j
    End If

    info.EstVide = (Not hasAA And Not hasAB And Not hasData)
    info.EstTitre = (hasAA And Not hasAB And Not hasData)
    info.EstSousTitre = (Not hasAA And hasAB And Not hasData)
    info.EstTableau = hasData
    
    Exit Sub
    
ErreurPretraitement:
    info.EstVide = True
End Sub

' =====================================================================
'  Ajout d'un titre/sous-titre
' =====================================================================
Private Sub AjouterTitre(WordDoc As Object, ByRef anchor As Object, ByVal texte As String, ByVal styleNom As String)
    On Error GoTo ErreurTitre
    
    If Len(Trim$(texte)) = 0 Then Exit Sub
    
    Dim rng As Object
    Set rng = WordDoc.Range(anchor.Start, anchor.Start)
    rng.Text = texte & vbCr
    rng.Style = styleNom
    anchor.SetRange rng.End, rng.End
    
    Exit Sub
    
ErreurTitre:
    Debug.Print "Erreur ajout titre '" & texte & "': " & Err.Description
End Sub

' =====================================================================
'  Exporter un bloc tableau
' =====================================================================
Private Function ExporterBlocTableau(ws As Worksheet, WordDoc As Object, WordApp As Object, _
                                     ByRef anchor As Object, ByVal ligneDebut As Long, _
                                     lignesData() As LigneInfo) As Long
    On Error GoTo ErreurTableau
    
    Dim ligneFin As Long, wordTable As Object
    
    ligneFin = ligneDebut
    Do While ligneFin <= LIGNE_FIN And ligneFin <= UBound(lignesData)
        If Not lignesData(ligneFin).EstTableau Then Exit Do
        ligneFin = ligneFin + 1
    Loop
    ligneFin = ligneFin - 1
    
    If ligneFin > LIGNE_FIN Then ligneFin = LIGNE_FIN
    If ligneFin < ligneDebut Then ligneFin = ligneDebut
    
    Dim nbLignes As Long
    nbLignes = ligneFin - ligneDebut + 1

    Application.CutCopyMode = False
    Application.Wait Now + TimeValue("00:00:01")
    
    Dim colonneCopieDebut As Long
    colonneCopieDebut = COLONNE_DEBUT + 2
    
    If colonneCopieDebut > COLONNE_FIN Then
        ExporterBlocTableau = -1
        Exit Function
    End If
    
    ws.Range(ws.Cells(ligneDebut, colonneCopieDebut), ws.Cells(ligneFin, COLONNE_FIN)).Copy
    
    WordApp.Selection.SetRange anchor.Start, anchor.Start
    DoEvents
    
    WordApp.Selection.PasteExcelTable False, False, False
    
    Set wordTable = WordDoc.Tables(WordDoc.Tables.Count)
    If Not wordTable Is Nothing Then
        With wordTable
            .AllowAutoFit = False
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
            .Range.ParagraphFormat.SpaceAfter = 0
            .Rows.HeightRule = wdRowHeightAuto
            .Rows.Height = 0
            .Range.Font.Size = 8
        End With
        
        anchor.SetRange wordTable.Range.End, wordTable.Range.End
        anchor.InsertParagraphAfter
        anchor.Collapse wdCollapseEnd
    End If

    Application.CutCopyMode = False
    ExporterBlocTableau = ligneFin + 1
    Exit Function

ErreurTableau:
    Application.CutCopyMode = False
    ExporterBlocTableau = -1
End Function


