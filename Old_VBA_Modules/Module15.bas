Attribute VB_Name = "Module15"
Option Explicit

' === Constantes ===
Private Const WORD_TEMPLATE As String = "PP_8002-FR.dotx"
Private Const SHEET_NAME As String = "2.4-PP & SOW Annexe 2"
Private Const START_ROW As Long = 11
Private Const END_ROW As Long = 672
Private Const MARKER_TEXT As String = "(Annexe 2)"

Private Const COL_TITRE2 As Long = 6   ' F
Private Const COL_TITRE3 As Long = 7   ' G
Private Const COL_TITRE4 As Long = 8   ' H
Private Const COL_TEXTE As Long = 15   ' O
Private Const COL_LANGUE As Long = 17  ' Q
Private Const COL_FLAG As Long = 24    ' X

' Objets globaux
Private WordApp As Object
Private WordDoc As Object
Private ExcelWS As Worksheet
Private InsertionRange As Object

' Mémorisation pour éviter doublons
Private PrevTitre2 As String
Private PrevTitre3 As String
Private PrevTitre4 As String

' Compteurs
Private TotalRows As Long
Private InsertedRows As Long

' Détection de langue Excel (FR/EN)
Private lngExcel As Long

'===============================================================================
' MACRO PRINCIPALE
'===============================================================================
Public Sub PP_SOW_8002_FR_Annexe_2()
    On Error GoTo GestionErreur
    
    Dim t0 As Single: t0 = Timer
    
    ' Vérifier Excel
    Set ExcelWS = Nothing
    On Error Resume Next
    Set ExcelWS = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ExcelWS Is Nothing Then
        MsgBox "Feuille '" & SHEET_NAME & "' introuvable.", vbCritical
        Exit Sub
    End If
    
    ' Vérifier Word
    If Not OuvrirWord Then Exit Sub
    
    ' Trouver le marqueur
    If Not TrouverEtPreparerMarqueur Then Exit Sub
    
    ' Initialiser langue Excel
    lngExcel = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
    ' Lancer traitement
    TotalRows = 0: InsertedRows = 0
    PrevTitre2 = "": PrevTitre3 = "": PrevTitre4 = ""
    
    TraiterDonnees
    
    ' Résumé
    If InsertedRows > 0 Then
        MsgBox "? Annexe 2 exportée avec succès !" & vbCrLf & _
               "Feuille : " & SHEET_NAME & vbCrLf & _
               "Lignes parcourues : " & TotalRows & vbCrLf & _
               "Éléments insérés : " & InsertedRows & vbCrLf & _
               "Durée : " & Format(Timer - t0, "0.00") & " s" & vbCrLf & vbCrLf & _
               "?? Le document Word est ouvert mais NON enregistré. " & _
               "Veuillez sauvegarder manuellement.", vbInformation
    Else
        MsgBox "? Aucune donnée insérée." & vbCrLf & _
               "Vérifiez :" & vbCrLf & _
               "- Que la colonne Q = 'FR'" & vbCrLf & _
               "- Que la colonne X = 'Utilisé'" & vbCrLf & _
               "- Que le marqueur " & MARKER_TEXT & " est présent dans Word", vbExclamation
    End If
    
    Exit Sub
    
GestionErreur:
    MsgBox "Erreur Annexe 2 : " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'===============================================================================
' OUVRIR WORD
'===============================================================================
Private Function OuvrirWord() As Boolean
    Dim cheminDoc As String
    cheminDoc = ThisWorkbook.Path & Application.PathSeparator & WORD_TEMPLATE
    
    If Dir(cheminDoc) = "" Then
        MsgBox "Modèle Word introuvable : " & cheminDoc, vbCritical
        Exit Function
    End If
    
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If WordApp Is Nothing Then
        MsgBox "Impossible de démarrer Word.", vbCritical
        Exit Function
    End If
    
    WordApp.Visible = True
    Set WordDoc = WordApp.Documents.Open(cheminDoc, ReadOnly:=False)
    OuvrirWord = True
End Function

'===============================================================================
' TROUVER LE MARQUEUR
'===============================================================================
Private Function TrouverEtPreparerMarqueur() As Boolean
    Dim rng As Object
    Set rng = WordDoc.Content
    With rng.Find
        .ClearFormatting
        .Text = MARKER_TEXT
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    If rng.Find.Found Then
        Set InsertionRange = rng
        InsertionRange.Text = ""
        InsertionRange.Collapse 0
        TrouverEtPreparerMarqueur = True
    Else
        MsgBox "Marqueur " & MARKER_TEXT & " introuvable dans Word.", vbCritical
    End If
End Function

'===============================================================================
' TRAITEMENT DES DONNÉES
'===============================================================================
Private Sub TraiterDonnees()
    Dim r As Long, lastRow As Long
    lastRow = WorksheetFunction.Min(END_ROW, ExcelWS.Cells(ExcelWS.Rows.Count, COL_LANGUE).End(xlUp).row)
    
    For r = START_ROW To lastRow
        TotalRows = TotalRows + 1
        If TraiterLigne(r) Then InsertedRows = InsertedRows + 1
        If r Mod 50 = 0 Then DoEvents
    Next r
End Sub

'===============================================================================
' TRAITER UNE LIGNE
'===============================================================================
Private Function TraiterLigne(ByVal r As Long) As Boolean
    Dim langue As String, flag As String, t2 As String, t3 As String, t4 As String, txt As String
    
    langue = UCase(Trim(ExcelWS.Cells(r, COL_LANGUE).Value))
    flag = Trim(ExcelWS.Cells(r, COL_FLAG).Value)
    
    ' Filtre strict : FR et "Utilisé"
    If langue <> "FR" Then Exit Function
    If Not EstUtiliseExact(flag) Then Exit Function
    
    t2 = Trim(ExcelWS.Cells(r, COL_TITRE2).Value)
    t3 = Trim(ExcelWS.Cells(r, COL_TITRE3).Value)
    t4 = Trim(ExcelWS.Cells(r, COL_TITRE4).Value)
    txt = Trim(ExcelWS.Cells(r, COL_TEXTE).Value)
    
    If t2 = "" And t3 = "" And t4 = "" And txt = "" Then Exit Function
    
    If lngExcel = 1036 Then   ' Français
        If t2 <> "" And t2 <> PrevTitre2 Then InsererTitre t2, "Titre 2": PrevTitre2 = t2: PrevTitre3 = "": PrevTitre4 = ""
        If t3 <> "" And t3 <> PrevTitre3 Then InsererTitre t3, "Titre 3": PrevTitre3 = t3: PrevTitre4 = ""
        If t4 <> "" And t4 <> PrevTitre4 Then InsererTitre t4, "Titre 4": PrevTitre4 = t4
        If txt <> "" Then InsererTexte txt

    ElseIf lngExcel = 1033 Then   ' Anglais
        If t2 <> "" And t2 <> PrevTitre2 Then InsererTitre t2, "Heading 2": PrevTitre2 = t2: PrevTitre3 = "": PrevTitre4 = ""
        If t3 <> "" And t3 <> PrevTitre3 Then InsererTitre t3, "Heading 3": PrevTitre3 = t3: PrevTitre4 = ""
        If t4 <> "" And t4 <> PrevTitre4 Then InsererTitre t4, "Heading 4": PrevTitre4 = t4
        If txt <> "" Then InsererTexte txt
    End If
    
    TraiterLigne = True
End Function

'===============================================================================
' OUTILS
'===============================================================================
Private Function EstUtiliseExact(ByVal t As String) As Boolean
    Dim n As String
    n = LCase(Trim(t))
    n = Replace(n, "é", "e")
    n = Replace(n, "è", "e")
    n = Replace(n, "ê", "e")
    n = Replace(n, "ë", "e")
    EstUtiliseExact = (n = "utilise")
End Function

Private Sub InsererTitre(ByVal txt As String, ByVal styleNom As String)
    InsertionRange.InsertAfter txt & vbCr
    Dim r As Object: Set r = WordDoc.Range(InsertionRange.Start, InsertionRange.End)
    On Error Resume Next
    r.Style = styleNom
    On Error GoTo 0
    InsertionRange.Collapse 0
End Sub

Private Sub InsererTexte(ByVal txt As String)
    InsertionRange.InsertAfter txt & vbCr
    InsertionRange.Collapse 0
End Sub


