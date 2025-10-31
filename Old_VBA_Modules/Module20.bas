Attribute VB_Name = "Module20"
'===============================================================================
' Module d'automatisation Excel vers Word - Insertion GC (col S = "X")
' Version GC - identique � Annexe 2 mais NE PREND QUE les lignes o� S = "X"
' Version corrig�e - Ouvre le fichier exact PP-8002-FR.dotx
' Version robuste avec gestion d'erreurs compl�te et journalisation
'===============================================================================

Option Explicit

' Configuration globale
Private Const EXCEL_FILE As String = "" ' Sera d�fini dynamiquement
Private Const WORD_TEMPLATE As String = "PP_8002-FR.dotx"
Private Const SHEET_NAME As String = "2.4-PP & SOW Annexe 2"
Private Const START_ROW As Long = 11
Private Const END_ROW As Long = 672
Private Const MARKER_TEXT As String = "(Annexe 2)"
Private Const MAX_RETRY_PASTE As Integer = 3
Private Const RETRY_DELAY_MS As Integer = 1000

' Colonnes Excel
Private Const COL_TITRE2 As Integer = 6   ' F
Private Const COL_TITRE3 As Integer = 7   ' G
Private Const COL_TITRE4 As Integer = 8   ' H
Private Const COL_TITRE5 As Integer = 9   ' I (ignor�)
Private Const COL_TEXTE  As Integer = 15  ' O
Private Const COL_LANGUE As Integer = 17  ' Q (conserv� mais non utilis� ici)
Private Const COL_FLAG   As Integer = 24  ' X (conserv� mais non utilis� ici)
Private Const COL_SELECT As Integer = 20  ' S (NOUVEAU - doit �tre "X")

' Variables globales
Private WordApp As Object
Private WordDoc As Object
Private ExcelWB As Workbook
Private ExcelWS As Worksheet
Private InsertionRange As Object

' Variables de m�morisation pour d�-duplication
Private PrevTitre2 As String
Private PrevTitre3 As String
Private PrevTitre4 As String

' Statistiques et journalisation
Private TotalRows As Long
Private FilteredRows As Long
Private InsertedRows As Long
Private ErrorsList As Collection
Private WarningsList As Collection
Private startTime As Date
Private LogFile As Integer

'===============================================================================
' POINT D'ENTR�E PRINCIPAL
'===============================================================================
Public Sub zzz_PP_SOW_8002_FR_Annexe_2_BM()
    Dim success As Boolean
    Dim rapport As String
    
    ' Initialiser
    Call InitialiserVariables
    
    ' Message de d�marrage
    Debug.Print String(60, "=")
    Debug.Print "D�BUT DE L'AUTOMATISATION EXCEL VERS WORD (GC - S = ""X"")"
    Debug.Print String(60, "=")
    Call EcrireLog("D�BUT DE L'AUTOMATISATION (GC) - " & Now)
    
    On Error GoTo GestionErreurGlobale
    
    ' �tape 1: Localiser les fichiers
    Debug.Print "�tape 1: Localisation des fichiers..."
    Call EcrireLog("�tape 1: Localisation des fichiers...")
    If Not LocaliserFichiers() Then
        Call GenererRapport("�CHEC: Fichiers non trouv�s")
        Exit Sub
    End If
    
    ' �tape 2: Ouvrir Word et le document
    Debug.Print "�tape 2: Ouverture de Word..."
    Call EcrireLog("�tape 2: Ouverture de Word...")
    If Not OuvrirWord() Then
        Call GenererRapport("�CHEC: Impossible d'ouvrir Word")
        Exit Sub
    End If
    
    ' �tape 3: Ouvrir Excel et la feuille
    Debug.Print "�tape 3: Ouverture de la feuille Excel..."
    Call EcrireLog("�tape 3: Ouverture de la feuille Excel...")
    If Not OuvrirExcel() Then
        Call GenererRapport("�CHEC: Impossible d'ouvrir la feuille Excel")
        Call Nettoyer
        Exit Sub
    End If
    
    ' �tape 4: Trouver et pr�parer le marqueur
    Debug.Print "�tape 4: Recherche du marqueur '" & MARKER_TEXT & "'..."
    Call EcrireLog("�tape 4: Recherche du marqueur...")
    If Not TrouverEtPreparerMarqueur() Then
        Call GenererRapport("�CHEC: Marqueur '" & MARKER_TEXT & "' non trouv�")
        Call Nettoyer
        Exit Sub
    End If
    
    ' �tape 5: Traiter les donn�es
    Debug.Print "�tape 5: Traitement des donn�es..."
    Call EcrireLog("�tape 5: Traitement des donn�es...")
    Call TraiterDonnees
    
    ' �tape 6: G�n�rer le rapport final
    Call GenererRapport("SUCC�S: Traitement termin� (GC)")
    
    ' Message final
    If InsertedRows > 0 Then
        MsgBox "Traitement termin� avec succ�s!" & vbCrLf & vbCrLf & _
               InsertedRows & " �l�ments ins�r�s dans le document Word." & vbCrLf & vbCrLf & _
               "IMPORTANT: Le document Word est ouvert et NON sauvegard�." & vbCrLf & _
               "Veuillez v�rifier et sauvegarder manuellement si le r�sultat est correct.", _
               vbInformation, "Automatisation termin�e (GC)"
    Else
        MsgBox "Aucune donn�e n'a �t� ins�r�e." & vbCrLf & vbCrLf & _
               "V�rifiez que la colonne S contient bien ""X"" sur certaines lignes.", _
               vbExclamation, "Aucune insertion (GC)"
    End If
    
    Exit Sub
    
GestionErreurGlobale:
    Call AjouterErreur("Erreur fatale: " & Err.Description)
    Call GenererRapport("�CHEC: Erreur fatale - " & Err.Description)
    Call Nettoyer
End Sub

'===============================================================================
' INITIALISATION
'===============================================================================
Private Sub InitialiserVariables()
    Set ErrorsList = New Collection
    Set WarningsList = New Collection
    TotalRows = 0
    FilteredRows = 0
    InsertedRows = 0
    PrevTitre2 = ""
    PrevTitre3 = ""
    PrevTitre4 = ""
    startTime = Now
    
    ' Ouvrir le fichier log
    Dim logFileName As String
    logFileName = ThisWorkbook.Path & "\automation_log_" & Format(Now, "yyyymmdd_hhmmss") & "_GC.txt"
    LogFile = FreeFile
    Open logFileName For Output As #LogFile
End Sub

'===============================================================================
' LOCALISATION DES FICHIERS
'===============================================================================
Private Function LocaliserFichiers() As Boolean
    On Error GoTo GestionErreur
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim excelFound As Boolean
    Dim wordFound As Boolean
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(ThisWorkbook.Path)
    
    ' V�rifier le mod�le Word
    If fso.FileExists(ThisWorkbook.Path & "\" & WORD_TEMPLATE) Then
        wordFound = True
        Call EcrireLog("Mod�le Word trouv�: " & WORD_TEMPLATE)
    End If
    
    ' V�rifier qu'on a la feuille requise dans ce classeur
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_NAME Then
            excelFound = True
            Call EcrireLog("Feuille Excel trouv�e: " & SHEET_NAME)
            Exit For
        End If
    Next ws
    
    LocaliserFichiers = excelFound And wordFound
    
    If Not wordFound Then
        Call AjouterErreur("Mod�le Word non trouv�: " & WORD_TEMPLATE)
    End If
    If Not excelFound Then
        Call AjouterErreur("Feuille '" & SHEET_NAME & "' non trouv�e dans ce classeur")
    End If
    
    Exit Function
    
GestionErreur:
    Call AjouterErreur("Erreur localisation fichiers: " & Err.Description)
    LocaliserFichiers = False
End Function

'===============================================================================
' OUVERTURE DE WORD - VERSION CORRIG�E
' Ouvre le fichier exact PP-8002-FR.dotx (pas une copie)
'===============================================================================
Private Function OuvrirWord() As Boolean
    On Error GoTo GestionErreur
    
    Dim cheminComplet As String
    Dim fso As Object
    
    ' Construire le chemin complet
    cheminComplet = ThisWorkbook.Path & "\" & WORD_TEMPLATE
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' V�rifier l'existence du fichier
    If Not fso.FileExists(cheminComplet) Then
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Le fichier " & WORD_TEMPLATE & " n'existe pas dans:" & vbCrLf & _
                        ThisWorkbook.Path & vbCrLf & vbCrLf & _
                        "Voulez-vous le cr�er maintenant ?", vbQuestion + vbYesNo, "Fichier manquant")
        
        If reponse = vbYes Then
            Call CreerFichierWordSiNecessaire
            ' Apr�s cr�ation, v�rifier � nouveau
            If Not fso.FileExists(cheminComplet) Then
                Call AjouterErreur("�chec de cr�ation du fichier " & WORD_TEMPLATE)
                OuvrirWord = False
                Exit Function
            End If
        Else
            Call AjouterErreur("Fichier " & WORD_TEMPLATE & " requis mais non trouv�")
            OuvrirWord = False
            Exit Function
        End If
    End If
    
    ' Cr�er l'application Word
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    WordApp.DisplayAlerts = False
    
    ' OUVRIR LE FICHIER EXACT
    Set WordDoc = WordApp.Documents.Open(cheminComplet)
    
    Call EcrireLog("Fichier Word ouvert directement: " & cheminComplet)
    OuvrirWord = True
    Exit Function
    
GestionErreur:
    Call AjouterErreur("Erreur ouverture Word: " & Err.Description)
    OuvrirWord = False
End Function

'===============================================================================
' CR�ATION DU FICHIER WORD SI N�CESSAIRE
'===============================================================================
Private Sub CreerFichierWordSiNecessaire()
    On Error GoTo GestionErreur
    
    Dim cheminComplet As String
    Dim WordAppTemp As Object
    Dim WordDocTemp As Object
    
    cheminComplet = ThisWorkbook.Path & "\" & WORD_TEMPLATE
    
    ' Cr�er Word temporairement
    Set WordAppTemp = CreateObject("Word.Application")
    WordAppTemp.Visible = True
    
    ' Cr�er un nouveau document
    Set WordDocTemp = WordAppTemp.Documents.Add
    
    ' Ajouter du contenu basique avec le marqueur
    WordDocTemp.Content.Text = "Document PP-8002-FR" & vbCrLf & vbCrLf & _
                              "Les donn�es seront ins�r�es ci-dessous:" & vbCrLf & _
                              MARKER_TEXT & vbCrLf & vbCrLf & _
                              "Fin du document"
    
    ' Sauvegarder comme template
    WordDocTemp.SaveAs2 cheminComplet, 16 ' 16 = wdFormatXMLTemplate (.dotx)
    
    Call EcrireLog("Fichier PP-8002-FR.dotx cr��: " & cheminComplet)
    
    ' Fermer le document temporaire
    WordDocTemp.Close
    WordAppTemp.Quit
    Set WordDocTemp = Nothing
    Set WordAppTemp = Nothing
    
    Exit Sub
    
GestionErreur:
    Call AjouterErreur("Erreur cr�ation fichier Word: " & Err.Description)
    On Error Resume Next
    If Not WordDocTemp Is Nothing Then WordDocTemp.Close
    If Not WordAppTemp Is Nothing Then WordAppTemp.Quit
    On Error GoTo 0
End Sub

'===============================================================================
' OUVERTURE D'EXCEL
'===============================================================================
Private Function OuvrirExcel() As Boolean
    On Error GoTo GestionErreur
    
    ' Utiliser ce classeur
    Set ExcelWB = ThisWorkbook
    
    ' Trouver la feuille
    Dim ws As Worksheet
    For Each ws In ExcelWB.Worksheets
        If ws.Name = SHEET_NAME Then
            Set ExcelWS = ws
            Exit For
        End If
    Next ws
    
    If ExcelWS Is Nothing Then
        Call AjouterErreur("Feuille '" & SHEET_NAME & "' non trouv�e")
        OuvrirExcel = False
        Exit Function
    End If
    
    ' D�sactiver les mises � jour pour performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call EcrireLog("Feuille Excel '" & SHEET_NAME & "' ouverte")
    OuvrirExcel = True
    Exit Function
    
GestionErreur:
    Call AjouterErreur("Erreur ouverture Excel: " & Err.Description)
    OuvrirExcel = False
End Function

'===============================================================================
' RECHERCHE ET PR�PARATION DU MARQUEUR - VERSION CORRECTE
'===============================================================================
Private Function TrouverEtPreparerMarqueur() As Boolean
    On Error GoTo GestionErreur
    
    Dim rngRecherche As Object
    Dim trouve As Boolean
    
    ' Cr�er un range de recherche sur tout le document
    Set rngRecherche = WordDoc.Range
    
    ' Rechercher le marqueur
    With rngRecherche.Find
        .ClearFormatting
        .Text = MARKER_TEXT
        .Forward = True
        .Wrap = 0 ' wdFindStop - ne pas boucler
        trouve = .Execute
    End With
    
    If trouve Then
        ' Le range rngRecherche contient maintenant EXACTEMENT le marqueur trouv�
        Set InsertionRange = rngRecherche
        
        ' Maintenant on efface JUSTE le marqueur et on garde la position
        InsertionRange.Text = ""
        
        Call EcrireLog("Marqueur '" & MARKER_TEXT & "' trouv� et supprim� - position d'insertion d�finie")
        TrouverEtPreparerMarqueur = True
    Else
        Call AjouterErreur("Marqueur '" & MARKER_TEXT & "' non trouv� dans le document")
        TrouverEtPreparerMarqueur = False
    End If
    
    Exit Function
    
GestionErreur:
    Call AjouterErreur("Erreur recherche marqueur: " & Err.Description)
    TrouverEtPreparerMarqueur = False
End Function

'===============================================================================
' TRAITEMENT DES DONN�ES
'===============================================================================
Private Sub TraiterDonnees()
    Dim row As Long
    Dim lastRow As Long
    
    On Error GoTo GestionErreur
    
    ' D�terminer la derni�re ligne
    lastRow = DeterminerDerniereLigne()
    Call EcrireLog("Traitement des lignes " & START_ROW & " � " & lastRow & " (filtre S=""X"")")
    
    ' Traiter chaque ligne
    For row = START_ROW To lastRow
        TotalRows = TotalRows + 1
        
        ' Afficher la progression
        If row Mod 50 = 0 Then
            Debug.Print "Progression: ligne " & row & "/" & lastRow
            DoEvents
        End If
        
        ' Traiter la ligne avec gestion d'erreur individuelle
        On Error Resume Next
        If TraiterLigne(row) Then
            FilteredRows = FilteredRows + 1
            InsertedRows = InsertedRows + 1
        End If
        If Err.Number <> 0 Then
            Call AjouterErreur("Ligne " & row & ": " & Err.Description)
            Err.Clear
        End If
        On Error GoTo GestionErreur
    Next row
    
    Call EcrireLog("Traitement termin�: " & InsertedRows & " lignes ins�r�es")
    Exit Sub
    
GestionErreur:
    Call AjouterErreur("Erreur traitement donn�es: " & Err.Description)
End Sub

'===============================================================================
' TRAITEMENT D'UNE LIGNE (filtrage: SEULEMENT S = "X")
'===============================================================================
Private Function TraiterLigne(row As Long) As Boolean
    On Error GoTo GestionErreur
    
    Dim titre2 As String
    Dim titre3 As String
    Dim titre4 As String
    Dim texte As String
    Dim selectS As String
    
    ' Filtrer UNIQUEMENT sur la colonne S = "X"
    selectS = UCase(Trim(ObtenirValeurCellule(row, COL_SELECT)))
    If selectS <> "X" Then
        TraiterLigne = False
        Exit Function
    End If
    
    ' R�cup�rer les donn�es
    titre2 = ObtenirValeurCellule(row, COL_TITRE2)
    titre3 = ObtenirValeurCellule(row, COL_TITRE3)
    titre4 = ObtenirValeurCellule(row, COL_TITRE4)
    texte = ObtenirValeurCellule(row, COL_TEXTE)
    
    ' V�rifier qu'il y a au moins une donn�e
    If titre2 = "" And titre3 = "" And titre4 = "" And texte = "" Then
        TraiterLigne = False
        Exit Function
    End If
    
    ' Traiter les titres avec d�-duplication
    If titre2 <> "" And titre2 <> PrevTitre2 Then
        Call InsererTitre(titre2, "Titre 2", 2)
        PrevTitre2 = titre2
        PrevTitre3 = ""
        PrevTitre4 = ""
    End If
    
    If titre3 <> "" And titre3 <> PrevTitre3 Then
        Call InsererTitre(titre3, "Titre 3", 3)
        PrevTitre3 = titre3
        PrevTitre4 = ""
    End If
    
    If titre4 <> "" And titre4 <> PrevTitre4 Then
        Call InsererTitre(titre4, "Titre 4", 4)
        PrevTitre4 = titre4
    End If
    
    ' Traiter le texte
    If texte <> "" Then
        Call InsererTexteAvecFormat(ExcelWS.Cells(row, COL_TEXTE), texte)
    End If
    
    TraiterLigne = True
    Exit Function
    
GestionErreur:
    Call AjouterErreur("Erreur ligne " & row & ": " & Err.Description)
    TraiterLigne = False
End Function

'===============================================================================
' OBTENIR VALEUR CELLULE (ROBUSTE)
'===============================================================================
Private Function ObtenirValeurCellule(row As Long, col As Integer) As String
    On Error Resume Next
    
    Dim valeur As Variant
    
    ' Essayer d'abord .Value
    valeur = ExcelWS.Cells(row, col).Value
    
    ' Si erreur ou vide, essayer .Text
    If Err.Number <> 0 Or IsError(valeur) Then
        Err.Clear
        valeur = ExcelWS.Cells(row, col).Text
    End If
    
    ' Convertir en string
    If Not IsError(valeur) And Not IsEmpty(valeur) Then
        ObtenirValeurCellule = CStr(valeur)
    Else
        ObtenirValeurCellule = ""
    End If
    
    On Error GoTo 0
End Function

'===============================================================================
' V�RIFICATION "UTILIS�" (TOL�RANT AUX ACCENTS) - conserv�e si besoin ailleurs
'===============================================================================
Private Function CommenceParUtilise(texte As String) As Boolean
    If texte = "" Then
        CommenceParUtilise = False
        Exit Function
    End If
    
    Dim texteNormalise As String
    texteNormalise = LCase(Trim(texte))
    
    ' Remplacer les accents
    texteNormalise = Replace(texteNormalise, "�", "e")
    texteNormalise = Replace(texteNormalise, "�", "e")
    texteNormalise = Replace(texteNormalise, "�", "e")
    texteNormalise = Replace(texteNormalise, "�", "e")
    
    CommenceParUtilise = (Left(texteNormalise, 7) = "utilise")
End Function

'===============================================================================
' CONVERSION DES RETOURS � LA LIGNE
'===============================================================================
Private Function ConvertirRetoursLigne(texte As String) As String
    If texte = "" Then
        ConvertirRetoursLigne = ""
        Exit Function
    End If
    
    ' Remplacer tous les types de retours par le saut manuel Word
    Dim resultat As String
    resultat = texte
    resultat = Replace(resultat, vbCrLf, Chr(11))
    resultat = Replace(resultat, vbLf, Chr(11))
    resultat = Replace(resultat, vbCr, Chr(11))
    resultat = Replace(resultat, Chr(10), Chr(11))
    resultat = Replace(resultat, Chr(13), Chr(11))
    
    ConvertirRetoursLigne = resultat
End Function

'===============================================================================
' INSERTION D'UN TITRE - VERSION CORRIG�E
'===============================================================================
Private Sub InsererTitre(texte As String, nomStyle As String, niveau As Integer)
    On Error GoTo GestionErreur
    
    If texte = "" Then Exit Sub
    
    ' Convertir les retours � la ligne
    texte = ConvertirRetoursLigne(texte)
    
    ' M�moriser la position de d�part
    Dim startPos As Long
    startPos = InsertionRange.Start
    
    ' Ins�rer le texte
    InsertionRange.InsertAfter texte
    
    ' Cr�er un range qui couvre EXACTEMENT le texte qu'on vient d'ins�rer
    Dim titreRange As Object
    Set titreRange = WordDoc.Range(startPos, InsertionRange.End)
    
    ' Appliquer le style sur ce range pr�cis
    Call AppliquerStyleSurRange(titreRange, nomStyle, niveau)
    
    ' Ajouter un saut de paragraphe
    InsertionRange.InsertParagraphAfter
    
    ' D�placer le point d'insertion
    InsertionRange.Collapse 0 ' wdCollapseEnd
    
    Exit Sub
    
GestionErreur:
    Call AjouterAvertissement("Erreur insertion titre: " & Err.Description)
End Sub

'===============================================================================
' APPLIQUER UN STYLE SUR UN RANGE PR�CIS
'===============================================================================
Private Sub AppliquerStyleSurRange(targetRange As Object, nomStyle As String, niveau As Integer)
    On Error Resume Next
    
    ' Essayer le style fran�ais
    targetRange.Style = nomStyle
    If Err.Number = 0 Then
        Call EcrireLog("Style '" & nomStyle & "' appliqu� avec succ�s")
        Exit Sub
    End If
    
    ' Essayer le style anglais
    Err.Clear
    Select Case nomStyle
        Case "Titre 2"
            targetRange.Style = "Heading 2"
        Case "Titre 3"
            targetRange.Style = "Heading 3"
        Case "Titre 4"
            targetRange.Style = "Heading 4"
        Case Else
            targetRange.Style = "Normal"
    End Select
    
    If Err.Number = 0 Then
        Call EcrireLog("Style anglais appliqu� pour '" & nomStyle & "'")
        Exit Sub
    End If
    
    ' Fallback sur Normal avec OutlineLevel et formatage manuel
    Err.Clear
    targetRange.Style = "Normal"
    
    ' Appliquer formatage de titre manuellement
    With targetRange.Font
        Select Case niveau
            Case 2
                .Bold = True
                .Size = 16
            Case 3
                .Bold = True
                .Size = 14
            Case 4
                .Bold = True
                .Size = 12
        End Select
    End With
    
    ' Essayer de d�finir le niveau de plan
    On Error Resume Next
    targetRange.Paragraphs(1).OutlineLevel = niveau
    On Error GoTo 0
    
    If Err.Number <> 0 Then
        Call AjouterAvertissement("Impossible d'appliquer le style '" & nomStyle & "' - formatage manuel appliqu�")
    Else
        Call EcrireLog("Formatage manuel appliqu� pour '" & nomStyle & "'")
    End If
    
    On Error GoTo 0
End Sub

'===============================================================================
' INSERTION DE TEXTE AVEC FORMATAGE - VERSION CORRIG�E
'===============================================================================
Private Sub InsererTexteAvecFormat(cellule As Range, texte As String)
    On Error GoTo GestionErreur
    
    If texte = "" Then Exit Sub
    
    ' M�moriser la position de d�part
    Dim startPos As Long
    startPos = InsertionRange.Start
    
    ' V�rifier si formatage mixte
    If CelluleFormatageMixte(cellule) Then
        ' Collage riche
        If Not CollerTexteRiche(cellule) Then
            ' Fallback sur insertion simple
            Call InsererTexteSimple(cellule, texte)
        End If
    Else
        ' Insertion simple avec format uniforme
        Call InsererTexteSimple(cellule, texte)
    End If
    
    ' Cr�er un range qui couvre le texte qu'on vient d'ins�rer
    Dim texteRange As Object
    Set texteRange = WordDoc.Range(startPos, InsertionRange.End)
    
    ' Appliquer le style Normal sur ce range pr�cis
    On Error Resume Next
    texteRange.Style = "Normal"
    If Err.Number <> 0 Then
        Call AjouterAvertissement("Impossible d'appliquer le style Normal au texte")
    End If
    On Error GoTo GestionErreur
    
    ' Ajouter un saut de paragraphe
    InsertionRange.InsertParagraphAfter
    InsertionRange.Collapse 0
    
    Exit Sub
    
GestionErreur:
    Call AjouterAvertissement("Erreur insertion texte: " & Err.Description)
End Sub

'===============================================================================
' V�RIFIER FORMATAGE MIXTE
'===============================================================================
Private Function CelluleFormatageMixte(cellule As Range) As Boolean
    On Error GoTo PasDeFormatageMixte
    
    Dim longueur As Integer
    Dim i As Integer
    Dim premierGras As Variant
    Dim premierItalique As Variant
    
    longueur = Len(cellule.Value)
    If longueur <= 1 Then
        CelluleFormatageMixte = False
        Exit Function
    End If
    
    ' V�rifier les premiers caract�res
    premierGras = cellule.Characters(1, 1).Font.Bold
    premierItalique = cellule.Characters(1, 1).Font.Italic
    
    ' Comparer avec les suivants (�chantillon)
    For i = 2 To WorksheetFunction.Min(longueur, 10)
        If cellule.Characters(i, 1).Font.Bold <> premierGras Or _
           cellule.Characters(i, 1).Font.Italic <> premierItalique Then
            CelluleFormatageMixte = True
            Exit Function
        End If
    Next i
    
PasDeFormatageMixte:
    CelluleFormatageMixte = False
End Function

'===============================================================================
' COLLER TEXTE RICHE (AVEC RETRY)
'===============================================================================
Private Function CollerTexteRiche(cellule As Range) As Boolean
    Dim tentative As Integer
    
    On Error GoTo EchecCollage
    
    For tentative = 1 To MAX_RETRY_PASTE
        ' Copier la cellule
        cellule.Copy
        
        ' Attendre un peu
        Application.Wait Now + TimeValue("00:00:01")
        
        ' Coller en conservant le formatage
        InsertionRange.PasteSpecial DataType:=7 ' wdPasteRTF
        
        ' Si succ�s, nettoyer et sortir
        Application.CutCopyMode = False
        CollerTexteRiche = True
        Exit Function
    Next tentative
    
EchecCollage:
    Application.CutCopyMode = False
    Call AjouterAvertissement("�chec du collage riche apr�s " & MAX_RETRY_PASTE & " tentatives")
    CollerTexteRiche = False
End Function

'===============================================================================
' INSERTION TEXTE SIMPLE
'===============================================================================
Private Sub InsererTexteSimple(cellule As Range, texte As String)
    On Error GoTo GestionErreur
    
    ' Ins�rer le texte
    InsertionRange.InsertAfter texte
    
    ' Appliquer le formatage uniforme si pr�sent
    On Error Resume Next
    Dim textRange As Object
    Set textRange = InsertionRange.Duplicate
    textRange.MoveStart 1, -Len(texte) ' wdCharacter
    
    If cellule.Font.Bold Then textRange.Font.Bold = True
    If cellule.Font.Italic Then textRange.Font.Italic = True
    If cellule.Font.Underline <> xlUnderlineStyleNone Then textRange.Font.Underline = True
    
    Exit Sub
    
GestionErreur:
    ' Au minimum, le texte est ins�r�
End Sub

'===============================================================================
' D�TERMINER LA DERNI�RE LIGNE
'===============================================================================
Private Function DeterminerDerniereLigne() As Long
    On Error GoTo UtiliserDefaut
    
    Dim derniereLigne As Long
    
    ' M�thode 1: UsedRange
    derniereLigne = ExcelWS.UsedRange.Rows.Count + ExcelWS.UsedRange.row - 1
    
    ' Limiter au maximum configur�
    If derniereLigne > END_ROW Then
        derniereLigne = END_ROW
    End If
    
    If derniereLigne < START_ROW Then
        derniereLigne = END_ROW
    End If
    
    DeterminerDerniereLigne = derniereLigne
    Exit Function
    
UtiliserDefaut:
    DeterminerDerniereLigne = END_ROW
End Function

'===============================================================================
' NETTOYAGE
'===============================================================================
Private Sub Nettoyer()
    On Error Resume Next
    
    ' R�activer les param�tres Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Ne PAS fermer Word (document reste ouvert pour l'utilisateur)
    
    ' Lib�rer les objets
    Set InsertionRange = Nothing
    Set ExcelWS = Nothing
    Set ExcelWB = Nothing
    
    On Error GoTo 0
End Sub

'===============================================================================
' JOURNALISATION
'===============================================================================
Private Sub EcrireLog(message As String)
    On Error Resume Next
    Print #LogFile, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & message
    On Error GoTo 0
End Sub

Private Sub AjouterErreur(message As String)
    ErrorsList.Add message
    Call EcrireLog("ERREUR: " & message)
    Debug.Print "ERREUR: " & message
End Sub

Private Sub AjouterAvertissement(message As String)
    WarningsList.Add message
    Call EcrireLog("AVERTISSEMENT: " & message)
    Debug.Print "AVERTISSEMENT: " & message
End Sub

'===============================================================================
' G�N�RATION DU RAPPORT FINAL
'===============================================================================
Private Sub GenererRapport(statut As String)
    Dim rapport As String
    Dim duree As Double
    Dim i As Long
    
    On Error Resume Next
    
    duree = DateDiff("s", startTime, Now)
    
    ' Construire le rapport
    rapport = String(60, "=") & vbCrLf
    rapport = rapport & "RAPPORT D'EX�CUTION (GC - filtre S=""X"")" & vbCrLf
    rapport = rapport & String(60, "=") & vbCrLf
    rapport = rapport & "Statut: " & statut & vbCrLf
    rapport = rapport & "D�but: " & Format(startTime, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    rapport = rapport & "Fin: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    rapport = rapport & "Dur�e: " & duree & " secondes" & vbCrLf & vbCrLf
    
    rapport = rapport & "STATISTIQUES:" & vbCrLf
    rapport = rapport & "  - Lignes parcourues: " & TotalRows & vbCrLf
    rapport = rapport & "  - Lignes filtr�es (S = ""X""): " & FilteredRows & vbCrLf
    rapport = rapport & "  - Lignes ins�r�es avec succ�s: " & InsertedRows & vbCrLf & vbCrLf
    
    ' Erreurs
    If ErrorsList.Count > 0 Then
        rapport = rapport & "ERREURS (" & ErrorsList.Count & " total):" & vbCrLf
        For i = 1 To WorksheetFunction.Min(ErrorsList.Count, 20)
            rapport = rapport & "  " & i & ". " & ErrorsList(i) & vbCrLf
        Next i
        If ErrorsList.Count > 20 Then
            rapport = rapport & "  ... et " & (ErrorsList.Count - 20) & " autres erreurs" & vbCrLf
        End If
        rapport = rapport & vbCrLf
    End If
    
    ' Avertissements
    If WarningsList.Count > 0 Then
        rapport = rapport & "AVERTISSEMENTS (" & WarningsList.Count & " total):" & vbCrLf
        For i = 1 To WorksheetFunction.Min(WarningsList.Count, 10)
            rapport = rapport & "  " & i & ". " & WarningsList(i) & vbCrLf
        Next i
        If WarningsList.Count > 10 Then
            rapport = rapport & "  ... et " & (WarningsList.Count - 10) & " autres avertissements" & vbCrLf
        End If
        rapport = rapport & vbCrLf
    End If
    
    rapport = rapport & "NOTE IMPORTANTE:" & vbCrLf
    rapport = rapport & "  Le document Word n'a PAS �t� enregistr� automatiquement." & vbCrLf
    rapport = rapport & "  Veuillez v�rifier le r�sultat et sauvegarder manuellement si satisfaisant." & vbCrLf
    rapport = rapport & String(60, "=") & vbCrLf
    
    ' Afficher dans la fen�tre d'ex�cution
    Debug.Print rapport
    
    ' �crire dans le fichier log
    Call EcrireLog(vbCrLf & rapport)
    
    ' Fermer le fichier log
    Close #LogFile
    
    ' Sauvegarder le rapport dans un fichier texte s�par�
    Dim rapportFile As Integer
    Dim rapportFileName As String
    rapportFileName = ThisWorkbook.Path & "\rapport_execution_" & Format(startTime, "yyyymmdd_hhmmss") & "_GC.txt"
    rapportFile = FreeFile
    Open rapportFileName For Output As #rapportFile
    Print #rapportFile, rapport
    Close #rapportFile
    
    Debug.Print vbCrLf & "Rapport sauvegard� dans: " & rapportFileName
    
    On Error GoTo 0
End Sub

'===============================================================================
' PROC�DURES UTILITAIRES SUPPL�MENTAIRES
'===============================================================================

' Test rapide pour v�rifier que le module fonctionne
Private Sub TesterConfiguration()
    Dim msg As String
    
    msg = "Configuration actuelle (GC - S=""X""):" & vbCrLf & vbCrLf
    msg = msg & "Mod�le Word: " & WORD_TEMPLATE & vbCrLf
    msg = msg & "Feuille Excel: " & SHEET_NAME & vbCrLf
    msg = msg & "Lignes � traiter: " & START_ROW & " � " & END_ROW & vbCrLf
    msg = msg & "Marqueur � chercher: " & MARKER_TEXT & vbCrLf & vbCrLf
    msg = msg & "Colonnes utilis�es:" & vbCrLf
    msg = msg & "  F (col " & COL_TITRE2 & ") = Titre 2" & vbCrLf
    msg = msg & "  G (col " & COL_TITRE3 & ") = Titre 3" & vbCrLf
    msg = msg & "  H (col " & COL_TITRE4 & ") = Titre 4" & vbCrLf
    msg = msg & "  O (col " & COL_TEXTE & ") = Texte" & vbCrLf
    msg = msg & "  S (col " & COL_SELECT & ") = S�lection (doit �tre ""X"")" & vbCrLf
    MsgBox msg, vbInformation, "Configuration de l'automatisation (GC)"
End Sub

' Diagnostic des fichiers
Private Sub DiagnostiquerFichiers()
    Dim fso As Object
    Dim cheminExcel As String
    Dim cheminWord As String
    Dim msg As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Chemin du fichier Excel actuel
    cheminExcel = ThisWorkbook.Path
    cheminWord = cheminExcel & "\" & WORD_TEMPLATE
    
    msg = "DIAGNOSTIC DES FICHIERS (GC)" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "R�pertoire du fichier Excel:" & vbCrLf
    msg = msg & cheminExcel & vbCrLf & vbCrLf
    
    msg = msg & "Fichier Excel actuel:" & vbCrLf
    msg = msg & ThisWorkbook.Name & vbCrLf & vbCrLf
    
    msg = msg & "Mod�le Word recherch�:" & vbCrLf
    msg = msg & cheminWord & vbCrLf
    
    ' V�rifier l'existence du mod�le Word
    If fso.FileExists(cheminWord) Then
        msg = msg & "? TROUV�" & vbCrLf & vbCrLf
    Else
        msg = msg & "? NON TROUV�" & vbCrLf & vbCrLf
        
        ' Lister les fichiers .dotx dans le r�pertoire
        msg = msg & "Fichiers .dotx dans ce r�pertoire:" & vbCrLf
        Dim folder As Object
        Dim file As Object
        Set folder = fso.GetFolder(cheminExcel)
        
        Dim trouve As Boolean
        trouve = False
        For Each file In folder.Files
            If LCase(Right(file.Name, 5)) = ".dotx" Then
                msg = msg & "  - " & file.Name & vbCrLf
                trouve = True
            End If
        Next file
        
        If Not trouve Then
            msg = msg & "  (Aucun fichier .dotx trouv�)" & vbCrLf
        End If
        msg = msg & vbCrLf
    End If
    
    ' V�rifier la feuille Excel
    msg = msg & "Feuille Excel recherch�e: '" & SHEET_NAME & "'" & vbCrLf
    
    Dim ws As Worksheet
    Dim feuilletrouvee As Boolean
    feuilletrouvee = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_NAME Then
            feuilletrouvee = True
            Exit For
        End If
    Next ws
    
    If feuilletrouvee Then
        msg = msg & "? TROUV�E" & vbCrLf & vbCrLf
    Else
        msg = msg & "? NON TROUV�E" & vbCrLf & vbCrLf
        msg = msg & "Feuilles disponibles:" & vbCrLf
        For Each ws In ThisWorkbook.Worksheets
            msg = msg & "  - '" & ws.Name & "'" & vbCrLf
        Next ws
        msg = msg & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Diagnostic des fichiers (GC)"
    Debug.Print msg
End Sub

' R�initialiser compl�tement avant une nouvelle ex�cution
Private Sub ReinitialiserAutomatisation()
    On Error Resume Next
    
    ' R�activer les param�tres Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Fermer le fichier log s'il est ouvert
    Close #LogFile
    
    ' R�initialiser les variables
    Set WordApp = Nothing
    Set WordDoc = Nothing
    Set InsertionRange = Nothing
    Set ExcelWB = Nothing
    Set ExcelWS = Nothing
    Set ErrorsList = Nothing
    Set WarningsList = Nothing
    
    MsgBox "R�initialisation termin�e." & vbCrLf & _
           "Vous pouvez maintenant relancer l'automatisation (GC).", _
           vbInformation, "R�initialisation"
    
    On Error GoTo 0
End Sub

' FONCTION DE TEST UNITAIRE
Private Sub TesterFonctionsUnitaires()
    Debug.Print String(60, "=")
    Debug.Print "TESTS UNITAIRES (GC)"
    Debug.Print String(60, "=")
    
    ' Test 1: Filtre S="X"
    Debug.Print "Test 1 - Filtre colonne S:"
    Debug.Print "  Valeur 'X' -> " & (UCase(Trim("X")) = "X")
    Debug.Print "  Valeur 'x' -> " & (UCase(Trim("x")) = "X")
    Debug.Print "  Valeur '' -> " & (UCase(Trim("")) = "X")
    Debug.Print ""
    
    ' Test 2: Conversion retours � la ligne
    Debug.Print "Test 2 - Conversion retours � la ligne:"
    Dim testTexte As String
    testTexte = "Ligne 1" & vbCrLf & "Ligne 2" & vbLf & "Ligne 3"
    Debug.Print "  Original: " & Replace(Replace(testTexte, vbCrLf, "[CRLF]"), vbLf, "[LF]")
    Debug.Print "  Converti: " & Replace(ConvertirRetoursLigne(testTexte), Chr(11), "[^l]")
    Debug.Print ""
    
    Debug.Print "Tests termin�s"
    Debug.Print String(60, "=")
End Sub




