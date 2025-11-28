Attribute VB_Name = "Module1"
'==============================================================================
' DÉCLARATIONS GLOBALES - À PLACER EN HAUT DU MODULE
'==============================================================================
' Variable globale pour mémoriser la langue de Word
Public g_WordLanguage As String

' Type personnalisé pour l'analyse des lignes Annexe 3b
Public Type LigneInfo
    EstVide As Boolean
    EstTitre As Boolean
    EstSousTitre As Boolean
    EstTableau As Boolean
    ValeurAA As String
    ValeurAB As String
End Type

'==============================================================================
' MACRO PRINCIPALE : GENERER_LIVRABLE
' Description: Génère automatiquement des livrables Word/Excel selon paramètres
'==============================================================================
Public Sub Generer_Livrable()
    Dim startTime As Double
    Dim wsMaster As Worksheet
    Dim nomLivrable As String
    Dim langue As String
    Dim annexe1 As Boolean, annexe2 As Boolean, annexe3 As Boolean, annexe4 As Boolean
    Dim templatePath As String
    Dim basePath As String
    Dim typeLivrable As String
    Dim messageRecap As String
    Dim annexesGenerees As String
    
    On Error GoTo ErrHandler
    
    ' Réinitialiser la détection de langue Word
    g_WordLanguage = ""
    
    '--- DÉBUT CHRONOMÈTRE ---
    startTime = Timer
    
    '--- LECTURE DES PARAMÈTRES ---
    Debug.Print "=== DÉBUT GÉNÉRATION LIVRABLE ==="
    Debug.Print "Date/Heure: " & Now
    
    Set wsMaster = ThisWorkbook.Sheets("Master Guide")
    nomLivrable = wsMaster.Range("N7").Value
    langue = UCase(wsMaster.Range("O7").Value)
    annexe1 = wsMaster.Range("Q7").Value
    annexe2 = wsMaster.Range("Q8").Value
    annexe3 = wsMaster.Range("Q9").Value
    annexe4 = wsMaster.Range("Q10").Value
    
    Debug.Print "Livrable choisi: " & nomLivrable
    Debug.Print "Langue: " & langue
    Debug.Print "Annexes à générer: " & IIf(annexe1, "1 ", "") & IIf(annexe2, "2 ", "") & _
                IIf(annexe3, "3 ", "") & IIf(annexe4, "4", "")
    
    '--- VALIDATION DES PARAMÈTRES ---
    If nomLivrable = "" Then
        MsgBox "Erreur: Aucun livrable sélectionné en cellule N7", vbCritical, "Erreur Paramètres"
        Exit Sub
    End If
    
    If langue <> "FR" And langue <> "ENG" Then
        MsgBox "Erreur: La langue doit être FR ou ENG (cellule O7)", vbCritical, "Erreur Langue"
        Exit Sub
    End If
    
    '--- DÉTERMINATION DU TYPE DE LIVRABLE ---
    If InStr(nomLivrable, "PS 8002") > 0 Then
        typeLivrable = "PS"
    ElseIf InStr(nomLivrable, "PP 8002") > 0 Then
        typeLivrable = "PP"
    ElseIf InStr(nomLivrable, "SOW 8002") > 0 Then
        typeLivrable = "SOW"
    Else
        MsgBox "Erreur: Type de livrable non reconnu", vbCritical, "Erreur Type"
        Exit Sub
    End If
    
    '--- CONSTRUCTION DU CHEMIN DU TEMPLATE ---
    basePath = ThisWorkbook.Path
    templatePath = TrouverTemplate(nomLivrable, langue, basePath)
    
    If templatePath = "" Then
        MsgBox "Erreur: Template introuvable pour " & nomLivrable & " (" & langue & ")", _
               vbCritical, "Template Manquant"
        Exit Sub
    End If
    
    Debug.Print "Template trouvé: " & templatePath
    
    '--- GÉNÉRATION SELON LE TYPE ---
    If typeLivrable = "PS" Then
        ' Cas PS : Excel sans annexes
        Call Generer_PS_Global(templatePath, nomLivrable, langue)
        messageRecap = "Livrable PS Excel généré avec succès"
        
    ElseIf typeLivrable = "PP" Or typeLivrable = "SOW" Then
        ' Cas PP/SOW : Word avec annexes
        Dim WordApp As Object
        Dim WordDoc As Object
        
        ' Démarrage de Word
        On Error Resume Next
        Set WordApp = GetObject(, "Word.Application")
        If Err.Number <> 0 Then
            Set WordApp = CreateObject("Word.Application")
        End If
        On Error GoTo ErrHandler
        
        WordApp.Visible = True
        Set WordDoc = WordApp.Documents.Open(templatePath)
        
        ' Détecter la langue de Word une fois pour toutes
        Call DetecterLangueWord(WordApp)
        
        ' Insertion des annexes selon les cases cochées
        annexesGenerees = ""
        If annexe1 Then
            Call InsererAnnexe(WordApp, WordDoc, nomLivrable, langue, 1)
            annexesGenerees = annexesGenerees & "1, "
        End If
        If annexe2 Then
            Call InsererAnnexe(WordApp, WordDoc, nomLivrable, langue, 2)
            annexesGenerees = annexesGenerees & "2, "
        End If
        If annexe3 Then
            Call InsererAnnexe(WordApp, WordDoc, nomLivrable, langue, 3)
            annexesGenerees = annexesGenerees & "3, "
        End If
        If annexe4 Then
            Call InsererAnnexe(WordApp, WordDoc, nomLivrable, langue, 4)
            annexesGenerees = annexesGenerees & "4"
        End If
        
        ' MISE À JOUR DU SOMMAIRE APRÈS INSERTION DES ANNEXES
        Call MettreAJourSommaire(WordDoc)
        
        If annexesGenerees <> "" Then
            annexesGenerees = "Annexes insérées: " & annexesGenerees
        Else
            annexesGenerees = "Aucune annexe insérée"
        End If
        
        messageRecap = "Livrable " & typeLivrable & " Word généré avec succès" & vbCrLf & annexesGenerees
    End If
    
    '--- MESSAGE RÉCAPITULATIF ---
    Dim tempsExecution As Double
    tempsExecution = Round(Timer - startTime, 2)
    
    MsgBox "? GÉNÉRATION TERMINÉE" & vbCrLf & vbCrLf & _
           "?? Livrable: " & nomLivrable & vbCrLf & _
           "?? Langue: " & langue & vbCrLf & _
           messageRecap & vbCrLf & _
           "?? Temps d'exécution: " & tempsExecution & " secondes", _
           vbInformation, "Succès"
    
    Debug.Print "=== FIN GÉNÉRATION LIVRABLE ==="
    Debug.Print "Temps total: " & tempsExecution & " secondes"
    
    Exit Sub

ErrHandler:
    MsgBox "? Erreur lors de la génération:" & vbCrLf & _
           "Code: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, "Erreur"
    Debug.Print "ERREUR: " & Err.Description
End Sub

'==============================================================================
' FONCTION: ViderPressePapier
' Vide complètement le presse-papier Excel et système
'==============================================================================
Private Sub ViderPressePapier()
    On Error Resume Next
    
    ' Méthode 1: Vider le presse-papier Excel
    Application.CutCopyMode = False
    
    ' Méthode 2: Copier une cellule vide pour nettoyer
    Dim wsTemp As Worksheet
    Set wsTemp = ThisWorkbook.Worksheets(1)
    wsTemp.Range("A1").Copy
    Application.CutCopyMode = False
    
    ' Méthode 3: Utiliser l'API Windows pour vider le presse-papier système
    #If VBA7 Then
        Dim result As LongPtr
    #Else
        Dim result As Long
    #End If
    
    ' Tentative d'ouvrir et vider le presse-papier Windows
    On Error Resume Next
    CreateObject("htmlfile").parentWindow.clipboardData.clearData
    On Error GoTo 0
    
    Debug.Print "? Presse-papier vidé"
    
    On Error GoTo 0
End Sub

'==============================================================================
' FONCTION: CollerTableauAvecRetry - VERSION CORRIGÉE
' Colle un tableau Excel dans Word avec recopie avant chaque tentative
'==============================================================================
Private Function CollerTableauAvecRetry(anchor As Object, rngSource As Range, Optional contextInfo As String = "") As Boolean
    Const MAX_TENTATIVES As Integer = 3
    Dim tentative As Integer
    Dim succes As Boolean
    Dim delai As Double
    
    succes = False
    
    For tentative = 1 To MAX_TENTATIVES
        On Error Resume Next
        
        ' RECOPIER à chaque tentative (crucial!)
        rngSource.Copy
        DoEvents
        
        ' Tentative de collage
        anchor.PasteExcelTable False, False, False
        
        ' Vérifier si le collage a réussi
        If Err.Number = 0 Then
            succes = True
            If tentative > 1 Then
                Debug.Print "? Collage réussi à la tentative " & tentative & IIf(contextInfo <> "", " [" & contextInfo & "]", "")
            End If
            Exit For
        Else
            Debug.Print "?? Tentative " & tentative & "/" & MAX_TENTATIVES & " échouée" & _
                       IIf(contextInfo <> "", " [" & contextInfo & "]", "") & " - Erreur: " & Err.Description
            Err.Clear
            
            ' Attendre avant de réessayer (délai progressif)
            delai = 0.5 * tentative  ' 0.5s, 1s, 1.5s
            Application.Wait Now + TimeValue("0:00:0" & Format(delai, "0"))
            DoEvents
        End If
        
        On Error GoTo 0
    Next tentative
    
    CollerTableauAvecRetry = succes
    
    If Not succes Then
        Debug.Print "? ÉCHEC DÉFINITIF: Impossible de coller après " & MAX_TENTATIVES & " tentatives" & _
                   IIf(contextInfo <> "", " [" & contextInfo & "]", "")
    End If
End Function

'==============================================================================
' FONCTION: DetecterLangueWord
' Détecte si Word est en français ou anglais et mémorise le résultat
'==============================================================================
Private Sub DetecterLangueWord(WordApp As Object)
    On Error Resume Next
    
    ' Tentative de détecter via LanguageSettings (Office 2007+)
    Dim langID As Long
    langID = WordApp.LanguageSettings.LanguageID(2) ' 2 = msoLanguageIDUI
    
    ' Français : 1036, Anglais : 1033 (US) ou 2057 (UK)
    If langID = 1036 Then
        g_WordLanguage = "FR"
        Debug.Print "? Word détecté en FRANÇAIS (langID: " & langID & ")"
    ElseIf langID = 1033 Or langID = 2057 Then
        g_WordLanguage = "EN"
        Debug.Print "? Word détecté en ANGLAIS (langID: " & langID & ")"
    Else
        ' Fallback : on essaiera les deux versions
        g_WordLanguage = ""
        Debug.Print "?? Langue Word non détectée (langID: " & langID & "), mode auto"
    End If
    
    On Error GoTo 0
End Sub

'==============================================================================
' FONCTION: MettreAJourSommaire
' Met à jour toutes les tables des matières dans le document Word
'==============================================================================
Private Sub MettreAJourSommaire(WordDoc As Object)
    On Error Resume Next
    Dim toc As Object
    Dim compteur As Integer
    
    compteur = 0
    
    ' Parcourir toutes les tables des matières du document
    For Each toc In WordDoc.TablesOfContents
        toc.Update
        compteur = compteur + 1
    Next toc
    
    ' Si aucune table des matières classique, essayer les champs TOC
    If compteur = 0 Then
        Dim fld As Object
        For Each fld In WordDoc.Fields
            If fld.Type = 13 Then ' 13 = wdFieldTOC
                fld.Update
                compteur = compteur + 1
            End If
        Next fld
    End If
    
    If Err.Number <> 0 Then
        Debug.Print "Avertissement: Impossible de mettre à jour le sommaire - " & Err.Description
        Err.Clear
    Else
        Debug.Print "? Sommaire mis à jour (" & compteur & " table(s) trouvée(s))"
    End If
    
    On Error GoTo 0
End Sub

'==============================================================================
' FONCTION: ObtenirNomStyle
' Retourne le nom du style adapté à la langue de Word
'==============================================================================
Private Function ObtenirNomStyle(styleGenerique As String) As String
    ' Si langue détectée, retourner directement
    If g_WordLanguage = "FR" Then
        Select Case styleGenerique
            Case "Heading 1": ObtenirNomStyle = "Titre 1"
            Case "Heading 2": ObtenirNomStyle = "Titre 2"
            Case "Heading 3": ObtenirNomStyle = "Titre 3"
            Case "Heading 4": ObtenirNomStyle = "Titre 4"
            Case "Heading 5": ObtenirNomStyle = "Titre 5"
            Case "Heading 6": ObtenirNomStyle = "Titre 6"
            Case Else: ObtenirNomStyle = styleGenerique
        End Select
    ElseIf g_WordLanguage = "EN" Then
        ' Déjà en anglais, retourner tel quel
        ObtenirNomStyle = styleGenerique
    Else
        ' Mode auto : retourner le style générique (Heading X)
        ObtenirNomStyle = styleGenerique
    End If
End Function

'==============================================================================
' FONCTION: AppliquerStyleRobuste
' Applique un style de manière robuste (essaie FR puis EN si échec)
'==============================================================================
Private Sub AppliquerStyleRobuste(rng As Object, styleGenerique As String)
    On Error Resume Next
    Dim styleNom As String
    
    styleNom = ObtenirNomStyle(styleGenerique)
    
    ' Tentative 1 : avec le style détecté ou générique
    rng.Style = styleNom
    
    ' Si échec et qu'on n'avait pas détecté la langue
    If Err.Number <> 0 And g_WordLanguage = "" Then
        Err.Clear
        
        ' Essayer la version française
        Select Case styleGenerique
            Case "Heading 1": styleNom = "Titre 1"
            Case "Heading 2": styleNom = "Titre 2"
            Case "Heading 3": styleNom = "Titre 3"
            Case "Heading 4": styleNom = "Titre 4"
            Case "Heading 5": styleNom = "Titre 5"
            Case "Heading 6": styleNom = "Titre 6"
        End Select
        
        rng.Style = styleNom
        
        ' Si ça marche, mémoriser que Word est en français
        If Err.Number = 0 Then
            g_WordLanguage = "FR"
            Debug.Print "? Langue Word détectée: FRANÇAIS (par test de style)"
        Else
            ' Sinon, considérer que c'est anglais
            Err.Clear
            rng.Style = styleGenerique
            If Err.Number = 0 Then
                g_WordLanguage = "EN"
                Debug.Print "? Langue Word détectée: ANGLAIS (par test de style)"
            End If
        End If
    End If
    
    On Error GoTo 0
End Sub

'==============================================================================
' FONCTION: TrouverTemplate
'==============================================================================
Private Function TrouverTemplate(nomLivrable As String, langue As String, basePath As String) As String
    Dim dossierLangue As String
    Dim nomFichier As String
    Dim extension As String
    Dim cheminComplet As String
    
    ' Déterminer le dossier selon la langue
    If langue = "FR" Then
        dossierLangue = "\3-Dossier - livrables\Fr\"
    Else
        dossierLangue = "\3-Dossier - livrables\Eng\"
    End If
    
    ' Déterminer l'extension selon le type
    If InStr(nomLivrable, "PS 8002") > 0 Then
        extension = ".xlsx"
    Else
        extension = ".docx"
    End If
    
    ' Mapper le nom du livrable au nom du fichier template
    Select Case True
        ' ============================================
        ' PS Templates - FRANÇAIS
        ' ============================================
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "FR" And _
             (InStr(nomLivrable, "Général") > 0 Or InStr(nomLivrable, "General") > 0)
            nomFichier = "XXXXXXX-XXX-PS-8200-XXXX-0 (Fr)"
            
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "FR" And InStr(nomLivrable, "E&I") > 0
            nomFichier = "XXXXXXX-XXX-PS-8200-ITC_Travaux_Preparatoires_E&I-XXXX-0 (Fr)"
            
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "FR" And _
             (InStr(nomLivrable, "Modulaire") > 0 Or InStr(nomLivrable, "Modular") > 0)
            nomFichier = "XXXXXXX-XXX-PS-8200-ITC_Batiment_Modulaire-XXXX-0 (Fr)"
            
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "FR" And _
             (InStr(nomLivrable, "GC") > 0 Or InStr(nomLivrable, "Civil Works") > 0)
            nomFichier = "XXXXXXX-XXX-PS-8200-ITC_Travaux_Preparatoires_GC-XXXX-0 (Fr)"
            
        ' ============================================
        ' PS Templates - ANGLAIS
        ' ============================================
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "ENG" And _
             (InStr(nomLivrable, "Général") > 0 Or InStr(nomLivrable, "General") > 0)
            nomFichier = "XXXXXXX-XXX-PS-8200-XXXX-0 (Eng)"
            
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "ENG" And InStr(nomLivrable, "E&I") > 0
            nomFichier = "XXXXXXX-XXX-PS-8200-TSF_E&I_Preparatory_Works-XXXX-0 (Eng)"
            
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "ENG" And _
             (InStr(nomLivrable, "Modulaire") > 0 Or InStr(nomLivrable, "Modular") > 0)
            nomFichier = "XXXXXXX-XXX-PS-8200-TSF_Modular_Building-XXXX-0 (Eng)"
            
        Case InStr(nomLivrable, "PS 8002") > 0 And langue = "ENG" And _
             (InStr(nomLivrable, "GC") > 0 Or InStr(nomLivrable, "Civil Works") > 0)
            nomFichier = "XXXXXXX-XXX-PS-8200-TSF_Civil_Works_Preparatory_Works-XXXX-0 (Eng)"
            
        ' ============================================
        ' PP Templates - FRANÇAIS
        ' ============================================
        Case InStr(nomLivrable, "PP 8002") > 0 And langue = "FR"
            nomFichier = "XXXXXXX-XXX-PP-8200-XXXX-0 (Fr)"
            
        ' ============================================
        ' PP Templates - ANGLAIS
        ' ============================================
        Case InStr(nomLivrable, "PP 8002") > 0 And langue = "ENG"
            nomFichier = "XXXXXXX-XXX-PP-8200-XXXX-0 (Eng)"
            
        ' ============================================
        ' SOW Templates - FRANÇAIS
        ' ============================================
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "FR" And _
             (InStr(nomLivrable, "Général") > 0 Or InStr(nomLivrable, "General") > 0)
            nomFichier = "XXXXXXX-XXX-SOW-8200-XXXX-0 (Fr)"
            
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "FR" And InStr(nomLivrable, "E&I") > 0
            nomFichier = "XXXXXXX-XXX-SOW-8200-ITC_Travaux_Preparatoires_E&I-XXXX-0 (Fr)"
            
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "FR" And _
             (InStr(nomLivrable, "Modulaire") > 0 Or InStr(nomLivrable, "Modular") > 0)
            nomFichier = "XXXXXXX-XXX-SOW-8200-ITC_Batiment_Modulaire-XXXX-0 (Fr)"
            
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "FR" And _
             (InStr(nomLivrable, "GC") > 0 Or InStr(nomLivrable, "Civil Works") > 0)
            nomFichier = "XXXXXXX-XXX-SOW-8200-ITC_Travaux_Preparatoires_GC-XXXX-0 (Fr)"
            
        ' ============================================
        ' SOW Templates - ANGLAIS
        ' ============================================
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "ENG" And _
             (InStr(nomLivrable, "Général") > 0 Or InStr(nomLivrable, "General") > 0)
            nomFichier = "XXXXXXX-XXX-SOW-8200-XXXX-0 (Eng)"
            
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "ENG" And InStr(nomLivrable, "E&I") > 0
            nomFichier = "XXXXXXX-XXX-SOW-8200-TSF_E&I_Preparatory_Works-XXXX-0 (Eng)"
            
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "ENG" And _
             (InStr(nomLivrable, "Modulaire") > 0 Or InStr(nomLivrable, "Modular") > 0)
            nomFichier = "XXXXXXX-XXX-SOW-8200-TSF_Modular_Building-XXXX-0 (Eng)"
            
        Case InStr(nomLivrable, "SOW 8002") > 0 And langue = "ENG" And _
             (InStr(nomLivrable, "GC") > 0 Or InStr(nomLivrable, "Civil Works") > 0)
            nomFichier = "XXXXXXX-XXX-SOW-8200-TSF_Civil_Works_Preparatory_Works-XXXX-0 (Eng)"
            
        Case Else
            Debug.Print "? Aucun template trouvé pour: " & nomLivrable & " (" & langue & ")"
            TrouverTemplate = ""
            Exit Function
    End Select
    
    cheminComplet = basePath & dossierLangue & nomFichier & extension
    
    ' Vérifier si le fichier existe
    If Dir(cheminComplet) <> "" Then
        TrouverTemplate = cheminComplet
        Debug.Print "? Template trouvé: " & nomFichier & extension
    Else
        Debug.Print "? Template non trouvé: " & cheminComplet
        TrouverTemplate = ""
    End If
End Function

'==============================================================================
' PROCÉDURE: InsererAnnexe
'==============================================================================
Private Sub InsererAnnexe(WordApp As Object, WordDoc As Object, nomLivrable As String, langue As String, numAnnexe As Integer)
    On Error GoTo ErrAnnexe
    
    ' Déterminer quelle macro appeler selon le livrable et l'annexe
    Select Case numAnnexe
        Case 1
            Call ExecuterAnnexe1(WordApp, WordDoc, nomLivrable, langue)
        Case 2
            Call ExecuterAnnexe2(WordApp, WordDoc, nomLivrable, langue)
        Case 3
            Call ExecuterAnnexe3(WordApp, WordDoc, nomLivrable, langue)
        Case 4
            Call ExecuterAnnexe4(WordApp, WordDoc, nomLivrable, langue)
    End Select
    
    Exit Sub

ErrAnnexe:
    Debug.Print "Erreur InsererAnnexe " & numAnnexe & ": " & Err.Description
End Sub

'==============================================================================
' ADAPTATEURS pour appeler les bonnes macros selon le contexte
'==============================================================================
Private Sub ExecuterAnnexe1(WordApp As Object, WordDoc As Object, nomLivrable As String, langue As String)
    Dim mode As String
    
    ' Déterminer le mode
    If InStr(nomLivrable, "Général") > 0 Or InStr(nomLivrable, "General") > 0 Then
        mode = ""
    ElseIf InStr(nomLivrable, "GC") > 0 Or InStr(nomLivrable, "Civil Works") > 0 Then
        mode = "GC"
    ElseIf InStr(nomLivrable, "Modulaire") > 0 Or InStr(nomLivrable, "Modular") > 0 Then
        mode = "BM"
    ElseIf InStr(nomLivrable, "E&I") > 0 Then
        mode = "EI"
    End If
    
    ' Appeler la fonction ExportAnnexe1 adaptée
    Call ExportAnnexe1_Adapted(WordApp, WordDoc, langue, mode)
    
    ' PAUSE DE 3 SECONDES après collage
    Debug.Print "? Pause de 3 secondes après Annexe 1..."
    Application.Wait Now + TimeValue("0:00:03")
    
    ' VIDER LE PRESSE-PAPIER
    Call ViderPressePapier
End Sub

' ==========================================================================================================================================================
' ================================================================== PP/SOW 8002 ANNEXE 1 ==================================================================
' ==========================================================================================================================================================
Private Sub ExportAnnexe1_Adapted(WordApp As Object, WordDoc As Object, lang As String, filtreType As String)
    Dim wbActuel As Workbook, wbTemp As Workbook
    Dim wsSource As Worksheet, wsFiltered As Worksheet
    Dim anchor As Object, wordTable As Object
    Dim rngSrc As Range, rngFinal As Range
    Dim lastRow As Long, currentLastRow As Long, t0 As Single
    Dim arrFiltre As Variant, arrAG As Variant, r As Long, idxCol As Long
    Dim filtreCol As String, titre As String
    
    ' Déterminer filtreCol et titre selon filtreType
    Select Case filtreType
        Case "GC": filtreCol = "X": titre = "GC"
        Case "BM": filtreCol = "Y": titre = "BM"
        Case "EI": filtreCol = "Z": titre = "EI"
        Case Else: filtreCol = "": titre = "AG"
    End Select
    
    t0 = Timer
    Application.ScreenUpdating = False: Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    
    Set wbActuel = ThisWorkbook
    
    ' Feuille source
    On Error Resume Next
    Set wsSource = wbActuel.Worksheets("2.3-PP & SOW Annexe 1")
    On Error GoTo 0
    If wsSource Is Nothing Then GoTo CleanUp
    
    ' Classeur temporaire
    Set wbTemp = Workbooks.Add: Set wsFiltered = wbTemp.Worksheets(1)
    wsFiltered.Name = "Donnees_Filtrees_" & titre
    
    ' Déterminer bornes
    lastRow = LastUsedRowInRange(wsSource, IIf(lang = "FR", "L", "AG"))
    If lastRow < 3 Then lastRow = 501
    
    ' Copie brute
    If lang = "FR" Then
        Set rngSrc = wsSource.Range("B3:L" & lastRow)
    Else
        Set rngSrc = wsSource.Range("W3:AG" & lastRow)
    End If
    
    rngSrc.Copy
    With wsFiltered.Range("A1")
        .PasteSpecial xlPasteAll
        .PasteSpecial xlPasteColumnWidths
    End With
    Application.CutCopyMode = False
    
    ' Filtrage en mémoire
    currentLastRow = LastUsedRow(wsFiltered)
    If currentLastRow > 1 Then
        arrAG = wsFiltered.Range("K1:K" & currentLastRow).Value2
        
        If filtreCol <> "" Then
            ' Mode spécifique (GC, BM, EI) - avec filtre
            idxCol = ColIndex(filtreCol, lang)
            arrFiltre = wsFiltered.Range(ColLetter(idxCol) & "1:" & ColLetter(idxCol) & currentLastRow).Value2
            For r = currentLastRow To 2 Step -1
                If ShouldDeleteRow(arrAG, r, arrFiltre) Then wsFiltered.Rows(r).Delete
            Next r
        Else
            ' Mode Général - sans filtre
            For r = currentLastRow To 2 Step -1
                If ShouldDeleteRow(arrAG, r) Then wsFiltered.Rows(r).Delete
            Next r
        End If
    End If
    
    ' Définir la plage finale
    currentLastRow = LastUsedRow(wsFiltered)
    Set rngFinal = wsFiltered.Range("A1:K" & IIf(currentLastRow >= 1, currentLastRow, 1))
    
    ' Trouver le placeholder
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 1)"
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    
    If Not anchor.Find.Found Then
        Debug.Print "? Placeholder (Annexe 1) introuvable"
        GoTo CleanUp
    End If
    
    anchor.Text = ""
    anchor.Collapse 0
    
    ' Coller avec retry
    If Not CollerTableauAvecRetry(anchor, rngFinal, "Annexe 1") Then
        Debug.Print "? Échec définitif du collage Annexe 1"
        GoTo CleanUp
    End If
    
    ' Mise en forme Word
    On Error Resume Next
    Set wordTable = WordDoc.Tables(WordDoc.Tables.Count)
    If Not wordTable Is Nothing Then
        With wordTable
            .AllowAutoFit = False: .PreferredWidthType = 2: .PreferredWidth = 100
            .Range.ParagraphFormat.SpaceAfter = 0: .Rows.HeightRule = 1: .Rows.Height = 0
            .Range.Font.Size = 6
        End With
        Debug.Print "? Annexe 1 mise en forme appliquée"
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True: Application.DisplayAlerts = True: Application.Calculation = xlCalculationAutomatic

CleanUp:
    On Error Resume Next
    If Not wbTemp Is Nothing Then wbTemp.Close False
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
End Sub

Private Sub ExecuterAnnexe2(WordApp As Object, WordDoc As Object, nomLivrable As String, langue As String)
    Dim mode As String
    Dim colFiltre As Long
    
    ' Déterminer le mode et la colonne de filtrage
    If InStr(nomLivrable, "Général") > 0 Or InStr(nomLivrable, "General") > 0 Then
        mode = "AG"
        colFiltre = 0  ' Pas de filtre pour Général
    ElseIf InStr(nomLivrable, "GC") > 0 Or InStr(nomLivrable, "Civil Works") > 0 Then
        mode = "GC"
        colFiltre = 19  ' Colonne S
    ElseIf InStr(nomLivrable, "Modulaire") > 0 Or InStr(nomLivrable, "Modular") > 0 Then
        mode = "BM"
        colFiltre = 20  ' Colonne T
    ElseIf InStr(nomLivrable, "E&I") > 0 Then
        mode = "EI"
        colFiltre = 21  ' Colonne U
    End If
    
    ' Appeler ExecuterAnnexe2_Adapted avec WordDoc déjà ouvert
    Call ExecuterAnnexe2_Adapted(WordApp, WordDoc, mode, langue, colFiltre)
    
    ' PAUSE DE 3 SECONDES après collage
    Debug.Print "? Pause de 3 secondes après Annexe 2..."
    Application.Wait Now + TimeValue("0:00:03")
    
    ' VIDER LE PRESSE-PAPIER
    Call ViderPressePapier
End Sub

Private Sub ExecuterAnnexe3(WordApp As Object, WordDoc As Object, nomLivrable As String, langue As String)
    ' Appeler les 3 sous-annexes
    Call ExecuterAnnexe3a(WordApp, WordDoc, langue)
    Call ExecuterAnnexe3b(WordApp, WordDoc, langue)
    Call ExecuterAnnexe3c(WordApp, WordDoc, langue)
    
    ' PAUSE DE 3 SECONDES après TOUTE l'Annexe 3
    Debug.Print "? Pause de 3 secondes après Annexe 3 complète (3a+3b+3c)..."
    Application.Wait Now + TimeValue("0:00:03")
    
    ' VIDER LE PRESSE-PAPIER APRÈS TOUTE L'ANNEXE 3
    Call ViderPressePapier
End Sub

' ==========================================================================================================================================================
' ================================================================== PP/SOW 8002 ANNEXE 4 ==================================================================
' ==========================================================================================================================================================

' ==========================================================================================================================================================
' ================================================================== PP/SOW 8002 ANNEXE 4 ==================================================================
' ==========================================================================================================================================================



' =====================================================================
' PP / SOW 8002 - ANNEXE 4
' =====================================================================
' =====================================================================
' PP / SOW 8002 - ANNEXE 4 (sans redimensionnement)
' =====================================================================
Private Sub ExecuterAnnexe4(WordApp As Object, WordDoc As Object, _
                            nomLivrable As String, langue As String)
    On Error GoTo ErrAnnexe4
    
    Dim ws As Worksheet
    Dim rngScreenshot As Range
    Dim anchor As Object
    Dim anchorPos As Long
    Dim tentative As Long
    Dim succes As Boolean
    Dim shp As Shape
    Dim shpCount As Long
    Dim ils As Object
    
    Debug.Print "=== Début Annexe 4 ==="
    
    ' 1) Exécuter la macro de dessin Excel
    On Error Resume Next
    Application.Run "TSF_Layout_Array"
    If Err.Number <> 0 Then
        Debug.Print "Erreur TSF_Layout_Array : " & Err.Number & " - " & Err.Description
        Err.Clear
    Else
        Debug.Print "TSF_Layout_Array exécutée correctement"
    End If
    On Error GoTo ErrAnnexe4
    
    ' 2) Récupérer la feuille Annexe 4
    Set ws = ThisWorkbook.Worksheets("2.6-PP & SOW Annexe 4")
    If ws Is Nothing Then
        Debug.Print "Feuille 2.6-PP & SOW Annexe 4 introuvable"
        Exit Sub
    End If
    
    ws.Activate
    ws.Range("A1").Select
    
    ' Zone à capturer (adapter si besoin)
    Set rngScreenshot = ws.Range("E1:AZ26")
    
    ' Info debug : shapes présents dans la zone ?
    shpCount = 0
    For Each shp In ws.Shapes
        If Not Application.Intersect(shp.TopLeftCell, rngScreenshot) Is Nothing Then
            shpCount = shpCount + 1
        End If
    Next shp
    Debug.Print "Shapes dans la zone E1:AZ26 : " & shpCount
    
    ' 3) Trouver le placeholder (Annexe 4) dans Word
    Set anchor = WordDoc.Content
    With anchor.Find
        .ClearFormatting
        .Text = "(Annexe 4)"
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    
    If Not anchor.Find.Found Then
        Debug.Print "Placeholder (Annexe 4) introuvable"
        Exit Sub
    End If
    
    ' Remplacer le placeholder par le futur collage
    anchorPos = anchor.Start
    anchor.Text = ""
    anchor.Collapse 0
    
    WordApp.Activate
    succes = False
    
    ' 4) Tentatives de CopyPicture + Paste
    For tentative = 1 To 3
        
        Application.CutCopyMode = False
        
        ' 4.1 Copier la capture depuis Excel
        rngScreenshot.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        DoEvents
        
        ' 4.2 Repositionner le point d'insertion dans Word
        WordApp.Selection.SetRange anchorPos, anchorPos
        
        ' 4.3 Coller dans Word
        On Error Resume Next
        WordApp.Selection.Paste
        If Err.Number <> 0 Then
            Debug.Print "Tentative " & tentative & " - erreur collage : " & Err.Description
            Err.Clear
            On Error GoTo ErrAnnexe4
        Else
            ' Vérifier qu'on a bien une image collée
            If WordApp.Selection.InlineShapes.Count > 0 Then
                Set ils = WordApp.Selection.InlineShapes(1)
                succes = True
            ElseIf WordDoc.InlineShapes.Count > 0 Then
                Set ils = WordDoc.InlineShapes(WordDoc.InlineShapes.Count)
                succes = True
            End If
        End If
        On Error GoTo ErrAnnexe4
        
        If succes Then
            Debug.Print "Collage Annexe 4 OK à la tentative " & tentative
            
            ' Optionnel : centrer le paragraphe (sans redimensionner)
            If Not ils Is Nothing Then
                ' 1 = wdAlignParagraphCenter en late binding
                WordApp.Selection.ParagraphFormat.Alignment = 1
            End If
            
            Exit For
        Else
            Debug.Print "Tentative " & tentative & " : aucun InlineShape détecté"
            Application.Wait Now + TimeSerial(0, 0, 1)
        End If
    Next tentative
    
    If Not succes Then
        Debug.Print "Echec définitif : pas de collage Annexe 4 après 3 tentatives"
        MsgBox "Annexe 4 : la capture n'a pas pu être collée automatiquement." & vbCrLf & _
               "Vérifier le dessin sur la feuille '2.6-PP & SOW Annexe 4'.", _
               vbExclamation, "Annexe 4"
    End If
    
    ' Pause de 3 secondes, puis nettoyage presse-papiers
    Debug.Print "Pause de 3 secondes après Annexe 4..."
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    Application.CutCopyMode = False
    Call ViderPressePapier
    
    Debug.Print "=== Fin Annexe 4 ==="
    Exit Sub

ErrAnnexe4:
    Debug.Print "Erreur Annexe 4 : " & Err.Number & " - " & Err.Description
    Application.CutCopyMode = False
    Call ViderPressePapier
End Sub








' ==========================================================================================================================================================
' ================================================================== PS 8002 ==================================================================
' ==========================================================================================================================================================
Private Sub ExecuterPS_Adapted(ByVal mode As String, ByVal langue As String, ByVal feuilleSource As String, _
                               ByVal cheminTemplate As String, ByVal colC As Long, ByVal colY As Long, _
                               Optional ByVal colSep As Long = 0)
    Dim wbActuel As Workbook, wbTemplate As Workbook
    Dim wsSource As Worksheet, wsNew As Worksheet
    Dim rngSrc As Range
    Dim lastRow As Long, lastCol As Long
    Dim arrC As Variant, arrY As Variant, arrSep As Variant
    Dim r As Long, c As Long
    
    On Error GoTo GestionErreur
    
    Set wbActuel = ThisWorkbook
    
    ' Vérifie template
    If Dir(cheminTemplate) = "" Then
        MsgBox "Erreur: Le fichier template '" & cheminTemplate & "' est introuvable.", vbCritical
        Exit Sub
    End If
    
    ' Vérifie source
    On Error Resume Next
    Set wsSource = wbActuel.Worksheets(feuilleSource)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "Erreur: La feuille '" & feuilleSource & "' est introuvable dans ce classeur.", vbCritical
        Exit Sub
    End If
    
    ' Optimisations
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' Ouvre template
    Set wbTemplate = Workbooks.Open(cheminTemplate)
    
    ' (Re)crée Cfinal
    On Error Resume Next
    wbTemplate.Worksheets("Cfinal").Delete
    On Error GoTo 0
    Set wsNew = wbTemplate.Worksheets.Add(After:=wbTemplate.Sheets(wbTemplate.Sheets.Count))
    wsNew.Name = "Cfinal"
    
    ' Copie brute
    Set rngSrc = wsSource.UsedRange
    rngSrc.Copy
    With wsNew.Range("A1")
        .PasteSpecial Paste:=xlPasteAll
        .PasteSpecial Paste:=xlPasteColumnWidths
    End With
    Application.CutCopyMode = False
    
    ' Détermination bornes
    lastRow = LastUsedRow(wsNew)
    If lastRow < 2 Then GoTo PostClean
    lastCol = LastUsedCol(wsNew)
    
    ' Conversion en valeurs sauf col F
    For c = 1 To lastCol
        If c <> 6 Then
            With wsNew.Range(wsNew.Cells(1, c), wsNew.Cells(lastRow, c))
                .Value = .Value
            End With
        End If
    Next c
    
    ' Charge colonnes mémoire
    arrC = wsNew.Range(wsNew.Cells(1, colC), wsNew.Cells(lastRow, colC)).Value2
    
    If lastCol >= colY Then
        arrY = wsNew.Range(wsNew.Cells(1, colY), wsNew.Cells(lastRow, colY)).Value2
    Else
        ReDim arrY(1 To lastRow, 1 To 1)
        For r = 1 To lastRow: arrY(r, 1) = vbNullString: Next r
    End If
    
    If colSep > 0 Then
        If lastCol >= colSep Then
            arrSep = wsNew.Range(wsNew.Cells(1, colSep), wsNew.Cells(lastRow, colSep)).Value2
        Else
            ReDim arrSep(1 To lastRow, 1 To 1)
            For r = 1 To lastRow: arrSep(r, 1) = vbNullString: Next r
        End If
    End If
    
    ' Suppression bottom-up
    For r = lastRow To 2 Step -1
        If ShouldDeleteRow_PS(arrC, arrY, arrSep, r, langue, colSep) Then
            wsNew.Rows(r).Delete
        End If
    Next r

PostClean:
    wsNew.UsedRange.Rows.AutoFit
    
    ' Restauration
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Erreur " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ==========================================================================================================================================================
' ================================================================== FONCTIONS UTILITAIRES ==================================================================
' ==========================================================================================================================================================
Private Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find("*", , , , xlByRows, xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = f.Row
    End If
End Function

Private Function LastUsedRowInRange(ws As Worksheet, ColLetter As String) As Long
    LastUsedRowInRange = ws.Cells(ws.Rows.Count, ColLetter).End(xlUp).Row
End Function

Private Function LastUsedCol(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                         LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If f Is Nothing Then
        LastUsedCol = 1
    Else
        LastUsedCol = f.Column
    End If
End Function

Private Function ToStringNoNbsp(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        ToStringNoNbsp = ""
    Else
        ToStringNoNbsp = Replace(CStr(v), Chr(160), " ")
    End If
End Function

Private Function NormalizeForComparison(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Or IsEmpty(s) Then
        NormalizeForComparison = "": Exit Function
    End If
    t = Replace(CStr(s), Chr(160), " ")
    NormalizeForComparison = LCase$(Trim$(t))
End Function

Private Function NormalizeBasic(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Or IsEmpty(s) Then
        NormalizeBasic = ""
        Exit Function
    End If
    t = CStr(s)
    t = Replace(t, Chr(160), " ")
    t = Trim$(t)
    NormalizeBasic = LCase$(t)
End Function

Private Function IsNonNoOrNA(ByVal normalized As String) As Boolean
    IsNonNoOrNA = (normalized = "non" Or normalized = "no" Or normalized = "n/a")
End Function

Private Function ShouldDeleteRow(arrAG As Variant, r As Long, Optional arrFiltre As Variant) As Boolean
    Dim raw1 As String, rawAG As String, normAG As String
    
    rawAG = ToStringNoNbsp(arrAG(r, 1))
    normAG = NormalizeForComparison(arrAG(r, 1))
    
    If Not IsMissing(arrFiltre) Then
        raw1 = ToStringNoNbsp(arrFiltre(r, 1))
        If raw1 = " " Then
            ShouldDeleteRow = True: Exit Function
        End If
    End If
    
    ShouldDeleteRow = (rawAG = " ") Or IsNonNoOrNA(normAG)
End Function

Private Function ShouldDeleteRow_PS(arrC As Variant, arrY As Variant, arrSep As Variant, r As Long, _
                                   ByVal langue As String, ByVal colSep As Long) As Boolean
    Dim cRaw As String, yRaw As String, sRaw As String
    
    cRaw = NormalizeBasic(arrC(r, 1))
    yRaw = NormalizeBasic(arrY(r, 1))
    
    If colSep > 0 Then
        sRaw = ToStringNoNbsp(arrSep(r, 1))
    End If
    
    If langue = "FR" And cRaw = "eng" Then
        ShouldDeleteRow_PS = True: Exit Function
    End If
    If langue = "ENG" And cRaw = "fr" Then
        ShouldDeleteRow_PS = True: Exit Function
    End If
    If yRaw = "non utilisé" Or yRaw = "non utilise" Then
        ShouldDeleteRow_PS = True: Exit Function
    End If
    If colSep > 0 And sRaw = " " Then
        ShouldDeleteRow_PS = True: Exit Function
    End If
    
    ShouldDeleteRow_PS = False
End Function

Private Function ColIndex(ColLetter As String, lang As String) As Long
    Select Case ColLetter
        Case "X": ColIndex = 2
        Case "Y": ColIndex = 3
        Case "Z": ColIndex = 4
    End Select
End Function

Private Function ColLetter(idx As Long) As String
    ColLetter = Split(Cells(1, idx).Address(True, False), "$")(0)
End Function

Private Function GetOrCreateWordApp() As Object
    On Error Resume Next
    Set GetOrCreateWordApp = GetObject(, "Word.Application")
    If GetOrCreateWordApp Is Nothing Then
        Set GetOrCreateWordApp = CreateObject("Word.Application")
    End If
    If Not GetOrCreateWordApp Is Nothing Then
        GetOrCreateWordApp.Visible = True
    End If
End Function

Private Function TraiterLigneAnnexe2(ws As Worksheet, WordDoc As Object, InsertionRange As Object, ByVal r As Long, _
                                    ByVal COL_LANGUE As Long, ByVal COL_FLAG As Long, _
                                    ByVal COL_TITRE2 As Long, ByVal COL_TITRE3 As Long, ByVal COL_TITRE4 As Long, _
                                    ByVal COL_TEXTE As Long, ByVal COL_FILTRE As Long, _
                                    ByRef PrevTitre2 As String, ByRef PrevTitre3 As String, ByRef PrevTitre4 As String, _
                                    ByVal langueAttendue As String) As Boolean
    Dim langue As String, flag As String, t2 As String, t3 As String, t4 As String, txt As String
    Dim valeurFiltre As String
    
    langue = UCase(Trim(ws.Cells(r, COL_LANGUE).Value))
    flag = Trim(ws.Cells(r, COL_FLAG).Value)
    
    ' Filtre 1: Vérifier la langue
    If langue <> UCase(langueAttendue) Then Exit Function
    
    ' Filtre 2: Vérifier le flag "Utilisé"
    If Not CommenceParUtilise(flag) Then Exit Function
    
    ' Filtre 3: Vérifier la colonne spécifique (GC/BM/EI)
    ' Si COL_FILTRE > 0, on doit vérifier que la valeur = "X"
    If COL_FILTRE > 0 Then
        valeurFiltre = UCase(Trim(ws.Cells(r, COL_FILTRE).Value))
        ' Si la valeur n'est PAS "X", on ignore cette ligne
        If valeurFiltre <> "X" Then Exit Function
    End If
    
    t2 = Trim(ws.Cells(r, COL_TITRE2).Value)
    t3 = Trim(ws.Cells(r, COL_TITRE3).Value)
    t4 = Trim(ws.Cells(r, COL_TITRE4).Value)
    txt = Trim(ws.Cells(r, COL_TEXTE).Value)
    
    If t2 = "" And t3 = "" And t4 = "" And txt = "" Then Exit Function
    
    If t2 <> "" And t2 <> PrevTitre2 Then
        InsererTitre WordDoc, InsertionRange, t2, "Heading 2": PrevTitre2 = t2: PrevTitre3 = "": PrevTitre4 = ""
    End If
    If t3 <> "" And t3 <> PrevTitre3 Then
        InsererTitre WordDoc, InsertionRange, t3, "Heading 3": PrevTitre3 = t3: PrevTitre4 = ""
    End If
    If t4 <> "" And t4 <> PrevTitre4 Then
        InsererTitre WordDoc, InsertionRange, t4, "Heading 4": PrevTitre4 = t4
    End If
    If txt <> "" Then
        InsererTexte InsertionRange, txt
    End If
    
    TraiterLigneAnnexe2 = True
End Function

Private Function CommenceParUtilise(ByVal t As String) As Boolean
    Dim n As String
    n = LCase(Trim(t))
    n = Replace(n, "é", "e"): n = Replace(n, "è", "e")
    n = Replace(n, "ê", "e"): n = Replace(n, "ë", "e")
    CommenceParUtilise = (Left(n, 7) = "utilise")
End Function

Private Sub InsererTitre(ByVal WordDoc As Object, ByVal InsertionRange As Object, ByVal txt As String, ByVal styleNom As String)
    InsertionRange.InsertAfter txt & vbCr
    Dim r As Object: Set r = WordDoc.Range(InsertionRange.Start, InsertionRange.End)
    
    ' Appliquer le style de manière robuste (gère FR/EN automatiquement)
    Call AppliquerStyleRobuste(r, styleNom)
    
    InsertionRange.Collapse 0
End Sub

Private Sub InsererTexte(ByVal InsertionRange As Object, ByVal txt As String)
    InsertionRange.InsertAfter txt & vbCr
    InsertionRange.Collapse 0
End Sub

'==============================================================================
' GÉNÉRATION PS GLOBALE
'==============================================================================
Private Sub Generer_PS_Global(templatePath As String, nomLivrable As String, langue As String)
    Dim mode As String
    Dim feuilleSource As String
    Dim colC As Long, colY As Long, colSep As Long
    
    colC = 3
    colY = 25
    colSep = 0
    
    If InStr(nomLivrable, "Modulaire") > 0 Or InStr(nomLivrable, "Modular") > 0 Then
        mode = "BM"
        feuilleSource = "2.8-PS ITC Bâtiment Modulaire"
    ElseIf InStr(nomLivrable, "GC") > 0 Or InStr(nomLivrable, "Civil Works") > 0 Then
        mode = "GC"
        feuilleSource = "2.7-PS ITC Global"
        colSep = 27
    ElseIf InStr(nomLivrable, "E&I") > 0 Then
        mode = "EI"
        feuilleSource = "2.7-PS ITC Global"
        colSep = 29
    Else
        mode = "AG"
        feuilleSource = "2.7-PS ITC Global"
    End If
    
    Call ExecuterPS_Adapted(mode, langue, feuilleSource, templatePath, colC, colY, colSep)
End Sub

' ==========================================================================================================================================================
' ================================================================== PP/SOW 8002 ANNEXE 2 ==================================================================
' ==========================================================================================================================================================
Private Sub ExecuterAnnexe2_Adapted(WordApp As Object, WordDoc As Object, mode As String, langue As String, colFiltre As Long)
    Const COL_TITRE2 As Long = 6
    Const COL_TITRE3 As Long = 7
    Const COL_TITRE4 As Long = 8
    Const COL_TEXTE As Long = 15
    Const COL_LANGUE As Long = 17
    Const COL_FLAG As Long = 24
    
    Dim ExcelWS As Worksheet
    Dim InsertionRange As Object
    Dim PrevTitre2 As String, PrevTitre3 As String, PrevTitre4 As String
    Dim TotalRows As Long, InsertedRows As Long
    Dim t0 As Single: t0 = Timer
    
    Set ExcelWS = Nothing
    On Error Resume Next
    Set ExcelWS = ThisWorkbook.Worksheets("2.4-PP & SOW Annexe 2")
    On Error GoTo 0
    If ExcelWS Is Nothing Then Exit Sub
    
    Dim rng As Object
    Set rng = WordDoc.Content
    With rng.Find
        .ClearFormatting
        .Text = "(Annexe 2)"
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    
    If rng.Find.Found Then
        Set InsertionRange = rng
        InsertionRange.Text = ""
        InsertionRange.Collapse 0
    Else
        Exit Sub
    End If
    
    TotalRows = 0: InsertedRows = 0
    PrevTitre2 = "": PrevTitre3 = "": PrevTitre4 = ""
    
    Dim r As Long, lastRow As Long
    lastRow = WorksheetFunction.Min(672, ExcelWS.Cells(ExcelWS.Rows.Count, COL_LANGUE).End(xlUp).Row)
    
    Debug.Print "Annexe 2 - Mode: " & mode & " | ColFiltre: " & colFiltre & " | Langue: " & langue
    
    For r = 11 To lastRow
        TotalRows = TotalRows + 1
        ' Passage de colFiltre au lieu de colSelect
        If TraiterLigneAnnexe2(ExcelWS, WordDoc, InsertionRange, r, _
                              COL_LANGUE, COL_FLAG, COL_TITRE2, COL_TITRE3, COL_TITRE4, COL_TEXTE, colFiltre, _
                              PrevTitre2, PrevTitre3, PrevTitre4, langue) Then
            InsertedRows = InsertedRows + 1
        End If
        If r Mod 50 = 0 Then DoEvents
    Next r
    
    Debug.Print "Annexe 2 terminée - Lignes traitées: " & TotalRows & " | Lignes insérées: " & InsertedRows
End Sub

' ==========================================================================================================================================================
' ================================================================== PP/SOW 8002 ANNEXE 3a ==================================================================
' ==========================================================================================================================================================
Private Sub ExecuterAnnexe3a(WordApp As Object, WordDoc As Object, langue As String)
    Dim ws As Worksheet
    Dim anchor As Object
    Dim wordTable As Object
    Dim startCell As Range, endCell As Range
    Dim rStart As Long, rEnd As Long, cStart As Long, cEnd As Long
    Dim rngToCopy As Range
    
    Set ws = ThisWorkbook.Worksheets("1.5-Office Layout (INPUT Anx 3)")
    
    If langue = "FR" Then
        With ws.Cells
            Set startCell = .Find("5 LIGNES AVANT DEBUT ANNEXE 3 INPUT", LookAt:=xlWhole)
            Set endCell = .Find("2 LIGNES APRES FIN ANNEXE 3 INPUT", LookAt:=xlWhole)
        End With
        
        If startCell Is Nothing Or endCell Is Nothing Then
            Set rngToCopy = ws.Range("D14:G123")
        Else
            rStart = startCell.Row + 5
            rEnd = endCell.Row - 2
            cStart = startCell.Column
            cEnd = endCell.Column
            Set rngToCopy = ws.Range(ws.Cells(rStart, cStart), ws.Cells(rEnd, cEnd))
        End If
    Else
        With ws.Cells
            Set startCell = .Find("5 ROWS BEFORE START OF ANNEXE 3A", LookAt:=xlWhole)
            Set endCell = .Find("2 ROWS AFTER THE END OF ANNEXE 3A", LookAt:=xlWhole)
        End With
        
        If startCell Is Nothing Or endCell Is Nothing Then
            Set rngToCopy = ws.Range("AS14:AV123")
        Else
            rStart = startCell.Row + 5
            rEnd = endCell.Row - 2
            cStart = startCell.Column
            cEnd = endCell.Column
            Set rngToCopy = ws.Range(ws.Cells(rStart, cStart), ws.Cells(rEnd, cEnd))
        End If
    End If
    
    ' Trouver le placeholder
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3a)"
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    
    If Not anchor.Find.Found Then
        Debug.Print "? Placeholder (Annexe 3a) introuvable"
        Exit Sub
    End If
    
    anchor.Text = ""
    anchor.Collapse 0
    
    ' Coller avec retry (en passant la plage source)
    If Not CollerTableauAvecRetry(anchor, rngToCopy, "Annexe 3a") Then
        Debug.Print "? Échec définitif du collage Annexe 3a"
        Application.CutCopyMode = False
        Exit Sub
    End If
    
    ' Mise en forme
    On Error Resume Next
    Set wordTable = WordDoc.Tables(WordDoc.Tables.Count)
    If Not wordTable Is Nothing Then
        With wordTable
            .AllowAutoFit = False
            .PreferredWidthType = 2
            .PreferredWidth = 100
            .Range.ParagraphFormat.SpaceAfter = 0
            .Rows.HeightRule = 1
            .Rows.Height = 0
        End With
        wordTable.Range.Font.Size = 8
        Debug.Print "? Annexe 3a mise en forme appliquée"
    Else
        Debug.Print "?? Impossible de récupérer le tableau pour mise en forme"
    End If
    On Error GoTo 0
    
    Application.CutCopyMode = False
End Sub

' ==========================================================================================================================================================
' ================================================================== PP/SOW 8002 ANNEXE 3b ==================================================================
' ==========================================================================================================================================================
Private Sub ExecuterAnnexe3b(WordApp As Object, WordDoc As Object, langue As String)
    Dim ws As Worksheet
    Dim anchor As Object
    Dim i As Long, t0 As Single
    Dim lignesData() As LigneInfo, dataRange As Variant
    Dim cellDebut As Range, cellFin As Range
    Dim LIGNE_DEBUT As Long, LIGNE_FIN As Long, COLONNE_DEBUT As Long, COLONNE_FIN As Long
    
    On Error GoTo ErreurGlobale
    
    t0 = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set ws = ThisWorkbook.Worksheets("2.5-PP & SOW Annexe 3")
    If ws Is Nothing Then GoTo CleanUp
    
    With ws.Cells
        If langue = "FR" Then
            Set cellDebut = .Find("Cellule 6 Lignes Avant Premiere Cellule Range Annexe 3b", LookAt:=xlWhole)
            Set cellFin = .Find("Cellule 2 Lignes Après Derniere Cellule Range Annexe 3b", LookAt:=xlWhole)
        Else
            Set cellDebut = .Find("cell 6 rows before start of range of annexe 3b", LookAt:=xlWhole)
            Set cellFin = .Find("cell 2 rows after end of range of annexe 3b", LookAt:=xlWhole)
        End If
    End With
    
    If cellDebut Is Nothing Or cellFin Is Nothing Then GoTo CleanUp
    
    LIGNE_DEBUT = cellDebut.Row + 6
    LIGNE_FIN = cellFin.Row - 2
    COLONNE_DEBUT = cellDebut.Column
    COLONNE_FIN = cellFin.Column
    
    If COLONNE_DEBUT = COLONNE_FIN Then COLONNE_FIN = COLONNE_DEBUT + 4
    
    If LIGNE_DEBUT >= LIGNE_FIN Or COLONNE_DEBUT >= COLONNE_FIN Then GoTo CleanUp
    
    On Error Resume Next
    ReDim lignesData(LIGNE_DEBUT To LIGNE_FIN)
    dataRange = ws.Range(ws.Cells(LIGNE_DEBUT, COLONNE_DEBUT), ws.Cells(LIGNE_FIN, COLONNE_FIN)).Value
    On Error GoTo ErreurGlobale
    
    For i = LIGNE_DEBUT To LIGNE_FIN
        PretraiterLigne3b i, dataRange, lignesData(i), LIGNE_DEBUT, COLONNE_DEBUT, COLONNE_FIN
    Next i
    
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3b)"
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    
    If Not anchor.Find.Found Then GoTo CleanUp
    
    anchor.Text = ""
    anchor.Collapse 0
    
    Dim nbTitres As Long, nbSousTitres As Long, nbTableaux As Long
    Dim debutBloc As Long, finBloc As Long
    Dim blocsTableaux As Collection
    Set blocsTableaux = New Collection
    
    i = LIGNE_DEBUT
    Do While i <= LIGNE_FIN
        If i > UBound(lignesData) Then Exit Do
        
        If lignesData(i).EstVide Then
            i = i + 1
        ElseIf lignesData(i).EstTitre Then
            AjouterTitre3b WordDoc, anchor, lignesData(i).ValeurAA, "Heading 3"
            nbTitres = nbTitres + 1
            i = i + 1
        ElseIf lignesData(i).EstSousTitre Then
            AjouterTitre3b WordDoc, anchor, lignesData(i).ValeurAB, "Heading 4"
            nbSousTitres = nbSousTitres + 1
            i = i + 1
        ElseIf lignesData(i).EstTableau Then
            debutBloc = i
            Do While i <= LIGNE_FIN And i <= UBound(lignesData)
                If Not lignesData(i).EstTableau Then Exit Do
                i = i + 1
            Loop
            finBloc = i - 1
            Call ExporterBlocTableauOptimise(ws, WordDoc, WordApp, anchor, debutBloc, finBloc, _
                                            COLONNE_DEBUT, COLONNE_FIN, nbTableaux)
            nbTableaux = nbTableaux + 1
        Else
            i = i + 1
        End If
    Loop

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Exit Sub

ErreurGlobale:
    GoTo CleanUp
End Sub

Private Sub ExporterBlocTableauOptimise(ws As Worksheet, WordDoc As Object, WordApp As Object, _
                                       ByRef anchor As Object, ByVal ligneDebut As Long, _
                                       ByVal ligneFin As Long, ByVal COLONNE_DEBUT As Long, _
                                       ByVal COLONNE_FIN As Long, ByVal numTableau As Long)
    On Error GoTo ErreurTableau
    Dim wordTable As Object
    Dim colonneCopieDebut As Long
    Dim rngSource As Range
    
    colonneCopieDebut = COLONNE_DEBUT + 2
    If colonneCopieDebut > COLONNE_FIN Then Exit Sub
    
    ' Définir la plage source
    Set rngSource = ws.Range(ws.Cells(ligneDebut, colonneCopieDebut), ws.Cells(ligneFin, COLONNE_FIN))
    
    ' Coller avec retry (en passant la plage source)
    If Not CollerTableauAvecRetry(anchor, rngSource, "Annexe 3b - Tableau #" & (numTableau + 1)) Then
        Debug.Print "? Échec définitif du collage tableau Annexe 3b #" & (numTableau + 1)
        Application.CutCopyMode = False
        Exit Sub
    End If
    
    ' Mise en forme
    On Error Resume Next
    Set wordTable = WordDoc.Tables(WordDoc.Tables.Count)
    If Not wordTable Is Nothing Then
        With wordTable
            .AllowAutoFit = False
            .PreferredWidthType = 2
            .PreferredWidth = 100
            .Range.Font.Size = 8
        End With
        anchor.SetRange wordTable.Range.End, wordTable.Range.End
        anchor.InsertParagraphAfter
        anchor.Collapse 0
    End If
    On Error GoTo 0
    
    Application.CutCopyMode = False
    
    Exit Sub

ErreurTableau:
    Debug.Print "Erreur export tableau: " & Err.Description
    Application.CutCopyMode = False
End Sub

Private Sub PretraiterLigne3b(ByVal ligne As Long, dataRange As Variant, ByRef info As LigneInfo, _
                             ByVal LIGNE_DEBUT As Long, ByVal COLONNE_DEBUT As Long, ByVal COLONNE_FIN As Long)
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
    Debug.Print "Erreur prétraitement ligne " & ligne & ": " & Err.Description
    info.EstVide = True
End Sub

Private Sub AjouterTitre3b(WordDoc As Object, ByRef anchor As Object, ByVal texte As String, ByVal styleNom As String)
    On Error GoTo ErreurTitre
    
    If Len(Trim$(texte)) = 0 Then Exit Sub
    
    Dim rng As Object
    Set rng = WordDoc.Range(anchor.Start, anchor.Start)
    rng.Text = texte & vbCr
    
    ' Appliquer le style de manière robuste (gère FR/EN automatiquement)
    Call AppliquerStyleRobuste(rng, styleNom)
    
    anchor.SetRange rng.End, rng.End
    Exit Sub

ErreurTitre:
    Debug.Print "Erreur ajout titre '" & texte & "': " & Err.Description
End Sub

' ==========================================================================================================================================================
' ================================================================== PP/SOW 8002 ANNEXE 3c ==================================================================
' ==========================================================================================================================================================
Private Sub ExecuterAnnexe3c(WordApp As Object, WordDoc As Object, langue As String)
    Dim ws As Worksheet
    Dim anchor As Object
    Dim wordTable As Object
    Dim startCell As Range, endCell As Range
    Dim rStart As Long, rEnd As Long, cStart As Long, cEnd As Long
    Dim rngToCopy As Range
    
    Set ws = ThisWorkbook.Worksheets("2.5-PP & SOW Annexe 3")
    
    If langue = "FR" Then
        With ws.Cells
            Set startCell = .Find("4 Lignes au dessus de debut Annexe 3c", LookAt:=xlWhole)
            Set endCell = .Find("Cellule 4 Lignes Après Dernière Cellule Range Annexe 3c", LookAt:=xlWhole)
        End With
    Else
        With ws.Cells
            Set startCell = .Find("cell 4 rows before start of range of annexe 3c", LookAt:=xlWhole)
            Set endCell = .Find("cell 4 rows after end of range of annexe 3c", LookAt:=xlWhole)
        End With
    End If
    
    If Not startCell Is Nothing And Not endCell Is Nothing Then
        rStart = startCell.Row + 4
        rEnd = endCell.Row - 4
        cStart = startCell.Column
        cEnd = endCell.Column
        
        If cStart > cEnd Then
            Dim tmp As Long: tmp = cStart: cStart = cEnd: cEnd = tmp
        End If
        
        If rStart <= rEnd Then
            Set rngToCopy = ws.Range(ws.Cells(rStart, cStart), ws.Cells(rEnd, cEnd))
        Else
            Debug.Print "? Plage invalide pour Annexe 3c"
            Exit Sub
        End If
    Else
        Debug.Print "? Markers introuvables pour Annexe 3c"
        Exit Sub
    End If
    
    ' Trouver le placeholder
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3c)"
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    
    If Not anchor.Find.Found Then
        Debug.Print "? Placeholder (Annexe 3c) introuvable"
        Exit Sub
    End If
    
    anchor.Text = ""
    anchor.Collapse 0
    
    ' Coller avec retry (en passant la plage source)
    If Not CollerTableauAvecRetry(anchor, rngToCopy, "Annexe 3c") Then
        Debug.Print "? Échec définitif du collage Annexe 3c"
        Application.CutCopyMode = False
        Exit Sub
    End If
    
    ' Mise en forme
    On Error Resume Next
    Set wordTable = WordDoc.Tables(WordDoc.Tables.Count)
    If Not wordTable Is Nothing Then
        With wordTable
            .AllowAutoFit = False
            .PreferredWidthType = 2
            .PreferredWidth = 100
            .Range.ParagraphFormat.SpaceAfter = 0
            .Rows.HeightRule = 1
            .Rows.Height = 0
            .Range.Font.Size = 8
        End With
        Debug.Print "? Annexe 3c mise en forme appliquée"
    Else
        Debug.Print "?? Impossible de récupérer le tableau pour mise en forme"
    End If
    On Error GoTo 0
    
    Application.CutCopyMode = False
End Sub

'```

'**? Modifications apportées:**

'### ?? **Pause de 3 secondes intégrée:**

'1. **ExecuterAnnexe1** - Pause **APRÈS** le collage, **AVANT** ViderPressePapier
'2. **ExecuterAnnexe2** - Pause **APRÈS** le collage, **AVANT** ViderPressePapier
'3. **ExecuterAnnexe3** - Pause **APRÈS** toutes les sous-annexes (3a+3b+3c), **AVANT** ViderPressePapier
'4. **ExecuterAnnexe4** - Pause **APRÈS** le collage, **AVANT** ViderPressePapier

'### ?? **Séquence exacte:**
'```
'1. Collage du contenu ?
'2. ? Pause 3 secondes (Application.Wait)
'3. ?? Vidage du presse-papier



