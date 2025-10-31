Attribute VB_Name = "Module39"
Option Explicit

' === Constantes Word (late binding) ===
Private Const wdCollapseEnd As Long = 0
Private Const wdPasteRTF As Long = 1
Private Const wdStory As Long = 6

' ---------- D�tection format mixte (B/I/U) ----------
' Renvoie True si la cellule contient un m�lange de B/I/U (non uniformes),
' ce qui impose un collage RTF pour pr�server les formats internes.
Private Function HasMixedBIU(ByVal xlCell As Range) As Boolean
    Dim b As Variant, it As Variant, u As Variant
    b = xlCell.Font.Bold
    it = xlCell.Font.Italic
    u = xlCell.Font.Underline
    HasMixedBIU = (IsNull(b) Or IsNull(it) Or IsNull(u))
End Function

' ---------- Append dans une plage cible (rapide) ----------
' - Ins�re le texte � la fin de targetRange.
' - Applique un style de PARAGRAPHE (stylePara) SUR LA PORTION AJOUT�E UNIQUEMENT.
' - Option isTitle: convertit les retours Excel en sauts manuels Word (Chr(11)) pour
'   �viter de cr�er plusieurs paragraphes de titre (=> un seul � Titre 2/3/4 �).
' - applyUniformBIU: si la cellule est uniform�ment B/I/U, applique ce format au bloc.
Private Sub AppendPlainToRange( _
    ByVal xlCell As Range, _
    ByVal targetRange As Object, _
    ByVal stylePara As String, _
    Optional ByVal applyUniformBIU As Boolean = True, _
    Optional ByVal isTitle As Boolean = False)

    Dim s As String
    Dim vU As Variant
    Dim rng As Object
    Dim startBefore As Long, endAfter As Long

    s = CStr(xlCell.Value2)
    If Len(s) = 0 Then Exit Sub

    ' --- TITRE : garder un seul paragraphe (convertir CR/LF en saut manuel ^l) ---
    If isTitle Then
        s = Replace$(s, vbCrLf, vbLf)
        s = Replace$(s, vbCr, vbLf)
        s = Replace$(s, vbLf, Chr$(11)) ' ^l (saut de ligne manuel Word)
    End If

    ' Ins�rer � la FIN de la plage cible (qui sert de point d'insertion �volutif)
    targetRange.Collapse wdCollapseEnd
    startBefore = targetRange.Start
    targetRange.InsertAfter s & vbCr  ' on ajoute 1 seul � � la fin du bloc ajout�

    ' Styliser UNIQUEMENT la portion qu�on vient d�ajouter
    endAfter = startBefore + Len(s) + 1
    Set rng = targetRange.Document.Range(startBefore, endAfter)
    rng.Style = stylePara   ' style de paragraphe ; n��crase pas les formats caract�re existants

    ' Appliquer B/I/U UNIFORMES si la cellule Excel est uniform�ment format�e
    If applyUniformBIU Then
        If Not IsNull(xlCell.Font.Bold) Then rng.Font.Bold = IIf(CBool(xlCell.Font.Bold), -1, 0)
        If Not IsNull(xlCell.Font.Italic) Then rng.Font.Italic = IIf(CBool(xlCell.Font.Italic), -1, 0)
        vU = xlCell.Font.Underline
        If Not IsNull(vU) Then
            If (vU = -4142 Or vU = 0) Then
                rng.Font.Underline = 0          ' wdUnderlineNone
            Else
                rng.Font.Underline = 1          ' wdUnderlineSingle
            End If
        End If
    End If

    ' Avancer le point d'insertion global (fin = nouvelle position de targetRange)
    targetRange.SetRange endAfter, endAfter
End Sub

' ---------- Coller via RTF dans une plage cible ----------
' Utilis� pour les cellules texte (col. O) qui contiennent du formatage mixte.
' - Colle en RTF pour pr�server gras/italique/soulignage/surlignage internes.
' - Applique ensuite le style de PARAGRAPHE "Normal" sur la portion coll�e,
'   ce qui ne d�truit pas les formats au niveau caract�re.
Private Sub PasteCellRTFToRange( _
    ByVal xlCell As Range, _
    ByVal appWord As Object, _
    ByVal targetRange As Object, _
    ByVal stylePara As String)

    Dim startBefore As Long, endAfter As Long
    Dim rng As Object

    If Len(CStr(xlCell.Value2)) = 0 Then Exit Sub

    ' Se positionner � la fin de la plage cible et m�moriser le d�but du bloc
    targetRange.Collapse wdCollapseEnd
    startBefore = targetRange.Start
    targetRange.Select

    ' Collage RTF (pr�serve formats internes) + un � s�parateur
    xlCell.Copy
    appWord.Selection.PasteSpecial DataType:=wdPasteRTF
    appWord.Selection.TypeParagraph

    ' Fin du bloc coll� = position actuelle de la s�lection
    endAfter = appWord.Selection.Range.End

    ' Styliser UNIQUEMENT ce qui vient d'�tre coll�
    Set rng = targetRange.Document.Range(startBefore, endAfter)
    rng.Style = stylePara

    ' Avancer le point d'insertion global
    targetRange.SetRange endAfter, endAfter
End Sub

' ---------- Texte : choisir voie rapide ou RTF ----------
' Si la cellule texte a des formats mixtes ? RTF ; sinon ? insertion simple + formats uniformes.
Private Sub AppendTextPreservingBIUToRange( _
    ByVal xlCell As Range, _
    ByVal appWord As Object, _
    ByVal targetRange As Object, _
    ByVal stylePara As String)

    If HasMixedBIU(xlCell) Then
        PasteCellRTFToRange xlCell, appWord, targetRange, stylePara
    Else
        AppendPlainToRange xlCell, targetRange, stylePara, True
    End If
End Sub

' ---------- MACRO PRINCIPALE ----------
Public Sub PP8002_Annexe2()

    Dim appWord As Object, docWord As Object
    Dim ws As Worksheet
    Dim i As Long
    Dim langue As String, flagUtilise As String

    ' M�moires des DERNIERS titres affich�s pour �VITER les r�p�titions
    ' (si F/G/H r�p�tent la m�me valeur sur la ligne suivante, on ne r�-affiche pas le m�me titre)
    Dim Prec2 As String, Prec3 As String, Prec4 As String, prec5 As String

    Dim cT2 As Range, cT3 As Range, cT4 As Range, cT5 As Range, cTX As Range, cTxt As Range, cLang As Range
    Dim startTime As Single, endTime As Single, elapsedTime As Single
    Dim targetRange As Object

    ' === Start Timer (pour info) ===
    startTime = Timer

    ' Colonnes (SOW Annexe 1 + PP8002 Annexe 2)
    Const COL_T2 As Long = 6    ' F ? Titre 2
    Const COL_T3 As Long = 7    ' G ? Titre 3
    Const COL_T4 As Long = 8    ' H ? Titre 4
    Const COL_T5 As Long = 9    ' I (ignor� ici)
    Const COL_TXT As Long = 15  ' O ? Texte (style Normal)
    Const COL_LANG As Long = 17 ' Q ? Langue ("FR")
    Const COL_USE As Long = 24  ' X ? "Utilis�..."

    ' Boost Excel (perf)
    Dim calcMode As XlCalculation, ev As Boolean, sb As Variant
    calcMode = Application.Calculation: Application.Calculation = xlCalculationManual
    ev = Application.EnableEvents: Application.EnableEvents = False
    sb = Application.DisplayStatusBar: Application.DisplayStatusBar = False

    On Error GoTo CleanExit

    ' === OUVERTURE DU MOD�LE WORD (aucune sauvegarde !) ===
    Dim fso As Object, cheminExcel As String, dossier As String, cheminModele As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    cheminExcel = ThisWorkbook.FullName
    dossier = fso.GetParentFolderName(cheminExcel)
    cheminModele = dossier & "\PP-8002-FR.dotx"

    If Not fso.FileExists(cheminModele) Then
        MsgBox "Mod�le Word introuvable : " & cheminModele, vbCritical
        GoTo CleanExit
    End If

    On Error Resume Next
    Set appWord = CreateObject("Word.Application")
    On Error GoTo CleanExit
    If appWord Is Nothing Then
        MsgBox "Impossible d'ouvrir Word.", vbCritical
        GoTo CleanExit
    End If
    appWord.Visible = True
    Set docWord = appWord.Documents.Open(cheminModele) ' on OUVRE le .dotx (pas de Save ici)

    ' === TROUVER le texte litt�ral "(Annexe 2)" et se placer EXACTEMENT l� ===
    ' On ne touche � rien d�autre dans le document.
    Dim findRange As Object
    Set findRange = docWord.Content

    With findRange.Find
        .ClearFormatting
        .Text = "(Annexe 2)"
        .Forward = True
        .Wrap = 0 ' wdFindStop
        If Not .Execute Then
            MsgBox "Texte '(Annexe 2)' introuvable dans le mod�le.", vbExclamation
            GoTo CleanExit
        End If
    End With

    ' Remplacer ce seul marqueur par une plage vide qui servira de point d'insertion �volutif
    findRange.Text = ""
    Set targetRange = findRange  ' cette range sera �tendue au fur et � mesure des insertions

    ' === FEUILLE SOURCE ===
    Set ws = ThisWorkbook.Sheets("2.4-PP & SOW Annexe 2")
    If ws Is Nothing Then
        MsgBox "Feuille introuvable.", vbCritical
        GoTo CleanExit
    End If

    ' === PARCOURS DES LIGNES ===
    For i = 11 To 672

        Set cT2 = ws.Cells(i, COL_T2)
        Set cT3 = ws.Cells(i, COL_T3)
        Set cT4 = ws.Cells(i, COL_T4)
        Set cT5 = ws.Cells(i, COL_T5)
        Set cTxt = ws.Cells(i, COL_TXT)
        Set cLang = ws.Cells(i, COL_LANG)
        Set cTX = ws.Cells(i, COL_USE)

        ' --- FILTRES ---
        ' 1) Langue = FR (col. Q)
        langue = Trim$(CStr(cLang.Value2))
        If UCase$(langue) <> "FR" Then GoTo NextRow

        ' 2) X commence par "Utilis�..." (tol�rance accents)
        flagUtilise = LCase$(Replace(Replace(Trim$(CStr(cTX.Value2)), "�", "e"), "�", "e"))
        If Left$(flagUtilise, 7) <> "utilise" Then GoTo NextRow

        ' 3) Ligne enti�rement vide sur F:O ? ignorer
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, COL_T2), ws.Cells(i, COL_TXT))) = 0 Then GoTo NextRow

        ' --- TITRES (F/G/H) ---
        ' Logique de non-r�p�tition :
        '  � Si T2 change ? on ins�re T2 et on r�initialise T3/T4 (pour respecter la hi�rarchie).
        '  � Si T3 change ? on ins�re T3 et on r�initialise T4.
        '  � T4 s�ins�re � chaque changement de valeur.
        ' Cette logique permet d�obtenir a1 ? b1 ? c1 ? c2 ? b2 � comme demand�.
        If Trim$(CStr(cT2.Value2)) <> "" And Trim$(CStr(cT2.Value2)) <> Prec2 Then
            AppendPlainToRange cT2, targetRange, "Titre 2", False, True   ' True => garder un seul paragraphe de titre
            Prec2 = Trim$(CStr(cT2.Value2))
            Prec3 = "": Prec4 = "": prec5 = ""
        End If

        If Trim$(CStr(cT3.Value2)) <> "" And Trim$(CStr(cT3.Value2)) <> Prec3 Then
            AppendPlainToRange cT3, targetRange, "Titre 3", False, True
            Prec3 = Trim$(CStr(cT3.Value2))
            Prec4 = "": prec5 = ""
        End If

        If Trim$(CStr(cT4.Value2)) <> "" And Trim$(CStr(cT4.Value2)) <> Prec4 Then
            AppendPlainToRange cT4, targetRange, "Titre 4", False, True
            Prec4 = Trim$(CStr(cT4.Value2))
            ' prec5 inutilis� ici, laiss� pour compat si besoin
        End If

        ' --- TEXTE (col. O) ---
        ' Si la cellule O a un mix de formats (gras/non-gras, etc.) ? collage RTF + Style "Normal" SUR LE BLOC,
        ' ce qui pr�serve les formats internes (gras/italique/soulign�/surlignage).
        ' Sinon ? insertion simple + application des formats uniformes si pr�sents.
        If Trim$(CStr(cTxt.Value2)) <> "" Then
            AppendTextPreservingBIUToRange cTxt, appWord, targetRange, "Normal"
        End If

NextRow:
        ' pas de DoEvents pour la perf
    Next i

    ' === Fin === (AUCUNE sauvegarde automatique, le document reste ouvert)
    endTime = Timer
    elapsedTime = endTime - startTime

    MsgBox "OK � Insertion � la place de '(Annexe 2)' effectu�e." & vbCrLf & _
           "� Titres : F?Titre 2, G?Titre 3, H?Titre 4 (sans r�p�tition, avec hi�rarchie)." & vbCrLf & _
           "� Texte (O) : style 'Normal' appliqu�, formats internes B/I/U/surlignage conserv�s." & vbCrLf & _
           "� Le reste du mod�le n'a pas �t� modifi�. Aucune sauvegarde automatique." & vbCrLf & _
           "Temps �coul� : " & Format(elapsedTime, "0.00") & " s.", vbInformation

CleanExit:
    On Error Resume Next
    Application.Calculation = calcMode
    Application.EnableEvents = ev
    Application.DisplayStatusBar = sb
End Sub




