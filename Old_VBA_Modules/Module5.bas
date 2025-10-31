Attribute VB_Name = "Module5"
Sub InsererTableauDansWord()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim ExcelWS As Object
    Dim tableData As Variant
    Dim wordTable As Object
    Dim i As Long, j As Long

    ' Initialisation
    Set ExcelWS = ThisWorkbook.Sheets("2.5-PP & SOW Annexe 3")

    ' Instance Word
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0

    ' Ouvrir le document Word (template dans le même dossier)
    Set WordDoc = WordApp.Documents.Open(ThisWorkbook.Path & "\PP-8002-FR.dotx")

    ' Récupération des données Excel (exemple simplifié)
    tableData = ExcelWS.Range("A1:C5").Value

    ' Création du tableau Word
    Set wordTable = WordDoc.Tables.Add(WordDoc.Range(0, 0), UBound(tableData, 1), UBound(tableData, 2))

    ' Remplissage des données
    For i = 1 To UBound(tableData, 1)
        For j = 1 To UBound(tableData, 2)
            wordTable.cell(i, j).Range.Text = tableData(i, j)
        Next j
    Next i

    ' ==========================
    ' ?? Formatage demandé ici
    ' ==========================
    Dim tgtNames As Variant, sty As Object, s As Object
    Dim nameLocal As String, k As Long
    Dim isTableStyle As Boolean

    ' 1) Tenter d’appliquer le style "Text in table"
    '    - On cherche d’abord un style de TABLE du nom "Text in table"/"Texte de tableau"
    '    - Sinon, on cherche un style de PARAGRAPHE du même nom pour l’appliquer cellule par cellule
    tgtNames = Array("Text in table", "Texte de tableau", "Text in Table", "Texte dans le tableau")

    Set sty = Nothing
    isTableStyle = False

    ' Recherche d'un style de table (Type = 3) par NameLocal
    On Error Resume Next
    For Each s In WordDoc.Styles
        nameLocal = LCase$(s.nameLocal)
        For k = LBound(tgtNames) To UBound(tgtNames)
            If nameLocal = LCase$(tgtNames(k)) Then
                If s.Type = 3 Then ' wdStyleTypeTable = 3
                    Set sty = s: isTableStyle = True
                    Exit For
                End If
            End If
        Next k
        If Not sty Is Nothing Then Exit For
    Next s
    On Error GoTo 0

    ' Si trouvé en tant que style de TABLE ? appliquer au tableau
    If Not sty Is Nothing And isTableStyle Then
        On Error Resume Next
        wordTable.Style = sty
        On Error GoTo 0
    Else
        ' Sinon, chercher un style de PARAGRAPHE de ce nom
        Set sty = Nothing
        On Error Resume Next
        For Each s In WordDoc.Styles
            nameLocal = LCase$(s.nameLocal)
            For k = LBound(tgtNames) To UBound(tgtNames)
                If nameLocal = LCase$(tgtNames(k)) Then
                    If s.Type = 1 Or s.Type = 4 Then ' wdStyleTypeParagraph=1, wdStyleTypeLinked=4
                        Set sty = s
                        Exit For
                    End If
                End If
            Next k
            If Not sty Is Nothing Then Exit For
        Next s
        On Error GoTo 0

        ' Si aucun style "Text in table" n'existe, on le crée en tant que style de paragraphe
        If sty Is Nothing Then
            On Error Resume Next
            Set sty = WordDoc.Styles.Add("Text in table", 1) ' 1 = wdStyleTypeParagraph
            On Error GoTo 0
        End If

        ' Appliquer le style de paragraphe cellule par cellule (sécurisé)
        If Not sty Is Nothing Then
            For i = 1 To wordTable.Rows.Count
                For j = 1 To wordTable.Columns.Count
                    On Error Resume Next
                    wordTable.cell(i, j).Range.Style = sty
                    On Error GoTo 0
                Next j
            Next i
        End If
    End If

    ' 2) Première colonne : gris clair + gras (application cellule par cellule)
    For i = 1 To wordTable.Rows.Count
        On Error Resume Next
        With wordTable.cell(i, 1)
            .Shading.BackgroundPatternColor = RGB(192, 192, 192) ' Remplissage
            .Range.Font.Bold = True                              ' Gras
        End With
        On Error GoTo 0
    Next i

End Sub
