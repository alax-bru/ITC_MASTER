Attribute VB_Name = "Module18"
Option Explicit
' =====================================================================
'  PP8002_Annexe1_PasteFast – Export Excel -> Word (Annexe 1)
' =====================================================================
Public Sub PP_SOW_8002_FR_Annexe_1()
    Dim ws As Worksheet
    Dim WordApp As Object, WordDoc As Object
    Dim cheminDoc As String
    Dim anchor As Object
    Dim t0 As Single
    Dim wordTable As Object
    
    t0 = Timer
    Set ws = ThisWorkbook.Worksheets("2.3-PP & SOW Annexe 1")

    ' --- Copier la plage Excel ---
    ws.Range("B3:K501").Copy

    ' --- Ouvrir Word ---
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

    ' --- Trouver le repère "(Annexe 1)" ---
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 1)"
        .Forward = True
        .Wrap = 1 ' wdFindStop
        .Execute
    End With

    If Not anchor.Find.Found Then
        MsgBox "Repère '(Annexe 1)' introuvable dans " & cheminDoc, vbCritical
        Exit Sub
    End If

    anchor.Text = ""
    anchor.Collapse 0   ' wdCollapseEnd

    ' --- Coller avec format source ---
    anchor.PasteExcelTable False, False, False

    ' --- Récupérer le tableau Word collé ---
    Set wordTable = WordDoc.Tables(WordDoc.Tables.Count)

    ' --- Ajuster le tableau ---
    With wordTable
        .AllowAutoFit = False
        .PreferredWidthType = 2  ' wdPreferredWidthPercent
        .PreferredWidth = 100
        .Range.ParagraphFormat.SpaceAfter = 0
        ' === Hauteur automatique (gère aussi vertical) ===
        .Rows.HeightRule = 1   ' wdRowHeightAuto
        .Rows.Height = 0
    End With

    wordTable.Range.Font.Size = 6

    MsgBox "Annexe 1 exportée (B2:K501) dans " & cheminDoc & vbCrLf & _
           "Temps : " & Format(Timer - t0, "0.00") & " s", vbInformation
End Sub

' ===============================================================================================
' ===============================================================================================
'                                      ANNEXE 3
' ===============================================================================================
' ===============================================================================================


' =====================================================================
'                            Annexe 3a
' =====================================================================
Public Sub PP_SOW_8002_FR_Annexe_3a()
    Dim ws As Worksheet
    Dim WordApp As Object, WordDoc As Object
    Dim cheminDoc As String
    Dim anchor As Object
    Dim t0 As Single
    Dim wordTable As Object
    
    t0 = Timer
    Set ws = ThisWorkbook.Worksheets("1.5-Office Layout (INPUT Anx 3)")

    ' --- Copier la plage Excel ---
    ws.Range("D14:G123").Copy

    ' --- Ouvrir Word ---
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

    ' --- Trouver le repère "(Annexe 3a)" ---
    Set anchor = WordDoc.Content
    With anchor.Find
        .Text = "(Annexe 3a)"
        .Forward = True
        .Wrap = 1 ' wdFindStop
        .Execute
    End With

    If Not anchor.Find.Found Then
        MsgBox "Repère '(Annexe 3a)' introuvable dans " & cheminDoc, vbCritical
        Exit Sub
    End If

    anchor.Text = ""
    anchor.Collapse 0   ' wdCollapseEnd

    ' --- Coller avec format source ---
    anchor.PasteExcelTable False, False, False

    ' --- Récupérer le tableau Word collé ---
    Set wordTable = WordDoc.Tables(WordDoc.Tables.Count)

    ' --- Ajuster le tableau ---
    With wordTable
        .AllowAutoFit = False
        .PreferredWidthType = 2
        .PreferredWidth = 100
        .Range.ParagraphFormat.SpaceAfter = 0
        ' === Hauteur automatique (gère aussi vertical) ===
        .Rows.HeightRule = 1   ' wdRowHeightAuto
        .Rows.Height = 0
    End With

    wordTable.Range.Font.Size = 8

    MsgBox "Annexe 3a exportée (C10:H143) dans " & cheminDoc & vbCrLf & _
           "Temps : " & Format(Timer - t0, "0.00") & " s", vbInformation
End Sub
