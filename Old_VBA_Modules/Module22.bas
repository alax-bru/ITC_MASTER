Attribute VB_Name = "Module22"
Option Explicit

'===============================================================================
' MACRO PRINCIPALE - Export complet PP-8002 Excel vers Word
' Exécute les 5 annexes dans l'ordre logique
'===============================================================================

Public Sub PP_8002()
    Dim startTime As Single
    Dim response As VbMsgBoxResult
    Dim errorCount As Long
    Dim successCount As Long
    Dim logMessage As String
    
    startTime = Timer
    errorCount = 0
    successCount = 0
    
    ' Message de confirmation avant démarrage
    response = MsgBox("Cette macro va exporter les 5 annexes vers le document Word PP_8002-FR.dotx :" & vbCrLf & vbCrLf & _
                     "• Annexe 1 (PP & SOW Annexe 1)" & vbCrLf & _
                     "• Annexe 2 (PP & SOW Annexe 2)" & vbCrLf & _
                     "• Annexe 3a (Office Layout)" & vbCrLf & _
                     "• Annexe 3b (PP & SOW Annexe 3)" & vbCrLf & _
                     "• Annexe 3c (PP & SOW Annexe 3)" & vbCrLf & vbCrLf & _
                     "Voulez-vous continuer ?", _
                     vbYesNo + vbQuestion, "Export PP-8002 - Confirmation")
    
    If response = vbNo Then Exit Sub
    
    ' Désactiver les alertes et optimiser pour la performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Debug.Print String(80, "=")
    Debug.Print "DÉBUT EXPORT PP-8002 COMPLET - " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Debug.Print String(80, "=")
    
    ' ===== ANNEXE 1 =====
    Debug.Print vbCrLf & "1/5 - Export Annexe 1..."
    On Error Resume Next
    Call PP_SOW_8002_FR_Annexe_1
    If Err.Number = 0 Then
        successCount = successCount + 1
        Debug.Print "? Annexe 1 exportée avec succès"
    Else
        errorCount = errorCount + 1
        Debug.Print "? Erreur Annexe 1: " & Err.Description
        logMessage = logMessage & "Erreur Annexe 1: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' ===== ANNEXE 2 =====
    Debug.Print vbCrLf & "2/5 - Export Annexe 2..."
    On Error Resume Next
    Call PP_SOW_8002_FR_Annexe_2
    If Err.Number = 0 Then
        successCount = successCount + 1
        Debug.Print "? Annexe 2 exportée avec succès"
    Else
        errorCount = errorCount + 1
        Debug.Print "? Erreur Annexe 2: " & Err.Description
        logMessage = logMessage & "Erreur Annexe 2: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' ===== ANNEXE 3a =====
    Debug.Print vbCrLf & "3/5 - Export Annexe 3a..."
    On Error Resume Next
    Call PP_SOW_8002_FR_Annexe_3a
    If Err.Number = 0 Then
        successCount = successCount + 1
        Debug.Print "? Annexe 3a exportée avec succès"
    Else
        errorCount = errorCount + 1
        Debug.Print "? Erreur Annexe 3a: " & Err.Description
        logMessage = logMessage & "Erreur Annexe 3a: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' ===== ANNEXE 3b =====
    Debug.Print vbCrLf & "4/5 - Export Annexe 3b..."
    On Error Resume Next
    Call PP_SOW_8002_FR_Annexe_3b
    If Err.Number = 0 Then
        successCount = successCount + 1
        Debug.Print "? Annexe 3b exportée avec succès"
    Else
        errorCount = errorCount + 1
        Debug.Print "? Erreur Annexe 3b: " & Err.Description
        logMessage = logMessage & "Erreur Annexe 3b: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' ===== ANNEXE 3c =====
    Debug.Print vbCrLf & "5/5 - Export Annexe 3c..."
    On Error Resume Next
    Call PP_SOW_8002_FR_Annexe_3c
    If Err.Number = 0 Then
        successCount = successCount + 1
        Debug.Print "? Annexe 3c exportée avec succès"
    Else
        errorCount = errorCount + 1
        Debug.Print "? Erreur Annexe 3c: " & Err.Description
        logMessage = logMessage & "Erreur Annexe 3c: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Restaurer les paramètres Excel
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Rapport final
    Dim totalTime As Single
    totalTime = Timer - startTime
    
    Debug.Print String(80, "=")
    Debug.Print "EXPORT PP-8002 TERMINÉ"
    Debug.Print "Succès: " & successCount & "/5"
    Debug.Print "Erreurs: " & errorCount & "/5"
    Debug.Print "Temps total: " & Format(totalTime, "0.00") & " secondes"
    Debug.Print String(80, "=")
    
    ' Message utilisateur final
    Dim finalMessage As String
    If errorCount = 0 Then
        finalMessage = "? EXPORT COMPLET RÉUSSI !" & vbCrLf & vbCrLf & _
                      "Les 5 annexes ont été exportées avec succès vers PP_8002-FR.dotx" & vbCrLf & _
                      "Temps d'exécution : " & Format(totalTime, "0.00") & " secondes" & vbCrLf & vbCrLf & _
                      "Le document Word est ouvert et prêt à être vérifié."
        
        MsgBox finalMessage, vbInformation, "Export PP-8002 - Succès complet"
        
    ElseIf successCount > 0 Then
        finalMessage = "? EXPORT PARTIEL" & vbCrLf & vbCrLf & _
                      "Réussis : " & successCount & "/5 annexes" & vbCrLf & _
                      "Erreurs : " & errorCount & "/5 annexes" & vbCrLf & vbCrLf & _
                      "Détails des erreurs :" & vbCrLf & logMessage & vbCrLf & _
                      "Temps d'exécution : " & Format(totalTime, "0.00") & " secondes"
        
        MsgBox finalMessage, vbExclamation, "Export PP-8002 - Partiel"
        
    Else
        finalMessage = "? ÉCHEC DE L'EXPORT" & vbCrLf & vbCrLf & _
                      "Aucune annexe n'a pu être exportée." & vbCrLf & vbCrLf & _
                      "Erreurs rencontrées :" & vbCrLf & logMessage & vbCrLf & _
                      "Vérifiez :" & vbCrLf & _
                      "- Que le fichier PP_8002-FR.dotx existe" & vbCrLf & _
                      "- Que les feuilles Excel sont présentes" & vbCrLf & _
                      "- Que Word peut être ouvert"
        
        MsgBox finalMessage, vbCritical, "Export PP-8002 - Échec"
    End If
    
End Sub
