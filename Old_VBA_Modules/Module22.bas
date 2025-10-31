Attribute VB_Name = "Module22"
Option Explicit

'===============================================================================
' MACRO PRINCIPALE - Export complet PP-8002 Excel vers Word
' Ex�cute les 5 annexes dans l'ordre logique
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
    
    ' Message de confirmation avant d�marrage
    response = MsgBox("Cette macro va exporter les 5 annexes vers le document Word PP_8002-FR.dotx :" & vbCrLf & vbCrLf & _
                     "� Annexe 1 (PP & SOW Annexe 1)" & vbCrLf & _
                     "� Annexe 2 (PP & SOW Annexe 2)" & vbCrLf & _
                     "� Annexe 3a (Office Layout)" & vbCrLf & _
                     "� Annexe 3b (PP & SOW Annexe 3)" & vbCrLf & _
                     "� Annexe 3c (PP & SOW Annexe 3)" & vbCrLf & vbCrLf & _
                     "Voulez-vous continuer ?", _
                     vbYesNo + vbQuestion, "Export PP-8002 - Confirmation")
    
    If response = vbNo Then Exit Sub
    
    ' D�sactiver les alertes et optimiser pour la performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Debug.Print String(80, "=")
    Debug.Print "D�BUT EXPORT PP-8002 COMPLET - " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Debug.Print String(80, "=")
    
    ' ===== ANNEXE 1 =====
    Debug.Print vbCrLf & "1/5 - Export Annexe 1..."
    On Error Resume Next
    Call PP_SOW_8002_FR_Annexe_1
    If Err.Number = 0 Then
        successCount = successCount + 1
        Debug.Print "? Annexe 1 export�e avec succ�s"
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
        Debug.Print "? Annexe 2 export�e avec succ�s"
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
        Debug.Print "? Annexe 3a export�e avec succ�s"
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
        Debug.Print "? Annexe 3b export�e avec succ�s"
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
        Debug.Print "? Annexe 3c export�e avec succ�s"
    Else
        errorCount = errorCount + 1
        Debug.Print "? Erreur Annexe 3c: " & Err.Description
        logMessage = logMessage & "Erreur Annexe 3c: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Restaurer les param�tres Excel
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Rapport final
    Dim totalTime As Single
    totalTime = Timer - startTime
    
    Debug.Print String(80, "=")
    Debug.Print "EXPORT PP-8002 TERMIN�"
    Debug.Print "Succ�s: " & successCount & "/5"
    Debug.Print "Erreurs: " & errorCount & "/5"
    Debug.Print "Temps total: " & Format(totalTime, "0.00") & " secondes"
    Debug.Print String(80, "=")
    
    ' Message utilisateur final
    Dim finalMessage As String
    If errorCount = 0 Then
        finalMessage = "? EXPORT COMPLET R�USSI !" & vbCrLf & vbCrLf & _
                      "Les 5 annexes ont �t� export�es avec succ�s vers PP_8002-FR.dotx" & vbCrLf & _
                      "Temps d'ex�cution : " & Format(totalTime, "0.00") & " secondes" & vbCrLf & vbCrLf & _
                      "Le document Word est ouvert et pr�t � �tre v�rifi�."
        
        MsgBox finalMessage, vbInformation, "Export PP-8002 - Succ�s complet"
        
    ElseIf successCount > 0 Then
        finalMessage = "? EXPORT PARTIEL" & vbCrLf & vbCrLf & _
                      "R�ussis : " & successCount & "/5 annexes" & vbCrLf & _
                      "Erreurs : " & errorCount & "/5 annexes" & vbCrLf & vbCrLf & _
                      "D�tails des erreurs :" & vbCrLf & logMessage & vbCrLf & _
                      "Temps d'ex�cution : " & Format(totalTime, "0.00") & " secondes"
        
        MsgBox finalMessage, vbExclamation, "Export PP-8002 - Partiel"
        
    Else
        finalMessage = "? �CHEC DE L'EXPORT" & vbCrLf & vbCrLf & _
                      "Aucune annexe n'a pu �tre export�e." & vbCrLf & vbCrLf & _
                      "Erreurs rencontr�es :" & vbCrLf & logMessage & vbCrLf & _
                      "V�rifiez :" & vbCrLf & _
                      "- Que le fichier PP_8002-FR.dotx existe" & vbCrLf & _
                      "- Que les feuilles Excel sont pr�sentes" & vbCrLf & _
                      "- Que Word peut �tre ouvert"
        
        MsgBox finalMessage, vbCritical, "Export PP-8002 - �chec"
    End If
    
End Sub
