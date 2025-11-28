Attribute VB_Name = "Module2"
'TSF CTR

Option Explicit

'----------------- Common
Public CTR_Storey As Integer
Public CTR_3rd_Access As Boolean
Public CTR_Total_Container As Integer

'--------------------------------------------------- Floor 1
'----------------- Row A / 12 Types
Public Floor_1_CTR_A_Senior_Single As Integer   ' Row A only
Public Floor_1_CTR_A_Senior_Double As Integer   ' Row A only
Public Floor_1_CTR_A_Site_Manager As Integer    ' Floor 1 / Row A only

Public Floor_1_CTR_A_Meeting_1 As Integer
Public Floor_1_CTR_A_Meeting_2 As Integer
Public Floor_1_CTR_A_WC As Integer
Public Floor_1_CTR_A_Locker As Integer
Public Floor_1_CTR_A_Printer As Integer
Public Floor_1_CTR_A_Restroom As Integer
Public Floor_1_CTR_A_Storage As Integer
Public Floor_1_CTR_A_Elec_Room As Integer       ' Row A only
Public Floor_1_CTR_A_Empty As Integer

'----------------- Row B / 10 Types
Public Floor_1_CTR_B_OpenSpace As Integer       ' Row B only

Public Floor_1_CTR_B_Meeting_1 As Integer
Public Floor_1_CTR_B_Meeting_2 As Integer
Public Floor_1_CTR_B_WC As Integer
Public Floor_1_CTR_B_Locker As Integer
Public Floor_1_CTR_B_Printer As Integer
Public Floor_1_CTR_B_Restroom As Integer
Public Floor_1_CTR_B_Storage As Integer
Public Floor_1_CTR_B_Access As Integer          ' Row B only
Public Floor_1_CTR_B_Empty As Integer

'--------------------------------------------------- Floor 2
'----------------- Row A / 11 Types
Public Floor_2_CTR_A_Senior_Single As Integer   ' Row A only
Public Floor_2_CTR_A_Senior_Double As Integer   ' Row A only

Public Floor_2_CTR_A_Meeting_1 As Integer
Public Floor_2_CTR_A_Meeting_2 As Integer
Public Floor_2_CTR_A_WC As Integer
Public Floor_2_CTR_A_Locker As Integer
Public Floor_2_CTR_A_Printer As Integer
Public Floor_2_CTR_A_Restroom As Integer
Public Floor_2_CTR_A_Storage As Integer
Public Floor_2_CTR_A_Elec_Room As Integer       ' Row A only
Public Floor_2_CTR_A_Empty As Integer

'----------------- Row B / 10 Types
Public Floor_2_CTR_B_OpenSpace As Integer       ' Row B only

Public Floor_2_CTR_B_Meeting_1 As Integer
Public Floor_2_CTR_B_Meeting_2 As Integer
Public Floor_2_CTR_B_WC As Integer
Public Floor_2_CTR_B_Locker As Integer
Public Floor_2_CTR_B_Printer As Integer
Public Floor_2_CTR_B_Restroom As Integer
Public Floor_2_CTR_B_Storage As Integer
Public Floor_2_CTR_B_Access As Integer          ' Row B only
Public Floor_2_CTR_B_Empty As Integer

'----------------- Variables
Dim Container_By_Floor_By_Row As Integer
Dim Left_Already_Assigned As Integer
Dim Right_Already_Assigned As Integer
Dim Mid_Insertion_Point As Integer
Dim Offices_Image_Height As Integer
Dim Image_Width As Integer

'----------------- Array declaration
Dim Floor_1_CTR_A() As String
Dim Floor_1_CTR_B() As String
Dim Floor_2_CTR_A() As String
Dim Floor_2_CTR_B() As String

'==========================================================
'  VIDAGE COMPLET DU PRESSE-PAPIERS (Excel + Windows)
'==========================================================
Private Sub ClearClipboardTSF()
    On Error Resume Next
    
    ' Vider le presse-papiers Excel
    Application.CutCopyMode = False
    
    ' Copier une cellule neutre pour écraser le contenu
    ThisWorkbook.Worksheets(1).Range("A1").Copy
    Application.CutCopyMode = False
    
    ' Tentative de vider le presse-papiers Windows
    CreateObject("htmlfile").parentWindow.clipboardData.clearData
    
    On Error GoTo 0
End Sub

Sub OfficesInitializeValues()
    With Worksheets("1.6-TSF CTR (Input Anx 4)")

        '----------------- Common
        If .Range("D69") = 3 Then
            CTR_3rd_Access = True
        Else
            CTR_3rd_Access = False
        End If

        CTR_Storey = .Range("D66")
        CTR_Total_Container = .Range("X109")

        '--------------------------------------------------- Floor 1
        '----------------- Row A / 12 Types
        Floor_1_CTR_A_Senior_Single = .Range("X81")   ' Row A only
        Floor_1_CTR_A_Senior_Double = .Range("X82")   ' Row A only
        Floor_1_CTR_A_Site_Manager = .Range("X83")    ' Floor 1 / Row A only

        Floor_1_CTR_A_Meeting_1 = .Range("X86")
        Floor_1_CTR_A_Meeting_2 = .Range("X87")
        Floor_1_CTR_A_WC = .Range("X88")
        Floor_1_CTR_A_Locker = .Range("X89")
        Floor_1_CTR_A_Printer = .Range("X90")
        Floor_1_CTR_A_Restroom = .Range("X91")
        Floor_1_CTR_A_Storage = .Range("X92")
        Floor_1_CTR_A_Elec_Room = .Range("X93")       ' Row A only
        Floor_1_CTR_A_Empty = .Range("X107")

        '----------------- Row B / 10 Types
        Floor_1_CTR_B_OpenSpace = .Range("Y80")       ' Row B only

        Floor_1_CTR_B_Meeting_1 = .Range("Y86")
        Floor_1_CTR_B_Meeting_2 = .Range("Y87")
        Floor_1_CTR_B_WC = .Range("Y88")
        Floor_1_CTR_B_Locker = .Range("Y89")
        Floor_1_CTR_B_Printer = .Range("Y90")
        Floor_1_CTR_B_Restroom = .Range("Y91")
        Floor_1_CTR_B_Storage = .Range("Y92")
        Floor_1_CTR_B_Access = .Range("Y94")          ' Row B only
        Floor_1_CTR_B_Empty = .Range("Y107")

        '--------------------------------------------------- Floor 2
        '----------------- Row A / 11 Types
        Floor_2_CTR_A_Senior_Single = .Range("Z81")   ' Row A only
        Floor_2_CTR_A_Senior_Double = .Range("Z82")   ' Row A only

        Floor_2_CTR_A_Meeting_1 = .Range("Z86")
        Floor_2_CTR_A_Meeting_2 = .Range("Z87")
        Floor_2_CTR_A_WC = .Range("Z88")
        Floor_2_CTR_A_Locker = .Range("Z89")
        Floor_2_CTR_A_Printer = .Range("Z90")
        Floor_2_CTR_A_Restroom = .Range("Z91")
        Floor_2_CTR_A_Storage = .Range("Z92")
        Floor_2_CTR_A_Elec_Room = .Range("Z93")       ' Row A only
        Floor_2_CTR_A_Empty = .Range("Z107")

        '----------------- Row B / 10 Types
        Floor_2_CTR_B_OpenSpace = .Range("AA80")      ' Row B only

        Floor_2_CTR_B_Meeting_1 = .Range("AA86")
        Floor_2_CTR_B_Meeting_2 = .Range("AA87")
        Floor_2_CTR_B_WC = .Range("AA88")
        Floor_2_CTR_B_Locker = .Range("AA89")
        Floor_2_CTR_B_Printer = .Range("AA90")
        Floor_2_CTR_B_Restroom = .Range("AA91")
        Floor_2_CTR_B_Storage = .Range("AA92")
        Floor_2_CTR_B_Access = .Range("AA94")         ' Row B only
        Floor_2_CTR_B_Empty = .Range("AA107")

    End With

    '----------------- Variables
    Container_By_Floor_By_Row = CTR_Total_Container / 2 / CTR_Storey
    Offices_Image_Height = 175  'Used to Resize Pictures
    Image_Width = 75            'Used to Align Pictures

    '----------------- Array declaration
    ReDim Floor_1_CTR_A(0)
    ReDim Floor_1_CTR_B(0)
    ReDim Floor_2_CTR_A(0)
    ReDim Floor_2_CTR_B(0)

End Sub

'==========================================================
'  SUPPRESSION DES IMAGES / OBJETS DANS LA ZONE DE DESSIN
'==========================================================
Private Sub Clear_TSF_Shapes_In_Range()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim ch As ChartObject
    Dim rng As Range

    Set ws = ThisWorkbook.Worksheets("2.6-PP & SOW Annexe 4")
    ' zone de dessin des containers (à adapter si besoin)
    Set rng = ws.Range("F1:AZ200")

    ' Shapes (images, formes, etc.)
    For Each shp In ws.Shapes
        If Not Application.Intersect(shp.TopLeftCell, rng) Is Nothing Then
            shp.Delete
        End If
    Next shp

    ' Graphes éventuels
    For Each ch In ws.ChartObjects
        If Not Application.Intersect(ch.TopLeftCell, rng) Is Nothing Then
            ch.Delete
        End If
    Next ch
End Sub

Sub TSF_Layout_Array()

    ' --- vidage complet du presse-papiers + suppression anciens dessins ---
    ClearClipboardTSF
    Clear_TSF_Shapes_In_Range

    ' reset des compteurs gauche/droite pour cette exécution
    Left_Already_Assigned = 0
    Right_Already_Assigned = 0

    Call OfficesInitializeValues
    Call OfficesSizeLibraryPicture

    Dim i As Integer
    Dim Number_Of_Container_By_Room As Integer
    Dim First_Left_Free_Position As Integer
    Dim First_Right_Free_Position As Integer

    '--------------------------------------------------- Floor 1
    ReDim Preserve Floor_1_CTR_A(Container_By_Floor_By_Row)
    ReDim Preserve Floor_1_CTR_B(Container_By_Floor_By_Row)

    '--------------------------------------------------- Floor 2
    ReDim Preserve Floor_2_CTR_A(Container_By_Floor_By_Row)
    ReDim Preserve Floor_2_CTR_B(Container_By_Floor_By_Row)

    '--------------------------------------------------- Floor 1
    'Floor 1 / Row A
    Call Assign_Room(Floor_1_CTR_A, Floor_1_CTR_A_Site_Manager, Floor_1_CTR_A_Senior_Single, Floor_1_CTR_A_Senior_Double, Floor_1_CTR_A_Elec_Room, Floor_1_CTR_A_Meeting_1, Floor_1_CTR_A_Meeting_2, Floor_1_CTR_A_WC, Floor_1_CTR_A_Locker, Floor_1_CTR_A_Printer, Floor_1_CTR_A_Restroom, Floor_1_CTR_A_Storage, Floor_1_CTR_A_Empty, 0, 0)
    Call TSF_Layout_Drawing(Floor_1_CTR_A, 1, "A")

    'Floor 1 / Row B
    Call Assign_Room(Floor_1_CTR_B, 0, 0, 0, 0, Floor_1_CTR_B_Meeting_1, Floor_1_CTR_B_Meeting_2, Floor_1_CTR_B_WC, Floor_1_CTR_B_Locker, Floor_1_CTR_B_Printer, Floor_1_CTR_B_Restroom, Floor_1_CTR_B_Storage, Floor_1_CTR_B_Empty, Floor_1_CTR_B_OpenSpace, Floor_1_CTR_B_Access)
    Call TSF_Layout_Drawing(Floor_1_CTR_B, 1, "B")

    '--------------------------------------------------- Floor 2
    If CTR_Storey = 2 Then
        'Floor 2 / Row A
        Call Assign_Room(Floor_2_CTR_A, 0, Floor_2_CTR_A_Senior_Single, Floor_2_CTR_A_Senior_Double, Floor_2_CTR_A_Elec_Room, Floor_2_CTR_A_Meeting_1, Floor_2_CTR_A_Meeting_2, Floor_2_CTR_A_WC, Floor_2_CTR_A_Locker, Floor_2_CTR_A_Printer, Floor_2_CTR_A_Restroom, Floor_2_CTR_A_Storage, Floor_2_CTR_A_Empty, 0, 0)
        Call TSF_Layout_Drawing(Floor_2_CTR_A, 2, "A")
        
        'Floor 2 / Row B
        Call Assign_Room(Floor_2_CTR_B, 0, 0, 0, 0, Floor_2_CTR_B_Meeting_1, Floor_2_CTR_B_Meeting_2, Floor_2_CTR_B_WC, Floor_2_CTR_B_Locker, Floor_2_CTR_B_Printer, Floor_2_CTR_B_Restroom, Floor_2_CTR_B_Storage, Floor_2_CTR_B_Empty, Floor_2_CTR_B_OpenSpace, Floor_2_CTR_B_Access)
        Call TSF_Layout_Drawing(Floor_2_CTR_B, 2, "B")
    End If

End Sub

Function First_Left_Free(Floor_Row() As String) As Integer
    Dim i As Integer

    For i = 1 To Container_By_Floor_By_Row
        If Floor_Row(i) = "" Then
            First_Left_Free = i
            Exit Function
        End If
    Next

End Function

Function First_Right_Free(Floor_Row() As String) As Integer
    Dim i As Integer

    For i = Container_By_Floor_By_Row To 1 Step -1
        If Floor_Row(i) = "" Then
            First_Right_Free = i
            Exit Function
        End If
    Next

End Function

Sub Assign_Room(ByRef Actual_Floor_Row() As String, Site_Manager As Integer, Senior_Single As Integer, Senior_Double As Integer, Elec_Room As Integer, Meeting_1 As Integer, Meeting_2 As Integer, WC As Integer, Locker As Integer, Printer As Integer, Restroom As Integer, Storage As Integer, Empty_Container As Integer, OpenSpace As Integer, Access As Integer)

    Call AssignSide(Actual_Floor_Row, "WC", WC, 2)
    Call AssignSide(Actual_Floor_Row, "Locker", Locker, 1)
    Call AssignSide(Actual_Floor_Row, "Storage", Storage, 1)
    Call AssignSide(Actual_Floor_Row, "Restroom", Restroom, 2)
    Call AssignSide(Actual_Floor_Row, "Printer", Printer, 1)

    Call AssignSide(Actual_Floor_Row, "Meeting_1", Meeting_1, 2)
    Call AssignSide(Actual_Floor_Row, "OpenSpace", OpenSpace, 1)
    Call AssignSide(Actual_Floor_Row, "Senior_Double", Senior_Double, 1)
    Call AssignSide(Actual_Floor_Row, "Meeting_2", Meeting_2, 2)
    Call AssignSide(Actual_Floor_Row, "Senior_Single", Senior_Single, 1)
    Call AssignSide(Actual_Floor_Row, "Site_Manager", Site_Manager, 2)

    Call AssignSide(Actual_Floor_Row, "Elec_Room", Elec_Room, 1)
    Call AssignSide(Actual_Floor_Row, "Access", Access, 1)
    Call AssignSide(Actual_Floor_Row, "Empty_Container", Empty_Container, 1)

End Sub

Sub AssignSide(ByRef Actual_Floor_Row() As String, RoomName As String, ContainerQuantity As Integer, Number_Of_Container_By_Room As Integer)
    Dim i As Integer, j As Integer, First_Left_Free_Position As Integer, First_Right_Free_Position As Integer

    If ContainerQuantity = 0 Then Exit Sub

    Do Until ContainerQuantity = 0

        If Left_Already_Assigned <= Right_Already_Assigned Then

            First_Left_Free_Position = First_Left_Free(Actual_Floor_Row)

            For j = 1 To Number_Of_Container_By_Room
                If j = 1 Then
                    Actual_Floor_Row(First_Left_Free_Position + j - 1) = RoomName
                Else
                    Actual_Floor_Row(First_Left_Free_Position + j - 1) = "X"
                End If
            Next    'j
             
            Left_Already_Assigned = Left_Already_Assigned + Number_Of_Container_By_Room
            ContainerQuantity = ContainerQuantity - Number_Of_Container_By_Room

            If ContainerQuantity > 0 Then
                First_Right_Free_Position = First_Right_Free(Actual_Floor_Row)
                
                For j = 1 To Number_Of_Container_By_Room
                    If j = Number_Of_Container_By_Room Then
                        Actual_Floor_Row(First_Right_Free_Position - j + 1) = RoomName
                    Else
                        Actual_Floor_Row(First_Right_Free_Position - j + 1) = "X"
                    End If
                Next 'j
                
                Right_Already_Assigned = Right_Already_Assigned + Number_Of_Container_By_Room
                ContainerQuantity = ContainerQuantity - Number_Of_Container_By_Room
            End If 'ContainerQuantity > 0
            
        Else 'Left_Already_Assigned <= Right_Already_Assigned
        
            First_Right_Free_Position = First_Right_Free(Actual_Floor_Row)
            
            For j = 1 To Number_Of_Container_By_Room
                If j = Number_Of_Container_By_Room Then
                    Actual_Floor_Row(First_Right_Free_Position - j + 1) = RoomName
                Else
                    Actual_Floor_Row(First_Right_Free_Position - j + 1) = "X"
                End If
            Next 'j
        
            Right_Already_Assigned = Right_Already_Assigned + Number_Of_Container_By_Room
        
            ContainerQuantity = ContainerQuantity - Number_Of_Container_By_Room
               
            If ContainerQuantity > 0 Then
                First_Left_Free_Position = First_Left_Free(Actual_Floor_Row)
                
                For j = 1 To Number_Of_Container_By_Room
                    If j = 1 Then
                        Actual_Floor_Row(First_Left_Free_Position + j - 1) = RoomName
                    Else
                        Actual_Floor_Row(First_Left_Free_Position + j - 1) = "X"
                    End If
                Next    'j
                      
                Right_Already_Assigned = Right_Already_Assigned + Number_Of_Container_By_Room
                ContainerQuantity = ContainerQuantity - Number_Of_Container_By_Room
            End If 'ContainerQuantity >0
            
        End If 'Left_Already_Assigned <= Right_Already_Assigned
    
    Loop 'ContainerQuantity

End Sub

'==========================================================
'  PRÉPARATION ANNEXE 4
'  - Déplace F:AV derrière ZZ (à partir de AAA)
'  - Copie AZ dans toutes les colonnes F:AZ, lignes 1 à 100
'  - Supprime les shapes dans la zone de dessin avant déplacement
'==========================================================
Private Sub Prepare_TSF_Annexe4_Layout()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim srcCol As Long
    Dim firstDestCol As Long
    Dim lastDestCol As Long
    Dim c As Long

    ' --- reset presse-papiers ---
    ClearClipboardTSF

    Set ws = ThisWorkbook.Worksheets("2.6-PP & SOW Annexe 4")

    ' on s'assure que la zone F:AZ est vidée de ses shapes
    Clear_TSF_Shapes_In_Range

    With ws
        ' 0) Ne déplacer F:AV qu'une seule fois
        If .Range("AAA1").Value = "" Then
            ' 1) Déplacer les colonnes F:AV derrière ZZ, à partir de AAA
            .Columns("F:AV").Cut Destination:=.Columns("AAA")
            .Range("AAA1").Value = "MOVED"
        End If

        ' 2) Copier / coller la colonne AZ sur toutes les colonnes F:AZ
        '    seulement de la ligne 1 à 100
        lastRow = 100
        srcCol = .Columns("AZ").Column
        firstDestCol = .Columns("F").Column
        lastDestCol = srcCol        ' = AZ

        For c = firstDestCol To lastDestCol
            .Range(.Cells(1, srcCol), .Cells(lastRow, srcCol)).Copy _
                Destination:=.Cells(1, c)
        Next c

        Application.CutCopyMode = False
    End With

End Sub

Sub TSF_Layout_Drawing(ByRef Actual_Floor_Row() As String, Actual_Floor As String, Actual_Row As String)

    ' --- reset presse-papiers avant copies d’images ---
    ClearClipboardTSF

    Application.ScreenUpdating = False
    Dim Toggle As Boolean
    Toggle = False

    Dim Floor_Stair_Insertion As Integer

    Dim DelayTime As Integer
    DelayTime = 250 'Create delay to avoid bug during paste

    ' --- Préparation de la feuille (déplacement des colonnes + copie AZ) une seule fois ---
    Static LayoutPrepared As Boolean
    If Not LayoutPrepared Then
        Prepare_TSF_Annexe4_Layout
        LayoutPrepared = True
    End If
    ' --- fin préparation ---

    Sheets("TSF Library").Visible = True
    Sheets("TSF Library").Select

    Dim i As Integer
    Dim Offices_Left_Mini_Insertion_Point As Integer, Offices_Left_Insertion_Point As Integer, Offices_Top_Insertion_Point As Integer

    Offices_Left_Mini_Insertion_Point = 400
    Offices_Left_Insertion_Point = Offices_Left_Mini_Insertion_Point

    If Actual_Floor = 1 Then
        Offices_Top_Insertion_Point = 20
    Else
        Offices_Top_Insertion_Point = 3 * Offices_Image_Height + 20
    End If

    If Actual_Row = "A" Then
        Offices_Top_Insertion_Point = Offices_Top_Insertion_Point + 0
    Else
        Offices_Top_Insertion_Point = Offices_Top_Insertion_Point + Offices_Image_Height
    End If

    For i = 1 To Container_By_Floor_By_Row

        If Actual_Floor = 1 And (Actual_Floor_Row(i)) = "Access" Then
            Mid_Insertion_Point = Offices_Left_Insertion_Point
        End If

        If (Actual_Floor_Row(i)) = "X" Then
            Offices_Left_Insertion_Point = Offices_Left_Insertion_Point + Image_Width
        Else
            Sheets("TSF Library").Select
            ActiveSheet.Shapes.Range(Array("Pic_" & Actual_Row & "_" & Actual_Floor_Row(i))).Select
            Selection.Copy
            Sheets("2.6-PP & SOW Annexe 4").Select
            ActiveWindow.Zoom = 40 'To avoid bug in image dimensions
            
            sov (DelayTime / 1000) 'Create delay to avoid bug during paste
            
            ActiveSheet.Pictures.Paste.Select
            
            Selection.Name = "Pic_" & Actual_Floor & "_" & Actual_Row & "_" & i
            
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.Height = Offices_Image_Height
                
            Selection.Left = Offices_Left_Insertion_Point
            Selection.Top = Offices_Top_Insertion_Point
            
            If Actual_Floor_Row(i) = "OpenSpace" Then
                Toggle = Not (Toggle)
                If Toggle = True Then Selection.ShapeRange.Flip msoFlipHorizontal
            End If

            Offices_Left_Insertion_Point = Offices_Left_Insertion_Point + Image_Width

        End If '(Actual_Floor_Row(i)) = "X"

    Next i
        
    'Stairs
    Dim Stair_Left As String, Stair_Mid As String, Stair_Right As String
        
    If CTR_Storey = 1 Then
        Stair_Left = "Pic_Left_1F_Stair"
        Stair_Mid = "Pic_Mid_1F_Stair"
        Stair_Right = "Pic_Right_1F_Stair"
        Floor_Stair_Insertion = 100
    Else
        Stair_Left = "Pic_Left_2F_Stair"
        Stair_Mid = "Pic_Mid_2F_Stair"
        Stair_Right = "Pic_Right_2F_Stair"
        Floor_Stair_Insertion = 10
    End If 'CTR_Storey = 1
        
    If Actual_Row = "A" Then
            
        'Left Stair
        Sheets("TSF Library").Select
        ActiveSheet.Shapes.Range(Array(Stair_Left)).Select
        Selection.Copy
        Sheets("2.6-PP & SOW Annexe 4").Select
            
        sov (DelayTime / 1000) 'Create delay to avoid bug during paste
            
        ActiveSheet.Pictures.Paste.Select
            
        Selection.Name = Stair_Left
            
        Selection.Left = Offices_Left_Mini_Insertion_Point - 130
        Selection.Top = Offices_Top_Insertion_Point + Floor_Stair_Insertion
            
        'Right Stair
        Sheets("TSF Library").Select
        ActiveSheet.Shapes.Range(Array(Stair_Right)).Select
        Selection.Copy
        Sheets("2.6-PP & SOW Annexe 4").Select
            
        sov (DelayTime / 1000) 'Create delay to avoid bug during paste
            
        ActiveSheet.Pictures.Paste.Select
            
        Selection.Name = Stair_Right
            
        Selection.Left = Offices_Left_Insertion_Point - 15
        Selection.Top = Offices_Top_Insertion_Point + Floor_Stair_Insertion
            
    Else
        'Mid Stair
        If CTR_3rd_Access = True Then
                 
            Sheets("TSF Library").Select
            ActiveSheet.Shapes.Range(Array(Stair_Mid)).Select
            Selection.Copy
            Sheets("2.6-PP & SOW Annexe 4").Select
                
            sov (DelayTime / 1000) 'Create delay to avoid bug during paste
                
            ActiveSheet.Pictures.Paste.Select
                
            Selection.Name = Stair_Mid
                
            Offices_Top_Insertion_Point = Offices_Top_Insertion_Point + Offices_Image_Height
                
            Selection.Left = Mid_Insertion_Point
            Selection.Top = Offices_Top_Insertion_Point - 20

        End If 'CTR_3rd_Access = true
            
    End If 'Actual_Row = "A"
        
    Sheets("TSF Library").Visible = False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

End Sub

Function sov(sekunder As Double) As Double
    Dim starting_time As Double

    starting_time = Timer

    Do
        DoEvents
    Loop Until (Timer - starting_time) >= sekunder

End Function

Sub OfficesSizeLibraryPicture()

    Sheets("TSF Library").Visible = True
    Sheets("TSF Library").Select
    ActiveWindow.Zoom = 40
        
    Dim Pics_List() As Variant
    Pics_List = Array( _
        "Pic_A_Printer", _
        "Pic_A_Senior_Single", _
        "Pic_A_Senior_Double", _
        "Pic_B_Empty_Container", _
        "Pic_A_Elec_Room", _
        "Pic_A_Locker", _
        "Pic_A_Storage", _
        "Pic_A_Restroom", _
        "Pic_A_WC", _
        "Pic_A_Meeting_1", _
        "Pic_A_Site_Manager", _
        "Pic_A_Meeting_2", _
        "Pic_B_Printer", _
        "Pic_B_OpenSpace", _
        "Pic_B_Empty_Container", _
        "Pic_B_Access", _
        "Pic_B_Locker", _
        "Pic_B_Storage", _
        "Pic_B_Restroom", _
        "Pic_B_WC", _
        "Pic_B_Meeting_1", _
        "Pic_B_Meeting_2")
        
    Dim i As Integer
    ' 6m Height
    For i = 0 To UBound(Pics_List)
        ActiveSheet.Shapes.Range(Pics_List(i)).Select
        Selection.ShapeRange.LockAspectRatio = msoTrue
        Selection.Height = Offices_Image_Height
    Next

End Sub

