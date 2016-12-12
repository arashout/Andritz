VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Material Info Command Window"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "cmdWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cmdWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Top = Application.Top + 125 '< change 125 to what u want
        .Left = Application.Left + 25 '< change 25 to what u want
    End With
    'I have to do all this because I want to use the the items in the collection to populate the list box
    Dim tempMat As CMaterial
    Dim session As Variant
    Set session = SAPFunctions.connect2SAP
    Set tempMat = factory.createCMaterial(sapNum:="123456789", currentSession:=session, rowI:=1, colI:=1, plantNum:="1105")
    Set session = Nothing
    
    'Populate the list box with collectionInfo
    With Me.listboxOptions
        Dim infoTuple As Variant
        For Each infoTuple In tempMat.collectionInfo
            .AddItem (infoTuple(1))
        Next infoTuple
    End With
    
    'Populate the combo box with list of plants
    With Me.cmbPlant
        .AddItem "1105"
        .AddItem "0303"
    End With
    Me.cmbPlant.value = Me.cmbPlant.List(0) 'Make 1105 the default
    
    'Tool-tips
    cmdWindow.listboxOptions.ControlTipText = "Select the items you would like to output." & vbCrLf & "NOTE: Use CTRL + Click to select multiple items"
    cmdWindow.chkAllStock.ControlTipText = "This selects all the stock options + moving price for you"
    cmdWindow.chkHeaders.ControlTipText = "Inserts a row right above the first SAP number with the chosen headers"
    cmdWindow.cmbPlant.ControlTipText = "Choose the plant that you want ZMATINFO information from"
    
End Sub
Private Sub labelLink_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://www.linkedin.com/in/arashout", NewWindow:=True
End Sub
Private Sub chkAllStock_Click()
    Dim itemsToChange() As String
    Dim val As String
    Dim i As Integer
    itemsToChange = Split("Moving Price,Stock,Safety Stock,Project Stock,Order Reservation,Product Order,Purchase Requisition,Purchase Order Item,Dependant Requisition,Planned Order", ",")
    If Me.chkAllStock.value Then
        For i = 0 To cmdWindow.listboxOptions.ListCount - 1
            val = cmdWindow.listboxOptions.List(i)
            If HelpFunctions.inArr(itemsToChange, val) Then
                cmdWindow.listboxOptions.Selected(i) = True
            End If
        Next i
    Else
        For i = 0 To cmdWindow.listboxOptions.ListCount - 1
            val = cmdWindow.listboxOptions.List(i)
            If HelpFunctions.inArr(itemsToChange, val) Then
                cmdWindow.listboxOptions.Selected(i) = False
            End If
        Next i
    End If
End Sub
Private Sub btnExecute_Click()
    Dim i As Integer
    Dim key As String
    Dim outputCollectionKeys As Collection
    Set outputCollectionKeys = New Collection
    
    'Figure out which items we need to output by adding to keys collection
    For i = 0 To cmdWindow.listboxOptions.ListCount - 1
        key = cmdWindow.listboxOptions.List(i)
        If cmdWindow.listboxOptions.Selected(i) Then
            outputCollectionKeys.Add (key)
        End If
    Next i
    
    'Make sure the user selected something
    If outputCollectionKeys.Count = 0 Then
        MsgBox ("You need to pick an option from the listbox")
        Exit Sub
    End If
    
    'Get the range the user has selected to determine the first and last cell to iterate between
    Dim sel As Range
    Set sel = Selection
    
    Call MaterialInfoRun.RunCollectMatInfo(sel, cmdWindow.cmbPlant.value, outputCollectionKeys, cmdWindow.chkHeaders.value)
End Sub
