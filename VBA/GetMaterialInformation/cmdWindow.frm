VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Material Info Command Window"
   ClientHeight    =   4995
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
    
    'Tool-tips
    cmdWindow.listboxOptions.ControlTipText = "Select the items you would like to output." & vbCrLf & "NOTE: Use CTRL + Click to select multiple items"
    cmdWindow.chkAllStock.ControlTipText = "This selects all the stock options + moving price for you"
    cmdWindow.chkHeaders.ControlTipText = "Outputs the headers of the chosen values on ROW 1 of the sheet" & vbCrLf & "NOTE: THIS WILL OVERWRITE ITEMS AT THE TOP"
    
End Sub
Private Sub chkAllStock_Click()
    Dim itemsToChange() As String
    Dim val As String
    Dim i As Integer
    itemsToChange = Split("Moving Price,Stock,Safety Stock,Project Stock,Order Reservation,Product Order,Purchase Requisition,Purchase Order Item,Dependant Requisition,Planned Order", ",")
    If Me.opAllStock.value Then
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
    
    Call MaterialInfoRun.DSConlyRun(outputCollectionKeys, cmdWindow.chkHeaders.value)
End Sub
