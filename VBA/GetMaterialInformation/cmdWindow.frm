VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Material Info Command Window"
   ClientHeight    =   3225
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
    With Me.listboxOptions
        .AddItem "Get Long Text"
        .AddItem "Get Most Recent Price Info"
        .AddItem "Get Moving Price/Stock/Safety Stock"
        .AddItem "Get ALL Stock Info"
    End With
     
End Sub

Private Sub btnExecute_Click()
    Dim statusListbox As String
    If Not IsNull(cmdWindow.listboxOptions.Value) Then
        statusListbox = cmdWindow.listboxOptions.Value
    Else
        MsgBox ("You need to pick an option from the listbox")
        Exit Sub
    End If
    
    Select Case statusListbox
        Case "Get Long Text"
            Call MaterialInfoRunScripts.DSConlyRun
        Case "Get Most Recent Price Info"
            Call MaterialInfoRunScripts.multiplePlantRun
        Case "Get Moving Price/Stock/Safety Stock"
            Call MaterialInfoRunScripts.DSConlyRun
        Case "Get ALL Stock Info"
            Call MaterialInfoRunScripts.DSConlyRun
    End Select
End Sub
