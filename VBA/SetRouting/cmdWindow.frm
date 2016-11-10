VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Set Routing Command Window"
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
        .AddItem "CA02 - Routing"
        .AddItem "C002 - PO Routing (Not Implemented)"
    End With
    listboxOptions.Selected(1) = True 'Mark Routing as selected to start with
End Sub

Private Sub btnExecute_Click()
    Dim statusListbox As String
    'Validate user input
    If Not IsNull(cmdWindow.listboxOptions.Value) Then
        statusListbox = cmdWindow.listboxOptions.Value
    Else
        MsgBox ("You need to pick an option from the listbox")
        Exit Sub
    End If
    
    Select Case statusListbox
        Case "CA02 - Routing"
            Call CA02.runscript
    End Select
End Sub
