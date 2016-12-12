VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Basic Command Window"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   OleObjectBlob   =   "template.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cmdWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
End Sub
Private Sub labelLink_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://www.linkedin.com/in/arashout", NewWindow:=True
End Sub
Private Sub btnExecute_Click()
    
End Sub
