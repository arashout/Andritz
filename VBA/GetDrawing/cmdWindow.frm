VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Get Drawing Command Window"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
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
    
End Sub
Private Sub labelLink_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://www.linkedin.com/in/arashout", NewWindow:=True
End Sub
Private Sub btnExecute_Click()
    Dim outCol As Long
    If CInt(cmdWindow.tbOutputCol) = 0 Then
        MsgBox ("Enter a column number to output to")
        Exit Sub
    End If
    outCol = CInt(cmdWindow.tbOutputCol)
    Call GetDrawing.getDrawingMain(outCol, cmdWindow.cbDownload.Value)
End Sub
