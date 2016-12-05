VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Set BOM Command Window"
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
 
Private Sub listboxOptions_Change()
    If listboxOptions.ListIndex = BomRun.CO02 Then
        cmdWindow.lbOpNum.Visible = True
        cmdWindow.tbOpNum.Visible = True
        cmdWindow.lbSeq.Visible = True
        cmdWindow.tbSeq.Visible = True
    Else
        cmdWindow.lbOpNum.Visible = False
        cmdWindow.tbOpNum.Visible = False
        cmdWindow.lbSeq.Visible = False
        cmdWindow.tbSeq.Visible = False
    End If
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Top = Application.Top + 125 '< change 125 to what u want
        .Left = Application.Left + 25 '< change 25 to what u want
    End With
    
    With Me.listboxOptions
        .AddItem "CS02 - BOM"
        .AddItem "CO02 - PO BOM"
    End With
    
    cmdWindow.listboxOptions.Selected(0) = True 'Mark CS02 as selected to start with
    
    'Initiliaze Control Tip Text
    Me.lbSAPNum.ControlTipText = "REQUIRED: Enter the column number that contains your SAP numbers"
    Me.lbQty.ControlTipText = "REQUIRED: Enter the column number that contains your quantity amounts"
    Me.lbOpNum.ControlTipText = "REQUIRED: Enter the column number that contains your operation numbers"
    Me.lbSeq.ControlTipText = "OPTIONAL: Enter the column number that contains your sequence numbers"
    
End Sub
Private Sub labelLink_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://www.linkedin.com/in/arashout", NewWindow:=True
End Sub
Private Sub btnExecute_Click()
    'Validate user input
    If IsNull(cmdWindow.listboxOptions.Value) Then
        MsgBox ("You need to pick an option from the listbox")
        Exit Sub
    End If
    
    'Get the correct columns and mode
    Call BomRun.runscriptBOM(CInt(cmdWindow.textSAPNumCol), CInt(cmdWindow.textQtyCol), CInt(cmdWindow.tbSeq), CInt(cmdWindow.tbOpNum), cmdWindow.listboxOptions.ListIndex)
    
End Sub
