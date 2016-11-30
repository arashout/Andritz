VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cmdWindow 
   Caption         =   "Set Routing Command Window"
   ClientHeight    =   3360
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

Private Sub helpBtn_Click()
    Dim purpose, requirements As String
    purpose = HelpFunctions.repeatString("-", 5) & "Purpose" & HelpFunctions.repeatString("-", 5) & vbCrLf & _
    "The aim of this macro is to stream-line the process of entering operations in CA02." & vbCrLf & _
    "It does this by copying operation information (Operation Numbers, Work Centers, Hours and Long Text) into SAP from Excel for you." & vbCrLf
    
    requirements = HelpFunctions.repeatString("-", 5) & "Requirements" & HelpFunctions.repeatString("-", 5) & vbCrLf & _
    "In your EXCEL Sheet" & vbCrLf & _
    "The macro can only handle 1 Operation per row MAXIMUM" & vbCrLf & _
    "The operation number, operation description, work center, operations hours MUST be in seperate columns" & vbCrLf & _
    "In SAP" & vbCrLf & _
    "Ensure that the alternative editor is enabled (Talk to the developer for more information)" & vbCrLf
    
    
    MsgBox purpose & requirements, vbInformation + vbOKOnly, "Set Routing Help"
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Top = Application.Top + 125 '< change 125 to what u want
        .Left = Application.Left + 25 '< change 25 to what u want
    End With
    With Me.listboxOptions
        .AddItem "CA02 - Enter Operations"
        .AddItem "C002 - Enter Operations"
    End With
    listboxOptions.Selected(0) = True 'Mark CA02 as selected to start with
    
    'Initiliaze Control Tip Text
    Me.OpNumLabel.ControlTipText = "REQUIRED: Enter the column number that contains your operation numbers"
    Me.DescLabel.ControlTipText = "REQUIRED: Enter the column number that contains operation descriptions"
    Me.WorkCtrLabel.ControlTipText = "REQUIRED: Enter the column number that contains work centers"
    Me.HoursLabel.ControlTipText = "REQUIRED: Enter the column number that contains operation hours"
End Sub
Private Sub labelLink_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://www.linkedin.com/in/arashout", NewWindow:=True
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
    
    If Not validInputs Then
        MsgBox ("You need to specify the numbers of the columns that store your information")
        Exit Sub
    End If
    
    Call RoutingRun.EnterOperations(CInt(cmdWindow.tbDesc.Value), CInt(cmdWindow.tbWorkCtr.Value), _
        CInt(cmdWindow.tbHours.Value), CInt(cmdWindow.tbOpNum.Value), _
        cmdWindow.listboxOptions.ListIndex)
End Sub

Private Function validInputs() As Boolean
    If Not Not IsNumeric(cmdWindow.tbDesc.Value) And Not IsNumeric(cmdWindow.tbDesc.Value) And Not IsNumeric(cmdWindow.tbOpNum.Value) Then
        validInputs = False
        Exit Function
    End If
    validInputs = True
End Function
