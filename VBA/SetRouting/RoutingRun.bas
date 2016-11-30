Attribute VB_Name = "RoutingRun"
'Lots of code duplication, will have to figure out how to avoid this
Option Explicit
'Define Constants
Public Const CA02 = 0
Public Const CO02 = 1
'SAP Window struct
Public Type sapWindow
    scrollBarPos As Long
    entriesPerPage As Integer
    
    currentIndex As Integer
End Type

Sub EnterOperations(descCol As Long, ctrCol As Long, hoursCol As Long, opCol As Long, mode As Byte)
    'Get the range the user has selected to determine the first and last cell to iterate between
    'And to determine columns
    Dim myRange As Range
    Set myRange = Selection
    
    Dim startRow As Long: startRow = myRange.Row 'Get the first row of the selection
    Dim lastRow As Long: lastRow = myRange.Rows.Count + startRow - 1 'Get the last row of the selection
    
    'Initialize custom type
    Dim wnd As sapWindow
    
    'ATTACH TO SAP
    Dim session As Variant
    Set session = SAPFunctions.connect2SAP() ' Get the first session (window) on that connection
    session.findById("wnd[0]").maximize
    
    Dim i As Long
    Dim matNum, qty As String
    
    
    Call scrollToBlankGen(session, wnd, mode)
    
    
    Dim curOP As Variant

    
    'Start looping through selected operations
    For i = startRow To lastRow
        If mode = CA02 Then
            Set curOP = factory.createCOperation(Cells(i, opCol), Cells(i, descCol), Cells(i, ctrCol), Cells(i, hoursCol), session)
        ElseIf mode = CO02 Then
            Set curOP = factory.createCOperationPO(Cells(i, opCol), Cells(i, descCol), Cells(i, ctrCol), Cells(i, hoursCol), session)
        End If
        If curOP.isValidOperation Then
            curOP.enterOperation (wnd.currentIndex)
            session.findById("wnd[0]").sendVKey 0
            'Deal with possible errors
            wnd.currentIndex = wnd.currentIndex + 1
        ElseIf curOP.hasError Then
            Rows(i).Interior.ColorIndex = 44
        End If
    Next i

End Sub

Private Sub scrollToBlankGen(session As Variant, wnd As sapWindow, mode As Byte)
    On Error GoTo notCorrectPage
    Dim scrollBarID, workCtrsID As String
    
    If mode = CA02 Then
        scrollBarID = "wnd[0]/usr/tblSAPLCPDITCTRL_1400"
    ElseIf mode = CO02 Then
        scrollBarID = "wnd[0]/usr/tblSAPLCOVGTCTRL_0100"
    End If
    
    wnd.entriesPerPage = session.findById(scrollBarID).VerticalScrollbar.pagesize
    
    wnd.scrollBarPos = 0 'Set the scrollbar position to 0 to start with
    session.findById(scrollBarID).VerticalScrollbar.Position = wnd.scrollBarPos
    
    Dim workCtr As String ' Check the workCenter info column, if empty it means we hit the end of the list
    
    'Scroll first blank item
    Do While True
        'Scroll the page when hit the end of the page size
        If wnd.currentIndex = wnd.entriesPerPage - 1 Then 'Subtract because of zero based indexing
            wnd.scrollBarPos = wnd.scrollBarPos + wnd.currentIndex + 1
            session.findById(scrollBarID).VerticalScrollbar.Position = wnd.scrollBarPos
            wnd.currentIndex = 0 'Reset index variable
        End If
        
        If mode = CA02 Then
            workCtrsID = "wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & wnd.currentIndex & "]"
        ElseIf mode = CO02 Then
            workCtrsID = "wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4," & wnd.currentIndex & "]"
        End If
        
        workCtr = session.findById(workCtrsID).text
        
        If workCtr = "" Then 'Break out once we hit first blank
            Exit Do
        End If
        
        wnd.currentIndex = wnd.currentIndex + 1
    Loop

Exit Sub
notCorrectPage:
    If mode = CA02 Then
        MsgBox ("You are not in the 'Change/Create material Routing' page" & Chr(10) & "Navigate to the correct window before running this script")
    ElseIf mode = CO02 Then
        MsgBox ("You are not in the 'Production Order Change' page" & Chr(10) & "Navigate to the correct window before running this script")
    End If
End Sub

