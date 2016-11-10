Attribute VB_Name = "CA02"
Option Explicit
Public Type sapWindow
    scrollBarPos As Long
    entriesPerPage As Integer
    
    currentIndex As Integer
End Type
Sub runscript()
    'Get the correct columns for this Excel sheet
    Dim descCol As Long: descCol = CInt(cmdWindow.tbDesc)
    Dim ctrCol As Long: ctrCol = CInt(cmdWindow.tbWorkCtr)
    Dim hoursCol As Long: hoursCol = CInt(cmdWindow.tbHours)
    
    Dim opCol As Long
    Dim errorCol As Long
    Dim useOpNums As Boolean
    Dim logError As Boolean
    
    If IsNumeric(cmdWindow.tbError) Then
        errorCol = CInt(cmdWindow.tbError)
        logError = True
    Else
        logError = False
    End If
    
    If IsNumeric(cmdWindow.tbOpNum) Then
        opCol = CInt(cmdWindow.tbOpNum)
        useOpNums = True
    Else
        useOpNums = False
    End If
    
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
    
    Call scrollToBlank(session, wnd)
    
    Dim curOP As COperation
    'Start looping through selected operations
    For i = startRow To lastRow
        If useOpNums Then
            Set curOP = factory.createCOperation(Cells(i, opCol), Cells(i, descCol), Cells(i, ctrCol), Cells(i, hoursCol), session)
        Else
            Set curOP = factory.createCOperation("", Cells(i, descCol), Cells(i, ctrCol), Cells(i, hoursCol), session)
        End If
        If curOP.isValidOperation Then
            curOP.enterOperation (wnd.currentIndex)
            session.findById("wnd[0]").sendVKey 0
            'Deal with possible errors
            wnd.currentIndex = wnd.currentIndex + 1
        ElseIf curOP.hasError Then
            Cells(i, ctrCol).Interior.ColorIndex = 44
        End If
    Next i
    

End Sub

Sub scrollToBlank(session As Variant, wnd As sapWindow)
    'Get information about CA02
    On Error GoTo notCA02Page
    wnd.entriesPerPage = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").VerticalScrollbar.pagesize
    
    wnd.scrollBarPos = 0 'Set the scrollbar position to 0 to start with
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").VerticalScrollbar.Position = wnd.scrollBarPos
    
    Dim workCtr As String ' Check the workCenter info column, if empty it means we hit the end of the list
    
    'Scroll first blank item
    Do While True
        'Scroll the page when hit the end of the page size
        If wnd.currentIndex = wnd.entriesPerPage - 1 Then 'Subtract because of zero based indexing
            wnd.scrollBarPos = wnd.scrollBarPos + wnd.currentIndex + 1
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").VerticalScrollbar.Position = wnd.scrollBarPos
            wnd.currentIndex = 0 'Reset index variable
        End If
        
        workCtr = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & wnd.currentIndex & "]").Text
        
        If workCtr = "" Then 'Break out once we hit first blank
            Exit Do
        End If
        
        wnd.currentIndex = wnd.currentIndex + 1
    Loop

Exit Sub
notCA02Page:
    MsgBox ("You are not in the 'change material Routing' page" & Chr(10) & "Navigate to the correct window before running this script")
    End
    
End Sub
