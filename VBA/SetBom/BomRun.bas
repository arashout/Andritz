Attribute VB_Name = "BomRun"
Option Explicit
'Define Constants
Public Const CS02 = 0
Public Const CO02 = 1

Sub runscriptBOM(SAPCol As Long, qtyCol As Long, seqCol As Long, opNumCol As Long, mode As Byte)
    'Get the range the user has selected to determine the first and last cell to iterate between
    'And to determine columns
    Dim myRange As Range
    Set myRange = Selection
    
    Dim startRow As Long: startRow = myRange.Row 'Get the first row of the selection
    Dim lastRow As Long: lastRow = myRange.Rows.Count + startRow - 1 'Get the last row of the selection
    
    'Initialize custom type
    Dim wnd As SAPFunctions.sapWindow
    
    'ATTACH TO SAP
    Dim session As Variant
    Set session = SAPFunctions.connect2SAP() ' Get the first session (window) on that connection
    
    session.findById("wnd[0]").maximize
    
    Dim i As Long
    Dim matNum, qty, seq, opNum As String
    
    If mode = CS02 Then
        wnd.scrollBarId = "wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT"
    ElseIf mode = CO02 Then
        wnd.scrollBarId = "wnd[0]/usr/tblSAPLCOMKTCTRL_0120"
    End If
    
    Call scrollToBlankGen(session, wnd, mode)
    
    For i = startRow To lastRow
        matNum = Cells(i, SAPCol).Value
        qty = Cells(i, qtyCol).Value
        
        If mode = CO02 Then
            If seqCol = 0 Then
                seq = 0
            Else
                seq = Cells(i, seqCol).Value
            End
            opNum = Cells(i, opNumCol).Value
        End If
        
        If Len(matNum) = 9 And IsNumeric(matNum) And IsNumeric(qty) Then
            If mode = CS02 Then
                session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[3," & wnd.currentIndex & "]").Text = matNum
                session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[5," & wnd.currentIndex & "]").Text = qty
                session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[2," & wnd.currentIndex & "]").Text = "L"
            ElseIf mode = CO02 Then
                session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & wnd.currentIndex & "]").Text = matNum
                session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-VORNR[2," & wnd.currentIndex & "]").Text = opNum
                session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRCOLS-APLFL[9," & wnd.currentIndex & "]").Text = seq
                session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[6," & wnd.currentIndex & "]").Text = qty
                session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-POSTP[8," & wnd.currentIndex & "]").Text = "L"
            End If
            
            session.findById("wnd[0]").sendVKey 0
            
            'For confirmed operations
            If session.findById("wnd[0]/sbar").messagetype = "W" Then
                session.findById("wnd[0]").sendVKey 0
            End If
    
            ' Error for incorrect material number or wrong unit or not extended to plant
            'NOTE: CO02 and CS02 handle errors differently
            If session.findById("wnd[0]/sbar").messagetype = "E" And mode = CS02 Then
                Cells(i, SAPCol).Interior.ColorIndex = 44
                session.findById("wnd[0]").sendVKey 12
            ElseIf session.findById("wnd[0]/sbar").messagetype = "E" And mode = CO02 Then
                Cells(i, SAPCol).Interior.ColorIndex = 44
            Else
                wnd.currentIndex = wnd.currentIndex + 1
                If wnd.currentIndex = wnd.entriesPerPage Then 'Zero-based indexing
                    Call scrollToNewPage(wnd)
                End If
            End If
            
        'Sometimes the user accidently mistypes material number
        ElseIf Len(matNum) > 7 And Len(matNum) < 11 And IsNumeric(matNum) And IsNumeric(qty) Then
            Cells(i, SAPCol).Interior.ColorIndex = 44
        End If
    Next i
    
End Sub
Private Sub scrollToBlankGen(session As Variant, wnd As SAPFunctions.sapWindow, mode As Byte)
    On Error GoTo notCorrectPage
    wnd.entriesPerPage = session.findById(wnd.scrollBarId).VerticalScrollbar.pageSize
    
    wnd.scrollBarPos = 0 'Set the scrollbar position to 0 to start with
    session.findById(wnd.scrollBarId).VerticalScrollbar.Position = wnd.scrollBarPos
    
    Dim matNum, matNumId As String
    
    'Scroll first blank item
    Do While True
        'Scroll the page when hit the end of the page size
        If wnd.currentIndex = wnd.entriesPerPage - 1 Then 'Subtract because of zero based indexing
            wnd.scrollBarPos = wnd.scrollBarPos + wnd.currentIndex + 1
            session.findById(wnd.scrollBarId).VerticalScrollbar.Position = wnd.scrollBarPos
            wnd.currentIndex = 0 'Reset index variable
        End If
        
        If mode = CS02 Then
            matNumId = "wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & wnd.currentIndex & "]"
        ElseIf mode = CO02 Then
            matNumId = "wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & wnd.currentIndex & "]"
        End If
        
        matNum = session.findById(matNumId).Text
        
        If matNum = "" Then 'Break out once we hit first blank
            Exit Do
        End If
        
        wnd.currentIndex = wnd.currentIndex + 1
    Loop

Exit Sub
notCorrectPage:
    If mode = CS02 Then
        MsgBox ("You are not in the 'Change/Create material BOM' page" & Chr(10) & "Navigate to the correct window before running this script")
    ElseIf mode = CO02 Then
        MsgBox ("You are not in the 'Production Order Change: Component Overview' page" & Chr(10) & "Navigate to the correct window before running this script")
    End If
    End
End Sub
Private Sub scrollToNewPage(wnd As SAPFunctions.sapWindow)
    'Scroll the page when hit the end of the list
    wnd.scrollBarPos = wnd.scrollBarPos + wnd.currentIndex + 1
    session.findById(wnd.scrollBarId).VerticalScrollbar.Position = wnd.scrollBarPos
    wnd.currentIndex = 0
End Sub
