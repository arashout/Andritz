Attribute VB_Name = "GetLongText"
Dim session
Sub getLongText()
    Dim answer As Integer
    answer = MsgBox("Replace cells adjacent to selection with long text?", vbYesNo + vbQuestion, "Replace Long Text")
    
    If answer = vbNo Then
        Exit Sub
    End If
    
    'Connect to SAP
    Dim session
    Set session = SAPFunctions.connect2SAPNew()
    
    'Declare variables to use
    Dim matNum, i As Long
    Dim matNumCol, descCol As Integer
    Dim longText As String
    
    'Get the range the user has selected to determine the first and last cell to iterate between
    'And to determine columns
    Dim myRange As Range
    Set myRange = Selection
    
    Dim startI As Long: startI = myRange.Row 'Get the first row of the selection
    Dim lastI As Long: lastI = myRange.Rows.count + startI - 1 'Get the last row of the selection
    
    matNumCol = myRange.Column
    descCol = matNumCol + 1
    
    For i = startI To lastI 'Start at active cell location
        If VarType(Cells(i, matNumCol).Value) = vbDouble Then
            matNum = Cells(i, matNumCol)
            session.SendCommand ("/nmm03")
            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = matNum
            session.findById("wnd[0]").sendVKey 0
            
            'If any errors pop-up - log them then move on
            If session.findById("wnd[0]/sbar").messagetype = "E" Then
                    'Escape out of the error message
                    'session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-POSNR[0," & j & "]").SetFocus
                    'session.findById("wnd[0]").sendVKey 12 'Press Escape
                    'Log it
                    Cells(i, matNumCol).Interior.ColorIndex = 3 'Color Row RED
            Else
                session.findById("wnd[1]/tbar[0]/btn[19]").press
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True
                session.findById("wnd[1]").sendVKey 0
                longText = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2005/subSUB3:SAPLZMM00_ASTMGD1:2002/txtZRAST-TEXTAST").Text
                Cells(i, descCol).Value = longText
            End If
            

        Else
        End If
        
    Next i
End Sub
