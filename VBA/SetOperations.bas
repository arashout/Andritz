Attribute VB_Name = "SetOperations"
Sub EnterOperations()
Attribute EnterOperations.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected
    Set session = SAPCon.Children(0)            ' Get the first session (window) on that connection
    
    'Get the range the user has selected to determine the first and last cell to iterate between
    Dim myRange As Range
    Set myRange = Selection
    
    Dim startI As Long: startI = myRange.Row 'Get the first row of the selection
    Dim lastI As Long: lastI = myRange.Rows.Count + i - 1 'Get the last row of the selection
    
    Dim opNumCol, descCol, hoursCol, workCol As Byte
    opNumCol = 3: descCol = 4: hoursCol = 5: workCol = 6
    
    Dim opNum As Long
    Dim desc As String
    Dim hours As Double
    Dim workCenter As String
    
    Dim SAPi As Byte: SAPi = 0
    
    For i = startI To lastI
        'Read the required values
        opNum = Cells(i, opNumCol).value
        desc = Cells(i, descCol).value
        hours = Cells(i, hoursCol).value
        workCenter = Cells(i, workCol).value
        If (desc <> "") Then 'If cell is not empty (Merged cells mess stuff up)
            'Enter these into the proper places in SAP
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0," & SAPi & "]").Text = opNum
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & SAPi & "]").Text = workCenter
            'SAP editor doesn't like when new lines in string
            If Right(desc, 1) = vbLf Then desc = Left(desc, Len(desc) - 1) 'Remove trailing new line
            If (Len(desc) > 40 Or InStr(desc, Chr(10))) Then 'If description longer than short text limit or has new lines in it
                Call EnterLongTextInEditor(desc, SAPi, session) 'Sub to enter descriptions in long text
            Else
                On Error GoTo InvalidMethod:
                session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & SAPi & "]").Text = desc
            End If
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19," & SAPi & "]").Text = hours
            SAPi = SAPi + 1
        End If
    Next i
    
    Exit Sub
InvalidMethod:
Debug.Print (desc)
    Call EnterLongTextInEditor(desc, SAPi, session)
    Resume Next
End Sub
Private Sub EnterLongTextInEditor(ByVal desc As String, ByVal currentRow As Byte, session As Variant)
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").getAbsoluteRow(currentRow).Selected = True 'Select row to edit long text
    session.findById("wnd[0]/tbar[1]/btn[16]").press 'Press the long text button
    
    Dim curPos As Long: curPos = 1
    Dim currentIndexEditor As Long: currentIndexEditor = 1
    Dim lineText As String: lineText = "" 'The current text to be enter
    Dim lineCount As Integer: lineCount = 1
    Dim curChar As String
    
    While curPos <= Len(desc)
        'In this loop the string is entered one character at a time
        curChar = popChar(curPos, desc)
        If (InStr(curChar, Chr(10))) Then
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & currentIndexEditor & "]").Text = lineText
            currentIndexEditor = currentIndexEditor + 1
            lineCount = 1
            lineText = ""
        Else
            lineText = lineText & curChar
            lineCount = lineCount + 1
        End If
        
        If lineCount >= 73 Then
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & currentIndexEditor & "]").Text = lineText
            currentIndexEditor = currentIndexEditor + 1
            lineCount = 1
            lineText = ""
        End If
        curPos = curPos + 1
    Wend
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").getAbsoluteRow(currentRow).Selected = False 'DeselectRow
End Sub
Private Function popChar(index As Long, theString As String) As String
    popChar = Mid(theString, index, 1)
End Function
