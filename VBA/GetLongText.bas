Attribute VB_Name = "GetLongText"
Dim session
Sub getLongText()
MsgBox ("This macro will get the long text from SAP starting from the ACTIVE cell row")
' SAP Automation Code
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected
    Set session = SAPCon.Children(0)            ' Get the first session (window) on that connection
    
    session.findById("wnd[0]").maximize 'Maximize the window so you know what's going on
    
    'Declare variables to use
    Dim matNum, i As Long
    Dim matNumCol, descCol, errorLogCol As Integer
    Dim longText, errorTextField As String
    
    matNumCol = CInt(InputBox("Enter an integer to indicate which column contains SAP #"))
    descCol = CInt(InputBox("Enter an integer to indicate which column you want the long text in"))
    errorLogCol = CInt(InputBox("Enter an integer to indicate which column you want error log information to be in"))
    
    'Chose which rows contain information
    'matNumCol = 2
    'descCol = 3
    'errorLogCol = 11
    
    For i = ActiveCell.Row To HelpFunctions.lastRow 'Start at active cell location
        If VarType(Cells(i, matNumCol).Value) = vbDouble Then
            matNum = Cells(i, matNumCol)
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmm03"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = matNum
            'If any errors pop-up - log them then move on
            If session.findById("wnd[0]/sbar").messagetype = "E" Then
                    errorTextField = session.findById("wnd[0]/sbar").Text
                    'Escape out of the error message
                    session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-POSNR[0," & j & "]").SetFocus
                    session.findById("wnd[0]").sendVKey 12 'Press Escape
                    'Log it
                    Cells(i, errorLogCol).Value = errorTextField
                    Cells(i, errorLogCol).EntireRow.Interior.ColorIndex = 3 'Color Row RED
            End If
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[1]").sendVKey 0
            longText = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2005/subSUB3:SAPLZMM00_ASTMGD1:2002/txtZRAST-TEXTAST").Text
            Cells(i, descCol).Value = longText
        Else
        End If
        
    Next i
End Sub
