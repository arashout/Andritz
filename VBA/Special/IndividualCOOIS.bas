Attribute VB_Name = "IndividualCOOIS"
Sub jobCOOIS()
    Dim user_input As String
    Dim choice As Integer
    user_input = InputBox("Enter a job number or WBS element")
    
    'First test user input before opening SAP
    
    'User input is job number
    If IsNumeric(user_input) And Len(user_input) = 8 Then
        choice = 1
    'User input is wbs element
    ElseIf Len(user_input) = 15 Or Len(user_input) = 20 Then
        choice = 2
    'User pressed cancel or entered nothing
    ElseIf Len(user_input) = 0 Then
        Exit Sub
    Else
        MsgBox ("Unrecogized user input, try again or contact developer")
        Exit Sub
    End If
    
    Dim session
    Set session = SAPFunctions.connect2SAPNew()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOO000"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/DELTA OPS"
    
    'Leaving possibility for other choices later on
    If choice = 1 Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_AUFNR-LOW").text = user_input
    Else
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PROJN-LOW").text = user_input
    End If
    
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    On Error GoTo NoMessageBox
    If InStr(session.findById("wnd[1]/usr/txtMESSTXT1").text, "There is no data for the selection") Then
        session.findById("wnd[0]").Close
        MsgBox ("(SAP says) There is no data for the selection. Double-Check your entry")
        Exit Sub
    End If
    
NoMessageBox:

    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&XXL"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    'Name the file
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = user_input & "_COOIS.xlsx"
    
    session.findById("wnd[1]").sendVKey 0
    
    session.findById("wnd[0]").Close
End Sub

