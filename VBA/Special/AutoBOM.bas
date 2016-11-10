Attribute VB_Name = "AutoBOM"
Dim session
Sub fillBOM()
Attribute fillBOM.VB_ProcData.VB_Invoke_Func = "q\n14"
    ' SAP Automation Code
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected
    Set session = SAPCon.Children(0)            ' Get the first session (window) on that connection
    
    session.findById("wnd[0]").maximize 'Maximize the window so you know what's going on
    
    'Declaring variables used
    Dim i As Long 'Excel Sheet Index
    Dim j As Integer 'SAP GUI Row Index
    Dim itemNum As Long 'SAP GUI item number
    Dim matNum As Long
    Dim quantity As Double
    Dim matNumCol As Integer
    Dim quantityCol As Integer
    Dim errorTextField As String
    Dim errorLogCol As Integer
    Dim countItems As Integer
    Dim scrollBarPos As Integer
    Dim iCt As String 'Item category
    
    'Set-up Variables
    countItems = 1
    matNumCol = 4
    quantityCol = 1
    errorLogCol = 14
    j = 0
    
    'Navigate to BOM
    'Command textbox
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCS02"
    session.findById("wnd[0]").sendVKey 0 'Press Enter
    'Picking BOM screen
    session.findById("wnd[0]/usr/ctxtRC29N-MATNR").Text = InputBox("Enter the SAP BOM number")
    session.findById("wnd[0]/usr/ctxtRC29N-STLAN").Text = "c"
    session.findById("wnd[0]").sendVKey 0 'Press Enter
    'Arrived at BOM
    
    'If the ID of the vertical scroll bar is not found, this means that this BOM has alternatives BOMs.
    
    'Initialize scroll position to zero for consistency
    On Error GoTo altBOM:
    session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").verticalScrollbar.Position = 0
    
    scrollBarPos = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").verticalScrollbar.Position
    
    'Read the first ICt entry, in the while loop below we increment 'j' (the SAP row index) until the first empty ICt entry
    iCt = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[2," & j & "]").Text
    'This loop determines which row to start on SAP, reading until the next empty ICt entry
    Do While Len(iCt) > 0
        'This if statement moves the scroll bar down if we have reached the bottom of the page
        If countItems > 24 Then
            scrollBarPos = scrollBarPos + countItems
            countItems = 1
            session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").verticalScrollbar.Position = scrollBarPos 'Change vertical scroll
            'Have to reset the row index on SAP when scrolling down
            j = 1
        Else
            j = j + 1
            countItems = 1 + countItems
        End If
        itemNum = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-POSNR[0," & j & "]").Text
        iCt = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[2," & j & "]").Text
    Loop

    'Main loop here
    'The Excel row index starts at which ever cell is selected
    For i = ActiveCell.Row To HelpFunctions.lastRow
        'Check that the row as a quantity and a SAP material number
        If VarType(Cells(i, quantityCol).Value) = vbDouble And VarType(Cells(i, matNumCol).Value) = vbDouble Then
            quantity = Cells(i, quantityCol).Value
            matNum = Cells(i, matNumCol).Value
            'Only add materials with a quantity higher than 0
            If quantity > 0 Then
                'Actual data entry happens here
                session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-POSNR[0," & j & "]").Text = itemNum
                session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[2," & j & "]").Text = "L"
                session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[3," & j & "]").Text = matNum
                session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[5," & j & "]").Text = quantity
                session.findById("wnd[0]").sendVKey 0
                
                'If any errors pop-up - log them then move on
                If session.findById("wnd[0]/sbar").messagetype = "E" Then
                    errorTextField = session.findById("wnd[0]/sbar").Text
                    'Escape out of the error message
                    session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-POSNR[0," & j & "]").SetFocus
                    session.findById("wnd[0]").sendVKey 12 'Press Escape
                    'Log it
                    Cells(i, errorLogCol).Value = errorTextField
                    Cells(i, errorLogCol).EntireRow.Interior.ColorIndex = 3 'Color Row RED
                'Else entering the entry was a success
                Else
                    Cells(i, matNumCol).Interior.ColorIndex = 4 'Color Cell green
                    itemNum = itemNum + 10
                    'This if statement moves the scroll bar down if we have reached the bottom of the page
                    If countItems > 24 Then
                        scrollBarPos = scrollBarPos + countItems
                        countItems = 1
                        session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").verticalScrollbar.Position = scrollBarPos 'Change vertical scroll
                        'Have to reset the row index on SAP when scrolling down
                        j = 1
                    Else
                        j = 1 + j
                        countItems = 1 + countItems
                    End If
                End If

            End If
        End If
    Next i
    'Ensure to put the project on hold and save here
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    'Default is to put the project on hold
    session.findById("wnd[0]/usr/tabsTS_HEAD/tabpKHPT/ssubSUBPAGE:SAPLCSDI:1110/ctxtRC29K-STLST").Text = "10"
    session.findById("wnd[0]/usr/tabsTS_HEAD/tabpKHPT/ssubSUBPAGE:SAPLCSDI:1110/ctxtRC29K-STLST").SetFocus
    'Let user save by themselves
    Exit Sub
    'Here the user is asked to choose which alternative BOM
altBOM:
    Dim choice As Integer
    choice = InputBox("Enter an integer to indicate which BOM you want to fill")
    session.findById("wnd[0]/usr/tblSAPLCSDITCALT/txtRC29K-STLAL[0," & choice - 1 & "]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").verticalScrollbar.Position = 0
    Resume Next
End Sub
