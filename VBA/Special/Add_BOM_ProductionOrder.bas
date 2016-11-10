Attribute VB_Name = "Add_BOM_ProductionOrder"
Sub SetProductionBOM()
    'ATTACH TO SAP
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected
    Set session = SAPCon.Children(0)            ' Get the first session (window) on that connection
    
    session.findById("wnd[0]").maximize
    
    Dim i, j, scrollBarPos As Long
    Dim matNum, opNum, desc, qty, seq As String
    
    Dim numEntriesPerPage As Integer: numEntriesPerPage = 29
    
    j = 0 'Index for SAP
    scrollBarPos = 0 'Each increment here = decrement to j: while j>0
    
    'Scroll first blank item
    Do While True
    
        matNum = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & j & "]").Text
        
        
        'Scroll the page when hit the end of the list
        If j = numEntriesPerPage - 1 Then 'Zero-based indexing
            scrollBarPos = scrollBarPos + j + 1
            session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120").verticalScrollbar.Position = scrollBarPos
            j = 0
        End If
        
        If matNum = "" Then 'Break out once we hit first blank
            Exit Do
        End If
        
        j = j + 1
    Loop
    
    i = 2
    
    Do While True
        matNum = Cells(i, 1).Value
        opNum = Cells(i, 2).Value
        qty = Cells(i, 3).Value
        seq = Cells(i, 4).Value
        
        If matNum = "" Then 'Break out once we hit first blank
            session.findById("wnd[0]").sendVKey 0
            Exit Do
        End If
        
        session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & j & "]").Text = matNum
        session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-VORNR[2," & j & "]").Text = opNum
        session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[6," & j & "]").Text = qty
        session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRCOLS-APLFL[9," & j & "]").Text = seq
        session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-POSTP[8," & j & "]").Text = "L"
        
        session.findById("wnd[0]").sendVKey 0
        ' For confirmed operations
        If session.findById("wnd[0]/sbar").messagetype = "W" Then
            session.findById("wnd[0]").sendVKey 0
        End If

        ' Error for incorrect material number or wrong unit or not extended to plant
        If session.findById("wnd[0]/sbar").messagetype = "E" Then
            MsgBox session.findById("wnd[0]/sbar").Text
            End
        End If
        
        i = i + 1
        j = j + 1
        
        'Scroll the page when hit the end of the list
        If j = numEntriesPerPage - 1 Then 'Zero-based indexing
            scrollBarPos = scrollBarPos + j + 1
            session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120").verticalScrollbar.Position = scrollBarPos
            j = 0
        End If
        

        
    Loop
    
    
End Sub

