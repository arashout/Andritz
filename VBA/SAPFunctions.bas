Attribute VB_Name = "SAPFunctions"
Function connect2SAPNew()
    On Error GoTo Pull_Error_General
    
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected
    
    Dim OriginalChildren As Integer
    OriginalChildren = SAPCon.Children.Count
    
    If OriginalChildren > 4 Then GoTo Pull_Error_Many_SAP
    
    Set session = SAPCon.Children(0)            ' Get the first session (window) on that connection
    session.createSession                       ' Create a new session above the current sessions
    
    Do While SAPCon.Children.Count = OriginalChildren
    Loop
    
    Set session = SAPCon.Children(SAPCon.Children.Count - 1)  ' set the session to be the new window
    Set connect2SAPNew = session
    
    Exit Function
    
Pull_Error_Many_SAP:
    MsgBox ("Please close at least one SAP window as creating a new one to run the COOIS from will hit the limit.")
    
    End
    
Pull_Error_General:
    MsgBox ("Error: Trouble communicating with SAP. Please make sure SAP is running and you are logged in to the ASAP Production System.")
    
    End
End Function

    
Function connect2SAP()
    On Error GoTo Pull_Error_General
    
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected
    Set session = SAPCon.Children(0)            ' Set the session
    
    Set connect2SAP = session
    Exit Function
    
Pull_Error_Many_SAP:
    MsgBox ("Please close at least one SAP window as creating a new one to run the COOIS from will hit the limit.")
    
    End
    
Pull_Error_General:
    MsgBox ("Error: Trouble communicating with SAP. Please make sure SAP is running and you are logged in to the ASAP Production System.")
    
    End
End Function
    
