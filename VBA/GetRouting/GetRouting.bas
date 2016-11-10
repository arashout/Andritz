Attribute VB_Name = "GetRouting"
'This sub gets all the operation and sequences from a routing
Sub GetRouting()
    'Confirm that the user wants to run this on THIS SHEET, if NO exit the sub
    confirmBox = MsgBox("This macro will overwrite data on THIS SHEET, do you want to continue?", _
        vbYesNo, "Run this macro?")
        If confirmBox = vbNo Then Exit Sub
        
    'Speed-Up Excel
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim materialNum As Variant
    Dim View As Boolean
    View = False
    
    materialNum = InputBox("Enter 9 Digit Material Number") 'Prompts user to input material number
    If (materialNum = "") Then
        MsgBox ("You need to enter something in the material number box")
        Exit Sub
    End If
    ' SAP Automation Code
    ' Needed for program to function
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected

    Dim OriginalChildren As Integer
    OriginalChildren = SAPCon.Children.count
    
    Set session = SAPCon.Children(0)            ' Get the first session (window) on that connection
    
    'Navigate to Routing
    'Step 1 - Navigate to Correct Transaction
    session.findById("wnd[0]").maximize 'Maximize the SAP window - Although I don't know if this is strictly necessary
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCA03" 'Type "/nCA03" in command field
    session.findById("wnd[0]").sendVKey 0 'Press the enter key
    
    'Navigate to Routing
    'Step 2 - Navigate to Correct Material Number Routing
    session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = materialNum   'Enter MaterialNumber
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").text = "1105"        'Enter Plant Number
    session.findById("wnd[0]/usr/ctxtRC271-STTAG").text = "01/01/2012"  'Enter Key Date
    session.findById("wnd[0]/tbar[1]/btn[7]").press                     'Navigate to list of routing
    RoutingCount = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text 'Get total number of routings
    
    'IMPORTANT:
    'If more than one routing, must prompt user to select the correct routing
    If RoutingCount > 1 Then
        MsgBox ("Please pick the routing you would like to pull in SAP")
    End If
        
    'Nav to Sequences - Count # of Seq - Get Title and Number of Sequence
    session.findById("wnd[0]/tbar[1]/btn[6]").press                         'Navigate to Sequences
    SequenceCount = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text    'Get total number of sequences
    Dim SequenceSpace() As Integer                                          'Create a variable to hold the row number of every sequence entry in spreadsheet
    ReDim SequenceSpace(SequenceCount - 1)                                  'Resize SequenceSpace variable
        
    'Navigate to operations in standard sequence and get title of routing from operation 1
    session.findById("wnd[0]/tbar[1]/btn[7]").press                                             'Navigate to operations in standard sequence
    TitleText = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,0]").text   'Get Title of Routing from Op 1 - assign to TitleText variable
    Call SetUpWorkbook(TitleText, materialNum)                                                  'Run SetUpWorkBook sequence, pass material number and TitleText as variable, return name of workbook
    session.findById("wnd[0]/tbar[1]/btn[29]").press                                            'Navigate back to list of sequences
    
    a = 3 'List the row of the current sequence description in excel spreadsheet
    
    'These nested loops do most of the important work
    'The outer loop pulls the sequence name and number and puts this in a row
    'The inner loop pulls the operation hours, description, and other info and puts them in a row
    For j = 0 To SequenceCount - 1
        SequenceSpace(j) = a 'Input the row of the current sequence into the SequenceSpace variable
        
        'Get Sequence Name and Number from SAP
        'Put SAP addresses in strings
        StringSeqNum = "wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtPLFLD-PLNFL[0," & j & "]"   'Sequence Number String address
        StringSeqDesc = "wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtPLFLD-LTXA1[7," & j & "]"  'Sequence Description address
        'Pull Name and Number from SAP using addresses from above
        SeqNum = session.findById(StringSeqNum).text    'Pull Sequence Number
        SeqDesc = session.findById(StringSeqDesc).text  'Pull Sequence Description
        
        'Format cells and input sequence name and number
        Range(Cells(a, 2), Cells(a, 4)).Merge                               'Merge cells for sequence description
        Cells(a, 2).Value = SeqNum & "/ " & SeqDesc                         'Put sequence name and number in sequence description rows
        Range(Cells(a, 2), Cells(a, 6)).Interior.Color = RGB(204, 192, 218) 'Set background colour of sequence cells
        Range(Cells(a, 2), Cells(a, 6)).Font.Bold = True                    'Bold sequence text
        
        'Prepare to pull operation information
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300").getAbsoluteRow(j).Selected = True  'Select current sequence in SAP
        session.findById("wnd[0]/tbar[1]/btn[7]").press                                         'Navigate to list of operations in SAP
        LastEntry = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text                        'Determine number of operation in sequence and assign to variable LastEntry
        
        'Pull operation info from SAP
        For i = 0 To LastEntry - 1
            'Create string with addresses for different info
            StringDesc = "wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & i & "]"     'Short Text Description
            StringLongText = "wnd[0]/usr/tblSAPLCPDITCTRL_1400/chkRC270-TXTKZ[7," & i & "]" 'Long Text Description
            StringOp = "wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0," & i & "]"       'Operation Number
            StringWC = "wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & i & "]"      'Work Centre
            StringHr = "wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19," & i & "]"      'Hours
            
            'Pull Operation Description
            'Not all operations have long text
            'Start by checking for long text
            'If no long text exists, use short text
            If session.findById(StringLongText).Selected = True Then    'Check if Long Text exists
                desc = ReadLongText(i, View)                                  'Use a function to pull long text if long text exists
            Else
                desc = session.findById(StringDesc).text                'Use short text if long text does not exist
            End If
            
            'Pull remaining operation info
            Op = session.findById(StringOp).text    'Pull Operation Number
            WC = session.findById(StringWC).text    'Pull Work Centre
            Hr = session.findById(StringHr).text    'Pull Hours
            
            'Put all operation info in excel (in appropriate column)
            Cells(a + 1 + i, 4).Value = desc
            Cells(a + 1 + i, 5).Value = Hr
            Cells(a + 1 + i, 6).Value = WC
            Cells(a + 1 + i, 3).Value = Op
            
            'Format the cells
            Cells(a + 1 + i, 4).WrapText = True 'Wrap Text
            Rows(a + 1 + i).AutoFit 'Autfit row height
            
        Next i 'Move onto next operation
        
        'Sum hours for the current sequence in the sequence row
        Cells(a, 5).Formula = "=SUM(" & Range(Cells(a + 1, 5), Cells(a + LastEntry, 5)).Address(False, False) & ")" 'Input formula into cell
        Cells(a, 5).NumberFormat = "#.##"" hrs""" 'Format hours cell to include some text at the end of the number " hrs"
        
        'Prepare for next sequence
        a = a + LastEntry + 1                                                                   'Assign new value for new sequence row
        session.findById("wnd[0]/tbar[1]/btn[29]").press                                        'Navigate back to list of sequences in SAP
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300").getAbsoluteRow(j).Selected = False 'Deselect current seq
    Next j
    
    'Sum total hours for every op and put at top of sheet
    Sumstring = "=E" 'Initialize equation
    For i = 0 To SequenceCount - 2
        Sumstring = Sumstring & SequenceSpace(i) & " + $E$" 'Use SequenceSpace array to populate sumstring equation with a sum of all sequence sum hours
    Next i
    
    Sumstring = Sumstring & SequenceSpace(SequenceCount - 1)    'Finish formula for last entry
    Cells(1, 5).Formula = Sumstring                             'Put formula in cell
    Cells(1, 5).NumberFormat = "#"" hrs"""                      'Format cell to include text
    
    Columns("B:C").AutoFit
    Columns("E:G").AutoFit
    
    'Reset Settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub SetUpWorkbook(TitleText, materialNum)
    'Sets up the title of the workbook in the spreadsheet
    'Sets up column headers in the spreadsheet
    'Sets column width of operation description coluumn
    'Sets fill colour of certain heading cells
    'Returns name of workbook as variable
    
    'Thought adding date to sheet name but reconsidered
    'Dim strDate As String
    'strDate = Format(Now(), "yyyymmdd")
    Sheets.Add.Name = TitleText
    
    'Set up column headers
    Cells(2, 2).Value = "SEQ"
    Cells(2, 3).Value = "Op #"
    Cells(2, 4).Value = "Description"
    Cells(2, 5).Value = "Hours"
    Cells(2, 6).Value = "Work Centre"
    
    'Set up spreadsheet header
    Range(Cells(1, 1), Cells(1, 4)).Merge 'Merge cells for title
    Cells(1, 1).Value = TitleText & ": " & materialNum 'Create title with name of routing and material number
    'Cells(1, 7).Value = MaterialNum 'Add material number in another cell as well
    
    'Set column widths and colours
    Columns("D:D").ColumnWidth = 78 'Set Column Width for material description column
    Range(Cells(1, 1), Cells(1, 6)).Interior.Color = RGB(253, 234, 218) 'Set colour of title cells
    Range(Cells(1, 1), Cells(1, 6)).Font.Bold = True 'Bold title cells
    Range(Cells(2, 2), Cells(2, 6)).Interior.Color = RGB(242, 242, 242) 'Set colour of column heading sells
    
End Sub
Function ReadLongText(OpNum, View)
    'Take operation number as input and returns the operations long text
    'Also takes a boolean operator, telling function whether or not we are in the correct view
    'Changes long text view to a different view from the default view - easier to interact with this new view
    
    Dim ParaVar As String       'Holds information regarding formatting of long text for current line (whether you skip a line and such, more info available in SAP)
    Dim longText As String      'Holds long text for current line
    Dim LongTextOut As String   'Holds long text for entire document - includes info about whether or not to skip a line and such (ParaVar)
    Dim searchString As String  'Holds address of LongText (in SAP)
    Dim ParaString As String    'Holds address of Paragraph info (in SAP)
    
    'This code work by starting from the top and working until there is no more text.
    'When there is no more text, the SAP will give a string = "___..."
    'We can use this let us know when we've reached the end of the document
    'We also use the far left column in SAP (ParaVar) to give us info about the formatting and whether or not to start a new line

    'SAP CODE THAT IS NEEDED FOR PROPER FUNCTION
    'NOT VERY MUCH FUN
    Set SapGuiAuto = GetObject("SAPGUI")        ' Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  ' Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             ' Get the first system that is currently connected

    Dim OriginalChildren As Integer
    OriginalChildren = SAPCon.Children.count
    
    Set session = SAPCon.Children(0)            ' Get the first session (window) on that connection
    
    'THIS IS WHERE THE FUN BEGINS!
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").getAbsoluteRow(OpNum).Selected = True  'Select the operation in SAP
    session.findById("wnd[0]/tbar[1]/btn[16]").press                                            'Open long text
    
    'Change to the useful view - the default view is completely useless
    'Only change if not already in the correct view
    If View = False Then
        session.findById("wnd[0]/mbar/menu[2]/menu[3]").Select                                                                              'Open the menu
        session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0800/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0800/chkRSEUMOD-GRA_EDITOR").Selected = False  'Select the correct setting
        session.findById("wnd[1]/tbar[0]/btn[0]").press                                                                                     'Close the menu
        View = True                                                                                                                         'We are now in correct view, don't need to this step again
    End If
    
    'Initialize some valeus for the first run through the while loop
    count = 1
    searchString = "wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & count & "]"
    ParaString = "wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & count & "]"
    ParaVar = session.findById(ParaString).text
    longText = session.findById(searchString).text
    
        
    While longText <> "________________________________________________________________________" 'Stop running the loop when we reach the end of the document
        If ParaVar = "/" Then
            LongTextOut = LongTextOut & Chr(10) & longText & " " 'Skip a line
        Else
            LongTextOut = LongTextOut & longText & " " 'Dont skip a line, insert a space (" ")
        End If
                
        'Get new values
        count = count + 1
        searchString = "wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & count & "]"
        ParaString = "wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & count & "]"
        ParaVar = session.findById(ParaString).text
        longText = session.findById(searchString).text
    Wend
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Exit long text
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").getAbsoluteRow(OpNum).Selected = False 'Deselect op
    
    ReadLongText = LongTextOut 'Return long text
End Function
