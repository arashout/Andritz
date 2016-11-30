VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COperationPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NOTE: VBA doesn't support inheritance so instead I will use interfaces
Option Explicit
'''''''''''''''
'Properties
'''''''''''''''
Public opNum As String
Public desc As String
Public workCenter As String
Public hours As String
Public session As Variant
Public isValidOperation As Boolean
Public hasError As Boolean
Public shortTextLimit As Integer


Private arrDesc() As String 'Property to hold array of strings to enter into editor
Private lineFormats() As String 'Property to hold array of line formats for each desc

''''''''''''''''''''''
' Pretend Constructor
''''''''''''''''''''''
Public Sub initCOperation(fopNum As String, fdesc As String, fworkCenter As String, fhours As String, fsession As Variant)
    shortTextLimit = 40
    'Have to use different variable names when assigning! Added 'f' in front of names for factory
    opNum = fopNum
    desc = fdesc
    workCenter = fworkCenter
    hours = fhours
    
    'VALIDATE
    Dim validWorkCenters() As String
    validWorkCenters = Split("139,148,153,218,252,330,333,364,410,420,506,519,601,602,720,801,1013,1108,1207,1901,2004,2005,2006,2007,2008,NCR", ",")
    
    If IsNumeric(hours) And HelpFunctions.inArr(validWorkCenters, workCenter) Then
        isValidOperation = True
    'Possible scenarios
    ElseIf workCenter = "999" Then
        hours = "0"
        isValidOperation = True
    ElseIf workCenter = "NCR" Then
        hours = "0"
        isValidOperation = True
    ElseIf workCenter = "SUB" Then
        hasError = True
        Exit Sub
    ElseIf workCenter <> "" Then
        hasError = True
        Exit Sub
    End If
    
    Set session = fsession
    
    If Len(desc) > shortTextLimit Then
        Call splitDescription
    End If
End Sub

'''''''''''''''''''''
'Public Methods
'''''''''''''''''''''
Public Sub enterOperation(j As Integer)
    'Enter an operation
    session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & j & "]").text = opNum
    
    If workCenter = "999" Then
        session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6," & j & "]").text = "ZP08"
    End If
    
    session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4," & j & "]").text = workCenter 'Work Center
        
    If Len(desc) > shortTextLimit Then
        Call enterHours
        'Press on long text checkbox
        session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkRC270-TXTKZ[9," & j & "]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        Call enterLongText
        'Check if the stupid 'Do you want to edit next long text' pops up and exit it
        If session.Children().Count = 2 Then
            session.findById("wnd[1]").sendVKey 12
        End If
        
    Else
        
        session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & j & "]").text = desc
        session.findById("wnd[0]").sendVKey 0 'Enter into operation
        Call enterHours
    End If
    
End Sub
'''''''''''''''''''''
'Private Methods
'''''''''''''''''''''
Private Sub enterHours()
    session.findById("wnd[0]").sendVKey 0 'Reach the standard hours screen
    session.findById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW02").text = hours
    session.findById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE02").text = "HUR"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0 'Finish will get us back to 'Production Order Change'
End Sub
Private Sub enterLongText()
    Dim i As Integer
    For i = LBound(arrDesc) To UBound(arrDesc)
        'looping through array here
        session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i + 1 & "]").text = arrDesc(i)
        If lineFormats(i) = "/" Then 'Line feed
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i + 1 & "]").text = lineFormats(i)
        End If
        session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i + 1 & "]").caretPosition = 72
        session.findById("wnd[0]").sendVKey 0
    Next i
    session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub
Private Sub splitDescription()
    Dim longTextLineLimit As Integer: longTextLineLimit = 71
    
    Dim i As Long
    Dim currentCharCount As Integer
    Dim currentChar As String
    Dim currentLine As String
    
    ReDim arrDesc(1) As String
    ReDim lineFormats(1) As String
    
    For i = 1 To Len(desc)
        currentChar = HelpFunctions.popChar(i, desc)
        If currentChar = Chr(10) Then
            'Add current line as usual
            arrDesc(UBound(arrDesc)) = currentLine
            ReDim Preserve arrDesc(0 To UBound(arrDesc) + 1) As String
            'The line after the current one needs a line feed at the beginning LOL WHAT
            ReDim Preserve lineFormats(0 To UBound(lineFormats) + 1) As String
            lineFormats(UBound(lineFormats)) = "/"
            'Reset current line
            currentLine = ""
        Else
            currentLine = currentLine + currentChar
        End If
        
        If Len(currentLine) > longTextLineLimit Then
            arrDesc(UBound(arrDesc)) = currentLine
            ReDim Preserve arrDesc(0 To UBound(arrDesc) + 1) As String
            
            ReDim Preserve lineFormats(0 To UBound(lineFormats) + 1) As String
            lineFormats(UBound(lineFormats)) = "*"
            
            currentLine = ""
        End If
    Next i
    'Whatever is left, put in array
    arrDesc(UBound(arrDesc)) = currentLine
    
    'Get rid of unnessary first values
    lineFormats = HelpFunctions.removeFirstElement(lineFormats)
    arrDesc = HelpFunctions.removeFirstElement(arrDesc)
    
End Sub



