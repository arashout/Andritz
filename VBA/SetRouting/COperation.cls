VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    validWorkCenters = Split("139,148,153,218,252,330,333,364,410,420,506,519,601,602,720,801,1013,1108,1207,1901,2004,2005,2006,2007,2008,NCR,999", ",")
    
    If IsNumeric(hours) And HelpFunctions.inArr(validWorkCenters, workCenter) Then
        isValidOperation = True
    'Possible scenarios
    ElseIf workCenter = "NCR" Then
        hours = 0
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
    If opNum <> "" Then 'Enter an operation if specified
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0," & j & "]").Text = opNum
    End If
    
    If workCenter = "999" Then
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4," & j & "]").Text = "ZP08"
    End If
    
    'Enter rest of information
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & j & "]").Text = workCenter 'Work Center
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19," & j & "]").Text = hours 'Hours
    
    If Len(desc) > shortTextLimit Then
        'Press on long text checkbox
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/chkRC270-TXTKZ[7," & j & "]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        Call enterLongText
        'Check if the stupid 'Do you want to edit next long text' pops up
        If session.Children().Count = 2 Then
            session.findById("wnd[1]").sendVKey 12
        End If
        
    Else
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & j & "]").Text = desc
    End If
    
End Sub
'''''''''''''''''''''
'Private Methods
'''''''''''''''''''''
Private Sub enterLongText()
    Dim i As Integer
    For i = LBound(arrDesc) To UBound(arrDesc)
        'looping through array here
        session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i + 1 & "]").Text = arrDesc(i)
        If lineFormats(i) = "/" Then 'Line feed
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i + 1 & "]").Text = lineFormats(i)
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
            currentLine = ""
        End If
    Next i
    'Whatever is left, put in array
    arrDesc(UBound(arrDesc)) = currentLine
    
    arrDesc = HelpFunctions.removeFirstElement(arrDesc)
    
End Sub


