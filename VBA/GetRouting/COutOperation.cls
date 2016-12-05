VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COutOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public workCtr As String
Public hours As String
Public opNum As String
Public desc As String

Public Sub collectInfo(ByVal j As Integer, session As Variant)
    Dim cbLongtext As Boolean
    opNum = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0," & j & "]").text
    workCtr = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & j & "]").text
    hours = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19," & j & "]").text
    cbLongtext = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/chkRC270-TXTKZ[7," & j & "]").Selected
    If cbLongtext Then
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/chkRC270-TXTKZ[7," & j & "]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        Call readLongText(session)
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press the back button
    Else
        desc = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & j & "]").text
    End If
            
End Sub

Private Sub readLongText(session)
    Dim i As Integer: i = 1
    Dim curText, nextText As String
    Dim newLineChar As String
    Do While True
        newLineChar = session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i & "]").text
        curText = session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text
        nextText = session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i + 1 & "]").text
        If Len(curText) < 70 And nextText = "" Then
            desc = desc & curText & vbCrLf
        ElseIf newLineChar = "/" Then
            desc = desc & vbCrLf & curText
        ElseIf newLineChar = "*" Then
            desc = desc & curText
        ElseIf newLineChar = "__" Then 'End of longtext
            Exit Do
        End If
        i = i + 1
    Loop
End Sub
Public Sub addHeaders(r As GetRoutingMain.Routing, ws As Worksheet)
    ws.Cells(r.row, r.col + r.offsetOpNum).Value = "Op #"
    ws.Cells(r.row, r.col + r.offsetWorkCtr).Value = "Work Ctr"
    ws.Cells(r.row, r.col + r.offsetDesc).Value = "Description"
    ws.Cells(r.row, r.col + r.offsetHours).Value = "Hours"
    'Sequence related headers
    ws.Cells(r.row, r.col + r.offsetBranch).Value = "Branch"
    ws.Cells(r.row, r.col + r.offsetReturn).Value = "Return"
    'Color headers
    ws.Range(Cells(r.row, r.col + r.offsetSeq), Cells(r.row, r.col + r.countCol)).Interior.ColorIndex = 15
End Sub
Public Sub ouputOperation(r As GetRoutingMain.Routing, ws As Worksheet)
    ws.Cells(r.row, r.col + r.offsetOpNum).Value = opNum
    ws.Cells(r.row, r.col + r.offsetWorkCtr).Value = workCtr
    ws.Cells(r.row, r.col + r.offsetDesc).Value = desc
    ws.Cells(r.row, r.col + r.offsetHours).Value = hours
End Sub

