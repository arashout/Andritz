VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public num As String
Public desc As String
Public seqType As String
Public branchOperation As String
Public returnOperation As String
Public countOperations As Integer
Public operations As Collection
Public rowPosition As Integer
Private Sub Class_Initialize()
    
End Sub

Public Sub collectInfo(ByVal i As Integer, session As Variant)
    'Get sequence information
    num = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtPLFLD-PLNFL[0," & i & "]").text
    desc = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtPLFLD-LTXA1[7," & i & "]").text
    seqType = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtTCA08-FLG_STXT[2," & i & "]").text
    branchOperation = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtPLFLD-VORNR1[5," & i & "]").text
    returnOperation = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtPLFLD-VORNR2[6," & i & "]").text
End Sub
Public Sub collectOperations(session As Variant)
    Dim j As Integer
    Dim curOp As COutOperation
    Set operations = New Collection
    countOperations = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text
    For j = 0 To countOperations - 1
        Set curOp = New COutOperation
        'TODO - Get Page Size
        Call curOp.collectInfo(j, session)
        operations.Add curOp
    Next j
    Set curOp = Nothing
End Sub
Public Sub outputAll(r As GetRoutingMain.Routing, ws As Worksheet)
    Call outputSeqHeaders(r, ws)
    r.row = r.row + 1
    Call outputOperations(r, ws)
End Sub
Private Sub outputSeqHeaders(r As GetRoutingMain.Routing, ws As Worksheet)
    rowPosition = r.row 'Used for keep track of where each sequence starts
    ws.Cells(r.row, r.col + r.offsetSeq).Value = num
    'ws.Cells(r.row, r.col + 1).Value = seqType
    ws.Cells(r.row, r.col + r.offsetDesc).Value = desc
    ws.Cells(r.row, r.col + r.offsetHours).Formula = "=SUM(" & Range(Cells(r.row + 1, r.col + r.offsetHours), Cells(r.row + countOperations, r.col + r.offsetHours)).Address(False, False) & ")"
    ws.Cells(r.row, r.col + r.offsetBranch).Value = branchOperation
    ws.Cells(r.row, r.col + r.offsetReturn).Value = returnOperation
    'Color sequences to make them stand out
    ws.Range(Cells(r.row, r.col + r.offsetSeq), Cells(r.row, r.col + r.countCol)).Interior.ColorIndex = 39
    'Also append the total hours formula for the routing
    ws.Cells(r.startRow, r.col + r.offsetHours).Formula = ws.Cells(r.startRow, r.col + r.offsetHours).Formula & "+" _
    & HelpFunctions.ColumnLetter(r.col + r.offsetHours) & rowPosition
End Sub

Private Sub outputOperations(r As GetRoutingMain.Routing, ws As Worksheet)
    Dim op As COutOperation
    Dim i As Integer
    rowPosition = r.row
    For i = 1 To operations.Count
        Set op = operations.Item(i)
        Call op.ouputOperation(r, ws)
        r.row = r.row + 1
    Next i
    r.row = r.row + 1 'Extra space between sequences
    Set op = Nothing
End Sub
