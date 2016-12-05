Attribute VB_Name = "GetRoutingMain"
Option Explicit
Public Type Routing
    materialNum As String
    plantNum As String
    title As String
    countSequences As Integer
    sequences As Collection
    
    offsetSeq As Integer
    offsetSeqType As Integer
    offsetOpNum As Integer
    offsetWorkCtr As Integer
    offsetDesc As Integer
    offsetHours As Integer
    offsetBranch As Integer
    offsetReturn As Integer
    startRow As Integer
    countCol As Integer
    
    row As Long
    col As Long
End Type
Public Sub GetRouting()
    Dim r As Routing
    r.materialNum = InputBox("Enter material number for routing")
    r.materialNum = Trim(r.materialNum)
    
    'End if the user didn't enter any input
    If r.materialNum = "" Then
        End
    End If
    
    r.plantNum = "1105"
    
    Call validateInput(r.materialNum)
    
    'ATTACH TO SAP
    Dim session As Variant
    Set session = SAPFunctions.connect2SAPNew ' Get the first session (window) on that connection
    
    Call navigateToSequences(r, session)
    
    r.title = session.findById("wnd[0]/usr/txtRC270-HEAD1").text
    r.countSequences = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text
    
    Dim i, j As Integer
    Dim curSeq As CSequence
    
    Set r.sequences = New Collection
    'Main loop to collect information
    For i = 0 To r.countSequences - 1
        Set curSeq = New CSequence
        Call curSeq.collectInfo(i, session)
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1300/txtPLFLD-PLNFL[0," & i & "]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        Call curSeq.collectOperations(session)
        session.findById("wnd[0]/tbar[1]/btn[29]").press 'Press on sequence view
        r.sequences.Add curSeq
    Next i
    
    session.findById("wnd[0]").Close
    Set session = Nothing
    
    'Output information
    Dim tempOp As COutOperation
    Dim seq As CSequence
    Dim wb As Workbook
    Dim newSheet As Worksheet
    Set wb = ActiveWorkbook
    Set newSheet = HelpFunctions.returnNewSheet(wb)
    
    'Enter offset of columns for outputting
    'Offset is relative to r.col
    r.offsetSeq = 0
    r.offsetOpNum = 1
    r.offsetWorkCtr = 2
    r.offsetDesc = 3
    r.offsetHours = 4
    r.offsetBranch = 5
    r.offsetReturn = 6
    r.countCol = 6
    
    'Check if check name exists
    Dim verCount As Integer: verCount = 0
    Dim finalSheetName As String
    Do While True
        finalSheetName = r.materialNum & "(" & verCount & ")"
        verCount = verCount + 1
        If Not HelpFunctions.SheetExists(finalSheetName, wb) Then
            Exit Do
        End If
    Loop
    newSheet.Name = finalSheetName
    
    r.startRow = 1
    r.row = r.startRow
    r.col = 1
    
    'Setup Title header
    newSheet.Cells(r.row, r.col + r.offsetDesc).Value = r.title
    newSheet.Range(Cells(r.row, r.col + r.offsetSeq), Cells(r.row, r.col + r.countCol)).Interior.ColorIndex = 40
    newSheet.Cells(r.row, r.col + r.offsetHours).Formula = "=0" 'Values will be added as sequences are added
    
    r.row = r.row + 2
    
    Set tempOp = New COutOperation
    Call tempOp.addHeaders(r, newSheet)
    r.row = r.row + 2
    Set tempOp = Nothing
    
    For Each seq In r.sequences
        Call seq.outputAll(r, newSheet)
    Next seq
    
    Set seq = Nothing
    'Format Cells
    newSheet.Cells.EntireColumn.AutoFit
    newSheet.Cells.EntireRow.AutoFit
    
    Set newSheet = Nothing
    Set wb = Nothing
End Sub

Private Sub navigateToSequences(r As Routing, session As Variant)
    Dim headerString As String
    session.sendCommand ("CA03")
    session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = r.materialNum
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").text = r.plantNum
    session.findById("wnd[0]").sendVKey 0
    
    'Handle errors here
    If session.findById("wnd[0]/sbar").messagetype = "E" Then
        session.findById("wnd[0]").Close
        MsgBox (session.findById("wnd[0]/sbar").text)
        End
    End If
    
    headerString = session.findById("wnd[0]/usr/txtRC270-HEAD1").text
    
    If InStr(1, headerString, "Grp.Count") = 0 Then 'There are alternative BOMs
        MsgBox ("There are multiple routings in this material." & Chr(10) & _
        "Double click on the one you want in SAP then press OK")
    End If
    
    session.findById("wnd[0]/tbar[1]/btn[29]").press 'Press on sequence view
    
    'Assert that we are on sequences page
    headerString = session.findById("wnd[0]/usr/txtRC270-HEAD1").text
    If InStr(headerString, "Grp.Count") = 0 Then 'Didn't reach sequences
        MsgBox ("Couldn't get to sequences page, macro will end")
        session.findById("wnd[0]").Close
        End
    End If
End Sub

Private Sub validateInput(materialNum As String)
    If IsNumeric(materialNum) And Len(materialNum) = 9 Then
    Else
        MsgBox ("Not a material number")
        End
    End If
End Sub
