VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaterialDwg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sapNum As String
Private dwgName As String
Private dwgPath As String

Public dwgFound As Boolean
Private session As Variant
Private itemType As Variant
Private byteKey As Byte
Private stringKey As String
Private downloadDwg As Boolean
Private errInfo As String

Public Function getDrawingIfExists(num As String, curSession As Variant, downloadMode As Boolean) As Boolean
    Set session = curSession
    'Set up variables
    sapNum = num
    byteKey = 2 'Start at node 2
    downloadDwg = downloadMode
    Call NavigateToDrawing
    getDrawingIfExists = dwgFound
End Function

Private Sub NavigateToDrawing()
    session.SendCommand ("/nmm03")
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = sapNum
    session.findById("wnd[0]").sendVKey 0
    If session.findById("wnd[0]/sbar").messagetype = "E" Then 'Check SAP Num
        dwgFound = False
        errInfo = "SAP # Wrong"
        Exit Sub
    End If
    session.findById("wnd[1]").sendVKey 0 'Select Basic 1 and 3 by default
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP03").Select 'Go to Basic 3
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMM:2005/subSUB2:SAPLZMM00_DMS:9001/btnZANDR_PUSH_SDP").press 'Click on STD Drawing PACKAGE
    
    Do While True
        stringKey = createKey(byteKey)
        On Error GoTo noData
        itemType = session.findById("wnd[1]/usr/cntl100_CTRL_TREE/shellcont/shell").getItemText(stringKey, "DOKAR")
        If itemType = "DRW" Then
            dwgName = session.findById("wnd[1]/usr/cntl100_CTRL_TREE/shellcont/shell").getItemText(stringKey, "DOKNR")
            Exit Do
        Else
            byteKey = byteKey + 1 'Do I need to error checking?
        End If
    Loop
    
    If downloadDwg Then
        session.findById("wnd[1]/usr/cntl100_CTRL_TREE/shellcont/shell").selectItem stringKey, "DOKNR"
        session.findById("wnd[1]/usr/cntl100_CTRL_TREE/shellcont/shell").ensureVisibleHorizontalItem stringKey, "DOKNR"
        
        session.findById("wnd[1]/usr/cntl100_CTRL_TREE/shellcont/shell").itemContextMenu stringKey, "DOKNR"
        session.findById("wnd[1]/usr/cntl100_CTRL_TREE/shellcont/shell").selectContextMenuItem ("DIR_DISP")
        
        session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").nodeContextMenu "          1"
        session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "CF_EXP_COPY"
        session.findById("wnd[1]/usr/ctxtDRAW-DTTRG").Text = "zpkk"
        dwgPath = session.findById("wnd[1]/usr/ctxtDRAW-FILEP").Text
    End If
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    dwgFound = True
    Exit Sub

noData:
    dwgFound = False
    errInfo = "No Data"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End Sub
Public Sub outputDwg(ws As Worksheet, row As Long, col As Long)
    If dwgFound Then
        ws.Cells(row, col).Value = dwgName
    Else
        ws.Cells(row, col).Value = errInfo
    End If
End Sub
Private Function createKey(key As Byte) As String
    createKey = "          " & key
End Function
