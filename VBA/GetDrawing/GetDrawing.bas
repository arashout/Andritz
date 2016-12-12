Attribute VB_Name = "GetDrawing"
Public Sub getDrawingMain(outputCol As Long, mode As Boolean)
    Dim session As Variant
    Set session = SAPFunctions.connect2SAP()
    
    Dim i As Long
    Dim dwgNum As String
    Dim curMat As CMaterialDwg
    
    Dim ws As Worksheet
    Dim sel As SelectionT
    sel = SelectionModule.GetSelection()
    
    Set ws = Application.ActiveSheet
    For i = sel.startRow To sel.endRow
        Set curMat = New CMaterialDwg
        dwgNum = Cells(i, sel.startCol).Value
        Call curMat.getDrawingIfExists(dwgNum, session, mode)
        Call curMat.outputDwg(ws, i, outputCol)
    Next i
End Sub
