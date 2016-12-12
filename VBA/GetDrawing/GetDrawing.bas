Attribute VB_Name = "GetDrawing"
Public Sub getDrawingMain()
    
    Dim i As Long
    Dim curMat As CMaterial
    Dim session As Variant
    Set session = SAPFunctions.connect2SAP()
    For i = 2 To HelpFunctions.lastRow
        Set curMat = New CMaterial
        Call curMat.getDrawingIfExists(201708627, session)
    Next i
End Sub
