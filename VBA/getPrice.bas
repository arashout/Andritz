Attribute VB_Name = "getPrice"

Public Type Material
    'Info about material
     matNum As String
     price As String
     currency As String
     quantity As String
     unit As String
     plant As String
     hasError As Boolean
     found As Boolean
     
     'The moving prices and safety stock
     movingPrice As String
     movingQuantity As String
     stock As String
     safetyStock As String
     
     'Info SAP session
     session As Variant
     
     'Info Excel Postion
     rowIndex As Long
     
End Type
'Use this function to clear curMat
Public Function GetBlankMaterial() As Material
End Function
Sub findPrice(mode As Integer)
    'Retrieves the moving price + stock or recent prices depending on the given mode
    '1 = Recent Price, 2 = Moving Price + Stock
    
    Dim cellRange As Range
    'Get the range the user has selected to determine the first and last cell to iterate between
    Dim myRange As Range
    Set myRange = Selection
    
    
    Dim startJ As Long: startJ = myRange.Row 'Get the first row of the selection
    Dim lastJ As Long: lastJ = myRange.Rows.count + startJ - 1 'Get the last row of the selection
    
    Dim k As Long 'Column index where to look for SAP numbers
    k = myRange.Column 'First column of selection
    
    'Connect to SAP
    Set session = SAPFunctions.connect2SAPNew()
    
    
    Dim curMat As Material
    For j = startJ To lastJ
        curMat = GetBlankMaterial() 'Clear curMat for next material
        'Assign values to Material
        Set curMat.session = session
        curMat.matNum = Cells(j, k).Value
        curMat.rowIndex = j
        curMat.plant = "1105" 'Initial plant - Want to check DSC first
        If Len(curMat.matNum) = 9 And IsNumeric(curMat.matNum) Then 'Simple check that entry is in fact a valid material number
            Do While True
                Call navigateZmatinfo(curMat)
                'If successful made it to material information page
                If Not curMat.hasError Then
                    Call collectPriceInfo(curMat)
                    
                    If curMat.plant = "1105" And Not curMat.found And mode = 1 Then 'No purchase history in 1105, try 0303
                        curMat.plant = "0303"
                    ElseIf (curMat.plant = "0303" And Not curMat.found) Or mode = 2 Then 'No purchase history in 0303, give up or if we only want stock give up
                        curMat.plant = "N/A"
                        Exit Do
                    Else 'Found
                        Exit Do
                    End If
                'If there is an error it mean it was not defined in the plant
                Else
                    If curMat.plant = "1105" And mode = 1 Then 'If not defined in 1105 try 0303 -> Only if you want recent prices
                        curMat.plant = "0303"
                    Else
                        Exit Do
                    End If
                End If
            Loop
        End If
        
        
        If curMat.found Then 'If price found then output to excel columns
             Call outputPrices(curMat, k, mode)
        End If
    Next j
    
    session.findById("wnd[0]").Close
End Sub
Private Sub navigateZmatinfo(ByRef mat As Material)
    Dim session As Variant
    Set session = mat.session
    mat.hasError = False
    
    session.SendCommand ("/nzmatinfo")
    session.findById("wnd[0]/usr/chkX_MOVE").Selected = False
    session.findById("wnd[0]/usr/chkX_SALES").Selected = False
    session.findById("wnd[0]/usr/chkX_BOM").Selected = False
    session.findById("wnd[0]/usr/chkX_PROJ").Selected = False
    session.findById("wnd[0]/usr/chkX_PURCH").Selected = True
    session.findById("wnd[0]/usr/chkP_DEF").Selected = True
    
    session.findById("wnd[0]/usr/ctxtP_WERKS").Text = mat.plant
    session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").Text = mat.matNum
    session.findById("wnd[0]").sendVKey 8
    
    'If it hasn't made it to the material info page there is a problem
    On Error GoTo wrongSAPNum
    If session.findById("wnd[0]/usr/lbl[0,0]").Text = "Material Info" Then 'Just a check to ensure we are on the right page
    End If
        
    
Exit Sub
wrongSAPNum:
        mat.hasError = True
        Resume Next
    
End Sub

Private Sub collectPriceInfo(ByRef mat As Material)
    Dim captured As String: capture = ""
    Dim maxI, i As Integer
    Dim session As Variant
    Set session = mat.session
    maxI = 90 'Gross price header isn't found after 90 iterations it's not there
    i = 10
    
    Do While True
        On Error GoTo noIDError
        capture = session.findById("wnd[0]/usr/lbl[20," & i & "]").Text 'Gross Price
        
        If capture = "Gross Price" Then
            i = i + 1
            Exit Do
        End If
        
        If i >= maxI Then
            Exit Do
        End If
        i = i + 1
    Loop
    'Get price information from sheet if they exist - otherwise error will ensure they are blank
    
    If i <> maxI Then 'If MaxI has been reached means price for sure not found
        mat.price = session.findById("wnd[0]/usr/lbl[19," & i & "]").Text
        mat.currency = session.findById("wnd[0]/usr/lbl[44," & i & "]").Text
        mat.quantity = session.findById("wnd[0]/usr/lbl[50," & i & "]").Text
        mat.unit = session.findById("wnd[0]/usr/lbl[53," & i & "]").Text
        mat.found = True
        
        'These items are always in the same place luckily
        mat.movingPrice = session.findById("wnd[0]/usr/lbl[19,16]").Text 'Moving price location
        mat.movingQuantity = session.findById("wnd[0]/usr/lbl[45,15]").Text 'Moving price quantity
        mat.stock = session.findById("wnd[0]/usr/lbl[94,10]").Text 'Stock quantity
        mat.safetyStock = session.findById("wnd[0]/usr/lbl[94,11]").Text 'Safety Stock quantity
        
    Else
        mat.found = False
    End If
    
    Exit Sub
noIDError:
        Resume Next
End Sub

Private Sub outputPrices(mat As Material, k As Long, mode As Integer)
    Dim cellRange As Range
    Dim j As Long: j = mat.rowIndex
    
    Set cellRange = Range(Cells(j, k + 1).Address)
    cellRange.Select
    
    If Not CmdWindow.cbReplace.Value And mode = 1 Then 'Mode 1 is most recent price
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ElseIf Not CmdWindow.cbReplace.Value And mode = 2 Then
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
    If mode = 1 Then
        Cells(j, k + 1) = mat.price
        Cells(j, k + 2) = mat.currency
        Cells(j, k + 3) = mat.quantity
        Cells(j, k + 3) = Cells(j, k + 3) & mat.unit
        Cells(j, k + 4) = mat.plant
    ElseIf mode = 2 Then
        Cells(j, k + 1) = mat.price
        Cells(j, k + 2) = mat.stock
        Cells(j, k + 3) = mat.safetyStock
    End If
End Sub
