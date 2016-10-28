Attribute VB_Name = "PullPrices4Mat"

Public Type Material
    'Info about material
     matNum As String
     price As String
     currency As String
     quantity As String
     unit As String
     plant As String
     errorText As String
     found As Boolean
     
     'Info SAP session
     session As Variant
     
     'Info Excel Postion
     rowIndex As Long
     
End Type
'Use this function to clear curMat
Public Function GetBlankMaterial() As Material
End Function
Sub findPrice()
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
                Call navigateToMD04(curMat)
                'If no errors faced at MD04 screen
                If curMat.errorText = "" Then
                    Call collectPriceInfo(curMat)
                    If curMat.plant = "1105" And curMat.price = "" Then 'No purchase history in 1105, try 0303
                        curMat.plant = "0303"
                    ElseIf curMat.plant = "0303" And curMat.price = "" Then 'No purchase history in 1105, give up
                        curMat.plant = "N/A"
                        Exit Do
                    Else 'Price found
                        curMat.found = True
                        Exit Do
                    End If
                Else
                    If curMat.plant = "1105" Then 'If not defined in 1105 try 0303
                        curMat.plant = "0303"
                    Else 'Give up if error in 0303
                        Exit Do
                    End If
                End If
            Loop
        End If
        
        
        If curMat.found Then 'If price found then output to excel columns
            Set cellRange = Range(Cells(j, k + 1).Address)
            cellRange.Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(j, k + 1) = curMat.price
            Cells(j, k + 2) = curMat.currency
            Cells(j, k + 3) = curMat.quantity
            Cells(j, k + 3) = Cells(j, k + 3) & curMat.unit
            Cells(j, k + 4) = curMat.plant
        End If
    Next j
    
    session.findById("wnd[0]").Close
End Sub

Private Sub navigateToMD04(ByRef mat As Material)
    Dim session As Variant
    Set session = mat.session
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = mat.matNum 'Enter a Material Number
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").Text = mat.plant 'Enter a Plant
    session.findById("wnd[0]").sendVKey 0
    'If any errors pop up
    If session.findById("wnd[0]/sbar").messagetype = "E" Then
        mat.errorText = session.findById("wnd[0]/sbar").Text 'Get the text from the error message
        Exit Sub
    End If
    session.findById("wnd[0]").sendVKey 41 'CTRL-SHIFT-F5

End Sub

Private Sub collectPriceInfo(mat As Material)
    Dim captured As String: capture = ""
    Dim maxI, i As Integer
    Dim session As Variant
    Set session = mat.session
    maxI = 90 'Gross price header isn't found after 90 iterations it's not there
    i = 0
    
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
    End If
    
    Exit Sub
noIDError:
        Resume Next
End Sub

