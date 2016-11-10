Attribute VB_Name = "MaterialInfoRunScripts"
Public Sub showCMD()
    Call cmdWindow.Show(vbvModeless) 'So you can select outside
End Sub
'This run script looks for materials defined in 1105
Sub DSConlyRun()
    'Get the range the user has selected to determine the first and last cell to iterate between
    Dim myRange As Range
    Set myRange = Selection
    
    Dim startJ As Long: startJ = myRange.Row 'Get the first row of the selection
    Dim lastJ As Long: lastJ = myRange.Rows.Count + startJ - 1 'Get the last row of the selection
    
    Dim j, k As Long 'Row and Column index where to look for SAP numbers
    k = myRange.Column 'First column of selection
    
    'Get settings from command window
    Dim statusListbox As String
    statusListbox = cmdWindow.listboxOptions.Value
    
    'Connect to SAP
    Dim session As Variant
    Set session = SAPFunctions.connect2SAPNew()
    
    Dim curMat As CMaterial
    For j = startJ To lastJ
        If Len(Cells(j, k)) = 9 And IsNumeric(Cells(j, k)) Then
            'Initialize new CMaterial
            Set curMat = factory.createCMaterial(sapNum:=Cells(j, k), currentSession:=session, rowI:=j, colI:=k, plantNum:="1105")
            
            If curMat.isValidSAPNum Then 'Property set in constructor
                Call curMat.navigateZmatinfo
            End If
            
            If Not curMat.hasError Then 'If we manage to get to the material info page
                'Chose correct output depending on cmd settings
                Select Case statusListbox
                    Case "Get Long Text"
                        Call curMat.outputDescription
                    Case "Get Moving Price/Stock/Safety Stock"
                        Call curMat.outputMovingPriceAndStock
                    Case "Get ALL Stock Info"
                        Call curMat.outputAllStockInfo
                    Case Else
                        MsgBox ("You need to pick an option from the listbox")
                        Exit Sub
                End Select
            End If
        End If
    Next j
    
    session.findById("wnd[0]").Close
End Sub

'This runner checks in plant 1105 and if not found 0303
Sub multiplePlantRun()
    'Get the range the user has selected to determine the first and last cell to iterate between
    Dim myRange As Range
    Set myRange = Selection
    
    Dim startJ As Long: startJ = myRange.Row 'Get the first row of the selection
    Dim lastJ As Long: lastJ = myRange.Rows.Count + startJ - 1 'Get the last row of the selection
    
    Dim j, k As Long 'Row and Column index where to look for SAP numbers
    k = myRange.Column 'First column of selection
    
    'Get settings from command window
    Dim statusListbox As String
    statusListbox = cmdWindow.listboxOptions.Value
    
    'Connect to SAP
    Dim session As Variant
    Set session = SAPFunctions.connect2SAPNew()
    
    Dim curMat As CMaterial
    For j = startJ To lastJ
        'Initialize new CMaterial
        Set curMat = factory.createCMaterial(sapNum:=Cells(j, k), currentSession:=session, rowI:=j, colI:=k, plantNum:="1105")
        If curMat.isValidSAPNum Then 'Property set in constructor
            Do While True
                Call curMat.navigateZmatinfo
                'Possible scenarios
                
                'Not defined in plant
                If curMat.hasError Then
                    If curMat.plant = "1105" Then 'Not defined in 1105, try 0303
                        curMat.plant = "0303"
                    
                    Else 'Not defined in 0303, giveup
                        Exit Do
                    
                    End If
                'Defined in plant so at material info page
                Else
                    If curMat.foundRecentPrice Then 'Found needed info
                        Call curMat.outputRecentPrice
                        Exit Do
                    
                    ElseIf Not curMat.foundRecentPrice And curMat.plant = "1105" Then 'Info not found in 1105
                        curMat.plant = "0303"
                    
                    Else 'Info not found in 0303
                        Exit Do
                    End If
                    
                End If
    
            Loop
        End If
    Next j
    
    session.findById("wnd[0]").Close
End Sub
