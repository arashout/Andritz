VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''
' Attributes
''''''''''''''''''''''

'Info about material
Private matNum As String
Private plant As String
Private desc As String

'Moving Price Information
Private movingPrice As String
Private movingQuantity As String

'Most Recent Purchase Information
Private recentPrice As String
Private recentCurrency As String
Private recentQuantity As String
Private recentUnit As String

'MRP info
Private stock As Long
Private safetyStock As Long
Private projectStock As Long
Private orderReservation As Long
Private productOrder As Long
Private purchaseReq As Long
Private POItem As Long
Private dependantReq As Long
Private plannedOrder As Long

'Validation Properties
Public isValidSAPNum As Boolean
Public hasError As Boolean
Public foundRecentPrice As Boolean 'Whether price information was found or not

'SAP session
Public session As Variant

'Spreadsheet position of the SAP number
Public rowIndex As Long
Public colIndex As Long

'Collection that holds information for initiliazing option list and outputting
'Note: Could have used a dictionary here but order was important
Public collectionInfo As Collection
''''''''''''''''''''''
' Pretend Constructor
''''''''''''''''''''''
Sub initCMaterial(sapNum As String, currentSession As Variant, rowI As Long, colI As Long, plantNum As String)
    If Len(sapNum) = 9 And IsNumeric(sapNum) Then 'Simple check that entry is in fact a valid material number
        isValidSAPNum = True
        matNum = sapNum
        Set session = currentSession
        rowIndex = rowI
        colIndex = colI
        plant = plantNum
    Else
        isValidSAPNum = False
        Exit Sub
    End If
    Call createCollectionEntries
End Sub

''''''''''''''''''''''
' Public Methods
''''''''''''''''''''''

Public Sub navigateZmatinfo()
    hasError = False
    
    session.SendCommand ("/nzmatinfo")
    'Turn off all unnecessary information so SAP loads faster
    session.findById("wnd[0]/usr/chkX_MOVE").Selected = False
    session.findById("wnd[0]/usr/chkX_SALES").Selected = False
    session.findById("wnd[0]/usr/chkX_BOM").Selected = False
    session.findById("wnd[0]/usr/chkX_PROJ").Selected = False
    session.findById("wnd[0]/usr/chkX_PURCH").Selected = True
    session.findById("wnd[0]/usr/chkP_DEF").Selected = True
    
    'Enter the plant and material number
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = plant
    session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").text = matNum
    session.findById("wnd[0]").sendVKey 8 'Execute
    
    'If it hasn't made it to the material info page there is a problem
    On Error GoTo wrongSAPNum
    If session.findById("wnd[0]/usr/lbl[0,0]").text = "Material Info" Then 'Check if we are on material info page
        ReadMaterialInfo
    End If
        
    
Exit Sub
wrongSAPNum:
        hasError = True
        Resume Next
    
End Sub
''''''''''''''''''''''
' Private Methods
''''''''''''''''''''''
Private Sub createCollectionEntries()
    Set collectionInfo = New Collection
    'Set-up collection so cmd window can be populated
    
    With collectionInfo
        'Unfortunately you have to add new items manually here
        'This collection item will be used to update command window listbox
        'Also note that I had to add an array because you can't iterate over keys in collection
        .Add Array(plant, "Plant Code"), "Plant Code"
        .Add Array(desc, "Description"), "Description"
        .Add Array(recentPrice, "Recent Price"), "Recent Price"
        .Add Array(recentCurrency, "Recent Price Currency"), "Recent Price Currency"
        .Add Array(recentQuantity, "Recent Price Quantity"), "Recent Price Quantity"
        .Add Array(recentUnit, "Recent Price Units"), "Recent Price Units"
        .Add Array(movingPrice, "Moving Price"), "Moving Price"
        .Add Array(movingQuantity, "Moving Price Quantity"), "Moving Price Quantity"
        .Add Array(stock, "Stock"), "Stock"
        .Add Array(safetyStock, "Safety Stock"), "Safety Stock"
        .Add Array(projectStock, "Project Stock"), "Project Stock"
        .Add Array(orderReservation, "Order Reservation"), "Order Reservation"
        .Add Array(productOrder, "Product Order"), "Product Order"
        .Add Array(purchaseReq, "Purchase Requisition"), "Purchase Requisition"
        .Add Array(POItem, "Purchase Order Item"), "Purchase Order Item"
        .Add Array(dependantReq, "Dependant Requisition"), "Dependant Requisition"
        .Add Array(plannedOrder, "Planned Order"), "Planned Order"
    End With

End Sub
Private Sub updateCollectionEntries()
    Call HelpFunctions.updateCollectionVal(Array(plant, "Plant Code"), "Plant Code", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(desc, "Description"), "Description", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(recentPrice, "Recent Price"), "Recent Price", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(recentCurrency, "Recent Price Currency"), "Recent Price Currency", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(recentQuantity, "Recent Price Quantity"), "Recent Price Quantity", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(recentUnit, "Recent Price Units"), "Recent Price Units", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(movingPrice, "Moving Price"), "Moving Price", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(movingQuantity, "Moving Price Quantity"), "Moving Price Quantity", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(stock, "Stock"), "Stock", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(safetyStock, "Safety Stock"), "Safety Stock", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(projectStock, "Project Stock"), "Project Stock", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(orderReservation, "Order Reservation"), "Order Reservation", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(productOrder, "Product Order"), "Product Order", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(purchaseReq, "Purchase Requisition"), "Purchase Requisition", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(POItem, "Purchase Order Item"), "Purchase Order Item", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(dependantReq, "Dependant Requisition"), "Dependant Requisition", collectionInfo)
    Call HelpFunctions.updateCollectionVal(Array(plannedOrder, "Planned Order"), "Planned Order", collectionInfo)
End Sub
'This sub-routine assumes that the session is currently on "Material Info" page
Private Sub ReadMaterialInfo()
    Dim currentTitle As String
    Dim mrpStart, mrpEnd, i As Integer
    For i = 0 To 90
        On Error GoTo noID
        currentTitle = session.findById("wnd[0]/usr/lbl[0," & i & "]").text
        
        Select Case currentTitle
            Case "M R P - I n f o"
                mrpStart = i
            Case "P u r c h a s i n g  -  I n f o"
                mrpEnd = i
        End Select
    Next i
    
    If plant = "1105" Then
        Call CollectMRP(mrpStart, mrpEnd)
        Call CollectFixed
        Call collectDescription
        Call collectPriceInfo
    Else
        Call collectPriceInfo
    End If
    
    Call updateCollectionEntries
    
    Exit Sub
    
noID:
    currentTitle = ""
    Resume Next
    
End Sub
'''''''''''''''''''''''''''''''
'Information Collection Subs - These subs collect information from the material info page and update the collection object
'''''''''''''''''''''''''''''''
Private Sub CollectFixed()
        movingPrice = session.findById("wnd[0]/usr/lbl[19,16]").text 'Moving price location
        movingQuantity = session.findById("wnd[0]/usr/lbl[45,15]").text 'Moving price quantity
End Sub
Private Sub CollectMRP(ByVal startRow As Integer, ByVal endRow As Integer)
    Dim element As String
    Dim j As Integer
    
    For j = startRow To endRow
        On Error GoTo noID
        element = session.findById("wnd[0]/usr/lbl[13," & j & "]").text 'Read the titles
        
        Select Case element
            Case "Stock"
                stock = CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update the stock property
            Case "SafeSt"
                safetyStock = safetyStock + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update the safety stock property
            Case "ProjSt"
                projectStock = projectStock + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update the project stock property
            Case "OrdRes"
                orderReservation = orderReservation + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update order reservation property
            Case "PrdOrd"
                productOrder = productOrder + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update product order property
            Case "POitem"
                POItem = POItem + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update PO Item property
            Case "DepReq"
                dependantReq = dependantReq + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update Dependant Requisition property
            Case "PldOrd"
                plannedOrder = plannedOrder + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update Planned Order property
            Case "PurRqs"
                purchaseReq = purchaseReq + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update Purchase Req property
            'Debug cases
            Case ""
            Case "Element"
            Case Else
                MsgBox ("Unaccounted for element: " & element & " -Notify developer if possible")
                Exit For
        End Select
    Next j
    
    Exit Sub
    
noID:
    element = ""
    Resume Next
    
End Sub
Private Sub collectPriceInfo()
    Dim capture As String: capture = ""
    Dim maxI, i As Integer

    maxI = 90 'Gross price header isn't found after 90 iterations it's not there
    i = 10
    
    Do While True
        On Error GoTo noIDError
        capture = session.findById("wnd[0]/usr/lbl[20," & i & "]").text 'Gross Price
        
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
        recentPrice = Trim(session.findById("wnd[0]/usr/lbl[19," & i & "]").text)
        recentCurrency = session.findById("wnd[0]/usr/lbl[44," & i & "]").text
        recentQuantity = session.findById("wnd[0]/usr/lbl[50," & i & "]").text
        recentUnit = session.findById("wnd[0]/usr/lbl[53," & i & "]").text
        'Final check to ensure prices are found
        If recentPrice <> "" Then
            foundRecentPrice = True
        Else
            foundRecentPrice = False
        End If
    Else
        foundRecentPrice = False
    End If
    
    Exit Sub
noIDError:
        Resume Next
End Sub
Private Sub collectDescription()
    Dim i As Integer
    For i = 34 To 200 'These are min and max positions where descriptions will be found
        'I have to filter out the junk in between
        On Error Resume Next
        desc = desc + "|" + session.findById("wnd[0]/usr/lbl[" & i & ",4]").text
        
    Next i
    If desc <> "" Then
        desc = Mid(desc, 2, Len(desc) - 2) 'Get rid of | at start
    End If
    
End Sub
