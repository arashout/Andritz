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
Public matNum As String
Public plant As String
Public desc As String

'Validation Properties
Public isValidSAPNum As Boolean
Public hasError As Boolean

'Moving Price Information
Public movingPrice As String
Public movingQuantity As String

'Most Recent Purchase Information
Public recentPrice As String
Public recentCurrency As String
Public recentQuantity As String
Public recentUnit As String

Public foundRecentPrice As Boolean 'Whether price information was found or not

'MRP info
Public stock As Long
Public safetyStock As Long
Public projectStock As Long
Public orderReservation As Long
Public productOrder As Long
Public purchaseReq As Long
Public POItem As Long
Public dependantReq As Long
Public plannedOrder As Long

'SAP session
Public session As Variant

'Spreadsheet position of the SAP number
Public rowIndex As Long
Public colIndex As Long

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
    End If
End Sub

''''''''''''''''''''''
' Public Methods
''''''''''''''''''''''

Public Sub navigateZmatinfo()
    hasError = False
    
    session.SendCommand ("/nzmatinfo")
    session.findById("wnd[0]/usr/chkX_MOVE").Selected = False
    session.findById("wnd[0]/usr/chkX_SALES").Selected = False
    session.findById("wnd[0]/usr/chkX_BOM").Selected = False
    session.findById("wnd[0]/usr/chkX_PROJ").Selected = False
    session.findById("wnd[0]/usr/chkX_PURCH").Selected = True
    session.findById("wnd[0]/usr/chkP_DEF").Selected = True
    
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = plant
    session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").text = matNum
    session.findById("wnd[0]").sendVKey 8
    
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
'Output Modes
''''''''''''''''''''''
Public Sub outputAllStockInfo()
    Cells(rowIndex, colIndex + 1).Value = movingPrice
    Cells(rowIndex, colIndex + 2).Value = stock
    Cells(rowIndex, colIndex + 3).Value = projectStock
    Cells(rowIndex, colIndex + 4).Value = safetyStock
    Cells(rowIndex, colIndex + 5).Value = orderReservation
    Cells(rowIndex, colIndex + 6).Value = dependantReq
    Cells(rowIndex, colIndex + 7).Value = productOrder
    Cells(rowIndex, colIndex + 8).Value = POItem
    Cells(rowIndex, colIndex + 9).Value = plannedOrder
End Sub
Public Sub outputMovingPriceAndStock()
    Cells(rowIndex, colIndex + 1).Value = movingPrice
    Cells(rowIndex, colIndex + 2).Value = stock
End Sub
Public Sub outputRecentPrice()
    Cells(rowIndex, colIndex + 1).Value = recentPrice
    Cells(rowIndex, colIndex + 2).Value = recentQuantity & " " & recentUnit
    Cells(rowIndex, colIndex + 3).Value = recentCurrency
    Cells(rowIndex, colIndex + 4).Value = plant
End Sub
Public Sub outputDescription()
    Cells(rowIndex, colIndex + 1).Value = desc
End Sub
''''''''''''''''''''''
' Private Methods
''''''''''''''''''''''
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
    
    Exit Sub
    
noID:
    currentTitle = ""
    Resume Next
    
End Sub
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
        recentPrice = session.findById("wnd[0]/usr/lbl[19," & i & "]").text
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
