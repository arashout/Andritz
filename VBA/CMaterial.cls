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
Public hasError As Boolean


'Purchase History Information
Public movingPrice As String
Public movingQuantity As String

Public price As String
Public quantity As String
Public unit As String

Public found As Boolean 'Whether price information was found or not

'MRP info
Public stock As Long
Public safetyStock As Long

'SAP session
Public session As Variant

'Spreadsheet position of the SAP number
Public rowIndex As Long
Public colIndex As Long

''''''''''''''''''''''
' Public Methods
''''''''''''''''''''''

'This sub-routine assumes that the session is currently on "Material Info" page
Public Sub ReadMaterialInfo()
    Dim currentTitle As String
    Dim mrpStart, mrpEnd As Integer
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
    
    Call CollectMRP(mrpStart, mrpEnd)
    
    Exit Sub
    
noID:
    currentTitle = ""
    Resume Next
    
End Sub

''''''''''''''''''''''
' Private Methods
''''''''''''''''''''''
Private Sub CollectMRP(ByVal startRow As Integer, ByVal endRow As Integer)

    Dim element As String
    
    For j = startRow To endRow
        On Error GoTo noID
        element = session.findById("wnd[0]/usr/lbl[13," & j & "]").text 'Read the titles
        
        Select Case element
            Case "Stock"
                stock = CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update the stock property
            Case "SafeSt"
                'TODO - Test this
                safetyStock = safetyStock + CInt(session.findById("wnd[0]/usr/lbl[85," & j & "]").text) 'Update the safety stock property
        End Select
    Next j
    
    Exit Sub
    
noID:
    element = ""
    Resume Next
    
End Sub

