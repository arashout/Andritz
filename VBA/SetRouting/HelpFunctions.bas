Attribute VB_Name = "HelpFunctions"
Function repeatString(text As String, repeat As Integer) As String
    'This function returns string which contains the text parameter repeated a given number of times
    Dim i As Integer
    i = 1
    While i < repeat
        text = text + text
        i = 1 + i
    Wend
    repeatString = text
End Function
Function removeFirstElement(ByVal arr As Variant) As Variant
    For i = LBound(arr) + 1 To UBound(arr)
      arr(i - 1) = arr(i)
    Next i
    ReDim Preserve arr(UBound(arr) - 1)
    removeFirstElement = arr
End Function
Function inArr(ByRef arr As Variant, ByVal searchString As Variant) As Boolean
    Dim flag As Boolean: flag = False
    Dim element As Variant
    
    For Each element In arr
        If element = searchString Then
            flag = True
        End If
    Next element
    inArr = flag
End Function
Public Function popChar(index As Long, theString As String) As String
    'This function pops out the character at the given index
    'IN integer index representing position of character you want, string for the string you want the characte from
    'OUT string with single character
    popChar = Mid(theString, index, 1)
End Function
