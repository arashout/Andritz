Attribute VB_Name = "HelpFunctions"
Option Explicit
Sub updateCollectionVal(aValue As Variant, aKey As Variant, aCol As Collection)
    'A work around sub for editing the values added to collections
    'Collections aren't meant to have there values edited
    Dim temp As Variant
    temp = aCol.Item(aKey)
    aCol.Remove (aKey)
    Call aCol.Add(aValue, aKey)
    
End Sub
Function inArr(ByRef arr As Variant, ByVal searchString As Variant) As Boolean
    'Simple function that returns a bolean true if a element is IN an array
    'SHOULD work for all types (NOT TESTED OUTSIDE OF STRINGS)
    Dim flag As Boolean: flag = False
    Dim element As Variant
    
    For Each element In arr
        If element = searchString Then
            flag = True
        End If
    Next element
    inArr = flag
End Function
