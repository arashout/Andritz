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
