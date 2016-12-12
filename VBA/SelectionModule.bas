Attribute VB_Name = "SelectionModule"
'Module to get which areas the user has selected
Option Explicit
Public Type SelectionT
    startRow As Long
    endRow As Long
    startCol As Long
    endCol As Long
End Type

Public Function GetSelection() As SelectionT
    Dim userSel As Range
    Set userSel = Selection
    
    GetSelection.startRow = userSel.row 'Get the first row of the selection
    GetSelection.endRow = userSel.Rows.Count + GetSelection.startRow - 1 'Get the last row of the selection
    GetSelection.startCol = userSel.Column 'First column of selection
    GetSelection.endCol = GetSelection.startCol + userSel.Rows.Count - 1
    
End Function
