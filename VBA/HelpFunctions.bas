Attribute VB_Name = "HelpFunctions"
Function lastRow() As Long
    lastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row
End Function

Function lastCol() As Long
    lastCol = Range("A1").SpecialCells(xlCellTypeLastCell).Column
End Function
Private Function popChar(index As Long, theString As String) As String
    'This function pops out the character at the given index
    popChar = Mid(theString, index, 1)
End Function
Public Sub FastWB(Optional ByVal opt As Boolean = True)
    'Sub to make excel run faster
    With Application
        .Calculation = IIf(opt, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not opt
        .DisplayStatusBar = Not opt
        .EnableAnimations = Not opt
        .EnableEvents = Not opt
        .ScreenUpdating = Not opt
    End With
    FastWS , opt
End Sub

Public Sub FastWS(Optional ByVal ws As Worksheet = Nothing, _
                  Optional ByVal opt As Boolean = True)
    If ws Is Nothing Then
        For Each ws In Application.ActiveWorkbook.Sheets
            EnableWS ws, opt
        Next
    Else
        EnableWS ws, opt
    End If
End Sub

Private Sub EnableWS(ByVal ws As Worksheet, ByVal opt As Boolean)
    With ws
        .DisplayPageBreaks = False
        .EnableCalculation = Not opt
        .EnableFormatConditionsCalculation = Not opt
        .EnablePivotTable = Not opt
    End With
End Sub

Sub IsEmptyRange()
Dim cell As Range
Dim bIsEmpty As Boolean

bIsEmpty = True
For Each cell In Range("BB2:BB5")
    If cell.Value <> "" Then
        bIsEmpty = False
        Exit For
    End If
Next cell

If bIsEmpty = True Then
    'There are empty cells in your range
    '**PLACE CODE HERE**
    MsgBox "All cells empty"
Else
    'There are NO empty cells in your range
    '**PLACE CODE HERE**
    MsgBox "Some have values"
End If
End Sub
