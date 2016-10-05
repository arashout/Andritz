Attribute VB_Name = "HelpFunctions"
Function lastRow() As Long
    lastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row
End Function

Function lastCol() As Long
'
    lastCol = Range("A1").SpecialCells(xlCellTypeLastCell).Column
End Function

Public Sub FastWB(Optional ByVal opt As Boolean = True)
    'VBA macro to turn off unnecessary futures during a macro
    'Uses two private sub-routines to achieve this
    'These macros aren't my work but from this thread: http://stackoverflow.com/questions/30959315/excel-vba-performance-1-million-rows-delete-rows-containing-a-value-in-less
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

Private Sub FastWS(Optional ByVal ws As Worksheet = Nothing, _
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

