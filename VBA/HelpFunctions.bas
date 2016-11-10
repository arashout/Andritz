Attribute VB_Name = "HelpFunctions"
Function removeFirstElement(ByVal arr As Variant) As Variant
    For i = LBound(arr) + 1 To UBound(arr)
      arr(i - 1) = arr(i)
    Next i
    ReDim Preserve arr(UBound(arr) - 1)
    removeFirstElement = arr
End Function

Function lastRow() As Long
    lastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row
End Function

Function lastCol() As Long
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
End Function
Public Function LastAuthor() As String
   Application.Volatile
   LastAuthor = ThisWorkbook.BuiltinDocumentProperties("Last Author")
End Function
Public Function getUser() As String
    Environ ("Username")
End Function
Public Sub hideColumnsBasedOnRow(ByRef strArr As Variant, Optional rowIndex As Long = 1, Optional delete As Boolean = False)
    Dim colI As Long: colI = 1
    Dim element As Variant
    Dim insideFlag As Boolean: insideFlag = False
    Dim lastColumn As Long: lastColumn = lastCol()
    While colI <= lastColumn
        If Not inArr(strArr, Cells(rowIndex, colI).Value) Then
            If delete Then
                Columns(colI).EntireColumn.delete
            Else
                v.EntireColumn.Hidden = True
            End If
        End If
        colI = colI + 1
    Wend
    
End Sub
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
Public Function GetMaxCell(Optional ByRef rng As Range = Nothing) As Range

    'Returns the last cell containing a value, or A1 if Worksheet is empty

    Const NONEMPTY As String = "*"
    Dim lRow As Range, lCol As Range

    If rng Is Nothing Then Set rng = Application.ActiveWorkbook.ActiveSheet.UsedRange
    If WorksheetFunction.CountA(rng) = 0 Then
        Set GetMaxCell = rng.Parent.Cells(1, 1)
    Else
        With rng
            Set lRow = .Cells.Find(What:=NONEMPTY, LookIn:=xlFormulas, _
                                        After:=.Cells(1, 1), _
                                        SearchDirection:=xlPrevious, _
                                        SearchOrder:=xlByRows)
            If Not lRow Is Nothing Then
                Set lCol = .Cells.Find(What:=NONEMPTY, LookIn:=xlFormulas, _
                                            After:=.Cells(1, 1), _
                                            SearchDirection:=xlPrevious, _
                                            SearchOrder:=xlByColumns)

                Set GetMaxCell = .Parent.Cells(lRow.Row, lCol.Column)
            End If
        End With
    End If
End Function
Public Function popChar(index As Long, theString As String) As String
    'This function pops out the character at the given index
    'IN integer index representing position of character you want, string for the string you want the characte from
    'OUT string with single character
    popChar = Mid(theString, index, 1)
End Function
Public Sub FastWB(Optional ByVal opt As Boolean = True)
    'From stackOverflow
    'Set this to true when you want to turn off all unnecessary stuff for macros to run faster
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

Function IsEmptyRange(rangeObj As Range) As Boolean
'Returns True if the given range only contains "" in cells
'IN range object
'OUT boolean representing whether range is empty or not
    Dim cell As Range
    
    IsEmptyRange = True
    For Each cell In rangeObj
        If cell.Value <> "" Then
            IsEmptyRange = False
            Exit For
        End If
    Next cell
    
End Function

Function ColumnLetter(ColumnNumber As Long) As String
'From StackOverflow
'IN number corresponding to column
'OUT column letter corresponding to given number
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function

 Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ActiveWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function
