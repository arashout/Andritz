Attribute VB_Name = "HelpFunctions"
Option Explicit
Function returnNewSheet(wb As Workbook) As Worksheet
    With wb
        Set returnNewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
    End With
End Function
Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
   'Simple function that returns boolean true if the currently active workbook
   'contains a sheet with a given name
   Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
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
