Attribute VB_Name = "factory"
Option Explicit
''''''''''''''
'This module is a factory to easily create objects with constructors
''''''''''''''
Function createCOperation(opNum As String, desc As String, workCenter As String, hours As String, session As Variant) As COperation
    Dim op_obj As COperation
    Set op_obj = New COperation
    
    Call op_obj.initCOperation(fopNum:=opNum, fdesc:=desc, fworkCenter:=workCenter, fhours:=hours, fsession:=session)
    Set createCOperation = op_obj
End Function

Function createCOperationPO(opNum As String, desc As String, workCenter As String, hours As String, session As Variant) As COperationPO
    Dim op_obj As COperationPO
    Set op_obj = New COperationPO
    
    Call op_obj.initCOperation(fopNum:=opNum, fdesc:=desc, fworkCenter:=workCenter, fhours:=hours, fsession:=session)
    Set createCOperationPO = op_obj
End Function

