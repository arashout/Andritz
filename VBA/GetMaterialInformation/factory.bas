Attribute VB_Name = "factory"
''''''''''''''
'This module is a factory to easily create objects with constructors
''''''''''''''
Function createCMaterial(ByVal sapNum As String, ByVal currentSession As Variant, ByVal rowI As Long, ByVal colI As Long, ByVal plantNum As String) As CMaterial
    Dim mat_obj As CMaterial
    Set mat_obj = New CMaterial

    Call mat_obj.initCMaterial(sapNum:=sapNum, currentSession:=currentSession, rowI:=rowI, colI:=colI, plantNum:=plantNum)
    Set createCMaterial = mat_obj

End Function
