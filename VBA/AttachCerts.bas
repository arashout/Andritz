Attribute VB_Name = "AttachCerts"
Option Explicit
Sub AttachCerts()
    
    Dim WshShell
    Dim fso, directory, files, NewFile, fileObj
    Dim exists As Boolean
    Dim fullFilePath, directoryPath, curPONum As String
    Set WshShell = CreateObject("WScript.Shell")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    directoryPath = InputBox("Enter the directory which you want to scan")
    
    exists = fso.FolderExists(directoryPath)
    Dim session
    Set session = SAPFunctions.connect2SAP()
    If (exists) Then
        Set NewFile = fso.CreateTextFile(directoryPath & "filesAttachedLog.txt", True)
        Set directory = fso.GetFolder(directoryPath)
        Set files = directory.files
    
        For Each fileObj In files
            If LCase(fso.GetExtensionName(fileObj.Name)) = "pdf" Then
                
                fullFilePath = fso.BuildPath(directoryPath, fileObj.Name)
                curPONum = ""
                curPONum = returnPONumber(fullFilePath)
    
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme22n"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/tbar[1]/btn[17]").press
                session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = curPONum
                session.findById("wnd[1]").sendVKey 0
                
                If session.Children().count = 2 Then
                    NewFile.WriteLine (fileObj.Name & " | " & "Not Authorized")
                Else
                    If errorText = "" Then
                        session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
                        session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_PCATTA_CREA"
                        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = directoryPath
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = fileObj.Name
                        session.findById("wnd[1]/tbar[0]/btn[0]").press
                        NewFile.WriteLine (fileObj.Name & " | " & session.findById("wnd[0]/sbar").Text) 'Success
                    Else
                        NewFile.WriteLine (fileObj.Name & " | " & "No PO") 'Error
                    End If
                    
                End If
            End If
        Next
        NewFile.Close
        Else
            Debug.Print ("Incorrect directory path parameter was passed")
            Exit Sub
    End If

End Sub


Function returnPONumber(fileName As String)
    Dim re, mc, mo As Object
    Set re = CreateObject("vbscript.regexp")
    With re
        .IgnoreCase = True
        .Global = False
        .Pattern = "(4\d{9})"
    End With
    Set mc = re.Execute(fileName)
    If mc.count = 1 Then
        returnPONumber = mc.Item(0)
    Else
        returnPONumber = 0
    End If
End Function

Function getFilePaths(directoryPath As String, ext As String) As Variant
    Dim WshShell
    Dim fso, directoryFso, files, fileObj
    Dim exists, fullFilePath
    Dim Result As Variant
    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    exists = fso.FolderExists(directoryPath)
    If (exists) Then
        Set directoryFso = fso.GetFolder(directoryPath)
        Set files = directoryFso.files
    
        For Each fileObj In files
            If LCase(fso.GetExtensionName(fileObj.Name)) = ext Then
                fullFilePath = fso.BuildPath(directoryPath, fileObj.Name)
                ReDim Preserve Result(UBound(Result) + 1)
                Result(UBound(Result)) = fullFilePath
            End If
        Next
        Else
            Debug.Print ("Incorrect directory path parameter was passed")
            Exit Function
    End If
    getFilePaths = Result
End Function
