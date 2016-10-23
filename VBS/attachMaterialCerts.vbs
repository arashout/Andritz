Option Explicit

Dim WshShell
Dim fso, directory, files, NewFile, directoryPath, fileObj
Dim exists, fullFilePath, curPONum

Set WshShell = CreateObject("WScript.Shell")

Set fso = CreateObject("Scripting.FileSystemObject")
directoryPath = "C:\Users\arash\Github\WorkRepos\Andritz\VBS"
'directoryPath = Inputbox("Enter the directory which you want to scan")

exists = fso.FolderExists(directoryPath)
If (exists) Then
    Set NewFile = fso.CreateTextFile(directoryPath&"\filesAttachedLog.txt", True)
    Set directory = fso.GetFolder(directoryPath)
    Set files = directory.Files

    For each fileObj In files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "pdf" Then
            NewFile.WriteLine(fileObj.Name)
            fullFilePath = fso.BuildPath(directoryPath, fileObj.Name)
            curPONum = returnPONumber(fileObj.Name)
            If curPONum <> 0 Then
                Wscript.Echo curPONum
            else
                WshShell.Run fullFilePath
                curPONum = Inputbox("Enter the po number")
                Wscript.Echo curPONum
            End If
        End If
    Next
    NewFile.Close
    else
        Wscript.Echo "Incorrect directory path parameter was passed"
        Wscript.Quit
End if

Function returnPONumber(fileName)
    Dim re, mc, mo
    Set re = New RegExp
    With re
        .IgnoreCase = True
        .Global = False
        .Pattern = "(\d{9})"
    End With
    Set mc = re.Execute(fileName)
    If mc.Count = 1 Then
        returnPONumber = mc.Item(0)
    else
        returnPONumber = 0
    End If
End Function

