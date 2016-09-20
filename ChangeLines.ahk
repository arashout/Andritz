#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
AutoTrim, off

!m::
filePath := "C:\Users\arash\Desktop\test.vbs"
tempFilePath := "C:\Users\arash\Desktop\AHKtemp.vbs"

;Read the contents from specified template file
file := FileOpen(filePath, "r")
readContent := file.Read()

;Enter clipboard contents into variable
materialNum := Clipboard


;These lines are vbs
linesToAdd := "Dim materialNum`nmaterialNum = " . """" . materialNum . """" . " `n" ;All the quotes are to make a string in VBS - WHAT A PAIN!

;Write new content to temp file
writeContent := linesToAdd . readContent
fileTemp := FileOpen(tempFilePath, "w")
fileTemp.Write(writeContent)

;Run vbsscript
Run, wscript.exe "%tempFilePath%",,
return
