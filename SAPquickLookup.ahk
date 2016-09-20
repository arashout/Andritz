#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
AutoTrim, off

^!m::
;Send Copy command
Clipboard = ;
Send, ^c
ClipWait, 0.5
;Enter clipboard contents into variable
materialNum := Clipboard

folderPath := "\\DELSMS015\Everyone\ArashOutadi\Github\WorkRepos\Andritz\VBS_SAP_Operations\"
operation := "mm03"
filePath := folderPath . operation . ".vbs"

;If copying from Excel - remove any possible new lines
StringReplace, materialNum, materialNum,`n,,A
StringReplace, materialNum, materialNum,`r,,A

;Run vbsscript with parameters
Run, cscript.exe //nologo "%filePath%" "%materialNum%",,Hide
return
