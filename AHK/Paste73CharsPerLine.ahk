#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

;Create array object to hold clipboard variables
global stringArray := Object()

!p::
	Clipboard = ;
	Send, ^c
	ClipWait
	multiLineString := Clipboard
	numCharsPerLine := 73
	numChars := StrLen(multiLineString)
	numLoop := Floor(numChars/numCharsPerLine) + 1
	pointer := 1
	Loop %numLoop% {
		line := SubStr(multiLineString, pointer , numCharsPerLine)
		stringArray.Insert(line)
		pointer := numCharsPerLine + pointer
	}
return

!o::
	numLoop := stringArray.MaxIndex()
	Loop %numLoop%{
		v := stringArray.RemoveAt(1)
		SendRaw, %v%
		Send, {Down}
	}
return