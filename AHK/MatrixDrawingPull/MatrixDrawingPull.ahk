#SingleInstance Force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Include FindClick.ahk

Show_Me(this) {
	Loop % this.MaxIndex()
		Content .= this[A_Index] "`n"
	Msgbox % Content
}
ClickImage(imgPath, xOffset:=0, yOffset:=0, ByRef x:=0, ByRef y:=0){
	ImageSearch, x, y, 0, 0, A_ScreenWidth, A_ScreenHeight, %imgPath%
	if ErrorLevel = 2
		Return 2
	else if ErrorLevel = 1
		Return 1
	else
		MouseClick, Left, x + xOffset, y + yOffset
		Return 0
}
GetDrawings(drawingPath, ByRef o){
	Loop, read, %drawingPath%
    {
        o.Push(A_LoopReadLine)
    }
	o.Show_Me := Func("Show_Me")
	return
}
GetImageCoordArr(imgPath){
	unparsedCoords := FindClick(imgPath, "e")
	StringReplace,unparsedCoords,unparsedCoords,`n,`,,All
	StringReplace,unparsedCoords,unparsedCoords,`r,,All
	return StrSplit(unparsedCoords,"`,")
}
ClearTextBox(){
	Send, ^a
	Send, {Delete}
}
drawingList := Object()
GetDrawings("Drawings.txt", drawingList)
WinActivate, Engineering Cockpit - Internet Explorer ahk_class IEFrame
;ClickImage("Images\DwgName.PNG",100, 5)
;ClearTextBox()
;Send, *701670114*
;ClickImage("Images\OldDwgNumber.PNG",100, 10)
;coords := GetImageCoordArr("Images\download")
	

