#SingleInstance Force

Show_Me(this) {
	Loop % this.MaxIndex()
		Content .= this[A_Index] "`n"
	TrayTip,, %Content%
}

;Create array object to hold clipboard variables
queue := Object()
queue.Show_Me := Func("Show_Me")

!Esc::
   Suspend, Permit
   SusToggle := !SusToggle
   If (SusToggle)
   {   
		Suspend, On
		TrayTip,, "ClipQueue Paused"
   }
   Else
   {   
		Suspend Off
		TrayTip,, "ClipQueue Unpaused"
   }
   Return

;Insert values from clipboard into array
!q::
	Clipboard = ;
	SendInput, ^c
	ClipWait, 0.5
	queue.Push(Clipboard)
	queue.Show_Me()
	Return

;Paste values from queue
!w::
	Clipboard := queue.RemoveAt(1) ;Remove first item
	SendInput, %Clipboard%
	queue.Show_Me()
	Return

;Reset queue
!r::
	queue := Object()
	queue.Show_Me := Func("Show_Me")
	TrayTip,, "Queue reset"
	Return