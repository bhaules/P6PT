#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
#installKeybdHook
#Persistent

;;;;;# - Win | ! - Alt | ^ - Ctrl | + - Shift
;;;;Some Stuff I found on the internet for visual and audio indication that the script was restarted
;Menu, Tray, Icon , Shell32.dll, 25, 1
;TrayTip, test, Started, 1
;SoundBeep, 300, 150

;
;;;Reload Script -- launch this with Ctrl+Win+Alt+R
; ^#!r::
; Send, ^s ; To save a changed script
; Sleep, 300 ; give it time to save the script
; Reload
; Return
; ^!#e::Edit
Menu, Tray, Icon, Icons/P6PT.ico

Gui, New, hwndhGui AlwaysOnTop Resize MinSize ; always on top
Gui, Add, CheckBox, Checked vEnterGoes1Row gCheckSub, Enter - Enter and Go Down 1 row
Gui, Add, CheckBox, Checked vActivatePredWindow gCheckSub, Alt+Q - activates predecessor window
Gui, Add, CheckBox, Checked vCheck2 gCheckSub Checked, F2
Gui, Show
Gui, +LastFound
Gui, Submit, NoHide
return

GuiClose:
ExitApp

CheckSub:
Gui, Submit, NoHide
return
;;=============================================================================================
; Enter presses Enter and goes dowon one row
;=============================================================================================
#If (Check2 && WinActive("ahk_exe PM.exe"))

F2::
MsgBox, F2 Hotkey was pressed in P6.
return
#If

#If (EnterGoes1Row && WinActive("ahk_exe PM.exe"))

Enter::
Send, {Enter}
sleep, 100
Send, {Down} ; Sends down after 100 ms
return
#IfWinActive

MsgBox, Enter Hotkey was pressed in P6.
return
#If
; #If (Check2 && WinActive("ahk_class TDevMainForm"))

; F2::
; MsgBox, F2 Hotkey was pressed in P6.
; return
; #If

; #If
;;=============================================================================================
; Alt+q goes to the predecessors window
;=============================================================================================

#If (ActivatePredeWindow && WinActive("ahk_exe P6.exe"))
!q::
ControlClick, TCVirtualQueryGrid2, ahk_class TDevxMainForm, , , 1 ; click the button
return
#If
;;=============================================================================================
; Alt+w goes to the successors window
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm ; activate the window
!w::
ControlClick, TCVirtualQueryGrid1, ahk_class TDevxMainForm, , , 1 ; click the button
return
#IfWinActive
;;=============================================================================================
; Alt+e goes to the main activity window
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm ; activate the window
!e::
SetControlDelay, -1 ; set control delay to zero
ControlClick, TCUltraButton27, ahk_class TDevxMainForm, , , 1 ; click the button
sleep, 100
Send, {Tab} ; activates the activity window
Sleep, 200
return
#IfWinActive
;;=============================================================================================
; Alt+1 goes to the next activity
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm ; activate the window
!1::
SetControlDelay, -1 ; set control delay to zero
ControlClick, TCUltraButton27, ahk_class TDevxMainForm, , , 1 ; click the button
sleep, 100
Send, {Tab} ; activates the activity window
Sleep, 200
ControlClick, TCVirtualQueryGrid2, ahk_class TDevxMainForm, , , 1 ; click the button
return
#IfWinActive

#IfWinActive, ahk_class TDevxMainForm ; activate the window
!2::
SetControlDelay, -1 ; set control delay to zero
ControlClick, TCUltraButton28, ahk_class TDevxMainForm, , , 1 ; click the button
sleep, 100
Send, {Tab} ; activates the activity window
ControlClick, TCVirtualQueryGrid2, ahk_class TDevxMainForm, , , 1 ; click the button
return
#IfWinActive
;=============================================================================================
; double click in the predecessors/ successors takes you to the pred/successor 
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm ; activate the window
~LButton::
If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 200)
	{
		MouseGetPos, , , , controlClassNN, 1
		;Msgbox , controlClassNN=%controlClassNN%
		If (controlClassNN = "TCVirtualQueryGrid2")
			{ControlClick, TCUltraButton34, ahk_class TDevxMainForm, , , 1 ; click the button
			return
		}
		Else If (controlClassNN = "TCVirtualQueryGrid1")
			{ControlClick, TCUltraButton29, ahk_class TDevxMainForm, , , 1 ; click the button
			return
		}
		Else {
			return
		}
	}
return
#IfWinActive
;=============================================================================================
;Tab goes to either predecessors or successors
;=============================================================================================
;#IfWinActive, ahk_class TDevxMainForm ; activate the window
;Tab::
	;If (controlClassNN = "TCVirtualQueryGrid2") ; Predecessors
		;{ControlClick, TCVirtualQueryGrid1, ahk_class TDevxMainForm, , , 1 ; Go to successors
		;return
	;}
	;Else If (controlClassNN = "TCVirtualQueryGrid1") ; Successors
		;{ControlClick, TCVirtualQueryGrid2, ahk_class TDevxMainForm, , , 1 ; Go to Predecessors
		;return
		;}
;=============================================================================================
;Ctrl+Alt+x goes to predecessors
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm ; Check if Window is active
^!z:: ;Ctrl + Alt + z - Go to Predecessors
SetControlDelay, -1 ; set control delay to zero
ControlClick, TCUltraButton34, ahk_class TDevxMainForm, , , 1 ; click the button
sleep, 1000
ControlClick, TCVirtualQueryGrid2, ahk_class TDevxMainForm, , , 1 ; click the button
return
#IfWinActive

;=============================================================================================
;Ctrl+Alt+x goes to successors
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm ; activate the window
^!x:: ;Ctrl + Alt + x - Go to Sucessors
SetControlDelay, -1 ; set control delay to zero
ControlClick, TCUltraButton29, ahk_class TDevxMainForm, , , 1 ; click the button
sleep, 1000
ControlClick, TCVirtualQueryGrid1, ahk_class TDevxMainForm, , , 1 ; click the button
return
#IfWinActive



;=============================================================================================
;Scroll Left/ Right in P6 with mouse
;=============================================================================================
; #If MouseIsOver("ahk_class TDevxMainForm")
; WheelRight::+WheelDown
; WheelLeft::+WheelUp
; #if
; return
;=============================================================================================
;Mouse Back and Forward button goes down/ up one page at a time 
;=============================================================================================
#IfWinActive ahk_class TDevxMainForm
XButton1::PgDn
XButton2::PgUp
return
#IfWinActive
;=============================================================================================
;Hold Right Click and scroll up/ down - scrolls a whole screen
;=============================================================================================
#IfWinActive ahk_class TDevxMainForm
;RButton::RButton
XButton1 & WheelDown::PgDn
XButton1 & WheelUp::PgUp
return
#IfWinActive
;=============================================================================================
;Keyboard Volume_Down scrolls down in P6
;=============================================================================================
#IfWinActive ahk_class TDevxMainForm
Volume_Down::
	Send, {CtrlDown}{WheelDown}
	Sleep 500
	Send, {Click}
return
Volume_Up::
	Send {CtrlDown}{WheelUp}
	Sleep, 500
	Send, {Click}
return
#IfWinActive

;=============================================================================================
;Ctrl+Shift+f opens the filters window (Alt+e+k )
;=============================================================================================


;=============================================================================================
;Ctrl+Shift+a opens the layouts window (Alt+e+k )
;=============================================================================================

	
;=============================================================================================
;ctrl+shift+d opens user preferences and changes duraton format in weeks 
;=============================================================================================
#ifwinactive ahk_exe pm.exe
^+w::
	send {altdown}e 
	sleep 100
	send u{altup}
	sleep 100
	send {tab 9}
	sleep 100
	send w
	sleep 100
	send {tab 4}
	sleep 100
	send {enter}
return
#ifwinactive
;=============================================================================================
;ctrl+shift+d opens user preferences and changes duraton format in days 
;=============================================================================================
#ifwinactive ahk_exe pm.exe
^+d::
	send {altdown}e 
	sleep 100
	send u{altup}
	sleep 100
	send {tab 9}
	sleep 100
	send d
	sleep 100
	send {tab 4}
	sleep 100
	send {enter}
return
#ifwinactive
;=============================================================================================
;Ctrl+Shift+a opens the layouts window (Alt+e+k )
;=============================================================================================
#IfWinActive ahk_class TDevxMainForm
^+a::
	Send {AltDown}v 
	Sleep 100
	Send o
	sleep 100
	Send {Down}{AltUp}
	Send {Enter}
return
#IfWinActive
;=============================================================================================
;F6:Disolves Activities (Alt+e+o )
;=============================================================================================
#IfWinActive ahk_class TDevxMainForm
F6::
	;Loop 10
	;{	Sleep 500
	;	Send {Down} ; you need to have selected the previous row
	;	Sleep 100 
	;	Send {ShiftDown}{Down}{ShiftUp}
	;	Sleep 500
		Send {Alt}e 
		Sleep 100
		Send o
	;}
return
#IfWinActive
;=============================================================================================
;Ctrl+q Links Activities (Alt+e+k )
;=============================================================================================
#IfWinActive ahk_class TDevxMainForm
^q::
	;Loop 10
	;{	Sleep 500
	;	Send {Down} ; you need to have selected the previous row
	;	Sleep 100 
	;	Send {ShiftDown}{Down}{ShiftUp}
	;	Sleep 500
		Send {AltDown}e 
		Sleep 100
		Send k{AltUp}
	;}
return
#IfWinActive

;=============================================================================================
;F1 text macro adds .COW to the end of the text
;=============================================================================================
;#IfWinActive ahk_class TDevxMainForm
;F1::
;	Send {F2} 
;	Sleep 100
;	Send {End}
;	Sleep 100
;	Send {.}COW
;	Sleep 100
;	Send {Enter}
;return

;=============================================================================================
;F2 goes to the end of the text rather the beginning
;=============================================================================================
;#IfWinActive ahk_class TDevxMainForm
;F2::
	;Send {F2} 
	;Sleep 100
	;Send {End}
;return
;#IfWinActive

;=============================================================================================
;Delete and press y when pressing delete
;=============================================================================================
#IfWinActive ahk_exe PM.exe
+Delete::
Send {Delete} 
	WinWait, ahk_class TfrmERMsg, 
	IfWinNotActive, ahk_class TfrmERMsg, WinActivate, ahk_class TfrmERMs
	WinWaitActive, ahk_class TfrmERMsg
	;Sleep 1000
Send {y}
return
#IfWinActive

;=============================================================================================
;Middle click sends click and ctrl+tab to expand and contract a node
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm 
MButton::
	Send {LButton} 
	Sleep 100
	Send {Ctrl Down}{Tab}
	KeyWait, Control ;Waits for the control key to be released. Not sure why this is needed but it worked
	Send {Ctrl Up}
return
#IfWinActive

;=============================================================================================
; Backtick sends Ctrl + Tab to expand & contract grouping bands
;=============================================================================================
#IfWinActive, ahk_exe PM.exe 
SC029:: ; Symbol for backtick ;'
	Send {Ctrl Down}{Tab}
	KeyWait, Control ;Waits for the control key to be released. Not sure why this is needed but it worked
	Send {Ctrl Up}
return
#IfWinActive
;;=============================================================================================
; Alt+ Backtick collapses all
;=============================================================================================
#IfWinActive, ahk_exe PM.exe 
!SC029:: ; Symbol for backtick ;'
	Send {Ctrl Down}{NumpadSub}
	KeyWait, Control ;Waits for the control key to be released. Not sure why this is needed but it worked
	Send {Ctrl Up}
return
#IfWinActive;Back Mouse Button click sends click and ctrl+tab to expand and contract a node
;#IfWinActive ahk_class TDevxMainForm
;XButton1::
	;Send {LButton} 
	;Sleep 100
	;Send {Ctrl Down}{Tab}
	;KeyWait, Control ;Waits for the control key to be released. Not sure why this is needed but it worked
	;Send {Ctrl Up}
;return

;=============================================================================================
;Deletes one word at a time when pressing CTRL+BackSpace
;=============================================================================================
#IfWinActive, ahk_class TDevxMainForm 
^BackSpace::
	Send {CtrlDown}{ShiftDown}{Left}
	Sleep, 100
	Send, {Del}
return
#IfWinActive

;=============================================================================================
;Presses F9 and schedules the project without going through the prompts
;=============================================================================================
#IfWinActive ahk_exe PM.exe
F9::
Send {F9} 
	WinWait, ahk_class TfrmScheduler, 
	IfWinNotActive, , , WinActivate, ahk_class TfrmScheduler, 
	WinWaitActive, ahk_class TfrmScheduler, 
Send {Enter} 
	WinWait, ahk_class TfrmERMsg, ,2 ;waits for 2 seconds and then stops the next actions 
	if ErrorLevel
	{
		; MsgBox, waited for 2 seconds. No other users were present in the project
		return
	}
	; MsgBox, activate window and press Y
	IfWinNotActive, ahk_class TfrmERMsg, , WinActivate, ahk_class TfrmERMsg, 
	WinWaitActive, ahk_class TfrmERMsg,
Send {y}
Return
#IfWinActive

;=============================================================================================
;Presses Ctrl+F9 calculate float paths to a copied value 
;=============================================================================================
#IfWinActive ahk_exe PM.exe
^F9::
Send {F9} 
	WinWait, ahk_class TfrmScheduler, 
	IfWinNotActive, , , WinActivate, ahk_class TfrmScheduler, 
	WinWaitActive, ahk_class TfrmScheduler, ;loads the schedule window and activates it
Send, {TAB 3}
Sleep, 100
;MsgBox, Schedule Options Open
Send, {SPACE} ;Opens Schedule Options
	WinWait, Schedule Options, 
	IfWinNotActive, Schedule Options, , WinActivate, Schedule Options, 
	WinWaitActive, Schedule Options, 
;MsgBox, Schedule Options Open
Send, {TAB 5} ; 5 tabs Goes to the advanced section
Sleep,100
Send, {RIGHT}{TAB}{CTRLDOWN}v{CTRLUP}{TAB} ; pastes the ID and loads the Activity Name into the window
Sleep, 100
Send, {TAB 3} ;goes to the Close button
Sleep, 1000
Send, {SPACE} ;This should exit the Options window 

;MsgBox, "Wait for message about saving to appear"
WinWait, ahk_class TfrmERMsg, ,2 ;if the message pops up about saving changes it will press Yes
	if ErrorLevel
	{
		return
	}
	IfWinNotActive, ahk_class TfrmERMsg, , WinActivate, ahk_class TfrmERMsg, 
	WinWaitActive, ahk_class TfrmERMsg,
;MsgBox, "Send Y"
Send {y}
WinWait, ahk_class TfrmScheduler, ; waits for the schedule window and tabs in order to press the schedule button
	IfWinNotActive, , , WinActivate, ahk_class TfrmScheduler, 
	WinWaitActive, ahk_class TfrmScheduler, ;loads the schedule window and activates it
	Sleep, 2000
Send, {TAB}
Sleep, 500
Send, {ENTER}
;MsgBox, "Sent 6 tabs"
	;this section is executed if there are multiple users in the project.waits for 2 seconds and then stops the next actions 
	WinWait, ahk_class TfrmERMsg, ,2 
	if ErrorLevel
	{
		return
	}
	IfWinNotActive, ahk_class TfrmERMsg, , WinActivate, ahk_class TfrmERMsg, 
	WinWaitActive, ahk_class TfrmERMsg,
Send {y}
return
#IfWinActive
;