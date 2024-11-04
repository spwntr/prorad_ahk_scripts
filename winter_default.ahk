#Requires AutoHotkey v2.0
#SingleInstance

/* variables */
ahk_scripts := "AHK scripts are "
active_ref := "A"
first_diag_image_view_window_id := 0
second_diag_image_view_window_id := 0
ps := "PowerScribe"
f_4 := "{F4}"
f_12 := "{F12}"
shift := "+"
period := "."
backspace := "{backspace}"
too_many_periods := period period
space := "{space}"
space_period := space period
too_many_spaces := space space
tab := "{tab}"
page_up_to_start_at_top := "{PgUp}"
find_and_replace := "^h"
enter := "{enter}"
escape := "{escape}"
lshift_for_keywait := "LShift"
select_forward := true
keypad_right := "{Right}"
keypad_left := "{Left}"
shift_down := "{Shift down}"
ctrl_down := "{Ctrl down}"
shift_ctrl_down := shift_down ctrl_down
shift_up := "{Shift up}"
ctrl_up := "{Ctrl up}"
shift_ctrl_up := shift_up ctrl_up
shift_ctrl_command := shift_ctrl_down "{1}" shift_ctrl_up

default_sleep := 200 ; in ms
min_sleep := 50 ; in ms
search_sleep := 1000 ; in ms

transparency_value := 175


#SuspendExempt
; Suspend hotkey
; Ctrl + Alt + Q
^!q:: { 
	Suspend(-1)
	suspended_status := A_IsSuspended ? "suspended." : "active."
	MsgBox ahk_scripts suspended_status 
}
#SuspendExempt False


/* PowerScribe dictaphone functions */
; dictation deadmanswitch
; LShift + Space
<+Space:: Switch_to_PowerScribe_and_Send_with_KeyWait(f_4, lshift_for_keywait)

; toggle dictation
; LShift + `
<+`:: 
; F13
F13:: {	
	Switch_to_PowerScribe_and_Send(f_4)
}

; backward field
; LShift + 1
<+1::
; F14
F14:: {
	Switch_to_PowerScribe_and_Send(shift tab)
}

; forward field
; LShift + 2
<+2::
; F15
F15:: {
	Switch_to_PowerScribe_and_Send(tab)
}

; backward select
; LShift + 3
<+3::
; F16
F16:: {
	Switch_to_PowerScribe_and_Select(!select_forward)
}

; delete last word
; F21
<F21:: {
	Switch_to_PowerScribe_and_Delete_Last_Word()
}

; forward select
; LShift + 4
<+4::
; F17
F17:: {
	Switch_to_PowerScribe_and_Select(select_forward)
}

; minimize/restore PowerScribe
; F18
F18:: (WinGetMinMax(ps) = -1) ? WinRestore(ps) : WinMinimize(ps)

; toggle Powerscribe transparency
; F20
F20:: {
	trans_value := WinGetTransparent(ps)
	new_transparency_value := trans_value == transparency_value ? "Off" : " " transparency_value
	WinSetTransColor(new_transparency_value, ps)
}


/* PowerScribe report functions */
; sign report
; F19
F19:: Switch_to_PowerScribe_and_Send(f_12)

; find and replace "..", "  ", and " ."
; LCtl + Q
<^q:: {
	Find_Replace_PS(too_many_periods, period)
	Find_Replace_PS(space_period, period)
}


/* PACS functions */
; get 1st image viewer window id
; LCtl + LAlt + Left mouse click
<^<!LButton:: {
	first_diag_image_view_window_id := WinExist(active_ref)
	MsgBox "1st image viewer window id = " first_diag_image_view_window_id 
}

; get second image viewer window id
; LCtl + LAlt + Right mouse click
<^<!RButton:: {
	second_diag_image_view_window_id := WinExist(active_ref)
	MsgBox "2nd image viewer window id = " second_diag_image_view_window_id 
}


; functions
Switch_to_Window(window_name) {
	active_window := WinActive(window_name)
	different_initial_window := (active_window != window_name) or !active_window
	if different_initial_window {
		different_initial_window := WinExist(active_ref)
		WinActivate window_name
	}
	return different_initial_window
}

Switch_to_PowerScribe_and_Send(command) {
	different_initial_window := Switch_to_Window(ps)
	Send_Command(command)
	if different_initial_window {
		Switch_to_Window(different_initial_window)
	}
}

Send_Command(command) {
	Send_Command_with_Sleep(command, default_sleep)
}

Send_Command_with_Sleep(command, sleep_time) {
	Send command
	Sleep sleep_time
}

Switch_to_PowerScribe_and_Send_with_KeyWait(command, key_to_wait_for) {
	Switch_to_PowerScribe_and_Send(command)
	KeyWait key_to_wait_for
	Switch_to_PowerScribe_and_Send(command)
}

Switch_to_PowerScribe_and_Select(select) { 
	direction := select ? keypad_right : keypad_left
	Switch_to_PowerScribe_and_Send(shift_ctrl_down direction shift_ctrl_up)
}

Switch_to_PowerScribe_and_Delete_Last_Word() {
	Switch_to_PowerScribe_and_send(shift_ctrl_down keypad_left shift_ctrl_up backspace)
}

Find_Replace_PS(find, replace) {
	different_initial_window := Switch_to_Window(ps)
	Go_to_top_of_Report()
	Open_Find_Replace_Dialogue()
	Fill_Find_String(find)
	Fill_Replace_String(replace)
	Select_Replace_All()
	Escape_Find_Replace_Dialogue()
	if different_initial_window {
		Switch_to_Window(different_initial_window)
	}
}

Go_to_Top_of_Report() {
	Send_Command_with_Sleep(page_up_to_start_at_top, min_sleep)
	Send_Command_with_Sleep(page_up_to_start_at_top, min_sleep)
	Send_Command_with_Sleep(page_up_to_start_at_top, min_sleep)
}

Open_Find_Replace_Dialogue() {
	Send_Command_with_Sleep(find_and_replace, default_sleep)
}

Fill_Box_With_String(string_text) {
	Send_Command_with_Sleep(tab, min_sleep) 
	Send_Command_with_Sleep(string_text, default_sleep)
}

Fill_Find_String(find) {
	Fill_Box_With_String(find)
}

Fill_Replace_String(replace) {
	Fill_Box_With_String(replace)
}

Select_Replace_All() {
	Send_Command_with_Sleep(tab, min_sleep)
	Send_Command_with_Sleep(tab, min_sleep)
	Send_Command_with_Sleep(tab, min_sleep)
	Send_Command_with_Sleep(enter, search_sleep)
}

Escape_Find_Replace_Dialogue() {
	Send_Command_with_Sleep(escape, min_sleep)
	Send_Command_with_Sleep(escape, min_sleep)
}
