;	bulk export spirometry results from SentrySuite as pdf or gdt
;	author: br-f
;	27 Feb 2024
;	requirements: 	AutoHotKey 1.1, SentrySuite 3.20 running maximized on 1920 x 1080 FHD resolution
;	instructions: 	Do not move any windows while the script is running.

global EXPORT_FORMAT := "pdf"							; default export format is pdf
global SOFTWARE_VERSION := "SentrySuite V3.20.8"				; insert software version you are running here

!^+g::
	global EXPORT_FORMAT := "gdt"
	MsgBox Export-Dateiformat GDT gewaehlt.
	return

!^+p::
	global EXPORT_FORMAT := "pdf"
	MsgBox Export-Dateiformat PDF gewaehlt.
	return

!^+e::										; execute main routine on Shift + Ctrl + Alt + e
	; requirements: SentrySuite running with main menu open and active
	software_specifier := "Home  " SOFTWARE_VERSION				; main SentrySuite window title					
	home_dir := "C:\path\to\dir\to\export\to\"				; insert path to directory to export to here
	
	if (A_Args.Length() < 1) or (A_Args[1] = ""){				; if no path to export list has been passed: terminate
		MsgBox Keine Liste zu exportierender Patienten erhalten. Kein Export moeglich.
		ExitApp
	}
	
	if !(EXPORT_FORMAT = "gdt" or EXPORT_FORMAT = "pdf"){
		MsgBox Kein Dateiformat fuer Export gewaehlt. Fuer gdt druecke Sondertasten + g, fuer pdf Sondertasten + p.
		Exit
	}
	
	if WinExist(software_specifier) {					; if SentrySuite is running
		WinActivate							; set focus to SentrySuite main window
	}
	WinWaitActive, %software_specifier%, , 5				; if not, throw error
	if ErrorLevel {
		MsgBox SentrySuite muss ausgefuehrt werden, um Befunde zu exportieren. Export abgebrochen.
		ExitApp
	} else {
		dir_to_export_to := home_dir "LUFU_Export_" A_NOW
		FileCreateDir %dir_to_export_to%				; create unique folder to export to using timestamp in
										; name conflicting with export
		list_of_patients_to_export := []				; list of patients by SentrySuite ID
		Loop, read, % A_Args[1]						; read passed export list file line by line
		{
			list_of_patients_to_export.Push(A_LoopReadLine)		; append each line (= one patient's SentrySuite ID) to array
		}
		
		for i, pat_id in list_of_patients_to_export {
			change_active_patient_to(pat_id)
			curr_visit := ""
			next_visit := ""
		
			Loop {
				curr_visit := get_visit()
				export(dir_to_export_to, pat_id, curr_visit)
				move_to_next_visit()
				next_visit := get_visit()
			}
			Until next_visit = curr_visit
			
			Click 1880, 80						; topright "Home" button, return to main screen
			Sleep 1000						; wait for measuring application to close
			if WinExist(software_specifier) {					
				WinActivate					; set focus to SentrySuite main window
			}
			WinWaitActive, %software_specifier%, , 5		; if unsuccessfull, throw error
			if ErrorLevel {
				MsgBox SentrySuite Hauptanwendung nicht gefunden. Export abgebrochen.
				ExitApp
			}
		}
		MsgBox Alle Befunde exportiert.					; success
		ExitApp
	}
	ExitApp

esc::
	MsgBox Export abgebrochen.
	ExitApp									; terminate script on pressing Escape
	
change_active_patient_to(sesuit_id) {
	; requirements: main menu open and active
	Click 1000, 500								; open "Patient"
	Sleep 200								; wait for "Patient" menu to load
	SendInput {Tab}{Tab}{Tab}{Tab}{Tab}					; tab forward to patient ID textbox
	SendInput %sesuit_id%							; insert patient ID
	SendInput {Tab}{Backspace}						; tab forward to last name field and clear it,
										; should there be any text left
	SendInput {Tab}{Backspace}						; tab forward to first name field and clear it
	SendInput {F1}								; search for the patient
	Sleep 1000
	SendInput {F2}								; ... and set it to be the active patient
	Sleep 1000
	Click 200, 450								; move selection to topmost (i.e. latest) visit
	SendInput {F1}								; set visit and return to main menu
	Sleep 500								; wait .5 sec
	Click 1400, 600								; open measurement application
	Sleep 15000								; wait 5 sec
	measure_window_specifier := "Bodyplethysmographie  " SOFTWARE_VERSION
	if WinExist(measure_window_specifier) {					; if measuring application is running
		WinActivate							; set focus to measuring application window
	}
	WinWaitActive, %measure_window_specifier%, , 5				; if not, throw error
	if ErrorLevel {
		MsgBox Mess-Anwendung nicht gefunden. Export abgebrochen.
		ExitApp
	}
}

get_visit() {
	; requirements: main window or measuring application window open, in focus and some visit is active
	curr_vis := ""
	wintext := ""
	WinGetText, wintext
	reached_date := False
	Loop, parse, wintext, %A_Space%, `n, `r,				; split retrieved string on linebreak or space and loop
	{									; through
		if(reached_date){
			curr_vis := StrSplit(A_Loopfield, "`n", "`r")[1]
			break
		}
		if(InStr(A_Loopfield, "Visite:") > 0) 
		{								; the "word" following "visite" is the date
			reached_date := True
		}
	}
	return curr_vis
}

move_to_next_visit() {
	; requirements: main window open, in focus ands some visit is active
	Click 100, 100								; open list of patient's vitis
	Sleep 1000								; wait for it to load
	SendInput {Tab 15}							; tab forward to list containing visits
	SendInput {Down}							; move selection to next (i.e. preceeding) visit
	SendInput {F1}								; set selected visit to be the active one
	Sleep 1500								; wait 1.5 sec
}

export(dir_path, pat_id, visit) {
	; requirements: Measuring application window open and active
	Click 1760, 160								; click "..." button
	Sleep 200								; wait for "select report" window to open
	Click 200, 500								; select "full view"
	Sleep 4000								; wait 5 sec for report view window to open
	visit_str := StrReplace(visit, ".")					; remove dots from visit date string
	export_path := dir_path "\LUFU_" pat_id "_" visit_str "." EXPORT_FORMAT
	if (EXPORT_FORMAT = "pdf"){
		SendInput {F7}							; "PDF output" button
		Sleep 1000							; wait for export dialog to open
		SendInput %export_path%						; specify path to save to
		SendInput {Enter}						; confirm
		Sleep 200							; wait for next dialog to open
		SendInput {Enter}						; "output" button
		Sleep 2000							; wait for export
	}
	if (EXPORT_FORMAT = "gdt"){
		SendInput {F10}							; "file output" button
		Sleep 200							; wait for "dropdown" menu to open
		SendInput {Enter}						; choose topmost entry in the "dropdown" menu, this is "GDT output"
		Sleep 1000							; wait for export dialog to open
		SendInput %export_path%						; specify path to save to
		SendInput {Enter}						; confirm
		Sleep 1000							; wait for export
	}
	Click 30, 40								; opens "Report" menu tab
	Click 50, 90								; exits the report window ("Beenden")
	Sleep 500								; wait for report window to close, this returns to the
										; measuring application window
	measure_window_specifier := "Bodyplethysmographie  " SOFTWARE_VERSION
	if WinExist(measure_window_specifier) {					; if measuring application is running
		WinActivate							; set focus to measuring application window
	}
	WinWaitActive, %measure_window_specifier%, , 5				; if not, throw error
	if ErrorLevel {
		MsgBox Mess-Anwendung nicht gefunden. Export abgebrochen.
		ExitApp
	}
}
