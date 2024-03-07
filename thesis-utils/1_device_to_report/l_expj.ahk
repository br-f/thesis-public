;	bulk export spirometry results from JLab 5 as pdf or lte
;	author: br-f
;	27 Feb 2024
;	requirements: AutoHotKey 1.1, JLab 5 running maximized on 1024 x 768 XGA resolution
;	instructions: Do not move any windows while the script is running. No test should be
;				  currently selected for any patient you wish to export reports for.

global EXPORT_FORMAT := "pdf"							; per default export as pdf

#!^+l::
	global EXPORT_FORMAT := "lte"
	MsgBox Export-Dateiformat LTE gewaehlt.
	return

#!^+p::
	global EXPORT_FORMAT := "pdf"
	MsgBox Export-Dateiformat PDF gewaehlt.
	return

#!^+e::										; execute main routine on Shift + Ctrl + Win + Alt + e
	; requirements: JLab running with main menu open and active, as reached by F12
	software_specifier := "LabManager V5.32.0"				; insert software version you are running here
	home_dir := "C:\Users\User\Desktop\Befunde PDF\"			; insert path to directory to export to here
	
	if (A_Args.Length() < 1) or (A_Args[1] = "") {				; if no path to export list has been passed: terminate
		MsgBox Keine Liste zu exportierender Patienten erhalten. Kein Export moeglich.
		ExitApp
	}
	
	if !(EXPORT_FORMAT = "lte" or EXPORT_FORMAT = "pdf"){
		MsgBox Kein Dateiformat fuer Export gewaehlt. Fuer lte druecke Sondertasten + l, fuer pdf Sondertasten + p.
		Exit
	}
	
	if WinExist(software_specifier) {					; if JLab is running
		WinActivate							; set focus to JLab
	}
	WinWaitActive, %software_specifier%, , 5				; if not, throw error
	if ErrorLevel {
		MsgBox JLab muss ausgefuehrt werden, um Befunde zu exportieren.
	} else {
		dir_to_export_to := home_dir "LUFU_Export_" A_NOW
		FileCreateDir %dir_to_export_to%				; create unique folder to export to using timestamp in
										; name conflicting with export
		list_of_patients_to_export := []				; list of patients by JLab ID
		Loop, read, % A_Args[1]						; read passed export list file line by line
		{
			list_of_patients_to_export.Push(A_LoopReadLine)		; append each line (= one patient's JLab ID) to array
		}
		
		for i, pat_id in list_of_patients_to_export {
			num_of_tests_to_export := 0
			exported_count := 0
		
			change_active_patient_to(pat_id)
			num_of_tests_to_export := get_num_of_tests_available()
		
			while exported_count < num_of_tests_to_export {
				move_to_next_visit(exported_count, num_of_tests_to_export)
				if (EXPORT_FORMAT = "lte"){
					export_lte(dir_to_export_to, pat_id, exported_count, num_of_tests_to_export)
				}
				else if (EXPORT_FORMAT = "pdf"){
					export_pdf(dir_to_export_to, pat_id, exported_count, num_of_tests_to_export)
				}
				else {
					MsgBox Export abgebrochen. Kein Dateiformat ausgewaehlt.
					ExitApp
				}
				exported_count++
			}
			
			SendInput {F12}						; return to main menu
			Sleep 1000						; wait for main menu to load
			
			Process, Close, TXTEDI32.exe				; garbage collect "Text Editor" (awfully heavy on CPU otherwise)
			Process, WaitClose, TXTEDI32.exe, 10			; wait a max of 10 sec for text editor to close
			if ErrorLevel {						; if it cant be closed, terminate program to prevent
				MsgBox Fehler. Export abgebrochen.		; unforeseeable behaviour due to CPU being fully occupied by
				ExitApp						; TXTEDI32.exe
			}
		}
		MsgBox Alle Befunde exportiert.					; success
		ExitApp
	}

esc::
	MsgBox Export abgebrochen.
	ExitApp									; terminate script on pressing Escape
	
change_active_patient_to(jlab_id) {
	; requirements: main menu open and active, as reached by F12
	SendInput p								; open "Patientendaten"
	Sleep 500								; wait for "Patientendaten" menu to load
	SendInput {Tab}{Tab}							; tab forward to patient ID textbox
	SendInput %jlab_id%							; insert patient ID
	SendInput {Enter}							; set current patient to desired one
	;Sleep 200
	Click 1240, 430								; open "Testverzeichnis"
	Click 150, 480								; set focus in list of visits by selecting topmost
	SendInput {F11}								; open "Programmauswahl"
	Click 100, 700								; open "Bildschirm Report"
	Sleep 4000								; wait 4 sec for "Bildschirm Report" screen to open
	if (EXPORT_FORMAT = "lte"){
		Click 50, 125							; select "AA-MODULATE-CF"
	}
	else if (EXPORT_FORMAT = "pdf"){
		Click 50, 190							; select "BODYPDF"
	}
	else {
		MsgBox Export abgebrochen. Kein Dateiformat ausgewaehlt.
		ExitApp
	}
	Click 300, 60								; "OK"
	Sleep 1000								; wait 1 sec for report to load
}

get_num_of_tests_available() {
	; requirements: "Bildschirm Report" window open and active
	num_of_tests_available := 0
	WinGetTitle, bildschirmreport_window_title, A				; remember current window title to return to later
	SendInput {F2}								; open "Text Editor"
	Sleep 500								; wait 1/2 sec for text editor window to open
	Click 20, 40								; open "Text"
	Click 50, 60								; open "Lesen", this opens new window "Testverzeichnis"
	DetectHiddenText On							; needs to be enabled to retrieve no. of tests
	Sleep 500
	WinGetText, test_count_parse_str, A					; retrieve text, this follows the format
										; "OK\nAbbrechen\n<NUMBER OF TESTS AVAILABLE> Tests"
	Loop, parse, test_count_parse_str, `n, `r, %A_Space%			; split retrieved string on linebreak or space and loop
	{									; through
		IfNotInString, OKAbbrechenTests, %A_Loopfield% 
		{								; by principle of exclusion, following the format as
										; described above, the substring which is none of the
										; specified is the number of visits available
			num_of_tests_available := A_LoopField			; assign current substring
			num_of_tests_available += 0				; by conducting a numeric operation, this forces the
										; variable to be converted to integer as AHK converts
										; string to int as needed
		}
	}
	
	Click 700, 80								; close "Testverzeichnis"
	Click 340, 10								; close "Text Editor" window, after that JLab will not
										; be in focus anymore
	if WinExist(bildschirmreport_window_title) {
		WinActivate							; resume focus
	}
	WinWaitActive, %bildschirmreport_window_title%, , 5
	if ErrorLevel {
		MsgBox JLab muss ausgefuehrt werden, um Befunde zu exportieren.
	;	ExitApp
	} else {
		return num_of_tests_available
	}
}

move_to_next_visit(exported_count, max_num) {
	; requirements: "Bildschirm Report" window open and active
	SendInput {F3}								; open "Testauswahl"
	Sleep 500								; wait for "Testauswahl" to open
	SendInput r								; "Reset", clear list of visits currently in use
	Click 640, 100								; set focus to list of test available
	SendInput {Up}								; this will select the second most up to date visit
	SendInput {Down}							; this will deselect the former and select the most
										; recent visit
										; only using down key does not lead to any selection
	i := 0
	while i < max_num {
		SendInput {Down}
		++i
	}
	i := 0
	While i < exported_count {						; skip as many tests as have already been exported
		SendInput {Up}							; and continue with the next one up in line
		++i
	} 
	Click 700, 70								; add selected visit to list of those currently in use
	Click 400, 430								; "OK", this will lead back to the "Bildschirm Report"
										; screen with "BODYPDF" or "AA-MODULATE-CF" settings, resp.
	Sleep 500								; wait 1/2 sec for it to load
}

export_lte(dir_path, pat_id, exported_count, max_num) {
	; requirements: "Bildschirm Report" window open and active with "AA-MODULATE-CF" settings
	SendInput {F9}								; open "Druckerausgabe"
	Sleep 500
	test_num_chronologically := max_num - exported_count			; as we export them in reversed order
	export_path := dir_path "\LUFU_" pat_id "_" test_num_chronologically ".lte"
	SendInput %export_path%							; write desired path to export to in printer dialog
	SendInput {Enter}							; execute export, this will lead back to "Bildschirm
										; Report" screen
	Sleep 500								; wait 1/2 sec for it to load
}

export_pdf(dir_path, pat_id, exported_count, max_num) {
	; requirements: "Bildschirm Report" window open and active with "BODYPDF" settings
	SendInput {F6}								; open "Druckerausgabe"
	Sleep 500
	SendInput {Enter}							; "AUSGEBEN"
	Sleep 2000								; wait for Windows system printer dialog to open
	SendInput ^a								; Ctrl + A to select whole predefined export path
	test_num_chronologically := max_num - exported_count			; as we export them in reversed order
	export_path := dir_path "\LUFU_" pat_id "_" test_num_chronologically ".pdf"
	SendInput %export_path%							; write desired path to export to in printer dialog
	SendInput {Enter}							; execute export, this will lead back to "Bildschirm
										; Report" screen
	Sleep 500								; wait 1/2 sec for it to load
}
