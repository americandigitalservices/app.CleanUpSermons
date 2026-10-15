on run
	set chosenFile to choose file with prompt "Select the Sermon Outline (.docx, .pdf, .eml, or .txt):" of type {"org.openxmlformats.wordprocessingml.document", "com.adobe.pdf", "com.apple.mail.email", "public.plain-text", "docx", "pdf", "eml", "txt"}
	processSermon(POSIX path of chosenFile)
end run

on open droppedItems
	set theFile to POSIX path of (item 1 of droppedItems)
	processSermon(theFile)
end open

on processSermon(theFile)
	display dialog "Enter Speaker Name (or leave as-is to use default / .eml sender):" default answer "Pastor Ernest Fink"
	set speakerName to text returned of result
	
	tell application "System Events"
		set scriptDir to POSIX path of (container of (path to me))
	end tell
	if scriptDir does not end with "/" then set scriptDir to scriptDir & "/"
	
	set scriptPath to scriptDir & "clean_sermon.py"
	set venvDir to scriptDir & "venv"
	set venvPy to venvDir & "/bin/python3"
	set venvPip to venvDir & "/bin/pip"
	set reqPath to scriptDir & "requirements.txt"
	
	set cmd to "export PATH=\"/opt/homebrew/bin:/usr/local/bin:$PATH\"; "
	set cmd to cmd & "if [ ! -x " & quoted form of venvPy & " ]; then "
	set cmd to cmd & "  SYS_PY=$(command -v python3); "
	set cmd to cmd & "  \"$SYS_PY\" -m venv " & quoted form of venvDir & "; "
	set cmd to cmd & "  if [ -f " & quoted form of reqPath & " ]; then "
	set cmd to cmd & "    " & quoted form of venvPip & " install --upgrade pip -q && "
	set cmd to cmd & "    " & quoted form of venvPip & " install -r " & quoted form of reqPath & " -q; "
	set cmd to cmd & "  fi; "
	set cmd to cmd & "fi; "
	set cmd to cmd & quoted form of venvPy & " " & quoted form of scriptPath & " " & quoted form of theFile & " " & quoted form of speakerName
	
	set cmdOutput to do shell script cmd
	
	set sermonTextPath to POSIX path of (path to desktop folder) & "SermonText.txt"
	
	set dialogResult to display dialog cmdOutput & return & return & "Import SermonText.txt into ProPresenter?" with title "Sermon Outline Cleaner" buttons {"Done", "Import to ProPresenter"} default button "Import to ProPresenter"
	
	if button returned of dialogResult is "Import to ProPresenter" then
		importFileIntoProPresenter(sermonTextPath)
	end if
end processSermon

on ensureProPresenterReady()
	if application "ProPresenter" is not running then
		set launchChoice to display dialog "ProPresenter is not currently open. Please open ProPresenter and wait for it to finish loading, or click 'Launch ProPresenter' below." with title "ProPresenter Not Running" buttons {"Cancel", "Launch ProPresenter"} default button "Launch ProPresenter"
		if button returned of launchChoice is "Cancel" then
			return false
		end if
		tell application "ProPresenter" to activate
	else
		tell application "ProPresenter" to activate
	end if
	
	tell application "System Events"
		repeat 60 times
			if exists process "ProPresenter" then
				tell process "ProPresenter"
					set frontmost to true
					if (count of windows) > 0 and (exists menu bar item "File" of menu bar 1) then
						delay 0.5
						return true
					end if
				end tell
			end if
			delay 0.5
		end repeat
	end tell
	
	display dialog "ProPresenter did not finish loading a main window in time. Please make sure ProPresenter is open and ready on screen, then try again." with title "ProPresenter Not Ready" buttons {"OK"} default button "OK"
	return false
end ensureProPresenterReady

on importFileIntoProPresenter(sermonFilePath)
	if not ensureProPresenterReady() then return
	
	tell application "System Events"
		tell process "ProPresenter"
			set frontmost to true
			
			try
				click menu item "File…" of menu 1 of menu item "Import" of menu 1 of menu bar item "File" of menu bar 1
			on error
				try
					click menu item "File..." of menu 1 of menu item "Import" of menu 1 of menu bar item "File" of menu bar 1
				on error
					keystroke "i" using {shift down, command down}
				end try
			end try
			delay 0.8
			
			keystroke "g" using {shift down, command down}
			delay 0.6
			
			set the clipboard to sermonFilePath
			keystroke "v" using {command down}
			delay 0.4
			key code 36
			delay 0.5
			key code 36
			delay 0.9
			key code 36
		end tell
	end tell
end importFileIntoProPresenter