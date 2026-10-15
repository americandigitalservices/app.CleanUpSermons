-- Triggered when double-clicking the app or running from Script Editor (no args passed)
on run
    set chosenFile to choose file with prompt "Select the Sermon Outline (.docx, .pdf, .eml, or .txt):" of type {"org.openxmlformats.wordprocessingml.document", "com.adobe.pdf", "com.apple.mail.email", "public.plain-text", "docx", "pdf", "eml", "txt"}
    processSermon(POSIX path of chosenFile)
end run

-- Triggered when dragging and dropping a file onto the app icon
on open droppedItems
    set theFile to POSIX path of (item 1 of droppedItems)
    processSermon(theFile)
end open

-- Shared handler that resolves paths, sets up venv if missing, prompts for speaker, and runs Python
on processSermon(theFile)
    -- 1. Prompt for Speaker Name
    display dialog "Enter Speaker Name (or leave as-is to use default / .eml sender):" default answer "Pastor Ernest Fink"
    set speakerName to text returned of result
    
    -- 2. Locate the directory where this AppleScript lives
    tell application "System Events"
        set scriptDir to POSIX path of (container of (path to me))
    end tell
    if scriptDir does not end with "/" then set scriptDir to scriptDir & "/"
    
    set scriptPath to scriptDir & "clean_sermon.py"
    set venvDir to scriptDir & "venv"
    set venvPy to venvDir & "/bin/python3"
    set venvPip to venvDir & "/bin/pip"
    set reqPath to scriptDir & "requirements.txt"
    
    -- 3. Create venv & install requirements.txt if missing, then run clean_sermon.py
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
    
    -- 4. Display the full multi-line output from the Python script
    display dialog cmdOutput with title "Sermon Outline Cleaner" buttons {"OK"} default button "OK"
end processSermon