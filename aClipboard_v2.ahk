

; @TODO consider using richedit to highlight search terms in the clipboard edit control? 

; ====== INFO ====== 
{
    ; ===================================================================
    ; Script Name:     [aClipboard]
    ; Author:          [https://github.com/alexofrhodes]
    ;
    ; Usage:
    ;   -   See the HOTKEYS section bellow
    ;
    ; Requirements
    ;   -   AnchorV2.ahk
    ;   -   GuiState.ahk
    ;
    ; Alternatives:
    ;     - [https://github.com/hluk/CopyQ]
    ;     - [Alternative 2]
    ;
    ; Changelog:
    ;   - [2025-02-17] - [Version 1.0.0]: [Initial Release]
    ; ===================================================================
}

Version := "1.0.0"

; ====== INCLUDE ======
{
    #Requires AutoHotkey v2.0
    #SingleInstance Force

    #Include lib\Author.ahk
    #Include lib\AnchorV2.ahk

    #Include lib\GuiState.ahk
    global iniFile := "GuiState.ini"
}
; ====== Hotkeys ======
{
    #c:: {    ; Toggle GUI
        if WinActive(myGui) {
            myGui.Hide()
        } else {
            showForm()
        }
    }
    $^c::       Copy()          ; Ctrl + C
    $^x::       Cut()           ; Ctrl + X
    $^!c::      Append()        ; Ctrl + Alt   + C 
    $^!x::      CutAppend()     ; Ctrl + Alt   + X 
    $^+c::      Prepend()       ; Ctrl + Shift + P 
    $^+x::      CutPrepend()    ; Ctrl + Shift + X 
    $!i::       Inject()        ; Alt  + I 
    $#z::       Undo()          ; Win  + Z 
    $#y::       Redo()          ; Win  + Y 
    $#v::       Paste()         ; Win  + V
    $^#V::      Format()        ; Ctrl + Win + V
    ^s:: { 
        if WinActive(myGui) 
            SaveEditClipboardChanges()
        }
}

; ====== GLOBAL VARIABLES ======
{
    global previousClipboard := ""
    global actionCheckboxes := [] ; Array to store transformation checkboxes

    global historyDir := A_ScriptDir "\log"
    if !DirExist(historyDir)
        DirCreate(historyDir)

    global LogFiles := []
    global historyLimit := 100          ; after which the oldest file will be deleted
    Loop Files, historyDir . "\*.*" 
        LogFiles.InsertAt(1, StrReplace(A_LoopFileName, ".txt", "")) 

    global currentIndex := 1    ; used for lbLogfiles

    global Datatype_Empty      := 0
    global Datatype_Text       := 1
    global Datatype_NonText    := 2
    global ClipboardDatatype   := Datatype_Empty
    OnClipboardChange ClipChanged, -1   

    ; OnClipboardChange Callback , [AddRemove]
    ;If omitted, it defaults to 1. Otherwise, specify one of the following numbers:
    ;  1 = Call the callback after any previously registered callbacks.
    ; -1 = Call the callback before any previously registered callbacks.    <==
    ;  0 = Do not call the callback.

    A_Clipboard := A_Clipboard
}
; ====== INITIALIZE ======
{
    myGui := Gui()
    
    myGui.Title := "aClipboard"
    myGui.Opt("+AlwaysOnTop +Resize" ) ;  -LastFound )+E0x08000000") ; Prevents GUI from stealing focus

    SB := myGui.AddStatusBar()
    UpdateStatusBar()

    myGui.OnEvent("Close", Gui_Close)
    myGui.OnEvent("Size", OnSize)
    OnMessage(0x404, AHK_NOTIFYICON)  

    CreateControls()
    showForm()
}

ClipChanged(DataType) {
    global
    ClipboardDatatype := DataType
}

; ====== GUI FUNCTIONS ======
{
    CreateControls(*){

        global

        ; ====== FIND/REPLACE SECTION ======
        btnFind := myGui.Add("Button", "w120", "Find")
        btnFind.Enabled := false

        editFind := myGui.Add("Edit", "w250 x+m veditFind", "") 

        btnSave := myGui.Add("Button", "ys w120", "Save")
        btnSave.OnEvent("Click", SaveRegexValues)

        btnReplace := myGui.Add("Button", "xm w120", "Replace")
        btnReplace.OnEvent("Click", (*)=> ExecuteFunction(Replace))

        editReplace := myGui.Add("Edit", "x+m section w250 veditReplace") 

        btnLoad := myGui.Add("Button", "ys w120", "Load")
        btnLoad.OnEvent("Click", LoadRegexValues)

        chRegex := myGui.Add("Checkbox", "xs vchRegex", "Regex")

        chCaseSensitive := myGui.Add("Checkbox", "x+m vchCaseSensitive" , "CaseSensitive")

        lblLimit := myGui.Add("Text", "x+m section", "Limit")
        editLimit := myGui.Add("Edit", "x+m w20 ys-1 h16 veditLimit", "")

        lblStartingPos := myGui.Add("Text", "ys", "Start")
        editStartingPos := myGui.Add("Edit", "x+m ys-1 w20 h16 veditStartingPos", "")

        Separator1 := myGui.Add("Text", "xm h1 w500 0x10") ; Separator

        ; ====== MENU BAR ======
        MyArray := [  "Copy"
                    , "Cut"
                    , "Paste"
                    , "Format"
                    , "Append"    
                    , "Prepend"
                    , "CutAppend"   
                    , "CutPrepend"     
                    , "Input"
                    , "Inject"
                    , "Undo" 
                    , "Redo"
                    , "Help"
                    ]

        MyMenuBar := MenuBar()
        For value in MyArray{
            MyMenuBar.Add(value, %value%)   
            targetFile := A_ScriptDir . "\icons\" . value . ".ico"
            if (FileExist(targetFile))
                MyMenuBar.SetIcon(value, targetFile) ; , IconNumber, IconWidth)
       }

        myGui.MenuBar := MyMenuBar

        addAuthorMenubar(myGui)

        ; ====== Append Options ======

        myGui.Add("Text", "xm section w120", "Append/Prepend Opts:")
        MyArray := [  "Linebreak"
                    , "Space"
                    , "Tab"
                    ]
        
        AppendBy := myGui.Add("DropDownList","w120 vAppendBy", MyArray)                
        myGui.Add("Text", "xs h1 w120 0x10")    ; separator

        MyGui.Add("Text", "xs", "Paste Method:")
        pasteMethods := ["Ctrl+V", "SendText", "SendInput"]
        ddlPasteMethod := MyGui.Add("DropDownList", "xs w120 vddlPasteMethod", pasteMethods)
        ddlPasteMethod.Choose(1)                ; Default to Ctrl+V
        
        MyGui.Add("Text", "xs", "After Paste:")
        ddlAfterPasteAction := myGui.Add("DropDownList", "xs w120 vddlAfterPasteAction", ["None", "Previous", "Next"])
        ddlAfterPasteAction.choose(1)
        
        myGui.Add("Text", "xs h1 w120 0x10")    ; separator

        ; ====== ACTION CHECKBOXES ====== 

        myGui.Add("Text", "xm w120", "Select Actions:")

        ; Button to execute checked functions
        btnGetChecked := myGui.Add("Button", "xs w120", "Apply Selected")
        btnGetChecked.OnEvent("Click", applyActionCheckboxes)

        AutoApply := myGui.Add("Checkbox", "xs vchAutoApply", "Auto Apply")
        myGui.Add("Text", "xs h1 w120 0x10") ; separator
        btnToggle := myGui.AddButton("w120 ", "Toggle Actions")
        btnToggle.OnEvent("Click", ToggleActions)
        
        ;===============================================================================
        ; @MODIFY (1) - Add the names of your clipboard modifying functions to this array. See @MODIFY(2)
        ;===============================================================================

        MyArray := [
                      "Replace"
                    , "Transclude"
                    , "BreakLinesToSpaces"
                    , "ToUpperCase"
                    , "ToLowerCase"
                    , "ReverseWords" 
                    ]
        
        ;vCh + Value is used to assign a name, to use with the GuiState functions
        For value in MyArray
            actionCheckboxes.Push(myGui.Add("Checkbox", "xs vch" value, value)) 
        myGui.Add("Text", "ys section w120", "Log Files")
        myGui.Add("Button", "xs w55", "Del Sel").OnEvent("Click", DeleteSelectedLogFile)
        myGui.Add("Button", "x+m w55", "Del All").OnEvent("Click", DeleteAllLogFiles)

        btnRename := myGui.Add("Button", "xs w120", "Rename")
        btnRename.OnEvent("Click", RenameLogFile)

        myGui.Add("Text", "xs w120", "Filter")
        editFilter := myGui.Add("Edit", "xs w120", "")
        editFilter.OnEvent("Change", FilterLogFiles)

        lbLogFiles := myGui.Add("ListBox", "xs h400 w150 HScroll", LogFiles) 
        lbLogFiles.OnEvent("Change", LoadSelectedLogFile)
        
        myGui.Add("Text", "ys section", "Clipboard Content")
        btnSaveEditClipboard := myGui.Add("Button", "ys w120", "Save Changes")
        btnSaveEditClipboard.OnEvent("Click", SaveEditClipboardChanges)        
        editClipboard := myGui.Add("Edit", "xs w240 h470 +HScroll -wrap") ;no vName, no need to save in the ini file

        if logfiles.Length > 0{
            lbLogFiles.Choose(1)  
            LoadSelectedLogFile()
        }
        if (A_Clipboard != editClipboard.Text)
            SaveClipboardToLogFile()
    }

    Help(*) {
        HelpHTML := "
        (
        <html>
        <body>
            <style>
                body { font-family: Arial, sans-serif; }
                table { border-collapse: collapse; width: 100%; }
                th, td { border: 1px solid black; padding: 5px; text-align: left; }
                th { background-color: #f2f2f2; }
            </style>
            <table>
                <tr><th>Hotkey</th><th>Action</th></tr>
                <tr><td>#C</td><td>Toggle GUI</td></tr>
                <tr><td>Ctrl + C</td><td>Copy</td></tr>
                <tr><td>Ctrl + X</td><td>Cut</td></tr>
                <tr><td>#V</td><td>Paste by chosen PasteMethod</td></tr>
                <tr><td>^#V</td><td>Format Selected Text</td></tr>
                <tr><td>Ctrl + Alt + C</td><td>Append</td></tr>
                <tr><td>Ctrl + Alt + X</td><td>Cut Append</td></tr>
                <tr><td>Ctrl + Shift + P</td><td>Prepend</td></tr>
                <tr><td>Ctrl + Shift + X</td><td>Cut Prepend</td></tr>
                <tr><td>Alt + I</td><td>Inject</td></tr>
                <tr><td>Win + Z</td><td>Undo</td></tr>
                <tr><td>Win + Y</td><td>Redo</td></tr>
                <tr><td>Ctrl + S</td><td>If gui active, Save Edit Changes</td></tr>
            </table>
        </body>
        </html>
        )"
    
        ; Create GUI
        helpGui := Gui(, "Hotkey Guide")
        browser := helpGui.AddActiveX("x10 y10 w500 h500", "Shell.Explorer")
    
        ; Wait until the ActiveX control is fully ready
        ComObject := browser.Value
        if (ComObject) {
            ComObject.Navigate("about:blank")  ; Navigate to an empty page first
            Loop {
                Sleep 50  ; Wait until the page is ready
            } Until !ComObject.Busy && ComObject.ReadyState = 4
    
            ComObject.document.write(HelpHTML)  ; Inject the HTML content
            ComObject.document.close()
        }
        helpGui.Opt("+AlwaysOnTop")
        helpGui.Show("AutoSize ")
    }
    
    showForm(*) {
        myGui.Show("NoActivate x5000 y5000") ;show offscreen to avoid fliccker from guiLoadState
        GuiLoadState(myGui.Title, iniFile)
    }

    toggleGUI(*){
        if IsSet(myGui) && myGui {
            if WinActive(myGui.Hwnd) {
                GuiSavePosition()
                myGui.Hide()
            } else {
                showForm()
            }
        } else {
            
        }
    }

    RenameLogFile(*) {
        global 
    
        if (currentIndex = 0) {
            MsgBox "Please select a file to rename."
            return
        }
    
        oldFileName := logFiles[currentIndex]
        oldFilePath := historyDir "\" oldFileName ".txt"
    
        ; Extract timestamp part (assuming it's always at the start)
        RegExMatch(oldFileName, "^\d{4}-\d{2}-\d{2}-\d{6}", &match)
        timestamp := match[0]
    
        if (!timestamp) {
            MsgBox "Invalid file format. Cannot rename."
            return
        }
    
        ; Ask for a new name
        myGui.Opt("-AlwaysOnTop")
        IB := InputBox("Enter a new name for the log file:", "Rename Log File")
        myGui.Opt("+AlwaysOnTop")
        newName := IB.Value
        if (IB.Result = "Cancel" || newName = "")
            return
        
        ; Remove old custom name if one exists
        newFileName := timestamp " " newName
        newFilePath := historyDir "\" newFileName ".txt"
    
        ; Ensure the new name does not exist
        while FileExist(newFilePath) {
            myGui.Opt("-AlwaysOnTop")
            IB := InputBox(newfilename " already exists.`nEnter a new name for the log file:","Rename Log File")
            myGui.Opt("+AlwaysOnTop")
            newName := IB.Value                
            if (IB.Result := "Cancel" || newName = "")
                return
            newFileName := timestamp " " newName
            newFilePath := historyDir "\" newFileName ".txt"
        } 
        FileMove(oldFilePath, newFilePath)
        logFiles[currentIndex] := newFileName
        lbLogFiles.Delete
        lbLogFiles.add(logFiles)
        lbLogFiles.choose(currentIndex)
    }
    
    FilterLogFiles(*) {
        global
        currentIndex := 0
        lbLogFiles.choose(0)
        searchTerm := editFilter.Text

        LogFiles := []
        Loop Files, historyDir . "\*.*"  
        {
            item := StrReplace(A_LoopFileName, ".txt", "")
            if searchTerm = ""
                LogFiles.InsertAt(1, item)
            else if (InStr(A_LoopFileName, searchTerm))  
                LogFiles.InsertAt(1, item)
        }
        lbLogFiles.delete
        if LogFiles.length > 0 {
            lbLogFiles.Add(LogFiles)
        }
        if LogFiles.length > 0{
            lbLogFiles.choose(1)
            LoadSelectedLogFile()
        }
    }

    LoadSelectedLogFile(*) {
        global 
        currentIndex := lbLogFiles.Value
        selectedFile := lbLogFiles.Text . ".txt"
        if (selectedFile != "") {
            filePath := historyDir "\" selectedFile
            editClipboard.Value := FileRead(filePath)
        }
    }
  
    DeleteSelectedLogFile(*) {
        global
        if (lbLogFiles.Value > 0) {
            if (LogFiles.length = 1) {
                DeleteAllLogFiles()
                return
            }
    
            selectedFile := lbLogFiles.text . ".txt"
            filePath := historyDir "\" selectedFile
            
            ; Check if the file has a manual name
            if RegExMatch(selectedFile, "^\d{4}-\d{2}-\d{2}-\d{6} (.+)") {
                userChoice := MsgBox("This file has a manual name. Do you still want to delete it?",, 3)  ; Yes/No/Cancel
                if (userChoice = "No" || userChoice = "Cancel") {
                    return  ; Do nothing if No or Cancel is chosen
                }
            }
    
            if FileExist(filePath) {
                FileDelete(filePath)
                LogFiles.RemoveAt(lbLogFiles.Value)
                lbLogFiles.Delete(lbLogFiles.Value)
    
                if (LogFiles.length > 0) {
                    currentIndex := Min(currentIndex, LogFiles.length)  ; Ensure valid index
                    lbLogFiles.Choose(currentIndex)
                    LoadSelectedLogFile()
                } else {
                    currentIndex := 0
                    editClipboard.Value := ""
                }
    
                UpdateStatusBar()
            }
        }
    }
    
    DeleteAllLogFiles(*) {
        global
        manualFilesExist := false
        deletedFiles := []  ; Store deleted files
        remainingFiles := []  ; Store files that were not deleted
        
        ; Check if ANY file has a manual name
        Loop Files, historyDir "\*.txt" {
            if RegExMatch(A_LoopFileName, "^\d{4}-\d{2}-\d{2}-\d{6} (.+)") {
                manualFilesExist := true
                break
            }
        }
    
        ; Ask user if they want to delete manually named files (Yes/No/Cancel)
        if manualFilesExist {
            userChoice := MsgBox("Some files have manual names. Do you want to delete them as well?",,"0x1000 3")
            if (userChoice = "Cancel") {
                return  ; Stop if Cancel is chosen or no file exists
            }
        } else {
            userChoice := "Yes"  ; If no manual files exist, proceed with deletion
        }
    
        ; Loop through files and delete based on user choice
        Loop Files, historyDir "\*.txt" {
            isManual := RegExMatch(A_LoopFileName, "^\d{4}-\d{2}-\d{2}-\d{6} (.+)")
            if (userChoice = "Yes" || !isManual) {
                FileDelete(A_LoopFilePath)
                deletedFiles.Push(A_LoopFileName)
            } else {
                remainingFiles.Push(StrReplace(A_LoopFileName, ".txt", ""))  ; Keep manually named files
            }
        }
    
        ; Update UI based on deleted files
        if deletedFiles.Length > 0 {
            LogFiles := remainingFiles
            lbLogFiles.Delete()
    
            ; Re-add remaining files (if any)
            if remainingFiles.Length > 0 {
                lbLogFiles.Add(LogFiles)
                lbLogFiles.Choose(1)
                currentIndex := 1
                LoadSelectedLogFile()
            } else {
                currentIndex := 0
                editClipboard.Value := ""
            }
    
            UpdateStatusBar()
        }
    }
    
    

    UpdateStatusBar(*){
        SB.SetText("There are " . LogFiles.Length . " items in Clipboard History.")
    }

    SaveState(*){
        GuiSaveState(myGui.Title, iniFile)
    }
    SaveEditClipboardChanges(*){
        if lbLogfiles.value =0
            return
        A_Clipboard := editClipboard.text
        targetFile := historyDir "\" . lbLogFiles.text . ".txt"
        FileDelete(targetFile)
        FileAppend(A_Clipboard, targetFile)
    }
}
; ====== GUI EVENTS ======
{
    Gui_Close(*){
        SaveState()               
        ExitApp()
    }

    OnSize(GuiObj, MinMax, Width, Height){    
        Anchor(editFind.Hwnd, "w")
        Anchor(editReplace.Hwnd, "w")    
        Anchor(btnSave.Hwnd, "x")
        Anchor(btnLoad.Hwnd, "x")    

        Anchor(Separator1.Hwnd, "w")

        Anchor(lbLogFiles.Hwnd, "h")
        Anchor(editClipboard.Hwnd,"hw")
    }

    AHK_NOTIFYICON(wParam, lParam,*){
        if (lParam = 0x201) ; WM_LBUTTONDOWN
        {
            showForm()
        }
    }
}
; ====== MAIN FUNCTIONS ======
{
    Input(*){
        myGui.opt("-AlwaysOnTop")
        userInput := InputBox("Enter text to copy:", "Clipboard Input")
        myGui.opt("+AlwaysOnTop")

        if userInput.Result = "OK" && userInput.Value != "" {
            A_Clipboard := userInput.Value
            SaveClipboardToLogFile()
            ; MsgBox "Text copied to clipboard: `n" userInput.Value
        } else {
            ; MsgBox "No text entered. Clipboard unchanged."
        }
    }

    Copy(*) {
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()
        
        previousClipboard := A_Clipboard
        A_Clipboard := ""
        Send("{ctrl Down}c{ctrl up}")
        
        if !ClipWait(3) || (ClipboardDatatype = !Datatype_Text) { 
            A_Clipboard := previousClipboard
            ; MsgBox "The attempt to copy text onto the clipboard failed."
        } else {
            SaveClipboardToLogFile()
            if AutoApply.Value 
                applyActionCheckboxes()
        }
        if flag
            toggleGUI()
    }

    Format(*){
        global
        ddlAfterPasteAction.choose(1) ; "None" = don't cycle
        AutoApply.value := 1
        Copy()
        Paste()
    }

    Cut(*) {
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()

        previousClipboard := A_Clipboard
        A_Clipboard := ""
        Send("{ctrl Down}x{ctrl up}")
        if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
            A_Clipboard := previousClipboard
            ; MsgBox "The attempt to copy text onto the clipboard failed."
        } else {
            SaveClipboardToLogFile()
            if AutoApply.Value 
                applyActionCheckboxes()            
        }
        if flag
            toggleGUI()
    }

    Append(*) { 
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()

        previousClipboard := A_Clipboard
        This := A_Clipboard
        A_Clipboard := ""
        Send("{ctrl Down}c{ctrl up}")
        if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
            A_Clipboard := previousClipboard
            ; MsgBox "The attempt to copy text onto the clipboard failed."
        } else {
            appendMethod := AppendBy.Text
            appendSeparator := "`r`n"  ; Default to Linebreak
            if (appendMethod = "Space")
                appendSeparator := " "
            else if (appendMethod = "Tab")
                appendSeparator := "`t"
    
            if (This != "") 
                A_Clipboard := This appendSeparator A_Clipboard
    
            SaveClipboardToLogFile()  
            if AutoApply.Value 
                applyActionCheckboxes()            
        }
        if flag
            toggleGUI()
    }

    CutAppend(*) {
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()

        previousClipboard := A_Clipboard        
        This := A_Clipboard
        A_Clipboard := ""
        Send("{ctrl Down}x{ctrl up}")
        if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
            A_Clipboard := previousClipboard
            ; MsgBox "The attempt to copy text onto the clipboard failed."
        } else {
            appendMethod := AppendBy.Text
            appendSeparator := "`r`n"  ; Default to Linebreak
    
            if (appendMethod = "Space")
                appendSeparator := " "
            else if (appendMethod = "Tab")
                appendSeparator := "`t"
    
            if (This != "") 
                A_Clipboard := This appendSeparator A_Clipboard
    
            SaveClipboardToLogFile()  
            if AutoApply.Value 
                applyActionCheckboxes() 
        }
        if flag
            toggleGUI()        
    } 

    Prepend(*) {
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()
    
        previousClipboard := A_Clipboard
        This := A_Clipboard
        A_Clipboard := ""
        Send("{ctrl Down}c{ctrl up}")
        if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
            A_Clipboard := previousClipboard
            ; MsgBox "The attempt to copy text onto the clipboard failed."
        } else {
            prependMethod := AppendBy.Text  ; Using the same dropdown selection
            prependSeparator := "`r`n"      ; Default to Linebreak
        
            if (prependMethod = "Space")
                prependSeparator := " "
            else if (prependMethod = "Tab")
                prependSeparator := "`t"
        
            if (This != "") 
                A_Clipboard := A_Clipboard prependSeparator This
        
            SaveClipboardToLogFile()
            if AutoApply.Value 
                applyActionCheckboxes()
        }    
        if flag
            toggleGUI()
    }
    
    CutPrepend(*) {
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()
    
        previousClipboard := A_Clipboard        
        This := A_Clipboard
        A_Clipboard := ""
        Send("{ctrl Down}x{ctrl up}")
        if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
            A_Clipboard := previousClipboard
            ; MsgBox "The attempt to copy text onto the clipboard failed."
        } else {
        prependMethod := AppendBy.Text
        prependSeparator := "`r`n"  ; Default to Linebreak
    
        if (prependMethod = "Space")
            prependSeparator := " "
        else if (prependMethod = "Tab")
            prependSeparator := "`t"
    
        if (This != "") 
            A_Clipboard := A_Clipboard prependSeparator This
    
        SaveClipboardToLogFile()
        if AutoApply.Value 
            applyActionCheckboxes()
        }       
        if flag
            toggleGUI()        
    }
    
    Undo(*) {
        LoopHistory(1)
    }

    Redo(*) {
        LoopHistory(-1)
    }

    SaveClipboardToLogFile(*) {
        global
        
        fileName := FormatTime(A_Now, "yyyy-MM-dd-HHmmss")
        FileAppend(A_Clipboard, historyDir "\" . fileName . ".txt")

        ; Add to history array
        LogFiles.InsertAt(1, fileName)

        ; If history exceeds the limit, delete the oldest file
        if (LogFiles.Length > historyLimit) {
            oldFile := LogFiles.RemoveAt(LogFiles.Length)  ; Remove from the array (we're storing them in reverse order)
            FileDelete(oldFile)                      
        }

        currentIndex := 1
        UpdateStatusBar()

        ; Add new file to the top of ListBox
        lbLogFiles.delete()
        lbLogFiles.add(LogFiles)
        lbLogFiles.Choose(1)  
        LoadSelectedLogFile()
    }

    Inject(*){
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()    previousClipboard := A_Clipboard
        A_Clipboard := ""
        Send("{ctrl Down}c{ctrl up}")
        if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
            A_Clipboard := previousClipboard
            ; MsgBox "The attempt to copy text onto the clipboard failed."
        } else {
            EditPaste A_Clipboard, editClipboard
            A_Clipboard := editClipboard.Text     
            SaveClipboardToLogFile()
        }
        if flag
            toggleGUI()
    }

    LoopHistory(direction) {
        global 
        newIndex := currentIndex + direction

        ; Check if the new index is within valid bounds
        if (newIndex >= 1 && newIndex <= LogFiles.Length) {
            currentIndex := newIndex ;update the global currentIndex
            fileToRead := historyDir . "\" . LogFiles[currentIndex] .  ".txt"
            A_Clipboard := FileRead(fileToRead)
            try editClipboard.Value := A_Clipboard
            lbLogFiles.Choose(currentIndex)
        } else {
            MsgBox("No more history in this direction.",,0x1000)
        }
    }

    Paste(*) {
        global
        flag := WinActive(myGui.hwnd)
        if flag
            toggleGUI()

        ; Get selected method from DropDownList
        pasteMethod := ddlPasteMethod.Text

        ; Use the selected method for pasting
        if (pasteMethod = "Ctrl+V") {
            Send("{ctrl Down}v{ctrl up}")
        } else if (pasteMethod = "SendText") {
            SendText(A_Clipboard) ; Send raw clipboard content
        } else if (pasteMethod = "SendInput") {
            SendInput(A_Clipboard) ; Simulate user typing
        }

        selectedMode := ddlAfterPasteAction.Text  ; Get selected option
        if (selectedMode = "Previous") {
            LoopHistory(1)
        } else if (selectedMode = "Next") {
            LoopHistory(-1)
        }

        if flag
            toggleGUI()
    }

}
; ====== MAIN clipboard mod functions  ====== 
{

    ToggleActions(*) {
        global 
        checkState := 0  ; Assume all are checked
    
        ; Check if any checkbox is unchecked
        for _, checkbox in actionCheckboxes {
            if (checkbox.Value = 0) {  
                checkState := 1  ; If any is unchecked, we set checkState to check all
                break
            }
        }
    
        ; Apply checkState to all checkboxes
        for _, checkbox in actionCheckboxes {
            checkbox.Value := checkState
        }
    }
    

    applyActionCheckboxes(*) {
        for ctrl in actionCheckboxes {
            if (ctrl.Value) { ; Check if checkbox is checked
                FunctionName := ctrl.Text
                ExecuteFunction(FunctionName)
            }
        }
        editClipboard.Value := A_Clipboard
    }
    ExecuteFunction(FunctionName) {
        previousClipboard := A_Clipboard
        if IsFunc(FunctionName) {
            %FunctionName%.Call() ; Call function dynamically
        } else {
            Switch FunctionName, false
            {
                Case "test"      : MsgBox("test")
                Default: MsgBox("Function '" FunctionName "' not found!")
            }
        }
        try editClipboard.Value := A_Clipboard
    }
    Replace(*) {
        ; Validate input: Ensure the "Find" field is not empty
        if (editFind.text = "") {
            MsgBox("Error: 'Find' field cannot be empty!", "Input Error", 0x30)
            return
        }
        ; Assign defaults if fields are empty
        if (editStartingPos.text = "")
            editStartingPos.text := "1"
        if (editLimit.text = "")
            editLimit.text := "-1"
        limit := editLimit.Value + 0
        startingPos := editStartingPos.Value + 0
        ; If Regex mode is enabled, use RegExReplace
        if chRegex.value {
            A_Clipboard := RegExReplace(A_Clipboard, editFind.text, editReplace.text, , limit, startingPos)
        } else {
            A_Clipboard := StrReplace(A_Clipboard, editFind.text, editReplace.text, chCaseSensitive.Value, , limit)
        }
        if (previousClipboard = A_Clipboard) {
            MsgBox("No matches found for '" editFind.text "'!", "No Change", 0x40)
        }
    }
    SaveRegexValues(*) {
        fileName := InputBox("Enter file name", "Enter the desired file name or leave it blank to select a file:", "").value
        
        if (fileName == "") 
            return
        Folder := A_ScriptDir "\regex"
        if !DirExist(Folder)
            DirCreate(Folder)
        filePath := Folder "\" StrReplace(fileName ".edit",".edit.edit",".edit")
        try 
            FileDelete(filePath) 
        FileAppend("Find: " editFind.Value "`nReplace: " editReplace.Value "`n", filePath)
        MsgBox("Saved regex values to " filePath)
    }
    LoadRegexValues(*) {
        myGui.Opt("-AlwaysOnTop")
        filePath := FileSelect(3, , "Open a file", "Text Documents (*.edit; *.doc)")
        myGui.Opt("+AlwaysOnTop")
        if filePath = ""
            return
        fileContent := FileRead(filePath)
        ; Assuming format "Find: <value>" and "Replace: <value>"
        findMatch := RegExMatch(fileContent, "Find: (.*)")
        if (findMatch) {
            editFind.Value := findMatch[1]
        }
        replaceMatch := RegExMatch(fileContent, "Replace: (.*)")
        if (replaceMatch) {
            editReplace.Value := replaceMatch[1]
        }
        MsgBox("Loaded regex values from " filePath)           
    }
}

; ====== USER clipboard mod functions ====== 

;===============================================================================
; @MODIFY(2) 
;===============================================================================

; Bellow this point add your new functions 
    ; they just need to modify the clipboard as a final result
    ; Remember to add the function name to the actionCheckboxes array at @MODIFY(1)

{
  
    Transclude(*) { 
        This := A_Clipboard
        FileContent := ""        
        FileRegex := "m)^C:\\.*\.(txt|ahk|bas|md|ahk|py|csv|log|ini|config)"
        pos := 1
        counter := 0
        while (RegExMatch(This, FileRegex, &Match, pos)) {
            FileName := Match[0]
            if (FileExist(FileName)) {
                FileContent := FileRead(FileName)
                This := StrReplace(This, FileName, FileContent)
            }
            counter += 1
            pos := Match.Pos + Match.Len
        }
        A_Clipboard := This

    }
    BreakLinesToSpaces() {
        A_Clipboard := StrReplace(A_Clipboard, "`r`n", " ")
    }
    ToUpperCase() {
        A_Clipboard := StrUpper(A_Clipboard)
    }
    ToLowerCase() {
        A_Clipboard := StrLower(A_Clipboard)
    }
    ReverseWords() {
        words := StrSplit(A_Clipboard, " ")
        A_Clipboard := ArrayToString(words, " ", true) ; Reverse order
    }
}
    ; ====== HELPER FUNCTIONS ======
{
    IsFunc(FunctionName){
        Try{
            return %FunctionName%.MinParams+1
        }
        Catch{
            return 0
        }
        return
    }

    ArrayToString(arr, delimiter := " ", reverse := false) {
        if reverse
            arr := arr.Clone(), arr.Reverse()
        return arr.Join(delimiter)
    }

}