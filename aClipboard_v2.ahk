

; @TODO consider using richedit to highlight search terms in the clipboard edit control 
; @TODO consider pasting options, eg ^v | sendinput | sendtext 
; @TODO Delete file / all 



; ===================================================================
; Script Name:     [aClipboard]
; Author:          [https://github.com/alexofrhodes]
;
; Description:
;     [Brief overview of what the script does, its purpose, and key features.]
;
; Features:
;     - [Feature 1]
;     - [Feature 2]
;     - [Feature 3]
;
; Requirements:
;     - [List any dependencies, e.g., AHK v2, external libraries]
;     - [Special setup instructions if needed]
;
; Usage:
;     - [Explain how to run the script, hotkeys, and any special commands]
;
; Notes:
;     - [Any additional information, limitations, or future improvements]
;   Notable alternatives:
;     - [https://github.com/hluk/CopyQ]
;     - [Alternative 2]
; Changelog:
;   - [YYYY-MM-DD] - [Version 1.0.0]: [Brief description of changes]
; ===================================================================

Version := "1.0.0"

; ====== INCLUDE ======

#Requires AutoHotkey v2.0
#SingleInstance Force

#Include lib\Author.ahk
#Include lib\AnchorV2.ahk

#Include lib\GuiState.ahk
global iniFile := "guistate.ini"

; ====== Hotkeys ======
#c:: {
    if WinActive(myGui) {
        myGui.Hide()
    } else {
        showForm()
    }
}
$^c::       Copy()
$^x::       Cut()
!c::        Append()
!i::        Inject()
$!z::       LoopHistory(+1)
$!y::       LoopHistory(-1)
; $!v::       Paste()

; ====== GLOBAL VARIABLES ======

global previousClipboard := ""
global actionCheckboxes := [] ; Array to store transformation checkboxes

global historyDir := A_ScriptDir "\log"
if !DirExist(historyDir)
    DirCreate(historyDir)
global LogFiles := []
global historyLimit := 100          ; after which the oldest file will be deleted
Loop Files, historyDir . "\*.*" 
    LogFiles.InsertAt(1, StrReplace(A_LoopFileName, ".txt", "")) 

global currentIndex := 1

; ====== INITIALIZE ======

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

; ------------------------

CreateControls(*){

    global

    ; ====== FIND/REPLACE SECTION ======
    btnFind := myGui.Add("Button", "w100", "Find")
    btnFind.Enabled := false

    editFind := myGui.Add("Edit", "w500 x+m veditFind", "") 

    btnSave := myGui.Add("Button", "ys w100", "Save")
    btnSave.OnEvent("Click", SaveRegexValues)

    btnReplace := myGui.Add("Button", "xm w100", "Replace")
    btnReplace.OnEvent("Click", Replace)

    editReplace := myGui.Add("Edit", "x+m section w500 veditReplace") 


    btnLoad := myGui.Add("Button", "ys w100", "Load")
    btnLoad.OnEvent("Click", LoadRegexValues)

    chRegex := myGui.Add("Checkbox", "xs vchRegex", "Regex")

    chCaseSensitive := myGui.Add("Checkbox", "x+m vchCaseSensitive" , "CaseSensitive")

    lblLimit := myGui.Add("Text", "x+m section", "Limit")
    editLimit := myGui.Add("Edit", "x+m w20 ys-1 h16 veditLimit", "")

    lblStartingPos := myGui.Add("Text", "ys", "Start")
    editStartingPos := myGui.Add("Edit", "x+m ys-1 w20 h16 veditStartingPos", "")

    Separator1 := myGui.Add("Text", "xm h1 w720 0x10") ; Separator
    ; @OBSOLETE 
    ; @TODO Hotkeys to Copy, Call Function, Paste 
            ; #Include lib\HotkeyRemap.ahk

            ; MyArray := [  "Copy"
            ;             , "Append"            
            ;             , "Undo" 
            ;             ]

            ; For value in MyArray
            ;     AddHotkey(value, value, value)
            ; HotkeyRemap()

    ; ====== MENU BAR ======
    ; MyArray := [  "Copy"
    ;             , "Append"            
    ;             , "Inject"
    ;             , "Undo" 
    ;             , "Clear"
    ;             ;, "SetHotkeys"
    ;             ]

    ; MyMenuBar := MenuBar()
    ; For value in MyArray
    ;     MyMenuBar.Add(value, %value%)   ; (*) => ExecuteFunction(FunctionName))
    ; myGui.MenuBar := MyMenuBar

    addAuthorMenubar(myGui)

    ; ====== Append Options ======

    myGui.Add("Text", "xm section w100", "Append Options:")
    MyArray := [  "Linebreak"
                , "Space"
                , "Tab"
                ]
    
    AppendBy := myGui.Add("DropDownList","w100 vAppendBy", MyArray)                
    myGui.Add("Text", "xs h1 w100 0x10") ; separator

    ; Add dropdown to choose paste method
    MyGui.Add("Text", "xs", "Paste Method:")
    pasteMethods := ["Ctrl+V (Send)", "SendText", "SendInput"]
    ddlPasteMethod := MyGui.Add("DropDownList", "xs w100 vddlPasteMethod", pasteMethods)
    ddlPasteMethod.Choose(1)  ; Default to Ctrl+V
    myGui.Add("Text", "xs h1 w100 0x10") ; separator

    ; ====== ACTION CHECKBOXES ======

    myGui.Add("Text", "xm w100", "Select Actions:")

    ; Button to execute checked functions
    btnGetChecked := myGui.Add("Button", "xs w100", "Run Checked")
    btnGetChecked.OnEvent("Click", applyActionCheckboxes)

    AutoApply := myGui.Add("Checkbox", "xs vchAutoApply", "AutoApply")
    myGui.Add("Text", "xs h1 w100 0x10") ; separator

    MyArray := [
                    "Replace"
                    , "Tranclude"
                    , "BreakLinesToSpaces"
                    , "ToUpperCase"
                    , "ToLowerCase"
                    , "ReverseWords" 
                ]

    For value in MyArray
        actionCheckboxes.Push(myGui.Add("Checkbox", "xs vch" value, value)) ;vCh + Value is used to assign a name, to use with the GuiState functions

    myGui.Add("Text", "ys section w100", "Log Files")
    myGui.Add("Button", "xs", "Del Sel").OnEvent("Click", DeleteSelectedLogFile)
    myGui.Add("Button", "x+m", "Del All").OnEvent("Click", DeleteAllLogFiles)

    lbLogFiles := myGui.Add("ListBox", "xs h550 w120", LogFiles)
    lbLogFiles.OnEvent("Change", LoadSelectedLogFile)
    
    myGui.Add("Text", "ys section", "Clipboard Content")
    
    editClipboard := myGui.Add("Edit", "xs w450 h570 +HScroll -wrap") ;no vName, no need to save in the ini file

    if logfiles.Length > 0{
        lbLogFiles.Choose(1)  
        ; LoadSelectedLogFile()
    }
    if (A_Clipboard != editClipboard.Text)
        SaveClipboardToLogFile()
}


showForm() {
    myGui.Show("NoActivate")
    GuiLoadState(myGui.Title, iniFile)
}

UpdateStatusBar(*){
    SB.SetText("There are " . LogFiles.Length . " items in Clipboard History.")
}
; ====== EVENTS ======

Gui_Close(*){
    GuiSaveState(myGui.Title, iniFile)                
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


; ====== MAIN FUNCTIONS ======

Copy(*) {
    global
    flag := WinActive(myGui.hwnd)
    if flag
        toggleGUI()
    
    previousClipboard := A_Clipboard
    A_Clipboard := ""
    Send("{ctrl Down}c{ctrl up}")
    ClipWait(3)
    SaveClipboardToLogFile()  

    if AutoApply.Value 
        applyActionCheckboxes()

    if flag
        toggleGUI()
}

Cut(*) {
    global
    flag := WinActive(myGui.hwnd)
    if flag
        toggleGUI()

    previousClipboard := A_Clipboard
    A_Clipboard := ""
    Send("{ctrl Down}x{ctrl up}")
    ClipWait(3)
    SaveClipboardToLogFile()  

    if AutoApply.Value 
        applyActionCheckboxes()

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
    ClipWait

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

    if flag
        toggleGUI()
}

toggleGUI(*){
    if IsSet(myGui) && myGui {
        if WinActive(myGui.Hwnd) {
            myGui.Hide()
        } else {
            showForm()
        }
    } else {
        
    }
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
        if FileExist(filePath) {
            FileDelete(filePath)
            LogFiles.RemoveAt(lbLogFiles.Value)
            lbLogFiles.Delete(lbLogFiles.Value)
            if logfiles.length = 1
                currentIndex := 1
            lbLogFiles.Choose(currentIndex)
            UpdateStatusBar()
        }
    }
}

DeleteAllLogFiles(*) {
    global
    Loop Files, historyDir "\*.txt" {
        FileDelete(A_LoopFilePath)
    }
    LogFiles := []
    lbLogFiles.Delete()
    editClipboard.Value := ""
    currentIndex := 0
    UpdateStatusBar()
}

Inject(*){
    global
    flag := WinActive(myGui.hwnd)
    if flag
        toggleGUI()    previousClipboard := A_Clipboard
    A_Clipboard := ""
    Send("{ctrl Down}c{ctrl up}")
    ClipWait    
    EditPaste A_Clipboard, editClipboard
    A_Clipboard := editClipboard.Text     
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
        myGui.Opt("-AlwaysOnTop")
        MsgBox("No more history in this direction.")
        myGui.Opt("+AlwaysOnTop")
    }
}

; placeholder
; maybe we want to do something while pasting
; like pasting a formatted text without modifying the clipboard

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

    if flag
        toggleGUI()
}

; @TODO are these necessary?
Clear(*) {
    previousClipboard := A_Clipboard
    A_Clipboard := ""
    editClipboard.text := ""
}

; ====== Functions that modify copied text ======

applyActionCheckboxes(*) {
    for ctrl in actionCheckboxes {
        if (ctrl.Value) { ; Check if checkbox is checked
            FunctionName := ctrl.Text
            ExecuteFunction(FunctionName)
            try editClipboard.Value := A_Clipboard
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
}

Replace(*) {
    previousClipboard := A_Clipboard
    
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
    try editClipboard.Value := A_Clipboard

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

Transclude(*) { 
    previousClipboard := A_Clipboard
    if WinActive("ahk_class EVERYTHING") {
        A_Clipboard := ""
        Sleep(100)    
        Send("^+c")
        ClipWait(1)
    }
    This := A_Clipboard
    A_Clipboard := ""
    FileContent := ""
    Sleep(100)
    
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
    ToolTip("Replaced " counter " file paths in the clipboard with their content")
    Sleep(1000)
    ToolTip()
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

; ====== HELPER ======

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
