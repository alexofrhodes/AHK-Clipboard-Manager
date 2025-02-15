; @TODO richedit


Version := "1.0.0"

; ====== INCLUDE ======

#Requires AutoHotkey v2.0
#SingleInstance Force

#Include lib\Author.ahk
#Include lib\AnchorV2.ahk

#Include lib\GuiState.ahk
global iniFile := "guistate.ini"

; ====== GLOBAL VARIABLES ======

global previousClipboard := ""
global actionCheckboxes := [] ; Array to store transformation checkboxes

global historyDir := A_ScriptDir "\log"
if !DirExist(historyDir)
    DirCreate(historyDir)
global clipboardHistory := []
global historyLimit := 100
Loop Files, historyDir . "\*.*" 
    clipboardHistory.Push(A_LoopFileFullPath)
global currentIndex := clipboardHistory.Length

; ====== INITIALIZE ======

myGUI := Gui()
myGUi.Title := "aClipboard"
myGUI.Opt("+AlwaysOnTop +Resize" ) ;  -LastFound )+E0x08000000") ; Prevents GUI from stealing focus

SB := MyGui.AddStatusBar()
UpdateStatusBar()

MyGui.OnEvent("Close", Gui_Close)
MyGui.OnEvent("Size", OnSize)
OnMessage(0x404, AHK_NOTIFYICON)  

CreateControls()
showForm()

; ------------------------

CreateControls(*){

    global

    ; ====== FIND/REPLACE SECTION ======
    btnFind := myGUI.Add("Button", "w80", "Find")
    btnFind.Enabled := false
    editFind := myGUI.Add("Edit", "w300 x+m veditFind", "") 

    btnSave := myGUI.Add("Button", "ys w80", "Save")
    btnSave.OnEvent("Click", SaveRegexValues)

    btnReplace := myGUI.Add("Button", "xm w80", "Replace")
    btnReplace.OnEvent("Click", Replace)

    editReplace := myGUI.Add("Edit", "x+m section w300 veditReplace") 

    btnLoad := myGUI.Add("Button", "ys w80", "Load")
    btnLoad.OnEvent("Click", LoadRegexValues)

    chRegex := myGUI.Add("Checkbox", "xs vchRegex", "Regex")

    chCaseSensitive := myGUI.Add("Checkbox", "x+m vchCaseSensitive" , "CaseSensitive")

    lblLimit := myGUI.Add("Text", "x+m section", "Limit")
    editLimit := myGUI.Add("Edit", "x+m w20 h16 veditLimit", "")

    lblStartingPos := myGUI.Add("Text", "x+m", "Start")
    editStartingPos := myGUI.Add("Edit", "x+m w20 h16 veditStartingPos", "")

    myGUI.Add("Text", "xm h1 w400 0x10") ; Separator

    ; ====== HOTKEYS for functions other than Copy, Append, Undo ======

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
    ; myGUI.MenuBar := MyMenuBar

    addAuthorMenubar(myGUI)

    ; ====== Append Options ======

    myGUI.Add("Text", "xm section w80", "Append Options:")
    MyArray := [  "Linebreak"
                , "Space"
                , "Tab"
                ]
    
    AppendBy := MyGui.Add("DropDownList","vAppendBy", MyArray)                
    myGUI.Add("Text", "xs h1 w80 0x10") ; separator

    ; ====== CHECKBOXES ======

    myGUI.Add("Text", "xm w80", "Select Actions:")

    ; Button to execute checked functions
    btnGetChecked := myGUI.Add("Button", "xs w80", "Run Checked")
    btnGetChecked.OnEvent("Click", applyActionCheckboxes)

    AutoApply := myGUI.Add("Checkbox", "xs vchAutoApply", "AutoApply")
    myGUI.Add("Text", "xs h1 w80 0x10") ; separator

    MyArray := [
                    "Replace"
                    , "Tranclude"
                    , "BreakLinesToSpaces"
                    , "ToUpperCase"
                    , "ToLowerCase"
                    , "ReverseWords" 
                ]

    For value in MyArray
        actionCheckboxes.Push(myGUI.Add("Checkbox", "xs vch" value, value)) ;vCh + Value is used to assign a name, to use with the GuiState functions

    editClipboard := myGUI.Add("Edit", "ys w500 h500", A_Clipboard) ;no vName, no need to save in the ini file

}

showForm() {
    myGUI.Show("NoActivate")
    GuiLoadState(MyGUI.Title, iniFile)
}

UpdateStatusBar(*){
    SB.SetText("There are " . clipboardHistory.Length . " items in Clipboard History.")
}
; ====== EVENTS ======

Gui_Close(*){
    GuiSaveState(myGui.Title, iniFile)                
    ExitApp()
}

OnSize(GuiObj, MinMax, Width, Height){
    Anchor(editClipboard.Hwnd,"hw")
}

AHK_NOTIFYICON(wParam, lParam,*){
    if (lParam = 0x201) ; WM_LBUTTONDOWN
    {
        showForm()
    }
}

; ====== MAIN FUNCTIONS ======

SaveClipboard() {
    global
    
    ; Save current clipboard to file
    fileName := historyDir "\" (currentIndex + 1) ".txt"
    FileAppend(A_Clipboard, fileName)

    ; Add to history array
    clipboardHistory.Push(fileName)

    ; If history exceeds the limit, delete the oldest file
    if (clipboardHistory.Length > historyLimit) {
        oldFile := clipboardHistory[1]
        FileDelete(oldFile)  ; Remove the oldest file
        clipboardHistory.RemoveAt(1)  ; Remove the entry from the array
    }

    ; Update the index
    currentIndex := clipboardHistory.Length
    UpdateStatusBar()
}

$^c:: Copy()
Copy(*) {
    global
    flag := WinActive(mygui.hwnd)
    if flag
        toggleGUI()
    
    previousClipboard := A_Clipboard
    A_Clipboard := ""
    Send("{ctrl Down}c{ctrl up}")
    ClipWait(3)
    SaveClipboard()  ; Save clipboard after copying
    try editClipboard.Value := A_Clipboard

    if AutoApply.Value {
        applyActionCheckboxes()
    }

    if flag
        toggleGUI()
}

$^x:: Cut()
Cut(*) {
    global
    flag := WinActive(mygui.hwnd)
    if flag
        toggleGUI()

    previousClipboard := A_Clipboard
    A_Clipboard := ""
    Send("{ctrl Down}x{ctrl up}")
    ClipWait(3)
    SaveClipboard()  ; Save clipboard after cutting
    try editClipboard.Value := A_Clipboard

    if AutoApply.Value {
        applyActionCheckboxes()
    }

    if flag
        toggleGUI()
}

!c:: Append()
Append(*) { 
    global
    flag := WinActive(mygui.hwnd)
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

    SaveClipboard()  ; Save clipboard after appending
    try editClipboard.Value := A_Clipboard

    if AutoApply.Value 
        applyActionCheckboxes()

    if flag
        toggleGUI()
}

toggleGUI(*){
    if IsSet(MyGui) && MyGui {
        if WinActive(MyGui.Hwnd) {
            MyGui.Hide()
        } else {
            showForm()
        }
    } else {
        
    }
}

; $!Z:: Undo()
; Undo(*) {
;     global
;     tmp := A_Clipboard
;     A_Clipboard := previousClipboard
;     previousClipboard := tmp
;     editClipboard.Value := A_Clipboard
; }

$!z::  LoopHistory(-1)
$!y::  LoopHistory(1)

LoopHistory(direction) {
    global 
    newIndex := currentIndex + direction

    ; Check if the new index is within valid bounds
    if (newIndex >= 1 && newIndex <= clipboardHistory.Length) {
        currentIndex := newIndex
        fileToRead := clipboardHistory[currentIndex]
        A_Clipboard := FileRead(fileToRead)
        try editClipboard.Value := A_Clipboard
    } else {
        MsgBox("No more history in this direction.")
    }
}

; @TODO are these necessary?
Clear(*) {
    previousClipboard := A_Clipboard
    A_Clipboard := ""
    editClipboard.text := ""
}
    
Paste(*){
    global
    flag := WinActive(mygui.hwnd)
    if flag
        toggleGUI()
    previousClipboard := A_Clipboard
    
    A_Clipboard := ""
    Send("{ctrl Down}c{ctrl up}")
    ClipWait    
    EditPaste A_Clipboard, editClipboard
    A_Clipboard := editClipboard.Text
    if flag
        toggleGUI()       
}

Inject(*){
    global
    flag := WinActive(mygui.hwnd)
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
        MyGui.Opt("-AlwaysOnTop")
        filePath := FileSelect(3, , "Open a file", "Text Documents (*.edit; *.doc)")
        MyGui.Opt("+AlwaysOnTop")
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
