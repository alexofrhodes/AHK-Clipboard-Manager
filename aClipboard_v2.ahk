

; @TODO consider using richedit to highlight search terms in the clipboard edit control? 

; ====== INFO ====== 
{
    ; ===================================================================
    ; Script Name:     [aClipboard]
    ; Author:          [https://github.com/alexofrhodes]
    ; ===================================================================
}

Version := "1.0.0"

; ====== Includes ======
{
    #Requires AutoHotkey v2.0
    #SingleInstance Force

    #Include lib\Author.ahk

    #Include lib\Array.ahk
    #Include lib\AnchorV2.ahk

    #Include lib\GuiState.ahk
    #Include lib\Translator.ahk

    #Include  lib\UseGDIP.ahk
    UseGDIP()
    #Include lib\CreateImageButton.ahk
    CreateImageButton("SetDefGuiColor", 0xFFF0F0F0)
    ; if true use GuiButtonIcon_v2, else use CreateImageButton
}
; ====== Global Variables ======
{
    global showOnStart := 1
    global previousClipboard := ""
    global actionCheckboxes := [] ; Array to store transformation checkboxes

    global historyDir := A_ScriptDir "\log"
    if !DirExist(historyDir)
        DirCreate(historyDir)

    global LogFiles := []
    global historyLimit := IniRead("Settings.ini", "General", "historyLimit", 100)          ; after which the oldest file will be deleted
    
    Loop Files, historyDir . "\*.txt" 
        LogFiles.InsertAt(1, StrReplace(A_LoopFileName, ".txt", "")) 

    global currentIndex := 1        ; used for lbLogfiles
    global activeControlIndex := 1  ; used for cycling focus

    global Datatype_Empty      := 0
    global Datatype_Text       := 1
    global Datatype_NonText    := 2
    global ClipboardDatatype   := Datatype_Empty
    OnClipboardChange(ClipChanged, -1)   
    ; A_Clipboard := A_Clipboard

    ; OnClipboardChange Callback , [AddRemove]
    ;If omitted, it defaults to 1. Otherwise, specify one of the following numbers:
    ;  1 = Call the callback after any previously registered callbacks.
    ; -1 = Call the callback before any previously registered callbacks.    <==
    ;  0 = Do not call the callback.

    ; languages := Map()
    ; languages.CaseSense := true
    languages := Map(
        "--Auto-Detect--", "auto",
        "Afrikaans", "af",
        "Albanian", "sq",
        "Amharic", "am",
        "Arabic", "ar",
        "Armenian", "hy",
        "Assamese", "as",
        "Aymara", "ay",
        "Azerbaijani", "az",
        "Bambara", "bm",
        "Basque", "eu",
        "Belarusian", "be",
        "Bengali", "bn",
        "Bhojpuri", "bho",
        "Bosnian", "bs",
        "Bulgarian", "bg",
        "Catalan", "ca",
        "Cebuano", "ceb",
        "Chinese (Simplified)", "zh-CN",
        "Chinese (Traditional)", "zh-TW",
        "Corsican", "co",
        "Croatian", "hr",
        "Czech", "cs",
        "Danish", "da",
        "Dhivehi", "dv",
        "Dogri", "doi",
        "Dutch", "nl",
        "English", "en",
        "Esperanto", "eo",
        "Estonian", "et",
        "Ewe", "ee",
        "Filipino (Tagalog)", "fil",
        "Finnish", "fi",
        "French", "fr",
        "Frisian", "fy",
        "Galician", "gl",
        "Georgian", "ka",
        "German", "de",
        "Greek", "el",
        "Guarani", "gn",
        "Gujarati", "gu",
        "Haitian Creole", "ht",
        "Hausa", "ha",
        "Hawaiian", "haw",
        "Hebrew", "he",
        "Hindi", "hi",
        "Hmong", "hmn",
        "Hungarian", "hu",
        "Icelandic", "is",
        "Igbo", "ig",
        "Ilocano", "ilo",
        "Indonesian", "id",
        "Irish", "ga",
        "Italian", "it",
        "Japanese", "ja",
        "Javanese", "jv",
        "Kannada", "kn",
        "Kazakh", "kk",
        "Khmer", "km",
        "Kinyarwanda", "rw",
        "Konkani", "gom",
        "Korean", "ko",
        "Krio", "kri",
        "Kurdish", "ku",
        "Kurdish (Sorani)", "ckb",
        "Kyrgyz", "ky",
        "Lao", "lo",
        "Latin", "la",
        "Latvian", "lv",
        "Lingala", "ln",
        "Lithuanian", "lt",
        "Luganda", "lg",
        "Luxembourgish", "lb",
        "Macedonian", "mk",
        "Maithili", "mai",
        "Malagasy", "mg",
        "Malay", "ms",
        "Malayalam", "ml",
        "Maltese", "mt",
        "Maori", "mi",
        "Marathi", "mr",
        "Meiteilon (Manipuri)", "mni-Mtei",
        "Mizo", "lus",
        "Mongolian", "mn",
        "Myanmar (Burmese)", "my",
        "Nepali", "ne",
        "Norwegian", "no",
        "Nyanja (Chichewa)", "ny",
        "Odia (Oriya)", "or",
        "Oromo", "om",
        "Pashto", "ps",
        "Persian", "fa",
        "Polish", "pl",
        "Portuguese", "pt",
        "Punjabi", "pa",
        "Quechua", "qu",
        "Romanian", "ro",
        "Samoan", "sm",
        "Sanskrit", "sa",
        "Scots Gaelic", "gd",
        "Sepedi", "nso",
        "Serbian", "sr",
        "Sesotho", "st",
        "Shona", "sn",
        "Sindhi", "sd",
        "Sinhala (Sinhalese)", "si",
        "Slovak", "sk",
        "Slovenian", "sl",
        "Somali", "so",
        "Spanish", "es",
        "Sundanese", "su",
        "Swahili", "sw",
        "Swedish", "sv",
        "Tagalog (Filipino)", "tl",
        "Tajik", "tg",
        "Tamil", "ta",
        "Tatar", "tt",
        "Telugu", "te",
        "Thai", "th",
        "Tigrinya", "ti",
        "Tsonga", "ts",
        "Turkish", "tr",
        "Turkmen", "tk",
        "Twi (Akan)", "ak",
        "Ukrainian", "uk",
        "Urdu", "ur",
        "Uyghur", "ug",
        "Uzbek", "uz",
        "Vietnamese", "vi",
        "Welsh", "cy",
        "Xhosa", "xh",
        "Yiddish", "yi",
        "Yoruba", "yo",
        "Zulu", "zu"
    )
    languageList := []
    for lang in languages
        languageList.Push(lang)
}
; ====== Create GUI ======
{
    ; ---- Main ----
    {
        myGui := Gui()
        MyGui.SetFont("w500 q5", "Segoe UI")
        myGui.Title := "aClipboard"
        myGui.Opt("+AlwaysOnTop +Resize" )

        SB := myGui.AddStatusBar()
        UpdateStatusBar()

        myGui.OnEvent("Close", Gui_Close)
        myGui.OnEvent("Size", OnSize)
        OnMessage(0x404, AHK_NOTIFYICON)  

        CreateControls()
        global guiManager := GuiState(myGui, "Settings.ini", "GUI") ; must be after the control creation
        if showOnStart{
            showForm()
            editFilter.focus()
        }
    }
  
    CreateControls(*){
        global
        ; Fold
        {
            ; ---- MenuBar ----
            {
                MyArray := [  "Copy"
                            , "Cut"
                            , "Paste"
                            , "myFormat"
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
                    targetFile := A_ScriptDir . "\lib\menubaricons\" . value . ".ico"
                    if (FileExist(targetFile))
                        MyMenuBar.SetIcon(value, targetFile) ; , IconNumber, IconWidth)
                }
                myGui.MenuBar := MyMenuBar
                addAuthorMenubar(myGui)
            }
            ; ---- Find/Replace ----
            {
                btnFind := myGui.Add("Button", "w40", "Find")
                btnFind.Enabled := false
                CreateImageButton(btnFind, 0, IBStyles["info-round"]*)

                btnExtractByRegex := myGui.Add("Button", "x+m w100", "ExtractByRegex")
                btnExtractByRegex.OnEvent("Click", (*) => ExecuteFunction(ExtractByRegex, editFind.text))
                CreateImageButton(btnExtractByRegex, 0, IBStyles["success-round"]*)

                editFind := myGui.Add("Edit", "w350 x+m veditFind", "") 
        
                btnSave := myGui.Add("Button", "ys w50", "Save")
                btnSave.OnEvent("Click", SaveRegexValues)
                CreateImageButton(btnSave, 0, IBStyles["info-round"]*)

                btnReplace := myGui.Add("Button", "xm w150 ", "Replace")
                btnReplace.OnEvent("Click", (*) => ExecuteFunction(Replace, A_Clipboard))
                CreateImageButton(btnReplace, 0, IBStyles["success-round"]*)

                editReplace := myGui.Add("Edit", "x+m w350 section veditReplace") 
        
                btnLoad := myGui.Add("Button", "x+m w50", "Load")
                btnLoad.OnEvent("Click", LoadRegexValues)
                CreateImageButton(btnLoad, 0, IBStyles["info-round"]*)
                      


                btnChangeHistoryLimit := myGui.Add("Button", "w150 xm", "History Limit")
                btnChangeHistoryLimit.OnEvent("Click", ChangeHistoryLimit)
                CreateImageButton(btnChangeHistoryLimit, 0, IBStyles["warning-round"]*)

                ; myGui.Add("Text", "xm h1 w150 0x10") ; Separator

                chRegex := myGui.Add("Checkbox", "x+m vchRegex", "By Regex")
                
                chCaseSensitive := myGui.Add("Checkbox", "x+m vchCaseSensitive" , "CaseSensitive")
        
                lblLimit := myGui.Add("Text", "x+m section", "Limit")
                editLimit := myGui.Add("Edit", "x+m w20 ys-1 h16 veditLimit", "")
        
                lblStartingPos := myGui.Add("Text", "ys", "Start")
                editStartingPos := myGui.Add("Edit", "x+m ys-1 w20 h16 veditStartingPos", "")
      
                
                Separator1 := myGui.Add("Text", "xm h1 w570 0x10") ; Separator
            }

            ; ---- Append/Prepend Options ----
            {
                myGui.Add("Text", "section w60", "Separator:")
                MyArray := [  "Linebreak"
                            , "Space"
                            , "Tab"
                            ]
                
                ddlSeparator := myGui.Add("DropDownList","x+m w80 vddlSeparator", MyArray)                
            
            }
            ; ---- Paste Options ----
            {
                MyGui.Add("Text", "xs w70", "Paste Method:")
                pasteMethods := ["Ctrl+V", "SendText", "SendInput"]
                ddlPasteMethod := MyGui.Add("DropDownList", "x+m w70 vddlPasteMethod", pasteMethods)
                ddlPasteMethod.Choose(1)                ; Default to Ctrl+V
                
                MyGui.Add("Text", "xs w65", "After Paste:")
                ddlAfterPasteAction := myGui.Add("DropDownList", "x+m w75 vddlAfterPasteAction", ["None", "Previous", "Next"])
                ddlAfterPasteAction.choose(1)

                myGui.Add("Text", "xm h1 w150 0x10") ; Separator
            }  
            ; ---- Translation ----
            {
                myGui.Add("Text", "xm w40", "Source")
                ddlSourceLanguage := myGui.Add("DropDownList", "x+m w100 vSourceLang Choose1 vddlSourceLanguage", languageList)
                myGui.Add("Text", "xm w40", "Target")
                ddlTargetLanguage := myGui.Add("DropDownList", "x+m w100 vTargetLang Choose2 vddlTargetLanguage", languageList)

                btnSwap := myGui.Add("Button", "xm w70", "Swap")
                btnSwap.OnEvent("Click", SwapLanguages)
                CreateImageButton(btnSwap, 0, IBStyles["info-round"]*)  

                btnTranslate := myGui.Add("Button","x+m w70", "Translate")
                btnTranslate.OnEvent("Click", (*) => ExecuteFunction("Translate", A_Clipboard))
                CreateImageButton(btnTranslate, 0, IBStyles["success-round"]*)

                myGui.Add("Text", "xm h1 w150 0x10") ; Separator
            }
    
        }
        ; UNFold 
        {
            ; ---- Actions ---- 
            { 
                btnToggle := myGui.AddButton("xs w70 ", "Toggle")
                btnToggle.OnEvent("Click", ToggleActions)
                CreateImageButton(btnToggle, 0, IBStyles["info-round"]*)

                btnGetChecked := myGui.Add("Button", "x+m w70", "Apply")
                btnGetChecked.OnEvent("Click", applyActionCheckboxes)
                CreateImageButton(btnGetChecked, 0, IBStyles["success-round"]*)

                chCurrentTabOnly := myGui.Add("Checkbox", "xm+80 vchCurrentTabOnly", "Only This Tab")
                chAutoApply := myGui.Add("Checkbox", "xm+80 vchAutoApply", "Auto")
            }   
            
            ;@bm arrays - @MODIFY (1) 
            ; !!! MAP your predefined tabs or add to the array of remaining checkboxes 
            ;     the names of your clipboard modifying functions 
            ;     See also @MODIFY(2)

            ; Predefined tabs with specific checkboxes stored in a map.
            ; predefinedTabs := Map(tabname1, arrayoffunctionnames, tabname2, arrayoffunctionnames)

            predefinedTabs := Map(
                "1. Main", ["Replace", "Transclude", "Translate"],
                "2. Common", ["ToUpperCase", "ToLowerCase","ToTitleCase", "ReverseWords", "TrimSpaces", 
                                "SpaceToUnderscore", "BreakLinesToSpaces", "RemoveHTMLTags","RemoveSpecialChars", "RemoveDigits",
                                "RemoveExcessLines", "SortAZ", "SortZA","SortByLengthAZ","SortByLengthZA",
                                "AddTimestamp", "NumberedList","BulletList"],
                "3. Regex",["ExtractWebsites", "ExtractPhoneNumbers", "ExtractEmails"],
            )
    
            ; Remaining checkboxes (to be distributed into dynamic tabs)
            ; remainingCheckboxes := [] ; array of function names
            remainingCheckboxes := [
                "Test1",  "Test2",  "Test3",  "Test4",  "Test5",  
                "Test6",  "Test7",  "Test8",  "Test9",  "Test10",  
                "Test11", "Test12", "Test13", "Test14", "Test15",  
                "Test16", "Test17", "Test18", "Test19", "Test20",  
                "Test21", "Test22", "Test23", "Test24", "Test25",  
                "Test26", "Test27", "Test28", "Test29", "Test30",  
                "Test31", "Test32", "Test33", "Test34", "Test35",  
                "Test36", "Test37", "Test38", "Test39", "Test40",  
                "Test41", "Test42", "Test43", "Test44", "Test45",  
                "Test46", "Test47", "Test48", "Test49", "Test50"

            ]
            {
                ItemsPerTab := 15  ; Max number of checkboxes per tab
        
                ; Create an array for tab names
                tabNames := []
                For Key , Value in predefinedTabs
                    tabNames.Push(Key)
        
                ; Calculate extra tabs needed for remaining checkboxes
                extraItems := remainingCheckboxes.Length
                extraTabs := Ceil(extraItems / ItemsPerTab)
                Index := tabNames.Length
                ; Generate additional tab names
                Loop extraTabs
                    tabNames.Push("Tab" A_Index + Index)  ; Naming dynamic tabs (Tab4, Tab5, etc.)
        
                ; Create the Tab control
                tabControl := myGui.Add("Tab3", "vTabControl xs w155 0x80 0x100 R" ItemsPerTab, tabNames)
        
                ; Store checkbox objects for later use
                actionCheckboxes := []
        
                ; Add checkboxes to predefined tabs
                For tabName in predefinedTabs {
                    tabControl.UseTab(tabName)  ; Switch to predefined tab
                    For checkboxLabel in predefinedTabs[tabName] {
                        actionCheckboxes.Push(myGui.AddCheckbox("vch" checkboxLabel, checkboxLabel)) 
                    }
                }
        
                ; Assign remaining checkboxes to dynamic tabs
                tabIndex := predefinedTabs.Count + 1  ; Start from the first dynamic tab
                itemCount := 0
        
                For checkboxLabel in remainingCheckboxes {
                    if (itemCount = 0)  ; Move to the next tab when a new tab starts
                        tabControl.UseTab(tabIndex++)
        
                    actionCheckboxes.Push(myGui.AddCheckbox("vch" checkboxLabel, checkboxLabel)) 
                    
                    itemCount++
                    if (itemCount >= ItemsPerTab)  ; Reset item count for new tab
                        itemCount := 0
                }
        
                tabControl.UseTab()  ; Reset tab selection
            }
        }
        ; Part 3 _LogFiles and Clipboard Preview
        {
            ; ---- Log Files ----
            {
            btnDeleteSelected:= myGui.Add("Button", "ys x+m section w70", "Del Sel")
            btnDeleteSelected.OnEvent("Click", DeleteSelectedLogFile)
            CreateImageButton(btnDeleteSelected, 0, IBStyles["critical-round"]*)   

            btnDeleteAll := myGui.Add("Button", "x+m w70", "Del All")
            btnDeleteAll.OnEvent("Click", DeleteAllLogFiles)
            CreateImageButton(btnDeleteAll, 0, IBStyles["critical-round"]*)   
        
            btnDuplicate := myGui.Add("Button", "xs w70", "Duplicate")
            btnDuplicate.OnEvent("Click", (*)=> ExecuteFunction("DuplicateLogFile",A_Clipboard))
            CreateImageButton(btnDuplicate, 0, IBStyles["info-round"]*)

            btnRename := myGui.Add("Button", "x+m w70", "Rename")
            btnRename.OnEvent("Click", RenameLogFile)
            CreateImageButton(btnRename, 0, IBStyles["success-round"]*)

            myGui.Add("Text", "xs  w30", "Filter")
            editFilter := myGui.Add("Edit", "x+m w110", "")
            editFilter.OnEvent("Change", FilterLogFiles)
    
            lbLogFiles := myGui.Add("ListBox", "xs h500 w150 HScroll", LogFiles) 
            lbLogFiles.OnEvent("Change", LoadSelectedLogFile)
            }
            ; ---- Clipboard Edit ----
            {
                myGui.Add("Text", "ys+5 section", "Clipboard")
                btnSaveEditClipboard := myGui.Add("Button", "ys-5 w100", "Save Changes")
                btnSaveEditClipboard.OnEvent("Click", SaveEditClipboardChanges)     
                CreateImageButton(btnSaveEditClipboard, 0, IBStyles["success-round"]*)

                editClipboard := myGui.Add("Edit", "xs w240 h520 +HScroll -wrap 0x100") ;no vName, no need to save in the ini file
    
                if logfiles.Length > 0{
                    lbLogFiles.Choose(1)  
                    LoadSelectedLogFile()
                }

                if (StrReplace(Trim(A_Clipboard), "`r`n", "`n") != StrReplace(Trim(editClipboard.value), "`r`n", "`n")) {
                    if (currentIndex < LogFiles.Length) {  
                        lastSavedText := FileRead(historyDir "\" logFiles[currentIndex + 1] ".txt")
                        if (StrReplace(Trim(A_Clipboard), "`r`n", "`n") != StrReplace(Trim(lastSavedText), "`r`n", "`n")) {
                            editClipboard.value := A_Clipboard  
                            SaveClipboardToLogFile()  
                        }
                    } else {  
                        editClipboard.value := A_Clipboard  
                        SaveClipboardToLogFile()
                    }
                }
            }
        }
        
    }    
    ; ---- Fold ----
    {    
        ; ---- Hotkeys ----
        {
            #c::        toggleGUI()                     ; Win  + C
            $^c::       Copy()                          ; Ctrl + C
            $^x::       Cut()                           ; Ctrl + X
            $^!c::      Append()                        ; Ctrl + Alt   + C 
            $^!x::      CutAppend()                     ; Ctrl + Alt   + X 
            $^+c::      Prepend()                       ; Ctrl + Shift + P 
            $^+x::      CutPrepend()                    ; Ctrl + Shift + X 
            $!i::       Inject()                        ; Alt  + I 
            $#z::       Undo()                          ; Win  + Z 
            $#y::       Redo()                          ; Win  + Y 
            $#v::       Paste()                         ; Win  + V
            $#F::       myFormat()                      ; Win  + F
            $^s::       SaveChanges()                   ; Ctrl + S
            $#T::       ExecuteFunction("Translate", A_Clipboard)    ; Win  + T

            $del::     {
                if !WinActive(mygui.hwnd){
                    Send "{Delete}"
                    return
                }
                GuiCtrlObj := MyGui.FocusedCtrl    
                if (guictrlobj != lbLogFiles) {
                    Send "{Delete}"
                    return
                }else{
                    DeleteSelectedLogFile()   
                }          
            } 

            $Tab::CycleFocus()
            $+Tab::CycleFocus()

            CycleFocus(*){
                global
                if !WinActive(mygui.hwnd){
                    if GetKeyState('Shift') {
                        Send "+{Tab}"
                    } else {
                        Send "{Tab}"
                    }
                    return
                }

                zControls := [editFilter, lbLogFiles, editClipboard]
                if GetKeyState('Shift') {
                    ; Cycle backwards if Shift is held
                    activeControlIndex -= 1
                    if activeControlIndex < 1
                        activeControlIndex := zControls.length
                } else {
                    ; Cycle forwards if Shift is not held
                    activeControlIndex += 1
                    if activeControlIndex > zControls.length
                        activeControlIndex := 1
                }
                
                zControls[activeControlIndex].Focus()
            }

            ; ----Encapsulation ----
            {
            ;     $(:: Encapsulate()
            ;     $[:: Encapsulate()
            ;     ${:: Encapsulate()
            ;     $":: Encapsulate()
            ;     $<:: Encapsulate()
            ;     ; $':: Encapsulate 
                
            ;     Encapsulate(*) {
            ;         if winactive(mygui.hwnd){
            ;             send substr(A_ThisHotkey,2)
            ;             return
            ;         }
            ;         Copy()
                    ; if (A_Clipboard = "") 
                    ;     return
            ;         switch substr(A_ThisHotkey,2) {
            ;         Case "[": A_Clipboard := "[" A_Clipboard "]"
            ;         Case "(": A_Clipboard := "(" A_Clipboard ")"  
            ;         Case "{": A_Clipboard := "{" A_Clipboard "}"
            ;         Case '"': A_Clipboard := '"' A_Clipboard '"'
            ;         Case "<": A_Clipboard := "<" A_Clipboard ">"
            ;         ; Case "'": A_Clipboard := "'" A_Clipboard "'"
            ;         }
            ;         paste()
            ;     }
            }        
        }
        ; ---- Gui Events ----
        {
            
            ClipChanged(DataType) {
                global
                ClipboardDatatype := DataType
            }

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
        ; ---- Gui Functions ----
        {
            SwapLanguages(*) {
                global 
                temp := ddlSourceLanguage.Text
                ddlSourceLanguage.Text := ddlTargetLanguage.Text
                ddlTargetLanguage.Text := temp
            }
            showForm(*) {
                
                myGui.Show("x5000 y5000") ;show offscreen to avoid flicker from guiLoadState
                guiManager.LoadState()

            }
            toggleGUI(*){
                guiManager.SaveState()
                if IsSet(myGui) && myGui {
                    if WinActive(myGui.Hwnd) {
                        myGui.Hide()
                    } else {
                        showForm()
                    }
                } else {
                    
                }
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
                        <tr><td>Win + C</td><td>Toggle GUI</td></tr>
                        <tr><td>Ctrl + C</td><td>Copy</td></tr>
                        <tr><td>Ctrl + X</td><td>Cut</td></tr>
                        <tr><td>Win + V</td><td>Paste by chosen PasteMethod</td></tr>
                        <tr><td>Win + F</td><td>Format Selected Text</td></tr>
                        <tr><td>Win + T</td><td>Translate Selected Text</td></tr>
                        <tr><td>Ctrl + Alt + C</td><td>Append</td></tr>
                        <tr><td>Ctrl + Alt + X</td><td>Cut Append</td></tr>
                        <tr><td>Ctrl + Shift + P</td><td>Prepend</td></tr>
                        <tr><td>Ctrl + Shift + X</td><td>Cut Prepend</td></tr>
                        <tr><td>Alt + I</td><td>Inject</td></tr>
                        <tr><td>Win + Z</td><td>Undo</td></tr>
                        <tr><td>Win + Y</td><td>Redo</td></tr>
                        <tr><td>Ctrl + S</td><td>If gui active, Save Edit Changes</td></tr>
                    </table>
                    </br></br>
                    <p>For advanced clipboard management, consider using <a href='https://hluk.github.io/CopyQ/' target='_blank'>CopyQ</a>.</p>
                    <p>For advanced translator with ocr, consider using <a href='https://apps.kde.org/en-gb/crowtranslate/' target='blank'>Crow Translate</a>.</p>
                </body>
                </html>
                )"
            
                ; Create GUI
                helpGui := Gui(, "Hotkey Guide")
                browser := helpGui.AddActiveX("x10 y10 w500 h600", "Shell.Explorer")
            
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
            DuplicateLogFile(*){
                global
                if (currentIndex = 0) {
                    MsgBox "Please select a file to duplicate."
                    return
                }
            
                oldFileName := logFiles[currentIndex]
                oldFilePath := historyDir "\" oldFileName ".txt"
            
                ; Extract timestamp part (assuming it's always at the start)
                RegExMatch(oldFileName, "^\d{4}-\d{2}-\d{2}-\d{6}", &match)
                timestamp := match[0]
            
                if (!timestamp) {
                    MsgBox "Invalid file format. Cannot duplicate."
                    return
                }
            
                newName := timestamp " Copy"
                newFilePath := historyDir . "\" . newName . ".txt"
            
                ; Ensure the new name does not exist
                counter := 1
                while FileExist(newFilePath) {
                    counter++
                    newName := newName " (" . counter . ")"
                    newFilePath := historyDir "\" . newName . ".txt"
                } 
                FileCopy(oldFilePath, newFilePath)
                LogFiles.InsertAt(currentIndex, newName)
                lbLogFiles.Delete
                lbLogFiles.add(logFiles)
                lbLogFiles.choose(currentIndex)
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
                Loop Files, historyDir . "\*.txt"  
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
                if (currentIndex != 0) {
                    filePath := historyDir "\" selectedFile
                    editClipboard.Value := FileRead(filePath)
                    A_Clipboard := editClipboard.Value
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
                        FileRecycle(filePath)
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
                        FileRecycle(A_LoopFilePath)
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
                guiManager.SaveState()
            }
            SaveEditClipboardChanges(*){
                if lbLogfiles.value =0
                    return
                A_Clipboard := editClipboard.value
                targetFile := historyDir "\" . lbLogFiles.text . ".txt"
                FileDelete(targetFile)
                FileAppend(A_Clipboard, targetFile)
            }
            
            SaveChanges(*){
                if WinActive(myGui) {
                    SaveEditClipboardChanges()
                    guiManager.SaveState()
                }else{
                    Send("{ctrl Down}s{ctrl up}")
                }
            }
            ToggleActions(*) {
                global 
            
                activeTab := tabControl.Value  
                checkState := 0  ; Assume all in the current tab are checked
                changeNeeded := false  ; Track if other tabs would be affected
                
                ; Check if any checkbox in the current tab is unchecked
                for ctrl in actionCheckboxes {
                    vis := ControlGetVisible(ctrl.Hwnd)
                    if (vis && ctrl.Value = 0) {  
                        checkState := 1  ; If any is unchecked, we set checkState to check all
                        break
                    }
                }

                ; Apply checkState only to checkboxes in the current tab
                for ctrl in actionCheckboxes {
                    vis := ControlGetVisible(ctrl.Hwnd)
                    if (vis)  
                        ctrl.Value := checkState
                }
            
                ; Check if applying to all tabs would change anything
                for ctrl in actionCheckboxes {
                    vis := ControlGetVisible(ctrl.Hwnd)
                    if (!vis && ctrl.Value != checkState) {
                        changeNeeded := true
                        break
                    }
                }
            
                ; Ask only if other tabs have checkboxes that would change
                if (changeNeeded && MsgBox("Apply this change to all tabs?", "Confirmation", "Y/N") = "Yes") {
                    for ctrl in actionCheckboxes {
                        ctrl.Value := checkState
                    }
                }
            }
            SaveRegexValues(*) {
                myGui.Opt("-AlwaysOnTop")
                fileName := InputBox("Enter file name", "Enter the desired file name or leave it blank to select a file:", "").value
                myGui.Opt("+AlwaysOnTop")
            
                if (fileName == "")
                    return
            
                Folder := A_ScriptDir "\lib\RegexPatterns"
                if !DirExist(Folder)
                    DirCreate(Folder)
            
                filePath := Folder "\" fileName ".ini"
            
                ; If the file exists, prompt the user for action
                while FileExist(filePath) {
                    choice := MsgBox("File already exists. Do you want to overwrite it?", "File Exists", "YesNoCancel")
                    if (choice = "Yes") {
                        break  ; Proceed with overwrite
                    } else if (choice = "No") {
                        fileName := InputBox("Enter new file name", "Choose a different name:", "").value
                        if (fileName == "")
                            return  ; If no name is provided, cancel
                        filePath := Folder "\" fileName ".ini"
                    } else {
                        return  ; Cancel action
                    }
                }
            
                ; Save values to the INI file
                IniWrite(editFind.Value, filePath, "Regex", "Find")
                IniWrite(editReplace.Value, filePath, "Regex", "Replace")
            
                MsgBox("Saved regex values to " filePath)
            }
            
            
            LoadRegexValues(*) {
                myGui.Opt("-AlwaysOnTop")
                filePath := FileSelect(3, A_ScriptDir "\lib\RegexPatterns\", "Open a file", "INI Files (*.ini)")
                myGui.Opt("+AlwaysOnTop")
            
                if filePath = ""
                    return
            
                editFind.Value := IniRead(filePath, "Regex", "Find", "")
                editReplace.Value := IniRead(filePath, "Regex", "Replace", "")
            
            }       
            applyActionCheckboxes(*) {
                global 
                guiManager.SaveState()
                currentTabOnly := chCurrentTabOnly.Value  ; Get checkbox state
                textToModify := A_Clipboard  ; Store initial clipboard content
            
                for ctrl in actionCheckboxes {
                    vis := ControlGetVisible(ctrl.Hwnd)
                    if (ctrl.Value && (!currentTabOnly || vis)) {  
                        FunctionName := ctrl.Text
                        textToModify := ExecuteFunction(FunctionName, textToModify)
                    }
                }
                
                ; Finally, update clipboard once at the end if necessary
                A_Clipboard := textToModify
                ; Optionally, save or log the final clipboard if needed
                SaveClipboardToLogFile()
            }
            ExecuteFunction(FunctionOrName, textToModify) {
                previousClipboard := textToModify  ; Store original clipboard content
                funcObj := ""
            
                ; Determine if the input is a function object or a function name
                if IsObject(FunctionOrName) && FunctionOrName is Func {
                    funcObj := FunctionOrName  ; Function reference
                } else if IsFunc(FunctionOrName) {
                    funcObj := %FunctionOrName%  ; Retrieve function by name
                }
            
                if funcObj {
                    if funcObj.MinParams = 0 {
                        funcObj.Call()  ; Call function without arguments
                    } else {
                        textToModify := funcObj.Call(textToModify)  ; Call function with textToModify as argument
                    }
                } else {
                    Switch FunctionOrName, false
                    {
                        Case "test": MsgBox("test")
                        Default: MsgBox("Function '" FunctionOrName "' not found!")
                    }
                }
                
                ; Return the modified clipboard after function execution
                return textToModify
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
            
            CutPrepend(*) {
                global
                flag := WinActive(myGui.hwnd)
                if flag
                    toggleGUI()
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                A_Clipboard := ""
                Send("{ctrl Down}x{ctrl up}")
                
                if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
                    ; Restore clipboard if cut and prepend operation fails
                    A_Clipboard := previousClipboard
                    MsgBox "The attempt to cut and prepend text onto the clipboard failed."
                } else {
                    A_Clipboard := A_Clipboard . selectedSeparator() . previousClipboard
                    SaveClipboardToLogFile()  ; Save the final clipboard state
                    if chAutoApply.Value 
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
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                A_Clipboard := ""
                Send("{ctrl Down}c{ctrl up}")
                
                if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
                    ; Restore clipboard if prepend operation fails
                    A_Clipboard := previousClipboard
                    MsgBox "The attempt to copy text onto the clipboard failed."
                } else {
                    A_Clipboard := A_Clipboard . selectedSeparator() . previousClipboard
                    SaveClipboardToLogFile()  ; Save the final clipboard state
                    if chAutoApply.Value 
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
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                A_Clipboard := ""
                Send("{ctrl Down}x{ctrl up}")
                
                if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
                    ; Restore clipboard if the cut and append operation fails
                    A_Clipboard := previousClipboard
                    MsgBox "The attempt to cut and copy text onto the clipboard failed."
                } else {
                    A_Clipboard := previousClipboard . selectedSeparator() . A_Clipboard
                    SaveClipboardToLogFile()  ; Save the final clipboard state
                    if chAutoApply.Value 
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
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                A_Clipboard := ""
                Send("{ctrl Down}c{ctrl up}")
                
                if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
                    ; Restore clipboard if the append operation fails
                    A_Clipboard := previousClipboard
                    MsgBox "The attempt to copy text onto the clipboard failed."
                } else {
                    A_Clipboard := previousClipboard . selectedSeparator() . A_Clipboard
                    SaveClipboardToLogFile()  ; Save the final clipboard state
                    if chAutoApply.Value 
                        applyActionCheckboxes()
                }
                
                if flag
                    toggleGUI()
            }
            
            Cut(*) {
                global
                flag := WinActive(myGui.hwnd)
                if flag
                    toggleGUI()
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                A_Clipboard := ""
                Send("{ctrl Down}x{ctrl up}")
                
                if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
                    ; Restore clipboard if cut fails
                    A_Clipboard := previousClipboard
                    MsgBox "The attempt to cut text from the clipboard failed."
                } else {
                    SaveClipboardToLogFile()  ; Save only after cutting
                    if chAutoApply.Value 
                        applyActionCheckboxes()
                }
                
                if flag
                    toggleGUI()
            }
            
            Copy(*) {
                global
                flag := WinActive(myGui.hwnd)
                if flag
                    toggleGUI()
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                A_Clipboard := ""
                Send("{ctrl Down}c{ctrl up}")
                
                if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) { 
                    ; Restore clipboard if the copy fails
                    A_Clipboard := previousClipboard
                    MsgBox "The attempt to copy text onto the clipboard failed."
                } else {
                    SaveClipboardToLogFile()  ; Save only after copying
                    if chAutoApply.Value 
                        applyActionCheckboxes()
                }
                
                if flag
                    toggleGUI()
            }
            myFormat(*){
                global
                ddlAfterPasteAction.choose(1) ; "None" = don't cycle
                tmp := chAutoApply.Value 
                chAutoApply.Value := false ;otherwise Copy() may apply format
                Copy()
                if (ClipboardDatatype != Datatype_Text)
                    return
                applyActionCheckboxes()
                Paste()
                chAutoApply.Value := tmp
            }
            Input(*) {
                myGui.opt("-AlwaysOnTop")
                userInput := InputBox("Enter text to copy:", "Clipboard Input")
                myGui.opt("+AlwaysOnTop")
            
                if userInput.Result = "OK" && userInput.Value != "" {
                    previousClipboard := A_Clipboard  ; Store the initial clipboard content
                    A_Clipboard := userInput.Value
                    SaveClipboardToLogFile()  ; Save only after updating the clipboard
                }
            }
            
            Replace(*) {
                ; Validate input: Ensure the "Find" field is not empty
                if (editFind.text = "") {
                    MsgBox("Error: 'Find' field cannot be empty!", "Input Error", 0x30 0x1000)
                    return
                }
                ; Assign defaults if fields are empty
                if (editStartingPos.text = "")
                    editStartingPos.text := "1"
                if (editLimit.text = "")
                    editLimit.text := "-1"
                
                limit := editLimit.Value + 0
                startingPos := editStartingPos.Value + 0
                
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                
                ; If Regex mode is enabled, use RegExReplace
                if chRegex.value {
                    A_Clipboard := RegExReplace(A_Clipboard, editFind.text, editReplace.text, , limit, startingPos)
                } else {
                    A_Clipboard := StrReplace(A_Clipboard, editFind.text, editReplace.text, chCaseSensitive.Value, , limit)
                }
                
                if (previousClipboard = A_Clipboard) {
                    MsgBox("No matches found for '" editFind.text "'!", "No Change", 0x40)
                }
                SaveClipboardToLogFile()  ; Save the final clipboard state
            }
            
            ;@bm
            Paste(*) {
                global
                flag := WinActive(myGui.hwnd)
                if flag
                    toggleGUI()
            
                originalClipboard := A_Clipboard  ; Store the current clipboard content
                selectedText := EditGetSelectedText(editClipboard.hwnd)
            
                if selectedText {
                    A_Clipboard := selectedText  ; Temporarily set clipboard to selected text
                    pasteContent := selectedText
                } else {
                    pasteContent := A_Clipboard
                }
            
                pasteMethod := ddlPasteMethod.Text
            
                ; Use the selected method for pasting
                if (pasteMethod = "Ctrl+V") {
                    Send("{ctrl Down}v{ctrl up}")
                } else if (pasteMethod = "SendText") {
                    SendText(pasteContent)  ; Send raw text
                } else if (pasteMethod = "SendInput") {
                    SendInput(pasteContent)  ; Simulate typing
                }
            
                selectedMode := ddlAfterPasteAction.Text
                if (selectedMode = "Previous") {
                    LoopHistory(1)
                } else if (selectedMode = "Next") {
                    LoopHistory(-1)
                }
            
                ; Restore the original clipboard if it was modified and no cycling is selected
                if (A_Clipboard != pasteContent and selectedMode = "None") {
                    A_Clipboard := originalClipboard
                }
            
                if flag
                    toggleGUI()
            }
            
            Inject(*) {
                global
                flag := WinActive(myGui.hwnd)
                if flag
                    toggleGUI()
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                A_Clipboard := ""
                Send("{ctrl Down}c{ctrl up}")
            
                if !ClipWait(3) || (ClipboardDatatype != Datatype_Text) {
                    A_Clipboard := previousClipboard  ; Restore clipboard if copy fails
                } else {
                    EditPaste A_Clipboard, editClipboard
                    A_Clipboard := editClipboard.value
                    SaveClipboardToLogFile()
                }
            
                if flag
                    toggleGUI()
            }
            
            SaveModifiedCopy(*) {
                if (currentIndex = 0) {
                    MsgBox "Please select a file to duplicate."
                    return
                }
            
                previousClipboard := A_Clipboard  ; Store initial clipboard content
                DuplicateLogFile()
                editClipboard.Value := A_Clipboard
                SaveEditClipboardChanges()
            
                A_Clipboard := previousClipboard  ; Restore original clipboard if modified
            }
            
            Undo(*) {
                LoopHistory(1)
            }
            Redo(*) {
                LoopHistory(-1)
            }
            LoopHistory(direction) {
                global
                newIndex := currentIndex + direction
            
                ; Check if the new index is within valid bounds
                if (newIndex >= 1 && newIndex <= LogFiles.Length) {
                    currentIndex := newIndex  ; Update the global currentIndex
                    fileToRead := historyDir . "\" . LogFiles[currentIndex] . ".txt"
                    previousClipboard := A_Clipboard  ; Store current clipboard
                    A_Clipboard := FileRead(fileToRead)
                    editClipboard.Value := A_Clipboard
                    lbLogFiles.Choose(currentIndex)
            
                    ; Restore clipboard if modified
                    A_Clipboard := previousClipboard
                } else {
                    MsgBox("No more history in this direction.",, 0x1000)
                }
            }
            
            selectedSeparator(*) {
                separator := ddlSeparator.text
                switch separator {
                    case "Linebreak": separator := "`r`n"
                    case "Space": separator := " "
                    case "Tab": separator := "`t"
                    default: separator := ""
                }
                return separator
            }        
            Translate(*) {
                target := Languages[ddlTargetLanguage.text]
                source := Languages[ddlSourceLanguage.text]
            
                previousClipboard := A_Clipboard  ; Store original clipboard content
                translation := Translator.Translate(A_Clipboard, target, source)
                A_Clipboard := translation
                SaveClipboardToLogFile()
            
                ; Restore the original clipboard if desired
                A_Clipboard := previousClipboard
            }
            

            ChangeHistoryLimit(*) {
                global 
                myGui.Opt("-AlwaysOnTop")
                newLimit := InputBox("Enter a new history limit (minimum: 1):", "Set History Limit")
                myGui.Opt("+AlwaysOnTop")

                if (newLimit = "") || (newLimit = false)
                    return
                if !(IsInteger(newLimit)) || (newLimit < 1) {
                    return
                }
                historyLimit := newLimit
                IniWrite(historyLimit, "Settings.ini", "General", "historyLimit")
            }
            
            
        }
    }    
}

; ==== @MODIFY(2) - USER clipboard mod functions ====
    ; Bellow this point add your own functions, they just need to modify the clipboard as a final result
    ; Remember to add the function name to the actionCheckboxes array at @MODIFY(1)
{
    ;@BM2
    Transclude(targetText) {
        FileContent := ""        
        FileRegex := "m)^C:\\.*\.(txt|ahk|bas|md|ahk|py|csv|log|ini|config)"
        pos := 1
        counter := 0
        while (RegExMatch(targetText, FileRegex, &Match, pos)) {
            FileName := Match[0]
            if (FileExist(FileName)) {
                FileContent := FileRead(FileName)
                targetText := StrReplace(targetText, FileName, FileContent)
            }
            counter += 1
            pos := Match.Pos + Match.Len
        }
        return targetText
    }
    
    BreakLinesToSpaces(targetText) {
        return StrReplace(targetText, "`r`n", " ")
    }
    
    ReverseWords(targetText) {
        words := StrSplit(targetText, " ")
        return words.reverse().join(" ")
    }
    
    ToUpperCase(targetText) {
        return StrUpper(targetText)
    }
    
    ToLowerCase(targetText) {
        return StrLower(targetText)
    }
    
    ToTitleCase(targetText) {
        return RegExReplace(targetText, "(\w)(\w*)", "$U1$2")
    }
    
    TrimSpaces(targetText) {
        return RegExReplace(targetText, "\s+", " ")
    }
    
    SpaceToUnderscore(targetText) {
        return StrReplace(targetText, " ", "_")
    }
    
    RemoveDigits(targetText) {
        return RegExReplace(targetText, "\d", "")
    }
    
    RemoveSpecialChars(targetText) {
        return RegExReplace(targetText, "[^\w\s]", "")
    }
    
    RemoveHTMLTags(targetText) {
        return RegExReplace(targetText, "<.*?>", "")
    }
    
    RemoveExcessLines(targetText) {
        return RegExReplace(targetText, "(\n\s*)+", "`n")
    }
    
    SortAZ(targetText) {
        lines := StrSplit(targetText, "`r`n") 
        lines.Sort("C0") 
        return lines.join("`r`n")
    }
    
    SortZA(targetText) {
        lines := StrSplit(targetText, "`r`n")
        arr := lines.Clone().sort("C0").Reverse()
        lines := arr
        return lines.join("`r`n")
    }
    
    SortByLengthAZ(targetText) {
        lines := StrSplit(targetText, "`r`n")
        lines.sort((a, b) => StrLen(a) - StrLen(b)) 
        return lines.join("`r`n")
    }
    
    SortByLengthZA(targetText) {
        lines := StrSplit(targetText, "`r`n")
        lines.sort((a, b) => StrLen(a) - StrLen(b)) 
        lines.Reverse()
        return lines.join("`r`n")
    }
    
    ExtractByRegex(targetText, pattern) {
        items := []
        pos := 1
        while (RegExMatch(targetText, pattern, &Match, pos)) {
            items.Push(match[0])
            pos := Match.Pos + Match.Len
        }
        return items.Length ? items.join("`n") : "No items found"
    }
    
    ExtractEmails(targetText) {
        pattern := "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
        return ExtractByRegex(targetText, pattern)
    }
    
    ExtractPhoneNumbers(targetText) {
        pattern := "\+?\d{1,3}[\s-]?(\(?\d{2,4}\)?[\s-]?)?\d{3,4}[\s-]?\d{3,4}"
        return ExtractByRegex(targetText, pattern)
    }
    
    ExtractWebsites(targetText) {
        pattern := "\b(?:https?://|www\.)[a-zA-Z0-9-]+\.[a-zA-Z]{2,}(?:\S*)\b"
        return ExtractByRegex(targetText, pattern)
    }
    
    AddTimestamp(targetText) {
        return A_Now . "`n" . targetText
    }
    
    NumberedList(targetText) {
        newText := ""
        index := 1
        Loop Parse, targetText, "`n", "`r" {
            if (A_LoopField != "") {
                newText .= index ". " . A_LoopField . "`r`n"
                index++
            } else {
                newText .= "`r`n"   
                index := 1
            }
        }
        return Trim(newText, "`r`n")
    }
    
    BulletList(targetText) {
        newText := ""
        Loop Parse, targetText, "`n", "`r" {
            if (A_LoopField != "") {
                newText .= "• " . A_LoopField . "`r`n"
            } else {
                newText .= "`r`n"   
            }
        }
        return Trim(newText, "`r`n")
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


}