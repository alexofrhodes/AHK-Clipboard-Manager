#Requires AutoHotkey v2.0
#SingleInstance Force

VERSION := 1.0

RegExMatch(A_ScriptName, "^(.*?)\.", &basename)
global WINTITLE := basename[1] " " VERSION
global defaultINIfile := A_ScriptDir "\" basename[1] ".ini" 
global guiName := ""
global iniFile := ""
global Preset := "Default"

ListLines(false)

; only guiName is required, may be myGui.Title (or myGui.name?)

GuiSaveState(guiName := "", inifile := "", Preset := 1, afterThisControl := "", beforeThisControl := "") {
    global
    if (inifile = "") 
        inifile := defaultINIfile
  
    if (guiName = ""){
        MsgBox "GuiSaveState: guiName is required, may be myGui.Title or myGui.name",, 0x1000
        return
    }
          
    oList_controls := WinGetControls(guiName)
    
    flag := (afterThisControl = "" ? 0 : 1)
    
    for n, ctrlName in oList_controls {
        ctrl := myGui[ctrlname]
        if !IsObject(ctrl) 
            continue
        if ctrl.Name = ""
            continue
        
        if (afterThisControl = ctrl.Name) {
            flag := 0
            continue
        }
        if (flag)
            continue
        if (beforeThisControl = ctrl.Name)
            break
        
        value := RegExReplace(ctrl.Value, "`n", "|")
        IniWrite(value, inifile, guiName "-" Preset, ctrl.Name)
    }
    GuiSavePosition()
}

GuiSavePosition(*){
    global
    winstate := WinGetMinMax(guiName)
    if (winstate != -1) {
        x := y := width := height := ""
        WinGetPos(&x, &y, &width, &height, guiName)
        IniWrite(x, inifile, guiName "-" Preset, "winx")
        IniWrite(y, inifile, guiName "-" Preset, "winy")
    }
}
GuiLoadState(guiName := "", inifile := "", Preset := 1) {
    global
    if (guiName = ""){
        MsgBox "GuiLoadState: guiName is required, may be myGui.Title or myGui.name",, 0x1000
        return
    }
    if (inifile = "") 
        inifile := defaultINIfile
    
    oList_controls := WinGetControls(guiName)
    
    for n, ctrlName in oList_controls {
        ctrl := myGui[ctrlname]
        if !IsObject(ctrl) 
            continue
        if ctrl.Name = ""
            continue
        value := IniRead(inifile, guiName "-" Preset, ctrl.Name, "ERROR")
        if IsNumber(value )
            Value := value + 0

        if (value != "ERROR") {
            ; value := RegExReplace(value, "\\|", "`n")
            if ctrl.Type = "ComboBox" || ctrl.Type = "ListBox" 
                ctrl.Choose(value)
            else if ctrl.Type = "CheckBox" 
                ControlSetChecked value, ctrl
            else
                ctrl.value := value
        }
    }
    guiLoadPosition()
}

guiLoadPosition(*){
    global
    winx := IniRead(inifile, guiName "-" Preset, "winx", 0)
    winy := IniRead(inifile, guiName "-" Preset, "winy", 0)
    
    VirtualWidth := SysGet(78)
    VirtualHeight := SysGet(79)
    if (winx < 0 || winx > VirtualWidth)
        winx := 0
    if (winy < 0 || winy > VirtualHeight)
        winy := 0

    myGui.Show("x" winx " y" winy)
}

GuiDeleteState(GuiName, IniFile := "" , Preset:=1) {
       
    if (inifile = "") 
        inifile := defaultINIfile

    IniDelete(inifile, GuiName "-" Preset)
}

GuiResetState(guiName := "") {
    if (guiName = "") {
        MsgBox "GuiResetState: guiName is required, may be myGui.Title or myGui.name",, 0x1000
        return
    }
    
    oList_controls := WinGetControls(guiName)
    
    for n, ctrlName in oList_controls {
        ctrl := myGui[ctrlname]
        if !IsObject(ctrl) 
            continue
        if ctrl.Name = ""
            continue
        
        ; Reset values based on control type
        if ctrl.Type = "ComboBox" || ctrl.Type = "ListBox" || ctrl.Type = "DropDownList" 
            ctrl.Choose(0) ; Select first (default) item
        else if ctrl.Type = "CheckBox" || ctrl.Type = "Radio" 
            ControlSetChecked(false, ctrl) ; Uncheck boxes and radio buttons
        else 
            ctrl.Value := "" ; Clear Edit fields and other controls
    }
}

HasVal(haystack, needle) {
	if !(IsObject(haystack)) || (haystack.Length() = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}