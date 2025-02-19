; Change Log:
; 2025-02-19: Changed to Class

;=======================================================================
; usasage:

; #include GuiState.ahk
;; create gui and controls first
; global guiManager := GuiState(myGui, "GuiState.ini", 1)      
;=======================================================================

class GuiState {
    myGui := ""
    guiName := ""   
    iniFile := ""
    preset := 1


    __New(myGui := "", iniFile := "", preset := 1) {
        this.myGui := myGui
        this.guiName := myGui.Title
        this.iniFile := (iniFile = "" ? A_ScriptDir "\" this.guiName ".ini" : iniFile)
        this.preset := preset
    }

    SaveState(afterThisControl := "", beforeThisControl := "") {
        if (this.guiName = "") {
            MsgBox "GuiSaveState: guiName is required, may be myGui.Title",, 0x1000
            return
        }

        flag := (afterThisControl = "" ? 0 : 1)

        For Hwnd, Ctrl in MyGui {
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
            IniWrite(value, this.iniFile, this.guiName "-" this.preset, ctrl.Name)
        }
        this.SavePosition()
    }

    SavePosition() {
        winstate := WinGetMinMax(this.myGui.hwnd)
        if (winstate != -1) {
            x := y := width := height := ""
            MyGui.GetPos(&x, &y, &width, &height)
            IniWrite(x, this.iniFile, this.guiName "-" this.preset, "winX")
            IniWrite(y, this.iniFile, this.guiName "-" this.preset, "winY")

            MyGui.GetClientPos(&x, &y, &width, &height)
            IniWrite(width, this.iniFile, this.guiName "-" this.preset, "winW")
            IniWrite(height, this.iniFile, this.guiName "-" this.preset, "winH")
        }
    }

    LoadState() {

        oList_controls := WinGetControls(this.myGui.hwnd)

        For Hwnd, Ctrl in MyGui{
            if ctrl.Name = ""
                continue

            value := IniRead(this.iniFile, this.guiName "-" this.preset, ctrl.Name, "ERROR")
            if IsNumber(value)
                value := value + 0

            if (value != "ERROR") {
                if ctrl.Type = "ComboBox" || ctrl.Type = "ListBox" 
                    ctrl.Choose(value)
                else if ctrl.Type = "CheckBox" 
                    ctrl.Value := value
                else
                    ctrl.Value := value
            }
        }
        this.LoadPosition()
    }

    LoadPosition() {
        winX := IniRead(this.iniFile, this.guiName "-" this.preset, "winX", 0)
        winY := IniRead(this.iniFile, this.guiName "-" this.preset, "winY", 0)
        winW := IniRead(this.iniFile, this.guiName "-" this.preset, "winW", 0)
        winH := IniRead(this.iniFile, this.guiName "-" this.preset, "winH", 0)

        VirtualWidth := SysGet(78)
        VirtualHeight := SysGet(79)
        if (winX < 0 || winX > VirtualWidth)
            winX := 0
        if (winY < 0 || winY > VirtualHeight)
            winY := 0

        if (winW = "0" || winH = "0") { 
            ; Only move the GUI without forcing size
            this.myGui.Show("x" winX " y" winY)
        } else { 
            ; Apply full position and size if valid
            this.myGui.Show("x" winX " y" winY " w" winW " h" winH)
        }
    }

    DeleteState() {
        IniDelete(this.iniFile, this.guiName "-" this.preset)
    }

    ResetState() {

        oList_controls := WinGetControls(this.mygui.hwnd)

        for n, ctrlName in oList_controls {
            ctrl := this.myGui[ctrlName]
            if !IsObject(ctrl) || ctrl.Name = ""
                continue

            if ctrl.Type = "ComboBox" || ctrl.Type = "ListBox" || ctrl.Type = "DropDownList" 
                ctrl.Choose(0)
            else if ctrl.Type = "CheckBox" || ctrl.Type = "Radio" 
                ctrl.Value := false
            else 
                ctrl.Value := ""
        }
    }
}
