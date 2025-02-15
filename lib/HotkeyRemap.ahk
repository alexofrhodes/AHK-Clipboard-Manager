/* Readme

Solution by Datapoint   https://www.autohotkey.com/boards/viewtopic.php?t=114887  

Changelog:
https://github.com/alexofrhodes  
anastasioualex@gmail.com  

2024/10/24

made gui scrollable in case we have many hotkeys
changed buttons to menubar

2024/10/01 

Modified so that 
- it can be used by #Include  
- it can handle multiple scripts 


Example:

    ; main.ahk
    #Include HotkeyRemap.ahk

    AddHotkey("FunctionA", "FuncA_Hotkey", "Activate Function A")
    AddHotkey("FunctionB", "FuncB_Hotkey", "Activate Function B")
    AddHotkey("CloseApp", "CloseApp_Hotkey", "Close Application")

    HotkeyRemap()

    FunctionA(*) {
        MsgBox "Function A triggered!"
    }

    FunctionB(*) {
        MsgBox "Function B triggered!"
    }

    CloseApp(*) {
        MsgBox "Closing application."
        ExitApp
    }

*/

#Requires AutoHotkey v2
;[edited to make it x64/x32 compatible]
;Scrollable Gui - Proof of Concept - Scripts and Functions - AutoHotkey Community
;https://autohotkey.com/board/topic/26033-scrollable-gui-proof-of-concept/#entry168174

global ScriptName := StrSplit(A_ScriptName, ".")[1]
global SavedHotkeys := []
global remapGui := Gui("+Resize +0x300000")  ; WS_VSCROLL | WS_HSCROLL

A_TrayMenu.Add
A_TrayMenu.Add("Set Hotkeys", SetHotkeys)
if FileExist("hotkey.ico")
    A_TrayMenu.SetIcon("Set Hotkeys", "hotkey.ico")

/**
 * 
 * @param FunctionName 
 * @param INI_Key 
 * @param GUI_Text 
 */
AddHotkey(FunctionName, INI_Key, GUI_Text) {
    global SavedHotkeys
    SavedHotkeys.Push({
        Function: FunctionName,
        INI_Key: INI_Key,
        GUI_Text: GUI_Text,
        GUI_CtrlObj: "",
        AssignedKey: ""
    })
}

HotkeyRemap() {
    global SavedHotkeys, remapGui, ScriptName
    HK_NOT_SET := false

    MyMenuBar := MenuBar()
    MyMenuBar.Add("&OK", OK_Click)
    MyMenuBar.Add("&Cancel", Cancel_Click)
    MyMenuBar.Add("&Reset", ResetHotkeys)
    remapGui.MenuBar := MyMenuBar

    SetHotkeyCount :=0
    for i, MyHotkey in SavedHotkeys {
        ; MyHotkey.AssignedKey := IniRead(A_ScriptDir "\Hotkeys.ini", "Hotkeys", MyHotkey.INI_Key, "")
        
        ; changed "Hotkeys" to ScriptName to have HotkeysRemap handle multiple scripts, eg having it in <LIB>
        MyHotkey.AssignedKey := IniRead(A_ScriptDir "\Hotkeys.ini", ScriptName, MyHotkey.INI_Key, "")
        remapGui.AddText("x10 y+m w200", MyHotkey.GUI_Text)
        MyHotkey.GUI_CtrlObj := remapGui.AddHotkey("x+10 w150", MyHotkey.AssignedKey)
        if MyHotkey.AssignedKey = ""
        {
            HK_NOT_SET := true
        }else{
            Hotkey MyHotkey.AssignedKey, %MyHotkey.Function%, "On"
            SetHotkeyCount++
        }
    }
    remapGui.Title := "Set " . SetHotkeyCount . " of " . SavedHotkeys.Length . " available hotkeys"
    remapgui.Opt("+AlwaysOnTop" )

    ; show immediately if there are hotkeys that are not set
    if HK_NOT_SET
        SetHotkeys()
}

SetHotkeys(*) {
    global remapGui
    
    remapGui.Show("h500")
    OnMessage(0x0115, OnScroll) ; WM_VSCROLL
    ; OnMessage(0x0114, OnScroll) ; WM_HSCROLL
    OnMessage(0x020A, OnWheel)  ; WM_MOUSEWHEEL
}

OK_Click(*) {
    global SavedHotkeys, ScriptName

    for i, MyHotkey in SavedHotkeys {
        if MyHotkey.AssignedKey != ""
            Hotkey MyHotkey.AssignedKey, "Off"
    }

    remapGui.Submit()
    for i, MyHotkey in SavedHotkeys {
        MyHotkey.AssignedKey := MyHotkey.GUI_CtrlObj.Value
        if MyHotkey.AssignedKey = ""
            continue
        Hotkey MyHotkey.AssignedKey, %MyHotkey.Function%, "On"
        ; IniWrite MyHotkey.AssignedKey, A_ScriptDir "\Hotkeys.ini", "Hotkeys", MyHotkey.INI_Key
        IniWrite MyHotkey.AssignedKey, A_ScriptDir "\Hotkeys.ini", ScriptName, MyHotkey.INI_Key
    }
}

Cancel_Click(*) {
    global 
    remapGui.Hide() ; Hide the GUI instead of destroying it
}

ResetHotkeys(*) {
    global 
    ; remapGui.Hide()
    ans := MsgBox("Cannot be undone. Proceed?", "Reset Hotkeys", "0x4 0x1000")
    ; remapGui.Show()
    if ans = "No"
        return

    ; Disable all current hotkeys and clear their assignments
    for i, MyHotkey in SavedHotkeys {
        if MyHotkey.AssignedKey != "" {
            Hotkey MyHotkey.AssignedKey, "Off"
            MyHotkey.AssignedKey := ""
            IniWrite MyHotkey.AssignedKey, A_ScriptDir "\Hotkeys.ini", ScriptName, MyHotkey.INI_Key
        }
        ; Clear the value of the GUI control
        MyHotkey.GUI_CtrlObj.Value := ""
    }

    ; MsgBox("All hotkeys have been reset.",, 0x1000)
    remapGui.Title := "Set 0 of " . SavedHotkeys.Length . " available hotkeys"
}


; ======================================================================================================================
remapGui_Size(GuiObj, MinMax, Width, Height) {
    If (MinMax != 1)
       UpdateScrollBars(GuiObj)
 }
 ; ======================================================================================================================
 remapGui_Close(*) {
    ExitApp
 }
 ; ======================================================================================================================
 UpdateScrollBars(GuiObj) {
    ; SIF_RANGE = 0x1, SIF_PAGE = 0x2, SIF_DISABLENOSCROLL = 0x8, SB_HORZ = 0, SB_VERT = 1
    ; Calculate scrolling area.
    WinGetClientPos( , , &GuiW, &GuiH, GuiObj.Hwnd)
    L := T := 2147483647   ; Left, Top
    R := B := -2147483648  ; Right, Bottom
    For CtrlHwnd In WinGetControlsHwnd(GuiObj.Hwnd) {
       ControlGetPos(&CX, &CY, &CW, &CH, CtrlHwnd)
       L := Min(CX, L)
       T := Min(CY, T)
       R := Max(CX + CW, R)
       B := Max(CY + CH, B)
    }
    L -= 8, T -= 8
    R += 8, B += 8
    ScrW := R - L ; scroll width
    ScrH := B - T ; scroll height
    ; Initialize SCROLLINFO.
    SI := Buffer(28, 0)
    NumPut("UInt", 28, "UInt", 3, SI, 0) ; cbSize , fMask: SIF_RANGE | SIF_PAGE
    ; Update horizontal scroll bar.
    NumPut("Int", ScrW, "Int", GuiW, SI, 12) ; nMax , nPage
    DllCall("SetScrollInfo", "Ptr", GuiObj.Hwnd, "Int", 0, "Ptr", SI, "Int", 1) ; SB_HORZ
    ; Update vertical scroll bar.
    ; NumPut("UInt", SIF_RANGE | SIF_PAGE | SIF_DISABLENOSCROLL, SI, 4) ; fMask
    NumPut("Int", ScrH, "UInt", GuiH,  SI, 12) ; nMax , nPage
    DllCall("SetScrollInfo", "Ptr", GuiObj.Hwnd, "Int", 1, "Ptr", SI, "Int", 1) ; SB_VERT
    ; Scroll if necessary
    X := (L < 0) && (R < GuiW) ? Min(Abs(L), GuiW - R) : 0
    Y := (T < 0) && (B < GuiH) ? Min(Abs(T), GuiH - B) : 0
    If (X || Y)
       DllCall("ScrollWindow", "Ptr", GuiObj.Hwnd, "Int", X, "Int", Y, "Ptr", 0, "Ptr", 0)
 }
 ; ======================================================================================================================
 OnWheel(W, L, M, H) {
    If !(HWND := WinExist()) || GuiCtrlFromHwnd(H)
       Return
    HT := DllCall("SendMessage", "Ptr", HWND, "UInt", 0x0084, "Ptr", 0, "Ptr", l) ; WM_NCHITTEST = 0x0084
    If (HT = 6) || (HT = 7) { ; HTHSCROLL = 6, HTVSCROLL = 7
       SB := (W & 0x80000000) ? 1 : 0 ; SB_LINEDOWN = 1, SB_LINEUP = 0
       SM := (HT = 6) ? 0x0114 : 0x0115 ;  WM_HSCROLL = 0x0114, WM_VSCROLL = 0x0115
       OnScroll(SB, 0, SM, HWND)
       Return 0
    }
 }
 ; ======================================================================================================================
 OnScroll(WP, LP, M, H) {
    Static SCROLL_STEP := 10
    If !(LP = 0) ; not sent by a standard scrollbar
       Return
    Bar := (M = 0x0115) ; SB_HORZ=0, SB_VERT=1
    SI := Buffer(28, 0)
    NumPut("UInt", 28, "UInt", 0x17, SI) ; cbSize, fMask: SIF_ALL
    If !DllCall("GetScrollInfo", "Ptr", H, "Int", Bar, "Ptr", SI)
       Return
    RC := Buffer(16, 0)
    DllCall("GetClientRect", "Ptr", H, "Ptr", RC)
    NewPos := NumGet(SI, 20, "Int") ; nPos
    MinPos := NumGet(SI,  8, "Int") ; nMin
    MaxPos := NumGet(SI, 12, "Int") ; nMax
    Switch (WP & 0xFFFF) {
       Case 0: NewPos -= SCROLL_STEP ; SB_LINEUP
       Case 1: NewPos += SCROLL_STEP ; SB_LINEDOWN
       Case 2: NewPos -= NumGet(RC, 12, "Int") - SCROLL_STEP ; SB_PAGEUP
       Case 3: NewPos += NumGet(RC, 12, "Int") - SCROLL_STEP ; SB_PAGEDOWN
       Case 4, 5: NewPos := WP >> 16 ; SB_THUMBTRACK, SB_THUMBPOSITION
       Case 6: NewPos := MinPos ; SB_TOP
       Case 7: NewPos := MaxPos ; SB_BOTTOM
       Default: Return
    }
    MaxPos -= NumGet(SI, 16, "Int") ; nPage
    NewPos := Min(NewPos, MaxPos)
    NewPos := Max(MinPos, NewPos)
    OldPos := NumGet(SI, 20, "Int") ; nPos
    X := (Bar = 0) ? OldPos - NewPos : 0
    Y := (Bar = 1) ? OldPos - NewPos : 0
    If (X || Y) {
       ; Scroll contents of window and invalidate uncovered area.
       DllCall("ScrollWindow", "Ptr", H, "Int", X, "Int", Y, "Ptr", 0, "Ptr", 0)
       ; Update scroll bar.
       NumPut("Int", NewPos, SI, 20) ; nPos
       DllCall("SetScrollInfo", "ptr", H, "Int", Bar, "Ptr", SI, "Int", 1)
    }
 }