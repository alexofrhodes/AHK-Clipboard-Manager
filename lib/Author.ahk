#Requires AutoHotkey v2

; Determine the directory of the currently running script file (Author.ahk)
AuthorDir := DirExist(A_LineFile) ? A_LineFile : RegExReplace(A_LineFile, "\\[^\\]+$")
IconsFolder := AuthorDir "\authorIcons"

; Define your links here
links := { 
    LinkedIn: "www.linkedin.com/in/alexofrhodes/",
    Github: "https://github.com/alexofrhodes/",
    WebSite: "https://alexofrhodes.github.io/",
    YouTube: "https://www.youtube.com/@alexofrhodes",
    InstaGram: "https://www.instagram.com/alexofrhodes/",
    BuyMeACoffee: "https://www.buymeacoffee.com/AlexOfRhodes",
    Gmail: "mailto:anastasioualex@gmail.com?subject=" A_ScriptName "&body=Hi! I would like to talk about ..."
}

; Create an ImageList for icons
imageList := Map()

; Load icons from the AuthorIcons folder
for name in links.ownprops() {
    iconPath := IconsFolder "\" name ".ico"
    if FileExist(iconPath)
        imageList[name] := iconPath
}

addAuthorTray()  ; Automatically add links to the tray menu

addAuthorTray(*) {
    TrayIcon := StrReplace(IconsFolder "\" . A_ScriptName, ".ahk", ".ico")
    if FileExist(TrayIcon)
        TraySetIcon(TrayIcon)

    Tray := A_TrayMenu
    Tray.Add()
    
    for name, url in links.ownprops() {
        Tray.Add(name, FollowLink.Bind(url))
        if imageList.Has(name)
            Tray.SetIcon(name, imageList[name])
    }
}

; Call this function in the main script to add the links to the GUI menu bar
addAuthorMenubar(gui) {
    if gui.HasProp("MenuBar") && IsObject(gui.MenuBar) {
        MyMenuBar := gui.MenuBar  ; Use existing menu bar
    } else {
        MyMenuBar := MenuBar()  ; Create new menu bar
        gui.MenuBar := MyMenuBar
    }

    ; Check if "Author" menu exists, otherwise create it
    if !MyMenuBar.HasProp("Author") {
        AuthorMenu := Menu()
        MyMenuBar.Add("Author", AuthorMenu)
    } else {
        AuthorMenu := MyMenuBar.Author
    }

    ; Add links under "Author" menu with icons
    for name, url in links.ownprops() {
        AuthorMenu.Add(name, FollowLink.Bind(url))
        if imageList.Has(name)
            AuthorMenu.SetIcon(name, imageList[name])
    }
}

FollowLink(url,*) {
    Run url
}
