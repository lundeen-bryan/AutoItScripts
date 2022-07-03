Opt('WinWaitDelay',1000)
Opt('WinDetectHiddenText',1)
Opt("WinTitleMatchMode", 4) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase
Opt("SendKeyDelay", 7) ;7 milliseconds

   ;start the temp repos folder
   TrayTip("start", "attempting to open temp repos", 0)
   Run("explorer.exe " & @HomePath & "\Documents\temp_repos\")
   Sleep(2000)
   Local $repos = WinGetTitle("[ACTIVE]","")
   If $repos = "" Then
    TrayTip("Repo not found", "Sorry the repo folder wasn't found", 3)
   Else
    TrayTip("Found it!", $repos & " successfully opened.", 3)
   EndIf
   WinActivate($repos, "")
   Send("#{LEFT}")
   TrayTip("")
   ;start the downloads folder
   Sleep(500)
   Run("explorer.exe " & @HomePath & "\Downloads")
   Sleep(2000)
   Local $downs = WinGetTitle("[ACTIVE]","")
   If $downs = "" Then
    TrayTip("Downloads folder not found", "Sorry the download folder wasn't opened", 3)
   Else
    TrayTip("Found it!", $downs & " successfully opened.", 3)
   EndIf
   WinActivate($downs, "")
   Send("#{RIGHT}")
   TrayTip("")
   ;start typertask
   Run(@MyDocumentsDir &  "\admin\typertask\typertask.exe")
   Run("C:\Program Files\Google\Chrome\Application\chrome.exe")

; #ABOUT# ==================================================================================
; Date: .............. 2022-07-03
; Title: ............. home_startup.au3
; Language: .......... English
; Website: ........... n/a
; Description: ....... runs just a couple apps automatically
; Installs to: ....... Documents/AutoIT
; Contact Author: .... lundeen-bryan
; Copyright © ........ n/a 2022. All rights reserved.
; Includes: .......... MsgBoxConstants.au3
; Syntax: ............ n/a
; Parameters: ........ n/a
; Return values: ..... n/a
;
;
; ==========================================================================================


