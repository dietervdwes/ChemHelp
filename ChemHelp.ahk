;ChemHelp v.1.12.1 - updated 22-07-2021
;Written by Dieter van der Westhuizen 2018-2021
;Inspired from TrakHelper by Chad Centner

#SingleInstance, force
;#NoTrayIcon
#NoEnv
;#Persistent
#MaxThreadsPerHotkey 2
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 85
SendMode, Input
SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
SetMouseDelay, 0
SetWinDelay, 500
^SPACE::  Winset, Alwaysontop, , A


;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Add Buttons ;;;;;;;;;;;;;;;;;;
Gui, Add, button, x2 y2 w40 h20 ,Log-on
Gui, Add, button, x44 y2 w40 h20, Mobile
Gui, add, button,x2 y22 w30 h20  ,Form
Gui, Add, button, x2 y42 w28 h20  ,EPR
Gui, Add, button, x30 y42 w45 h20, EPRxlsx
Gui, Add, button, x2 y64 w33 h20  ,FPSA
Gui, Add, button, x36 y64 w45 h20, dCDUM
Gui, Add, button, x2 y86 w45 h20  ,Verified
Gui, Add, button, x2 y108 w57 h20  ,KeepOpen
Gui, Add, button, x60 y108 w20 h20  ,Ex
Gui, Add, button, x2 y130 w36 h20  ,More
Gui, Add, button, x40 y130 w36 h20, VPN
;Gui, Add, button, x40 y130 w45 h20, ExMRN
Gui, Add, button, x2 y152 w32 h20 ,Close
Gui, Add, button, x45 y152 w20 h20 ,_i
;Gui, Add, Button, x6 y17 w100 h30 , Ok
;Gui, Add, Button, x116 y17 w100 h30 , Cancel
;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Set Window Options   ;;;;;;;;;;;;;;
;Gui, +AlwaysOnTop
Gui, -sysmenu +AlwaysOnTop
Gui, Show, , ChemHelp1.12
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
height := A_ScreenHeight-270
width := A_ScreenWidth-88
Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%width% y%height% w84




titleoffice = Microsoft Office Activation Wizard
classoffice = NUIDialog
settimer, office_activation_watch, 150
FileRead, chemhelpsettingsraw, %A_MyDocuments%\chemhelp_settings.txt
settingsarray := StrSplit(chemhelpsettingsraw, ",")
citrix_username := settingsarray[1]
citrix_password := settingsarray[2]
trakcare_username := settingsarray[3]
trakcare_password := settingsarray[4]
tabs_citrix := settingsarray[5]

return

/*
;Scheduling a task:
Loop,
{
Sleep, 60000; check time every minet
if %A_Hour% = 17 ; 5 o clock french time
    {
        ; your script here
    }
}
*/


;                                                          Generic script to close a pop-up window.
office_activation_watch:
IfWinExist, %titleoffice% ahk_class %classoffice% 
{
  WinActivate, %titleoffice%
  WinWaitActive, %titleoffice%
  sleep, 200
  send, {AltDown}c{AltUp}
  sleep, 100
settimer, office_activation_watch, off
}
else
  tooltip    ; remove the tooltip
return

;--                                                              Enabling and disapbling the proxy
;Write to the registry; set autoproxy to this value
+!p::
RegWrite, REG_DWORD, HKCU, Software\Microsoft\Windows\CurrentVersion\Internet Settings, ProxyEnable, 1
if errorlevel
	MsgBox % A_LastError
else
    MsgBox, Proxy enabled
return

;Write to the registry; set autoproxy to this value
+!r::
RegWrite, REG_DWORD, HKCU, Software\Microsoft\Windows\CurrentVersion\Internet Settings, ProxyEnable, 0
if errorlevel
	MsgBox % A_LastError
else
    MsgBox, Proxy disabled
return

; inetcpl.cpl <- Run this command in Run prompt to set proxy settings

buttonVPN:
run, "C:\Program Files\ShrewSoft\VPN Client\ipseca.exe"
WinWaitActive, VPN Access Manager
WinActivate, VPN Access Manager
MouseClick, left, 46, 121, 2, 100
WinWaitActive, VPN Connect
send, %citrix_username%
sleep, 200
send, {TAB down}{TAB up}
sleep, 200
send, %citrix_password%
sleep, 300
send, {TAB down}{TAB up}
send, {Return}
sleep, 2000
WinMinimize, VPN Connect
WinMinimize, VPN Access Manager
return

#c::
Run, Calc.exe
return


ButtonMobile:
run, ChemHelpMobile.ahk
Return

;---------------------------------             This is to LOG-ON with username and password.    

ButtonLog-on:
settimer, win_proxy_logon, 500
settimer, citrix_proxy_logon, 500
settimer, miscellaneous_user_config_file_notice, 500

url :="http://nhlslisapps02.nhls.ac.za/Citrix/XenApp/auth/login.aspx"
run, %url%
WinActivate, Citrix XenApp
WinWaitActive, Citrix XenApp
sleep, 2500
;FileRead, citrix_username, %A_MyDocuments%\citrix_username.txt
;FileRead, citrix_password, %A_MyDocuments%\citrix_password.txt
;FileRead, trakcare_username, %A_MyDocuments%\trakcare_username.txt
;FileRead, trakcare_password, %A_MyDocuments%\trakcare_password.txt
;FileRead, tabs_citrix, %A_MyDocuments%\tabs_citrix.txt
;FileRead, trakcare_workarea_tabs, %A_MyDocuments%\trakcare_workarea_tabs.txt
send, %citrix_username%
sleep, 500
send, {TAB down}{TAB up}
sleep, 200
send, %citrix_password%
sleep, 300
send, {Return}
sleep, 3000
/*
Loop %tabs_citrix%
{
    Send {Tab down}
	Send {Tab up}
    Sleep 80  ;      The number of milliseconds between keystrokes (or use SetKeyDelay).
}
*/
ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big100.PNG
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for the TrakCare Icon. Possibly the files in the folder trakcare_icons is missing in your "Documents" folder.
else if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big095.PNG
    if (ErrorLevel = 1)
        ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big090.PNG
        if (ErrorLevel = 1)
            ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big085.PNG
            if (ErrorLevel = 1)
                ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big080.PNG
                if (ErrorLevel = 1)
                    ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big105.PNG
                    if (ErrorLevel = 1)
                        ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big110.PNG
                        if (ErrorLevel = 1)
                            ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big115.PNG
                            if (ErrorLevel = 1)
                                ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big120.PNG
                                if (ErrorLevel = 1)
                                    ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_big125.PNG
if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small100.PNG
if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small095.PNG
    if (ErrorLevel = 1)
        ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small090.PNG
        if (ErrorLevel = 1)
            ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small085.PNG
            if (ErrorLevel = 1)
                ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small080.PNG
                if (ErrorLevel = 1)
                    ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small105.PNG
                    if (ErrorLevel = 1)
                        ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small110.PNG
                        if (ErrorLevel = 1)
                            ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small115.PNG
                            if (ErrorLevel = 1)
                                ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small120.PNG
                                if (ErrorLevel = 1)
                                    ImageSearch, FoundX, FoundY, 336,280, 1241, 850, %A_MyDocuments%\trakcare_icons\trakcarelab_small125.PNG
else if (ErrorLevel = 1)
MsgBox TrakCare Icon could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x`;%FoundY%y.
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
sleep, 300

send, {Enter}
sleep, 2000

settimer, labtrakstart_login, 1000

;;;;Windows Blue Proxy logon-screen sign-in
win_proxy_logon:
IfWinExist, Windows sign-in
{
WinActivate, Windows sign-in
send, {tab}
sleep, 200
send, %citrix_password%
sleep, 200
send, {enter}
settimer, win_proxy_logon, Off
}
else
    tooltip ;nothing
    return

citrix_proxy_logon:
IfWinExist, Proxy authentication required
{
WinActivate, Proxy authentication required
    sleep, 100
    send, %citrix_username%
    sleep, 200
    send, {Tab}
    sleep, 100
    send, %citrix_password%
    sleep, 200
    send, {Enter}
    sleep, 1000
    send, %citrix_password%
    sleep, 200
    ;settimer, citrix_proxy_logon, Off
}
else
    return

labtrakstart_login:
IfWinExist,  (Current Site : NHLS)
{
    settimer, labtrakstart_login, Off
    settimer, win_proxy_logon, Off
    ;settimer, citrix_proxy_logon, Off
    settimer, miscellaneous_user_config_file_notice, Off
    WinActivate,  (Current Site : NHLS)
    WinWaitActive,  (Current Site : NHLS)
    sleep, 200
    send, %trakcare_username%    
    sleep, 500
    send, {TAB down}{TAB up}
    send, %trakcare_password%                 
    sleep, 50

sleep, 1000
return
;Loop, %trakcare_workarea_tabs%
;{
;send, {TAB down}{TAB up}
;sleep, 500
;}
;send, {Down}
;sleep, 200
;send, {Enter}
;Return
}
else
  tooltip Waiting for TrakCare Window
return

miscellaneous_user_config_file_notice:
IfWinExist, LabTrakStart
{
WinActivate, LabTrakStart
WinWaitActive, LabTrakStart
    Send, {Enter}
    settimer, miscellaneous_user_config_file_notice, Off
    return
}
else
return

;;;;;;;;;;;;;;;;;;;;;                                                                           CREF error C900 entered but test is registered on TrakCare   
!y::
EpResultSingle()
sleep, 100
FileAppend, CREF-error`n%txt%`n, C:\TrakCare-Errors\TrakCare-Entry-Errors.txt
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                         Button KeepOpen  
ButtonKeepOpen:
;MsgBox, This script clicks the refresh button on "Results Entry - Verify" window every 10 minutes to keep TrakCare open.`nPlease make sure that "Results Entry - Verify" window is open when you leave your computer.`nClick OK to start.`nHit Esc when you are back on your computer.

MsgBox, 1, , This script clicks the refresh button on "Results Entry - Verify" window every 10 minutes to keep TrakCare open.`nPlease make sure that "Results Entry - Verify" window is open when you leave your computer.`nClick OK to start.`nHit Esc when you are back on your computer.`nThis MsgBox will time out in 5 seconds.  Continue?, 5
sleep, 300
SetTimer, refreshtimer, 600000 ;every 10 minutes
;SetTimer, refreshtimer, 300000 ;every 5 minutes
;SetTimer, refreshtimer, 5000 ;every 10 sec
;MsgBox, 1, , Timer has been set.`nThis Message will disappear automatically in 2 seconds, 2
Return

;This is the script to click refresh on Lab Results verify window. 
refreshtimer:
if WinExist("Result Entry - Verify ahk_class Transparent Windows Client")  
{
  WinActivate, Result Entry - Verify ahk_class Transparent Windows Client
  WinWaitActive, Result Entry - Verify ahk_class Transparent Windows Client
  sleep, 100
  WinGetPos,,,winw, winh
  mouseclick, left, winw-60 , 125, 1
  sleep, 150
  WinMinimize, Result Entry - Verify ahk_class Transparent Windows Client
  sleep, 500
  tooltip, Trak is being kept open automatically. `n Hit ESC to stop.
  sleep, 1000
  return
}
else if WinExist(" NHLS")
{
    If !WinActive(" NHLS")
    WinActivate,  NHLS 
    WinWaitActive,  NHLS 
    Sleep, 50
    send, {Alt}
    sleep, 50
    Send, v
    sleep, 50
    send, u
    WinWaitActive, User Audit Trail
    sleep, 100
    send, {Enter}
    sleep, 100
    WinMinimize,  NHLS
}
return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                           Alt+N Shortcut to copy Episode 
!n::
if WinActive("Medical Validation :   (Authorise By Episode)")
    {
    EpMedVal()
    Return
    }
else
    if WinActive("Result Entry - Single -")
    {
    EpResultSingle()
    Return
    }

;------------------------------------------                                           Ctrl+R Navigates to result entry for episode
ButtonRESULTENTRY:
WinActivate, Medical Validation :   (Authorise By Episode)
 
#IfWinActive, Medical Validation :   (Authorise By Episode)
!r::
EpMedVal()
sleep, 200
WinActivate Medical Validation :   (Authorise By Episode)
sleep, 200
   send, {altdown}<{altup}
sleep, 200
 
if !WinExist("Result Entry, , Result Entry - Verify")
{
                If !WinActive(" NHLS,")
    WinActivate,  NHLS, 
                WinWaitActive,  NHLS, 
                Sleep, 100
                Send, a{TAB}
                Send, Result Entry{Enter}
                WinWaitActive, Result Entry  
    sleep, 200
}
 
 
{
WinActivate, Result Entry, , Result Entry - Verify 
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 200
send, {AltDown}l{AltUp}
sleep, 200
send, {CtrlDown}v{CtrlUp}
sleep, 300
send, {Enter}
sleep, 400
}
return
#IfWinActive


;-------------------------                                                                  Alt+P Enter PHONC
ButtonPHONC:
WinActivate, Medical Validation :   (Authorise By Episode)

#IfWinActive, Medical Validation :   (Authorise By Episode)
!p::
EpMedVal()
sleep, 200
WinActivate Medical Validation :   (Authorise By Episode)
sleep, 200
   send, {altdown}<{altup}
sleep, 200

if !WinExist("Result Entry, , Result Entry - Verify")
{
	If !WinActive(" NHLS,")
    WinActivate,  NHLS, 
	WinWaitActive,  NHLS, 
	Sleep, 100
	Send, a{TAB}
	Send, Result Entry{Enter}
	WinWaitActive, Result Entry  
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify 
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 100
if WinExist("Search Unsuccessful")
{
    WinClose, Search Unsuccessful
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify 
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 200
send, {AltDown}l{AltUp}
sleep, 200
send, {CtrlDown}v{CtrlUp}
sleep, 300
send, {Enter}
sleep, 300
MouseClick, Left, 340, 133, 2, 100
sleep, 100
Send, PHONC
sleep, 100
send, {Enter}
sleep, 400
if WinExist("Search Unsuccessful")
{
    MsgBox, 4,,  PHONC does not appear for this Episode. `nDo you want to add a PHONC to this episode?
    IfMsgBox Yes
    {
        WinClose, Search Unsuccessful
        sleep, 200
        WinActivate, Result Entry, , Result Entry - Verify 
        WinWaitActive, Result Entry, , Result Entry - Verify
        sleep, 200
        send, {AltDown}l{AltUp}
        sleep, 400
        send, {CtrlDown}v{CtrlUp}
        sleep, 300
        send, {Enter}
        sleep, 300
        MouseClick, Left, 215, 300
        sleep, 200
        send, {altdown}d{altup}
        sleep, 200
        send, s
        sleep, 200
        WinWaitActive, Test Set Maintenance
        sleep, 200
        send, PHONC
        sleep, 200
        send, {tab}
        sleep, 100
        send, {enter}
        return
    }
    else IfMsgBox No
    { 
        sleep, 300
        WinClose, Search Unsuccessful
        Return 
    }
}
else
MouseClick, left, 70, 300, 1
sleep, 200
send, {AltDown}e{AltUp}
WinWaitActive, Result Entry - Single
sleep, 200
send, t{tab}
sleep, 100
send, n{tab}

return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                            Ctrl+Alt+P Enter CCOM with Dr. update links


#IfWinActive, Medical Validation :   (Authorise By Episode)
^!p::
EpMedVal()
sleep, 100
WinActivate Medical Validation :   (Authorise By Episode)
sleep, 200
   send, {altdown}<{altup}
sleep, 200

IfWinNotExist, Result Entry, , Result Entry - Verify 
{
	IfWinNotActive,  NHLS,
    WinActivate,  NHLS, 
	WinWaitActive,  NHLS, 
	Sleep, 100
	Send, a{TAB}
	Send, Result Entry{Enter}
	WinWaitActive, Result Entry  
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify 
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 100
IfWinExist, Search Unsuccessful
{
    WinClose, Search Unsuccessful
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify 
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 150
send, {AltDown}l{AltUp}
sleep, 150
send, {CtrlDown}v{CtrlUp}
sleep, 200
send, {Enter}
sleep, 200
/*
MouseClick, Left, 340, 133, 2, 100
sleep, 100
Send, CCOM
sleep, 100
send, {Enter}
sleep, 250
IfWinExist, Search Unsuccessful
{
    MsgBox, 4,,  CCOM does not appear for this Episode. `nDo you want to add a CCOM to this episode?
    IfMsgBox Yes
    {
        WinClose, Search Unsuccessful
        sleep, 200
        WinActivate, Result Entry, , Result Entry - Verify 
        WinWaitActive, Result Entry, , Result Entry - Verify
        sleep, 200
        send, {AltDown}l{AltUp}
        sleep, 400
        send, {CtrlDown}v{CtrlUp}
        sleep, 300
        send, {Enter}
        sleep, 300
        MouseClick, Left, 215, 300
        sleep, 200
        send, {altdown}d{altup}
        sleep, 200
        send, s
        sleep, 200
        WinWaitActive, Test Set Maintenance
        sleep, 200
        send, CCOM
        sleep, 200
        send, {tab}
        sleep, 100
        send, {enter}
        return
    }
    else IfMsgBox No
    { 
        sleep, 300
        WinClose, Search Unsuccessful
        Return 
    }
}
else
*/
MouseClick, left, 70, 300, 1
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 200
WinWaitActive, Test Set Maintenance
sleep, 500
send, CCOM
sleep, 500
send, {tab}
sleep, 300
send, {enter}
sleep, 300
WinClose, Test Set Maintenance
WinWaitClose, Test Set Maintenance
sleep, 500

MouseClick, Left, 340, 133, 2, 100
sleep, 100
Send, CCOM
sleep, 100
send, {Enter}
sleep, 250

IfWinExist, Search Unsuccessful
{
    MsgBox, 4,,  CCOM does not appear for this Episode. `nDo you want to add a CCOM to this episode?
    IfMsgBox Yes
    {
        WinClose, Search Unsuccessful
        sleep, 200
        WinActivate, Result Entry, , Result Entry - Verify 
        WinWaitActive, Result Entry, , Result Entry - Verify
        sleep, 200
        send, {AltDown}l{AltUp}
        sleep, 400
        send, {CtrlDown}v{CtrlUp}
        sleep, 300
        send, {Enter}
        sleep, 300
        MouseClick, Left, 215, 300
        sleep, 200
        send, {altdown}d{altup}
        sleep, 200
        send, s
        sleep, 200
        WinWaitActive, Test Set Maintenance
        sleep, 150
        send, CCOM
        sleep, 150
        send, {tab}
        sleep, 100
        send, {enter}
        return
    }
    else IfMsgBox No
    { 
        sleep, 300
        WinClose, Search Unsuccessful
        Return 
    }
}
else

MouseClick, Left, 400, 300, 1
sleep, 150
send, {AltDown}e{AltUp}
WinWaitActive, Result Entry - Single - 
sleep, 100
send, {F6}
WinWaitActive, Comments
sleep, 50
send, Clinician contact details may not be coded in our database.  Please follow the link below to update clinician contact details:`ntinyurl.com/nhls-update       ;<<-----Put link in here after the "`n"
sleep, 1800
WinClose, Comments
sleep, 100
send, {AltDown}a{AltUp}
sleep, 800
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation
sleep, 200
send, {AltDown}>{AltUp}

return

#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                 ALT+W for EPR      ALT+W for EPR   ALT+W for EPR      ALT+W for EPR  

ButtonEPR:
send, {AltDown}{Tab}{AltUp}
sleep, 200 

!w::
sleep, 300
IfWinActive, Result Entry - Single - 
{
MRNSingle()
txt := Clipboard
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
return
}
    
else IfWinActive, Medical Validation :   (Authorise By Episode)
{
MRNMedVal()
txt := Clipboard
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
return
}

else IfWinActive, Result Verify - Single - 
{
MRNResultVerSingle()
txt := Clipboard
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
return
}

else
{
MsgBox, No TrakCare window active.  `nPlease ensure Result Entry or Result Verify windows are open and visible.
sleep, 200
return
}
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Alt+F for Form    Alt+F for Form    Alt+F for Form    Alt+F for Form    Alt+F for Form 
ButtonForm:
sleep, 200
send, {AltDown}{Tab}{AltUp}
sleep, 200

!F::
;sleep, 300
IfWinActive, Medical Validation :   (Authorise By Episode)
{
EpMedVal()
txt := ClipBoard
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
}
else
IfWinActive,  Result Entry - Single - 
{
EpResultSingle()
txt := ClipBoard
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
}
else
IfWinActive, Result Verify - Single -
{
EpResultVerSingle()
txt := ClipBoard
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
}
else
{
MsgBox, No TrakCare window active.  `nPlease ensure Result Entry or Result Verify windows are open and visible.
sleep, 200
send, {AltDown}e{AltUp}
return
}

iehalfscreen:
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
return



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                          Button FPSA 
ButtonFPSA:
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 200
WinWaitActive, Test Set Maintenance
ImageSearch, FoundX, FoundY, 15,86, 523, 379, %A_MyDocuments%\cdum.png
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for CDUM. Possibly the file cdum.png is missing in "My Documents" folder.
else if (ErrorLevel = 1)
    MsgBox CDUM could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x%FoundY%.
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
    send, {AltDown}d{AltUp}
sleep, 300
WinActivate, Test Set Maintenance
sleep, 100
MouseClick, Left, 830, 308 
sleep, 100
send, FPSA
sleep, 200
send, {tab down}{tab up}
send, {enter}
Return


ButtondCDUM:
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 200
WinWaitActive, Test Set Maintenance
sleep, 800
ImageSearch, FoundX, FoundY, 15,86, 523, 379, %A_MyDocuments%\cdum.png
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for CDUM. `nPossibly the file cdum.png is missing in "My Documents" folder.`n If this is the issue, find it on github.com/dietervdwes/chemhelp
else if (ErrorLevel = 1)
    MsgBox CDUM could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x%FoundY%.
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
    send, {AltDown}d{AltUp}
sleep, 300
WinClose, Test Set Maintenance
MsgBox, Please confirm that CDUM2 has been removed.
Return
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                               Button Verified to insert staff note "Transcription Verified"  
ButtonVerified:
sleep, 200
send, !{Tab}
sleep, 200
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                Alt + V for Staff note in Med Val or Single: Transcription Verified 

!v::
IfWinActive, Medical Validation :   (Authorise By Episode)
{
;WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 100
send, {altdown}d{altup}
sleep, 100
send, f
WinWaitActive, Staff Notes
sleep, 100
send, Transcription Verified.
sleep, 100
send, {alt down}o{alt up}
sleep, 50
Return
}
else
IfWinActive, Result Verify - Single
{
WinActivate, Result Verify - Single
WinWaitActive, Result Verify - Single
sleep, 200
MouseClick, left, 904, 164, 1
sleep, 500
MouseClick, left, 19, 292, 1
sleep, 500
send, Transcription Verified.
sleep, 50
send, {alt down}o{alt up}
sleep, 50
Return
}
else
{
return
}

;;;;;;;;;;;;;;;;;;;                                                                             Alt + / for Staff note in Med Val or Single: Transcription Verified 
!/::
IfWinActive, Medical Validation :   (Authorise By Episode)
{
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, f
sleep, 200
send, Not yet scanned.
sleep, 100
send, {alt down}o{alt up}
sleep, 50
Return
}
else
IfWinActive, Result Verify - Single
{
WinActivate, Result Verify - Single
WinWaitActive, Result Verify - Single
sleep, 200
MouseClick, left, 904, 164, 1
sleep, 500
MouseClick, left, 19, 292, 1
sleep, 500
send, Not yet scanned.
sleep, 50
send, {alt down}o{alt up}
sleep, 50
Return
}
else
{
return
}

;;;;;;;;;;;;;;;;;;;;;;                                                                                     Page Down when Alt + z is pressed.
#IfWinActive Medical Validation :   (Authorise By Episode)
!z::
   mouseclick, left, 433,  231, 1
   sleep, 100
   send, {PgDn down}{PgDn up}
Return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;                                                                                                     Page Up when Alt + q is pressed.
#IfWinActive Medical Validation :   (Authorise By Episode)
!q::
   mouseclick, left, 433,  231, 1
   sleep, 100
   send, {PgUp down}{PgUp up}
Return
#IfWinActive


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                   Alt + B for Back 
#IfWinNotActive Medical Validation :   (Authorise By Episode)
!b::
WinActivate Medical Validation :   (Authorise By Episode)
#IfWinNotActive
#IfWinActive Medical Validation :   (Authorise By Episode)
!b::
   send, {altdown}<{altup}
sleep, 50
Return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                   Alt + x for Next
#IfWinNotActive Medical Validation :   (Authorise By Episode)
!x::
WinActivate Medical Validation :   (Authorise By Episode)
#IfWinNotActive

#IfWinActive Medical Validation :   (Authorise By Episode)
!x::
   send, {altdown}>{altup}
sleep, 50
Return
#IfWinActive

;                                                                                                                  Alt + I to invalidate a Potassium
#IfWinActive Result Entry
!i::
  send, {AltDown}d{AltUp}
  sleep, 150
  send, a
  sleep,100
  WinWaitActive, Assign Reason
  sleep, 80
  loop, 9
    {
    send i
    sleep, 70
    }
    
Return
#IfWinActive

;                                                                                                                  Alt + c to cancel a PHONC
#IfWinActive Result Entry
!c::
  send, {AltDown}d{AltUp}
  sleep, 150
  send, a
  sleep,100
  WinWaitActive, Assign Reason
  sleep, 80
  loop, 10
    {
    send c
    sleep, 70
    }
    
Return
#IfWinActive



;                                                                                              Mouse Wheel Actions in Verify   UP and DOWN SCROLL
#IfWinActive Medical Validation :   (Authorise By Episode)

WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-17 , winh-41, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-17 , 200, 1
Return
#IfWinActive

#IfWinActive Result Entry - Verify
WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , winh-26, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , 243, 1
Return
#IfWinActive

#IfWinActive Result Entry
WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , winh-26, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , 243, 1
Return
#IfWinActive

#IfWinActive Patient Audit - 
WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-22 , 581, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-22 , 80, 1
Return
#IfWinActive

#IfWinActive Visit Events - 
WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, 761 , 300, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, 761 , 61, 1
Return
#IfWinActive

#IfWinActive Test Set Maintenance - 
WheelDown::
mouseclick, left, 799 , 367
return

WheelUp::
mouseclick, left, 799 , 62
return
#IfWinActive

#IfWinActive Comment Text - 
WheelDown::
mouseclick, left, 547 , 272, 1
Return

WheelUp::
mouseclick, left, 547 , 65, 1
Return
#IfWinActive


;                                                        ALT-RIGHT CLICK ALT MENU: cobbled together somehow, by Samuel Murray - 26 May 2018, edited by Dieter (18/09/2019)
;;;;;;                                    for testing: Episode Number: SA01912457

!RButton::

MouseGetPos, xpos, ypos, window

Gui, New , , ChemHelp

Gui, Add, Button, gScript1, [&1] Open in Patient Entry
Gui, Add, Button, gScript2, [&2] Open in Result Entry
Gui, Add, Button, gScript3, [&3] Open in Specimen Info
Gui, Add, Button, gScript4, [&4] Open form on Equation
Gui, Add, Button, gExitApp, Reload

Gui, -Border
Gui, Show, x%xpos% y%ypos%
return

; Test Episode Number: SA01912457
Script1:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500

IfWinExist, Patient Entry
{
	WinActivate, Patient Entry
	sleep, 500
	send, {altdown}c{altup}
	sleep, 2000
	send, {enter}
	sleep, 2000
}

Else IfWinNotExist, Patient Entry
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS, 
	WinWaitActive,  NHLS, 
	Sleep, 100
	Send, pa{ENTER}
	WinWaitActive, Patient Entry  
}

sleep, 500
send, {CtrlDown}v{CtrlUp}
sleep, 200
send, {TAB}
return


Script2:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500

IfWinExist, Result Entry
{
	WinActivate, Result Entry
	sleep, 500
	send, {altdown}l{altup}
	sleep, 300
}

Else IfWinNotExist, Result Entry
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS, 
	WinWaitActive,  NHLS, 
	Sleep, 100
	Send, a{TAB}
	Send, Result Entry{Enter}
	WinWaitActive, Result Entry  
}

sleep, 500
send, {CtrlDown}v{CtrlUp}
sleep, 200
send, {Enter}
return


Script3:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500

IfWinExist, Specimen Information
{
	WinActivate, Specimen Information
	sleep, 500
	MouseClick, Left, 77, 82, 2, 100
	sleep, 500
}

Else IfWinNotExist, Specimen Information
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS, 
	WinWaitActive,  NHLS, 
	Sleep, 100
	Send, a{TAB}
	Send, Specimen Information{Enter}
	WinWaitActive, Specimen Information
}

sleep, 500
send, {CtrlDown}v{CtrlUp}
sleep, 200
send, {Tab}
return

Script4:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500
txt := Clipboard
sleep, 300
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return

ExitApp:
Reload
return




;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                       This section defines the different FUNCTIONS 
EpMedVal()
{
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 250
send, {altdown}i{altup}
sleep, 200
send, c
sleep, 100
WinActivate, Clinical History
WinWaitactive, Clinical History
sleep, 100
MouseClick, left, 122, 75, 2, 100
ClipSaved := clipboardall
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 100
txt := Clipboard
sleep, 200
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

EpResultSingle()
{
  WinActivate, Result Entry - Single -
WinWaitActive, Result Entry - Single -
sleep, 250
send, {altdown}3{altup}
sleep, 200
send, c
sleep, 250
WinActivate, Clinical History
WinWaitactive, Clinical History
sleep, 300
mouseclick, left, 122, 75, 2, 100
ClipSaved := clipboardall
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 200
txt := Clipboard
sleep, 200
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

EpResultVerSingle()
{
    WinActivate, Result Verify - Single - 
WinWaitActive, Result Verify - Single - 
sleep, 500
send, {altdown}3{altup}
sleep, 200
send, c
sleep, 200
WinActivate, Clinical History
WinWaitActive, Clinical History
sleep, 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved := clipboardall
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 200
txt := Clipboard
sleep, 200
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

MRNSingle()
{
    WinActivate, Result Entry - Single - 
WinWaitActive, Result Entry - Single - 
sleep, 500
send, {altdown}3{altup}
sleep, 500
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved := clipboardall
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 200
txt := Clipboard
sleep, 200
WinClose, Patient History
WinWaitClose, Patient History
sleep, 50
}


MRNMedVal()
{
    sleep, 500
send, {altdown}i{altup}
sleep, 500
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved := clipboardall
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 200
txt := Clipboard
sleep, 200
WinClose, Patient History
WinWaitClose, Patient History
sleep, 50
}

MRNResultVerSingle()
{
    sleep, 500
send, {altdown}3{altup}
sleep, 500
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved := clipboardall
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 200
txt := Clipboard
sleep, 200
WinClose, Patient History
WinWaitClose, Patient History
sleep, 50
}


;---------------------------------------------------------- Script to Loop Extraction by pre-defined Extracting Criteria with a pre-formatted configuration CSV.
#IfWinActive, Monthly Statistics (Not Responding) - \\Remote
!e::
settimer, extraction_watch, 10000
#IfWinActive

extraction_watch:
if WinExist("Monthly Statistics (Not Responding) - \\Remote")
{
    ToolTip, Waiting for Monthly Statistics Window to finish extraction
    sleep, 500
    ToolTip
    Return
}

if !WinExist("Monthly Statistics (Not Responding) - \\Remote")
{
    Extract()
    return
}
Return


ButtonEPRxlsx:
;--------------------------------Saving the patient's cumulative history in an Excel file: requires Node installed with Puppeteer module and the SavePatientEPR.js script.              

send, {AltDown}{Tab}{AltUp}
sleep, 200 

!+w::
sleep, 300
IfWinActive, Result Entry - Single - 
{
MRNSingle()
txt := Clipboard
run node %A_MyDocuments%\SavePatientEPR.js %txt%
if (ErrorLevel = "ERROR")
    MsgBox There was an error, likely Node is not installed.
else
return
}
    
else IfWinActive, Medical Validation :   (Authorise By Episode)
{
MRNMedVal()
txt := Clipboard
run node %A_MyDocuments%\SavePatientEPR.js %txt%
;Make sure Node.js is installed and that the SavePatientEPRs.js file, the node_modules folder, package.json and package-lock.json is present in "My Documents" folder.
if (ErrorLevel != 0)
    MsgBox There was an error, likely Node is not installed.
else
return
}

else IfWinActive, Result Verify - Single - 
{
MRNResultVerSingle()
txt := Clipboard
run node %A_MyDocuments%\SavePatientEPR.js %txt%
if (ErrorLevel = "ERROR")
    MsgBox There was an error, likely Node is not installed.
else
return
}

else
{
MsgBox, No TrakCare window active.  `nPlease ensure Result Entry or Result Verify windows are open and visible.
sleep, 200
return
}
Return


ButtonEx:
Extract()
Return

Extract()    {
FileAppend, time_started`,begin_date`,end_date`,loc_prov`,test_set_test_item`,email`,non_reportable`,sort_direction`n, extract_list.txt    
settimer, extraction_watch, Off
settimer, refreshtimer, Off
Loop, read, extract_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{
    LoopNumber := A_Index ;A_Index stores the number of the current look, i.e. in the case of 'loop, read' it will be the line number, as AHK reads files line by line.
    Looplinecontents := A_LoopReadLine ;This code stores the contents of the current line into the variable "Looplinecontents" or %Looplinecontents%.
    LineArray := StrSplit(Looplinecontents, ",")
    ;This portion configures which columns is present in the CSV
    begin_date := LineArray[1]
    end_date := LineArray[2]
    loc_prov := LineArray[3]
    test_set_test_item := LineArray[4]
    email := LineArray[5]
    non_reportable := LineArray[6]
    sort_direction := LineArray[7]
  sleep, 2000
   if !WinExist("Monthly Statistics") ;The exclamation mark before the WinExist specifies the negative of the statement, i.e. if win"doesnot"exist, execute the function below it in curly brackets.
    {
        MsgBox, Monthly Statistics Window does not exist.,4 ; Issue here needs to be resolved as sometimes it doesn't see the window, even though it's active.
        Return
    }

    if WinExist("Monthly Statistics (Not Responding)")
    {
        MsgBox, Monthly Statistics Window is "Not Responding" and likely busy extracting.  `nThe timer will now be set so that it can continue extraction when it starts responding.`nHit Esc to abort this process.
        while WinExist("Monthly Statistics (Not Responding)") {
            sleep, 10000
            ToolTip, Extracting...
            sleep, 500
            ToolTip,
        }
    }
  WinActivate, Monthly Statistics - \\Remote
  WinWaitActive, Monthly Statistics - \\Remote
  sleep, 2000
  MouseClick, Left, 18, 87
  sleep, 500
  send, results
  sleep, 500
  send, {Enter}
  sleep, 3000
  WinWaitActive, Report Listing - \\Remote
  sleep, 3000
  PixelGetColor, color, 178, 142
  While !(color = 0xFFFF00)
        {
        PixelGetColor, color, 178, 142
        ToolTip, Color is %color%
        sleep, 300
        Tooltip
        }
  sleep, 1000
  send, %begin_date%
  sleep, 1000
  send, {Tab}
  sleep, 500
  send, %end_date%
  sleep, 500
  send, {Tab}
  sleep, 500
  send, %loc_prov% ; Location code (province)
  sleep, 500
  send, {Down}
  sleep, 1000
  loop, 8
    {
        send, {Tab}
        sleep, 200
    }
  send, %test_set_test_item%
  sleep, 500
  loop, 9
    {
        send, {Tab}
        sleep, 200
    }
  send, {Space}
  sleep, 500
  send, {Tab}
  sleep, 500
  send, %email%
  sleep, 500
  send, {Tab}
  sleep, 500
  if (non_reportable = 1) {
    send, {space}
    sleep, 1000
    send, {Enter}
    sleep, 500
    send, {Tab}
    }
  else {
    send, {Tab}
       }
  sleep, 200
  send, {Tab}
  sleep, 500
  if (sort_direction = 1) {
    send, {Space}
    sleep, 500
    send, {Tab}
    sleep, 500
    }
  else {
    send, {Tab}
    sleep, 500
    send, {Space}
    sleep, 500
    }
   sleep, 500
   /*
   MsgBox, 4, , Please confirm that the correct parameters are entered. `nThis MsgBox will *auto-destruct* in 007 seconds and continue extracting.  Continue?, 7
        IfMsgBox No
            Return
        else IfMsgBox Timeout
            sleep, 500
  */
  WinActivate, Report Listing
  WinWaitActive, Report Listing
  sleep, 800
  MouseClick, Left, 808, 59 ; Click on "Print" button
  FileAppend, %A_Now%`,%begin_date%`,%end_date%`,%loc_prov%`,%test_set_test_item%`,%email%`,%non_reportable%`,%sort_direction%`n, extract_list.txt
  sleep, 10000
  settimer, refreshtimer, 600000 ; This is to keep TrakCare open
  While WinExist("Monthly Statistics (Not Responding)")
  {
    Sleep, 5000
    ToolTip, Extracting
  }
  settimer, refreshtimer, off ; This is to prevent the refreshtimer from interferring in the script.
  sleep, 5000
}
settimer, refreshtimer, 600000
  tooltip    ; remove the tooltip
}
Return
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                   SPE canned text 

::INFLAM::The alpha-1, alpha-2 and beta-2 (complement) regions are elevated. The gamma region measures _ g/L (8 - 14 g/L). `nNo monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::INFLAMP::The alpha-1, alpha-2 and beta-2 (complement) regions are elevated and there is a polyclonal hypergammaglobulinaemia at  g/L (8 - 14 g/L). `nNo monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::INFLAMS::The alpha-1, -2 and beta-2 (complement) region is elevated and the gamma region measures _ g/L (8 - 14 g/L). `nNo monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::PREV::The previously typed monoclonal Ig_ persists in _ gamma at _ g/L.  Immunoparesis is _.
::PROM::A prominent peak is present in the mid-gamma region measuring g/L. The remainder of the gamma region measures g/L (8-14 g/L).  Immunotyping will be performed. Please see results below, to follow.
::SURINE::Please send urine (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) for Bence Jones protein electrophoresis and serum for free light chain analysis.
::BGB::Beta-gamma bridging is present. This is consistent with a chronic inflammatory process associated with an IgA response.  Causes may include cirrhosis, and cutaneous or mucosal inflammation.
::SUSP::If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::REPEATSPE::Repeat SPE is recommended in 3-6 months or when the inflammatory condition has subsided.
::NORMP::Normal protein electrophoresis pattern.  No monoclonal peaks are visible. The gamma region measures _ g/L (8-14 g/L). `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::NEPHR::Hypoalbuminaemia is present.  The alpha-2 (macroglobulin) region is significantly increased at _ g/L (5-9 g/L).  The gamma region measures _ g/L (8-14 g/L). No monoclonal peaks are visible. `nThis pattern suggests nephrotic syndrome. If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended. 
::A-1::The alpha-1 peak is biphasic, suggesting alpha-1-antitrypsin heterozygosity.
::CSFELEC::
(Total protein concentration……………   g/L
Samples with high total protein concentrations >16.8 g/L will not be run due to the increased likelihood of false negative results
)
::CLINCONT::Clinician contact details may not be coded in our database.  Please go to the link below to update clinician contact details:`ntinyurl.com/nhls-update
::text1::
(
Any text
Newline preserved in this way
)
::USPCOMMENT::
(
~BOLD
URINE STEROID PROFILE:
Laboratory information
~NORMAL
SOPs used for analysis: SOP-SPT-003: Analysis of Anabolic Steroids
Samples were analysed according to standard SOPs.  There were no
deviations from the test methods.  Final results were reviewed and
approved by a Technical Signatory(ies).
Key to Coding: GC-MS/MS = Gas Chromatography/Tandem Mass Spectrometry
               LOQ = Limit of Quantification
Compound                               Method        Conc. ng/mL
5a-androstane-2a,17�-diol (5a-diol)    GC-MS/MS        
5�-androstane-2a,17�-diol (5�-diol)    GC-MS/MS        
Androsterone (Andro)                   GC-MS/MS        
Epitestosterone (E)                    GC-MS/MS        
Etiocholanolone (Etio)                 GC-MS/MS        
Testosterone (T)                       GC-MS/MS        
DHT (5a-Dihydrotestosterone)           GC-MS/MS        
~BOLD
Ratios
~NORMAL
T/E                                    GC-MS/MS        
5a-diol/5�-diol                        GC-MS/MS        
5�-diol/EpiT                           GC-MS/MS        
5a-diol/EpiT                           GC-MS/MS        
Andro/Etio                             GC-MS/MS        
Andro/T                                GC-MS/MS        
Andro/EpiT                             GC-MS/MS        
T/DHT                                  GC-MS/MS        
DHT/EpiT                               GC-MS/MS        

Comments:
The above steroid profile may not be suitable for diagnostic purposes, as some ratios could not be measured accurately as the concentrations were below the limit of quantification of the assay.  

The sample shows signs of extensive degradation as the 5a-androstanedione/Androsterone and/or 5�-androstanedione/Etiocholanolone ratio are >/= 0.1 in the sample. 

Signed
Director: JL du Preez
Test performed at a referral laboratory:
South African Doping Control Laboratory - SADoCoL
)

#IfWinNotActive, ahk_class XLMAIN ahk_exe EXCEL.EXE

::STN::Stable negative bias.
::STP::Stable positive bias.
::CTM::Close to mean.  Stable.

;If you have a long name like me, you can enter a shortcode like below, without the preceeding ";".  
;Type "d1" then space or tab and the text as coded will be typed automatically.

::D1::Dr Dieter van der Westhuizen
;::R1::Dr. Ronald Dalmacio
;::C1::Dr. Careen Hudson
;::H.::Dr. Heleen Vreede
;::J1::Dr. Jody Rusch
;::J2::Dr. Justine Cole
;::JL::Dr. Jarryd Lunn
;;::RA::Dr. Razia Banderker
;::TH::Dr. Thando Gcingca

#IfWinNotActive


;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                        Button Info 

Button_i:
run, http://nhls-results.co.za/chemhelp-readme/
return

;!RButton::
;MouseGetPos, xpos, ypos, window
;Gui, New , , ChemHelp
;Gui, Add, Button, gScript1, [&1] Open in Patient Entry
;Gui, Add, Button, gScript2, [&2] Open in Result Entry
;Gui, Add, Button, gScript3, [&3] Open in Specimen Info
;Gui, Add, Button, gScript4, [&4] Open form on Equation
;Gui, Add, Button, gExitApp, Reload
;
;Gui, -Border
;Gui, Show, x%xpos% y%ypos%
;return

;Script1:
;Gui, Hide
;sleep, 600
;send, {CtrlDown}c{CtrlUp}
;sleep, 500


;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                            Button More
ButtonMore:

Gui, New , , ChemHelp
Gui, Add, Edit, r1 vEpisode w135, EpisodeOrMRN
Gui, Add, Button, gMoreScript1, [&1] Open form on Equation
Gui, Add, Button, gMoreScript2, [&2] Open MRN in EPR view
Gui, Add, Button, gMoreScript3, [&3] Locate Episode Storage position
Gui, Add, Button, gMoreScript4, [&4] Open in Patient Entry 
Gui, Add, Button, gMoreScript5, [&5] Extract MRN to ExcelFile
Gui, Add, Button, gMoreScript6, [&6] SPE Comment
Gui, Add, Button, gMoreScript7, [&7] IFE Comment
Gui, Add, Button, gMoreScript8, [&8] Histo Comment
Gui, Add, Button, gMoreScript9, [&9] UOA Comment
Gui, Add, Button, x150 y265 w35 h20 gExitApp,  Close 
;Gui, -Border
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
ypos := A_ScreenHeight-370
xpos := A_ScreenWidth-290
Gui, Margin, 0, 0
Gui, Show, x%xpos% y%ypos% w190, ChemHelp - More
return

MoreScript1:
        Gui, Submit
		url := "http://172.22.4.40/multipdfsearch.php?file=" . Episode . ""
		run, %url%
		sleep, 1500
		WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
		Return
		
MoreScript2:
        Gui, Submit
		url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . Episode . ""
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
		Return
MoreScript3:
		Gui, Submit
		sleep, 100
        Specinfolookup(Episode)
		Return
MoreScript4:
        Gui, Submit
		sleep, 100
        OpenEpInPatientEntry(Episode)
		Return
MoreScript5:
        Gui, Submit
        run node %A_MyDocuments%\SavePatientEPR.js %Episode%
        if (ErrorLevel = "ERROR")
            MsgBox There was an error, likely Node is not installed.
        else
        Return
MoreScript6:
        Gui, Submit
		url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/eprajax.csp?chunk=" . Episode . "^C112^1^CF112^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript7:
        Gui, Submit
		url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/eprajax.csp?chunk=" . Episode . "^C113^1^CF113^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript8:
        Gui, Submit
		url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/eprajax.csp?chunk=" . Episode . "^A030^1^T^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
        Return
MoreScript9: 
        Gui, Submit
		url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/eprajax.csp?chunk=" . Episode . "^C440^1^C3200^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
ButtonClose2:
sleep, 500
send, {altdown}{F4 down}{F4 up}{altup}
sleep, 100
Return

;SA04972912
;MRN73839380
Specinfolookup(epis)
{
IfWinExist, Specimen Information
{
	WinActivate, Specimen Information
	sleep, 500
	MouseClick, Left, 77, 82, 2, 100
	sleep, 500
}
Else IfWinNotExist, Specimen Information
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS, 
	WinWaitActive,  NHLS, 
	Sleep, 100
	Send, a{TAB}
	Send, Specimen Information{Enter}
	WinWaitActive, Specimen Information
}
sleep, 500
send, %epis%
sleep, 200
send, {Tab}
sleep, 200
MouseClick, Left, 271, 271
send, {Altdown}d{AltUp}
return
}

OpenEpInPatientEntry(epis)
{
    IfWinExist, Result Entry - Single
    {
        MsgBox, Result Entry - Single window is already open.  It needs to be closed for this function to work.
        Return
    }
    Else IfWinNotExist, Result Entry - Single
    {
        IfWinNotActive,  NHLS, , WinActivate,  NHLS,
        WinWaitActive,  NHLS,
        Sleep, 200
        Send, {Home}
        sleep, 100
        Send, Result Entry
        sleep, 100
        Send, {Enter}
        WinWaitActive, Result Entry - \\Remote
        sleep, 200
        send, %epis%
        send, {Enter}
        Return
    }
}
/*
;-----------------------------   Logging each Authorize
#IfWinActive, Medical Validation :   (Authorise By Episode) - \\Remote
!a::
FileAppend, %A_now%`n, %A_MyDocuments%\chemhelp_log.txt
sleep, 100
If WinNotActive("Medical Validation :   (Authorise By Episode) - \\Remote")
{
    WinActivate, Medical Validation :   (Authorise By Episode) - \\Remote)
    send, {AltDown}
    sleep, 100
    send, a
    sleep, 100
    send, {AltUp}
    return
}
return
*/
;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                        Button Close    
ButtonClose:
WinActivate, ChemHelp
WinWaitActive, ChemHelp
ExitApp
sleep, 100
Return
;send, {altdown}{F4 down}{F4 up}{altup}
;sleep, 100
;Return

Escape::Reload
;ExitApp
Return

^!r::Reload  ; Assign Ctrl-Alt-R as a hotkey to restart the script.
