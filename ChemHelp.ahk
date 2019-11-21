;ChemHelp v.1.6 - updated 20-09-2019
;Written by Dieter van der Westhuizen 2018-2019
;Inspired from TrakHelper by Chad Centner

#SingleInstance, force
#NoTrayIcon
#NoEnv
;#Persistent
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 85
SendMode, Input
SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
SetMouseDelay, 0
SetWinDelay, 500


;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Add Buttons ;;;;;;;;;;;;;;;;;;
Gui, Add, button, x2 y2 w40 h20 ,Log-on
Gui, add, button,x2 y22 w30 h20  ,Form
Gui, Add, button, x2 y42 w28 h20  ,EPR
Gui, Add, button, x2 y64 w33 h20  ,FPSA
Gui, Add, button, x2 y86 w45 h20  ,Verified
Gui, Add, button, x2 y108 w57 h20  ,KeepOpen
;Gui, Add, button, w57 , VolVeri
Gui, Add, button, x2 y130 w45 h20  ,More
Gui, Add, button, x2 y152 w32 h20 ,Close
Gui, Add, button, x45 y152 w20 h20 ,_i
;Gui, Add, Button, x6 y17 w100 h30 , Ok
;Gui, Add, Button, x116 y17 w100 h30 , Cancel
;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Set Window Options   ;;;;;;;;;;;;;;
;Gui, +AlwaysOnTop
Gui, -sysmenu +AlwaysOnTop
Gui, Show, , ChemHelp1.6
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
height := A_ScreenHeight-270
width := A_ScreenWidth-85
Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%width% y%height% w80




titleoffice = Microsoft Office Activation Wizard
classoffice = NUIDialog
settimer, office_activation_watch, 150
return

;;;;;;;;;;;;;;;;;;;;;Generic script to close a pop-up window.
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


#c::
Run, Calc.exe
return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                      This is to LOG-ON with username and password.    

ButtonLog-on:
settimer, win_proxy_logon, 500
settimer, citrix_proxy_logon, 500
settimer, miscellaneous_user_config_file_notice, 500

url :="http://nhlslisapps02.nhls.ac.za/Citrix/XenApp/auth/login.aspx"
run, %url%
WinActivate, Citrix XenApp
WinWaitActive, Citrix XenApp
sleep, 2000
FileRead, citrix_username, %A_MyDocuments%\citrix_username.txt
FileRead, citrix_password, %A_MyDocuments%\citrix_password.txt
FileRead, trakcare_username, %A_MyDocuments%\trakcare_username.txt
FileRead, trakcare_password, %A_MyDocuments%\trakcare_password.txt
FileRead, tabs_citrix, %A_MyDocuments%\tabs_citrix.txt
;FileRead, trakcare_workarea_tabs, %A_MyDocuments%\trakcare_workarea_tabs.txt
send, %citrix_username%
sleep, 500
send, {TAB down}{TAB up}
sleep, 200
send, %citrix_password%
sleep, 200
send, {Return}
sleep, 3000

Loop %tabs_citrix%
{
    Send {Tab down}
	Send {Tab up}
    Sleep 80  ;      The number of milliseconds between keystrokes (or use SetKeyDelay).
}

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

/*
;;;;;;;;;;;;;;;;;;;;;;;                                                                         Episode button to open scanned form ; also works with Alt+E. 
#IfWinNotActive, Result Entry
!e::
#IfWinNotActive, 
ButtonEpisode:
InputBox, episodenumber, Episode, -sysmenu, , 100, 100, A_ScreenWidth-210 , A_ScreenHeight-275,
if ErrorLevel
   Return
else
url := "http://172.22.4.40/multipdfsearch.php?file=" . episodenumber
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
*/

;;;;;;;;;;;;;;;;;;;;;                                                                           CREF error C900 entered but test is registered on TrakCare   
!y::
EpResultSingle()
sleep, 100
FileAppend, CREF-error`n%txt%`n, C:\TrakCare-Errors\TrakCare-Entry-Errors.txt
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                         Button KeepOpen  
ButtonKeepOpen:
MsgBox, This script clicks the refresh button on "Results Entry - Verify" window every 10 minutes to keep TrakCare open.`nPlease make sure that "Results Entry - Verify" window is open when you leave your computer.`nClick OK to start.`nHit Esc when you are back on your computer.
SetTimer, refreshtimer, 600000 ;every 10 minutes
;SetTimer, refreshtimer, 300000 ;every 5 minutes
;SetTimer, refreshtimer, 10000 ;every 10 sec
Return

;This is the script to click refresh on Lab Results verify window. 
refreshtimer:
IfWinExist, Result Entry - Verify ahk_class Transparent Windows Client 
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
return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                           Alt+N Shortcut to copy Episode 
!n::
IfWinActive, Medical Validation :   (Authorise By Episode)
{
EpMedVal()
Return
}
else
    IfWinActive, Result Entry - Single -
    {
    EpResultSingle()
    Return
}


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                            Alt+P Enter PHONC
#IfWinActive, Medical Validation :   (Authorise By Episode)
!p::
EpMedVal()
sleep, 200
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
IfWinExist, Search Unsuccessful
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
}
    
    else IfWinActive, Medical Validation :   (Authorise By Episode)
    {
    MRNMedVal()
    }

else IfWinActive, Result Verify - Single - 
{
MRNResultVerSingle()
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
sleep, 200
send, FPSA
sleep, 200
send, {tab down}{tab up}
send, {enter}
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
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, f
sleep, 200
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

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                      Mouse Wheel Actions in Verify   UP and DOWN SCROLL
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

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;       RIGHT CLICK ALT MENU: cobbled together somehow, by Samuel Murray - 26 May 2018, edited by Dieter (18/09/2019)
;;;;;; for testing: Episode Number: SA01912457

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
sleep, 500
send, {altdown}i{altup}
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
sleep, 100
Return
}

EpResultSingle()
{
  WinActivate, Result Entry - Single -
WinWaitActive, Result Entry - Single -
sleep, 400
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
sleep, 100
Return  
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
sleep, 100
Return
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
sleep, 200
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
Return
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
sleep, 200
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
Return
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
sleep, 200
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
Return
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                   SPE canned text 

::INFLAM::The alpha-1, -2 and beta-2 (complement) regions are elevated and there is a polyclonal hypergammaglobulinaemia at  g/L (8 - 14 g/L). `nNo monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::INFLAMS::The alpha-1, -2 and beta-2 (complement) region is elevated and the gamma region measures _ g/L (8 - 14 g/L). `nNo monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::PREV::The previously typed monoclonal Ig_ persists in _ gamma at _ g/L.  Immunoparesis is _.
::PROM::A prominent peak is present in the mid-gamma region measuring g/L. The remainder of the gamma region measures g/L (8-14 g/L).  Immunotyping will be performed. Please see results below.
::SURINE::Please send urine (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) for Bence Jones protein electrophoresis.
::BGB::Beta-gamma bridging is present. This is consistent with a chronic inflammatory process associated with an IgA response.  Causes may include cirrhosis, and cutaneous or mucosal inflammation.
::SUSP::If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::REPEATSPE::Repeat SPE is recommended in 3-6 months or when the inflammatory condition has subsided.
::NORMP::Normal protein electrophoresis pattern.  No monoclonal peaks are present. The gamma region measures _ g/L (8-14 g/L). `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::NEPHR::Hypoalbuminaemia is present.  The alpha-2 (macroglobulin) region is significantly increased at _ g/L (5-9 g/L).  The gamma region measures _ g/L (8-14 g/L). No monoclonal peaks are visible. `nThis pattern suggests nephrotic syndrome. If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended. 
::A-1::The alpha-1 peak is biphasic, suggesting alpha-1-antitrypsin heterozygosity.
::text1::
(
Any text
Newline preserved in this way
)

#IfWinNotActive, ahk_class XLMAIN ahk_exe EXCEL.EXE

::STN::Stable negative bias.
::STP::Stable positive bias.
::CTM::Close to mean.  Stable.
::D1::Dr. Dieter van der Westhuizen
::R1::Dr. Ronald Dalmacio
::C1::Dr. Careen Hudson
::H.::Dr. Heleen Vreede
::J1::Dr. Jody Rusch
::J2::Dr. Justine Cole
::JL::Dr. Jarryd Lunn
;::RA::Dr. Razia Banderker
::TH::Dr. Thando Gcingca

#IfWinNotActive


;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                        Button Info 

Button_i:
run, http://nhls-results.co.za/chemhelp-readme/
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                            Button More
ButtonMore:
Gui, ChemHelpMore:New , -SysMenu , ChemHelp - More
Gui, Add, Edit, r1 vEpisode w135, Episode-or-MRN
Gui, Add, Edit, r1 vRefno w135, Lab Ref.
Gui, Add, button, x2 y64 h20  ,Form1
Gui, Add, button, x2 y86 h20  ,EPR_from_MRN
Gui, Add, button, x2 y108 h20  ,Open_Specimen_info
Gui, Add, button, x2 y130 h20  ,Enter_Private_Result
Gui, Add, button, x2 y152 h20 ,Close2
;Gui, Add, button, x45 y152 h20 ,_i_
;Gui, Add, Button, x6 y17 w100 h30 , Ok
;Gui, Add, Button, x116 y17 w100 h30 , Cancel
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
ypos := A_ScreenHeight-270
xpos := A_ScreenWidth-250
Gui, Margin, 0, 0
Gui, Show, x%xpos% y%ypos% w155, ChemHelp - More
Return

ButtonForm1:
        Gui, Submit
		url := "http://172.22.4.40/multipdfsearch.php?file=" . Episode . ""
		run, %url%
		sleep, 1500
		WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
		Return
		
ButtonEPR_from_MRN:
        Gui, Submit
		url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . Episode . ""
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
		Return
ButtonOpen_Specimen_info:
		Gui, Submit
		sleep, 200
        Specinfolookup()
		Return
ButtonEnter_Private_Result:
        Gui, Submit
		MsgBox, The episode number is: %Episode%
		Return
ButtonClose2:
sleep, 500
send, {altdown}{F4 down}{F4 up}{altup}
sleep, 100
Return

return

Specinfolookup()
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
send, %Episode%
sleep, 200
send, {Tab}
return
}

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

    

