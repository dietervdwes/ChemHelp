;ChemHelp v.1.4 - updated 06-05-2019
;Written by Dieter van der Westhuizen 2018
;Inspired form TrakHelper by Chad Centner

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


;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Add Buttons ;;;;;;;;;;;;;;;;;;
Gui, Add, button, w57 ,Log-on
Gui, add, button, w57 , Form
Gui, Add, button, w57 , EPR
Gui, Add, button, w57 , FPSA
Gui, Add, button, w57 , Verified
Gui, Add, button, w57 , KeepOpen
;Gui, Add, button, w57 , VolVeri
Gui, Add, button, w40, Close

;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Set Window Options   ;;;;;;;;;;;;;;
;Gui, +AlwaysOnTop
Gui, -sysmenu +AlwaysOnTop
Gui, Show, , ChemHelp1.4
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
height := A_ScreenHeight-275
width := A_ScreenWidth-85
Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%width% y%height%




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

;;;;;;;;;;;;;;;;;;;;;;;;;;Script to Resize Phoresis;;;;;;;;;;;;;;;;;;;;;

phoresis_resize:
IfWinActive,  PHORESIS REL. 8.6.3 - MINICAP N° 782
    WinMove,  PHORESIS REL. 8.6.3 - MINICAP N° 782,, A_ScreenWidth-300, 0, 545, A_ScreenHeight-25
sleep, 200
IfWinActive, Modify curves - Browsing 'Curve mosaic'
    WinMaximize, Modify curves - Browsing 'Curve mosaic'


;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Button Close    ;;;;;;;;;;;;;;;;;;;
ButtonClose:
WinActivate, ChemHelp
WinWaitActive, ChemHelp
sleep, 100
send, {altdown}{F4 down}{F4 up}{altup}
sleep, 100
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;   Button KeepOpen    ;;;;;;;;;;;;;;;;
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
;else
;  tooltip, Result Entry - Verify window not found. `n Hit ESC to exit.
;settimer, refreshtimer, 1000
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;   Alt+A activates Medical Validation Window ;;;;;;;;;;;
;#IfWinNotActive Medical Validation :   (Authorise By Episode)
;!a::
;WinActivate Medical Validation :   (Authorise By Episode)
;send {altdown}a{altup}
;sleep, 100
;Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Button EPR in Single    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
ButtonEPR:
sleep, 200
send, {AltDown}{Tab}{AltUp}
sleep, 200

IfWinActive, Single
{
sleep, 500
send, {altdown}3{altup}
sleep, 500
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 200
txt := Clipboard
sleep, 200
send, {altdown}c{altup}
sleep, 200
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
Return
}
else IfWinActive, Medical Validation :   (Authorise By Episode)
    {
        sleep, 200
        MsgBox EPR function not available in Medical Validation Window, yet.`nStill waiting for TrakCare programmers to fix  some things.
        sleep, 50
        Return
    }

else
{
    sleep, 200
    return
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Alt+W to open EPR in default web browser    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

!w::
sleep, 100
IfWinActive, Result Entry - Single - 
{
sleep, 500
send, {altdown}3{altup}
sleep, 500
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 200
txt := Clipboard
sleep, 200
send, {altdown}c{altup}
sleep, 200
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.location='%url%'`;</script></body></html>
sleep, 800
;WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25 
Return
}
    
    else IfWinActive, Medical Validation :   (Authorise By Episode)
    {
        /*
    MsgBox EPR function not available in Medical Validation Window, yet.`nStill waiting for TrakCare programmers to fix some things.
    Return
    
    sleep, 100
    MouseClick, Right, 579, 176
    sleep, 200
    MouseClick, Left, 98, 79
    sleep, 200
    Tooltip, Using Normal EPR until Trakcare gets fixed.
    return
    */
    sleep, 500
send, {altdown}i{altup}
sleep, 500
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 200
txt := Clipboard
sleep, 200
send, {altdown}c{altup}
sleep, 200
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.location='%url%'`;</script></body></html>
sleep, 800
;WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
;Winmove, 
Return
    }

else IfWinActive, Result Verify - Single - 
{
sleep, 500
send, {altdown}3{altup}
sleep, 500
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
mouseclick, left, 93, 75, 2, 100
sleep, 50
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 200
txt := Clipboard
sleep, 200
send, {altdown}c{altup}
sleep, 200
url := "http://trakdb-prod.nhls.ac.za:57772/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.location='%url%'`;</script></body></html>
sleep, 800
;WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-0, 0, 0, A_ScreenHeight-0
Return
}

else
{
return
}


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Shortcut to copy Episode     ;;;;;;;;;;;;;;;;;;;
!n::
IfWinActive, Medical Validation :   (Authorise By Episode)
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
mouseclick, left, 122, 75, 2, 100
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 200
txt := Clipboard
sleep, 200
send, {altdown}c{altup}
;txt := Clipboard
;ClipSaved:=clipboardall
;clipboard=
WinClose, Clinical History
sleep, 100
Return
}
else
    IfWinActive, Result Entry - Single -
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
mouseclick, left, 122, 75, 2, 100
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 200
txt := Clipboard
sleep, 200
send, {altdown}c{altup}
;txt := Clipboard
;ClipSaved:=clipboardall
;clipboard=
WinClose, Clinical History
sleep, 100
Return
}
;;;;;;;;;;;;;;;;;;;;;;;;;    Alt+F for Form from SINGLE or VALIDATION WINDOW   ;;;;;;;;;;;;;;;;

!F::
IfWinActive, Medical Validation :   (Authorise By Episode)
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
sleep, 200
mouseclick, left, 122, 75, 2, 100
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 400
txt := Clipboard
sleep, 300
send, {altdown}c{altup}
WinClose, Clinical History
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt
MyWidth = 545
MyHeight = A_ScreenHeight-25
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.resizeTo(%MyWidth%`,%MyHeight%)`;window.location='%url%'`;</script></body></html>
Return
}
else
IfWinActive,  Result Entry - Single - 
{
WinActivate, Result Entry - Single - 
WinWaitActive, Result Entry - Single - 
sleep, 500
send, {altdown}3{altup}
sleep, 200
send, c
sleep, 200
WinActivate, Clinical History
WinWaitActive, Clinical History
sleep, 100
mouseclick, left, 93, 75, 2, 100
sleep, 400
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 400
txt := Clipboard
sleep, 300
send, {altdown}c{altup}
WinClose, Clinical History
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt
MyWidth = 200
MyHeight = A_ScreenHeight-10
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.location='%url%'`;</script></body></html>
Return
}
else
IfWinActive, Result Verify - Single -
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
sleep, 200
mouseclick, left, 93, 75, 2, 100
sleep, 400
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 400
txt := Clipboard
sleep, 300
send, {altdown}c{altup}
WinClose, Clinical History
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt
MyWidth = 545
MyHeight = A_ScreenHeight-25
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.resizeTo(%MyWidth%`,%MyHeight%)`;window.location='%url%'`;</script></body></html>
Return
}
else
{
    return
}


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Button Form     ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
ButtonForm:
sleep, 100
IfWinExist, Medical Validation :   (Authorise By Episode)
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
sleep, 200
mouseclick, left, 122, 75, 2, 100
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 400
txt := Clipboard
sleep, 300
send, {altdown}c{altup}
WinClose, Clinical History
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt
MyWidth = 545
MyHeight = A_ScreenHeight-25
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.resizeTo(%MyWidth%`,%MyHeight%)`;window.location='%url%'`;</script></body></html>
Return
}
else
IfWinActive,  Result Entry - Single - 
{
WinActivate, Result Entry - Single - 
WinWaitActive, Result Entry - Single - 
sleep, 500
send, {altdown}3{altup}
sleep, 200
send, c
sleep, 200
WinActivate, Clinical History
WinWaitActive, Clinical History
sleep, 200
mouseclick, left, 93, 75, 2, 100
sleep, 400
ClipSaved:=clipboardall
clipboard=
send, {ctrldown}c{ctrlup}
sleep, 400
txt := Clipboard
sleep, 300
send, {altdown}c{altup}
WinClose, Clinical History
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt
MyWidth = 545
MyHeight = A_ScreenHeight-25
Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=data:text/html`,<html><body><script>window.location='%url%'`;</script></body></html>


/*
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
*/
Return
}
else
{
    return
}

iehalfscreen:
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;This is to LOG-ON with username and password. ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

ButtonLog-on:
settimer, win_proxy_logon, 500
settimer, citrix_proxy_logon, 500
settimer, miscellaneous_user_config_file_notice, 500

url :="http://nhlslisapps.nhls.ac.za/Citrix/XenApp/auth/login.aspx"
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


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; Button FPSA ;;;;;;;;;;;;;;;;;;
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

;;;;;;;;;;;;;;;;;;;;;;;;  Button Verified to insert staff note "Transcription Verified"
ButtonVerified:
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, f
sleep, 200
send, Transcription Verified.
sleep, 80
send, {alt down}o{alt up}
sleep, 50
Return

;;;;;;;;;;;;;;;;;;;  Alt + v for Staff note in Med Val or Single: Transcription Verified ;;;;;;;;;;;

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


;;;;;;;;;;;;;;;;;;;;;;  Button VolVeri
;ButtonVolVeri:
;WinActivate, Medical Validation :   (Authorise By Episode)
;WinWaitActive, Medical Validation :   (Authorise By Episode)
;sleep, 20
;send, {altdown}d{altup}
;sleep, 50
;send, f
;sleep, 200
;send, Volume verified as on form.
;sleep, 50
;send, {alt down}o{alt up}
;sleep, 50
;Return

;;;;;;;;;;;;;;;;;;;;;;    Page Down when Alt + z is pressed.
#IfWinActive Medical Validation :   (Authorise By Episode)
!z::
   mouseclick, left, 433,  231, 1
   sleep, 100
   send, {PgDn down}{PgDn up}
Return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;   Page Up when Alt + q is pressed.
#IfWinActive Medical Validation :   (Authorise By Episode)
!q::
   mouseclick, left, 433,  231, 1
   sleep, 100
   send, {PgUp down}{PgUp up}
Return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;   Activate first line when Alt + A is pressed
;#IfWinActive Medical Validation :   (Authorise By Episode)
;!a::
;   send, {alt down}a{alt up}
;   sleep, 3000
;   mouseclick, left, 433,  231, 1
;   sleep, 50
;   Return
;#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;  Alt + b for Back 
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

;;;;;;;;;;;;;;;;;;;;;;  Alt + x for Next ;;;;;;;;;;
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

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  Mouse Wheel Actions in Verify ;;;;;;;;;;;;;;;;
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
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; SPE canned text  ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

::INFLAM::The alpha-1 and -2 and beta-2 (complement) regions are elevated and there is a polyclonal hypergammaglobulinaemia at  g/L (8 - 14 g/L). No monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
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


::STN::Stable negative bias.
::STP::Stable positive bias.
::CTM::Close to mean.  Stable.
::D1::Dr. Dieter van der Westhuizen
::R1::Dr. Ronald Dalmacio
::C1::Dr. Careen Hudson
::H.::Dr. Heleen Vreede
::JO::Dr. Jody Rusch
::JU::Dr. Justine Cole
::JL::Dr. Jarryd Lunn
;::RA::Dr. Razia Banderker
::TH::Dr. Thando Gcingca



; Context sensitive help
; The hotkey below uses the clipboard to provide compatibility with the maximum
; number of editors (since ControlGet doesn't work with most advanced editors).
; It restores the original clipboard contents afterward, but as plain text,
; which seems better than nothing.

$^2::
; The following values are in effect only for the duration of this hotkey thread.
; Therefore, there is no need to change them back to their original values
; because that is done automatically when the thread ends:
SetWinDelay 10
SetKeyDelay 0
AutoTrim, On

if A_OSType = WIN32_WINDOWS  ; Windows 9x
    Sleep, 500  ; Give time for the user to release the key.

C_ClipboardPrev = %clipboard%
clipboard =
; Use the highlighted word if there is one (since sometimes the user might
; intentionally highlight something that isn't a command):
Send, ^c
ClipWait, 0.1
if ErrorLevel <> 0
{
    ; Get the entire line because editors treat cursor navigation keys differently:
    Send, {home}+{end}^c
    ClipWait, 0.2
    if ErrorLevel <> 0  ; Rare, so no error is reported.
    {
        clipboard = %C_ClipboardPrev%
        return
    }
}
C_Cmd = %clipboard%  ; This will trim leading and trailing tabs & spaces.
clipboard = %C_ClipboardPrev%  ; Restore the original clipboard for the user.
Loop, parse, C_Cmd, %A_Space%`,  ; The first space or comma is the end of the command.
{
    C_Cmd = %A_LoopField%
    break ; i.e. we only need one interation.
}
IfWinNotExist, AutoHotkey Help
{
    ; Determine AutoHotkey's location:
    RegRead, ahk_dir, HKEY_LOCAL_MACHINE, SOFTWARE\AutoHotkey, InstallDir
    if ErrorLevel  ; Not found, so look for it in some other common locations.
    {
        if A_AhkPath
            SplitPath, A_AhkPath,, ahk_dir
        else IfExist ..\..\AutoHotkey.chm
            ahk_dir = ..\..
        else IfExist %A_ProgramFiles%\AutoHotkey\AutoHotkey.chm
            ahk_dir = %A_ProgramFiles%\AutoHotkey
        else
        {
            MsgBox Could not find the AutoHotkey folder.
            return
        }
    }
    Run %ahk_dir%\AutoHotkey.chm
    WinWait AutoHotkey Help
}
; The above has set the "last found" window which we use below:
WinActivate
WinWaitActive
StringReplace, C_Cmd, C_Cmd, #, {#}
Send, !n{home}+{end}%C_Cmd%{enter}
return


Escape::Reload
;ExitApp
Return

^!r::Reload  ; Assign Ctrl-Alt-R as a hotkey to restart the script.

    

