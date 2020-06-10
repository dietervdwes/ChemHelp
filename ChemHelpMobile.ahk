;ChemHelp Mobile v.1.0 - updated 06-06-2020
;Written by Dieter van der Westhuizen 2018-2020
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

Gui, Add, Button, x2 y-1 w250 h160 , <          Previous
Gui, Add, Button, x262 y-1 w-1772 h-600 , Button
Gui, Add, Button, x282 y-1 w240 h160 , Next           >
Gui, Add, Button, x22 y169 w200 h120 , PageUp
Gui, Add, Button, x22 y299 w200 h120 , PageDown
Gui, Add, Button, x262 y169 w250 h260 , Authorize
Gui, Add, Button, x262 y439 w250 h140 , Form Mobile Version
Gui, Add, Button, x262 y589 w123 h120 , EPR
Gui, Add, Button, x389 y589 w123 h120 , KeepOpen
Gui, Add, Button, x2 y429 w250 h80 , Verified
Gui, Add, Button, x2 y509 w250 h80 , FPSA
Gui, Add, Button, x129 y589 w123 h70 , Alt + F4 (Close Window)
Gui, Add, Button, x2 y589 w123 h70 , PHONC
Gui, Add, Button, x2 y659 w250 h70 , Alt + TAB

; Generated using SmartGUI Creator 4.0

;------------------------   Set Window Options
Gui, +AlwaysOnTop
Gui, Show, , ChemHelp Mobile
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
ypos := A_ScreenHeight-850
xpos := A_ScreenWidth-530
;Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%xpos% y%ypos%
sleep, 200
WinActivate, Medical Validation :   (Auth
sleep, 200
WinGetPos, xposloc, yposloc,,, Medical Validation :
newxposloc := xposloc + 50
newyposloc := yposloc + 20
;MsgBox, %xposloc%, %yposloc%,`n %newxposloc%, %newyposloc%
;MouseMove, %newxposloc%, %newyposloc%
sleep, 100
MouseClickDrag, Left, newxposloc , newyposloc , - 50, - 50, 100, R

return



;;;;;;;;;;;;;;;;;;;;;;                                                                                     Page Down when Alt + z is pressed.
ButtonPageDown:
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
   mouseclick, left, 433,  231, 1
   sleep, 100
   send, {PgDn down}{PgDn up}
Return

;;;;;;;;;;;;;;;;;;;;;;                                                                                                     Page Up when Alt + q is pressed.
ButtonPageUp:
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
   mouseclick, left, 433,  231, 1
   sleep, 100
   send, {PgUp down}{PgUp up}
Return


Button<Previous:
#IfWinNotActive, Medical Validation :   (Authorize By Episode)
WinActivate Medical Validation :   (Authorise By Episode)
   send, {altdown}<{altup}
sleep, 50
Return

;---------------------------------------------------------------------                                    Alt + x for Next
ButtonNext>:
#IfWinNotActive Medical Validation :   (Authorise By Episode)
WinActivate Medical Validation :   (Authorise By Episode)
   send, {altdown}>{altup}
sleep, 50
Return



/*
ButtonVerify:
WinActivate, Medical Validation :   (Authorise By Episode)
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
*/

ButtonAuthorize:
WinActivate, Medical Validation :   (Authorise By Episode)
;WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 50
send, {altdown}a{altup}
sleep, 100
return

ButtonFormMobileVersion:
sleep, 200
EpMedVal()
txt := ClipBoard
sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . txt . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-1060, 0, 545, A_ScreenHeight-25
Return

ButtonAlt+F4(CloseWindow):
sleep, 200
send, {AltDown}{Tab}
sleep, 100
send, {AltUp}
sleep, 100
send, {AltDown}{F4}{AltUp}

Return





Escape::Reload
;ExitApp
Return


GuiClose:
ExitApp

#Include, ChemHelp.ahk
