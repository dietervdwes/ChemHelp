;~ ;ChemHelp v.2.0 - updated 2022-02-04
;Written by Dieter van der Westhuizen 2018-2022
;Inspired from TrakHelper by Chad Centner

/* Changes from previous version 1.11 added and to be incorporated into Github
1. added 2 columns to the extraction script: branch and user_site which should be updated on Excel template and uploaded.
2. Added Button AutoA (need to uncomment it to get it active - use with caution)
3. Increased FPSA test set maintenance opening wait time
4. Added mousewheel settings for Cumulative view window - useful when doing SPE's by "F3" function viewing previous results (on Tuesdays).
5. Added comment: ::IMMFIX::Immunofixation will be performed.  See result below, to follow.
6. Added comment: ::CTT::Close to target. Stable.
7. Added rSPE: to place SPE on hold, then refer the result to Chem Registrar VQ list so it is visible.
8. Added UPE and UIFE buttons in the more menu (lines)
9. Removed messagebox after dCDUM button pressed.
10. Changed EPR URL - removed port number.
11. Changed EPR URL - https for secure connection.

Changes in v2.0 (coded from 2022-01-11)
-Alt+B and Alt+V shortcuts works more efficiently in Result Entry - Single and Verify windows - Used the new If WinActive() - syntax
-Alt+Q and Alt+Z time delay shortened to 50ms from 100ms to page up / down faster.
-Alt+Q and Alt+Z added for Cumulative View - useful when viewing SPE histories. It has also been made active for non-TrakCare windows such as Chrome / Internet explorer.
-Tried to change mouse click position for login (adding to the variable lengths - unable to know if it will work)
-Added Equation-ip to the configuration file in ChemHelp-directory to make the form viewing compatible with other sites' Equation server IP. Your user site's Equation server IP should be the 6th item in chemhelp-config.txt file in ChemHelp-directory (comma separated from the other fields): 172.22.4.40 (GSH), 172.22.8.50 (RXH), 172.22.12.40 (TBH), 172.22.16.63 (?Green Point Lab)
-Changed the sleep (wait) times in the Alt+F and Alt+W functions - original sleep times are in comments at the EpMedVal() and MRNMedVal() function definitions.
-Added a way of logging the current episode number with Ctrl+Alt+N or Ctrl+Alt+L - it saves the episode number and the time in ChemHelp-directory\ChemHelp\episode_log.txt - may be useful for case logging.
-Added Window Capture button. Note you need to add a Screenshots folder in "ChemHelp-directory" - may be useful for case logging.
-Alt+W and Alt+F now works from within "Result Entry (Total Count)" and "Result Entry - Verify" window - no need to open the particular episode first - you can just select it by clicking and highlighting it.
-Keep-open button now doesn't cause the active window and mouse position to be lost after keeping TrakCare alive - thus more user-friendly.
-_i (info) link updated to point to Github.
-Directory structure changed: ChemHelp.ahk MUST now be in a separate folder along with its dependency files and folders else it will produce various errors.
-Configuration file structure changed. It now has fields for the Equation server IP and your own name (Hotstring d1)
-d1 Hotstring should now type your name (as in config file).
-Added 2 new hotkeys: Ctrl + Alt + F: Adds a Free-PSA and removes the CDUM.  Ctrl + Alt + D: Removes a CDUM.
-Changed 1 hotkey: Ctrl + Alt + W (from Shift + Alt + W) - to be more consistent with above new hotkeys.
-Removed shortcut for Shift+Alt+R - escape works better.
-Added shortcut to navigate to Results Entry - by using Ctrl+Alt+P - similar to Alt+P but without the "PHONC"-part.
-Added HYPOA and INFLAMOLIGO comments for SPE reporting.

To-do in v2.1
-https://www.autohotkey.com/board/topic/81915-solved-gui-control-tooltip-on-hover/ - to show hotkeys when hovering mouse over the app
-API which sends the size of VQ list over Telegram / Whatsapp (use tooltip scrape or an OCR program, such as command: C:\Program Files\Capture2Text_v4.6.0_64bit\Capture2Text>Capture2Text_CLI.exe --screen-rect "931 79 957 91" -o "C:\test.txt" --trim-capture --blacklist abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.! --scale-factor 2.5
-Ass uFix button & changed rFix to sFix.
-Do something cool with OCR using Capture-2-text's CLI.


*/


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


;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Add Buttons ;;;;;;;;;;;;;;;;;;
Gui, Add, button, x2 y2 w40 h20, Log-on
Gui, Add, button, x44 y2 w40 h20, Mobile
Gui, add, button,x2 y22 w30 h20, Form
;Gui, add, button,x34 y22 w35 h20, AutoA ;Hidden by default
Gui, add, button,x34 y22 w45 h20, Capture
Gui, Add, button, x2 y42 w28 h20, EPR
Gui, Add, button, x30 y42 w45 h20, EPRxlsx
Gui, Add, button, x2 y64 w33 h20, FPSA
Gui, Add, button, x36 y64 w45 h20, dCDUM
Gui, Add, button, x2 y86 w43 h20, Verified
Gui, Add, button, x46 y86 w35 h20, rFix
Gui, Add, button, x2 y108 w57 h20, KeepOpen
Gui, Add, button, x60 y108 w20 h20, Ex
Gui, Add, button, x2 y130 w36 h20, More
Gui, Add, button, x40 y130 w36 h20, VPN
;Gui, Add, button, x40 y130 w45 h20, ExMRN
Gui, Add, button, x2 y152 w32 h20, Close
Gui, Add, button, x40 y152 w20 h20, _i
Gui, Add, button, x60 y152 w20 h20, _t
;Gui, Add, Button, x6 y17 w100 h30 , Ok
;Gui, Add, Button, x116 y17 w100 h30 , Cancel
;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Set Window Options   ;;;;;;;;;;;;;;
;Gui, +AlwaysOnTop
Gui, -sysmenu +AlwaysOnTop
Gui, Show, , ChemHelp2.0
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
height := A_ScreenHeight-400
width := A_ScreenWidth-98
Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%width% y%height% w84



titleoffice = Microsoft Office Activation Wizard
classoffice = NUIDialog
settimer, office_activation_watch, 150
;The following part of initialisation reads a configuration file located at the user's "ChemHelp-directory", and saves the variables, 
;saved as comma separated values into the various variables below when the script is loaded.  
FileRead, chemhelpsettingsraw, %A_ScriptDir%\chemhelp_settings.txt
if ErrorLevel
    MsgBox, Failed to load the chemhelp_settings.txt from "ChemHelp-directory" folder. `nMake sure you place a file chemhelp_settings.txt formatted as the following with the commas: citrix.username,citrixpassword,trakcareusername,trakcarepassword,amnt-of-tabs-to-get-to-citrix-icon,equationserver-ip
else
settingsarray := StrSplit(chemhelpsettingsraw, ",")
citrix_username := settingsarray[1]
citrix_password := settingsarray[2]
trakcare_username := settingsarray[3]
trakcare_password := settingsarray[4]
full_name := settingsarray[5]
tabs_citrix := settingsarray[6]
equation_ip := settingsarray[7]
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

url :="https://nhlslisapps.nhls.ac.za/Citrix/XenApp/auth/login.aspx"
run, %url%
WinActivate, Citrix XenApp
WinWaitActive, Citrix XenApp
sleep, 2500
send, %citrix_username%
sleep, 500
send, {TAB down}{TAB up}
sleep, 200
send, %citrix_password%
sleep, 300
send, {Enter}
WinWaitActive, Citrix XenApp - Applications - 
WinWaitActive, Citrix XenApp - Applications - 
sleep, 3000
/*
Loop %tabs_citrix%
{
    Send {Tab down}
	Send {Tab up}
    Sleep 80  ;      The number of milliseconds between keystrokes (or use SetKeyDelay).
}
*/
ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big100.PNG
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for the TrakCare Icon. Possibly the files in the folder ChemHelp\trakcare_icons is missing in your "Documents" folder.
else if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big095.PNG
    if (ErrorLevel = 1)
        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big090.PNG
        if (ErrorLevel = 1)
            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big085.PNG
            if (ErrorLevel = 1)
                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big080.PNG
                if (ErrorLevel = 1)
                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big105.PNG
                    if (ErrorLevel = 1)
                        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big110.PNG
                        if (ErrorLevel = 1)
                            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big115.PNG
                            if (ErrorLevel = 1)
                                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big120.PNG
                                if (ErrorLevel = 1)
                                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big125.PNG
if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small100.PNG
if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small095.PNG
    if (ErrorLevel = 1)
        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small090.PNG
        if (ErrorLevel = 1)
            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small085.PNG
            if (ErrorLevel = 1)
                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small080.PNG
                if (ErrorLevel = 1)
                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small105.PNG
                    if (ErrorLevel = 1)
                        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small110.PNG
                        if (ErrorLevel = 1)
                            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small115.PNG
                            if (ErrorLevel = 1)
                                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small120.PNG
                                if (ErrorLevel = 1)
                                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small125.PNG
else if (ErrorLevel = 1)
MsgBox TrakCare Icon could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x`;%FoundY%y. `nIt will now try clicking on it.
    FoundX += 20
    FoundY += 20
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
;SetTimer, refreshtimer, 10000 ;every 10 sec
;MsgBox, 1, , Timer has been set.`nThis Message will disappear automatically in 2 seconds, 2
Return

;This is the script to click refresh on Lab Results verify window. 
refreshtimer:
if WinExist("Result Entry - Verify ahk_class Transparent Windows Client")  
{
  WinGetTitle, last_window_title, A
  MouseGetPos, xmousepos, ymousepos 
  WinActivate, Result Entry - Verify ahk_class Transparent Windows Client
  WinWaitActive, Result Entry - Verify ahk_class Transparent Windows Client
  sleep, 100
  WinGetPos,,,winw, winh
  mouseclick, left, winw-60 , 125, 1
  sleep, 50
    WinActivate, %last_window_title%
    MouseMove, %xmousepos%, %ymousepos%
  sleep, 150
  WinMinimize, Result Entry - Verify ahk_class Transparent Windows Client
  sleep, 500
  ToolTip, Trak is being kept open automatically. `n Hit ESC to stop.  ;This tooltip doesn't want to work well for some reason. I only get it working with a while loop.
  SetTimer, RemoveToolTip, -1000
  sleep, 1000
  return
}
else if WinExist(" NHLS")
{
    If !WinActive(" NHLS")
    WinGetTitle, last_window_title, A
    MouseGetPos, xmousepos, ymousepos 
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
    sleep, 50
    WinActivate, %last_window_title%
    MouseMove, %xmousepos%, %ymousepos%
    sleep, 100
    WinMinimize,  NHLS
    sleep, 500
  ToolTip, Trak is being kept open automatically. `n Hit ESC to stop.   ;This tooltip doesn't want to work well for some reason
  SetTimer, RemoveToolTip, -1000
  sleep, 1000
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
else
    if WinActive("Result Verify - Single -")
    {
    EpResultVerSingle()
    Return
    }
else
    if WinActive("Result Entry - Verify")
    {
        EpResEntrVer()
        Return
    }
else
    if WinActive("Result Entry (Total Count")
    {
        EpResEntrTot()
        Return
    }
else
    MsgBox, No relevant TrakCare Window active.
    Return

^!n:: ;                                                                                 Ctrl+Alt+N or Ctrl+Alt+L to log episode in a text file - useful for logging of episodes.
^!l::
log_episode:
if WinActive("Medical Validation :   (Authorise By Episode)")
    {
    EpMedVal()
    LogEpisode()
    }
else
    if WinActive("Result Entry - Single -")
    {
    EpResultSingle()
    LogEpisode()
    }
else
    if WinActive("Result Entry - Verify")
    {
    EpResEntrVer()
    LogEpisode()
    }
else
    if WinActive("Result Verify - Single -")
    {
    EpResultVerSingle()
    LogEpisode()
    }
else
    if WinActive("Result Entry (Total")
    {
    EpResEntrTot()
    LogEpisode()
    }
else
    MsgBox,1,, No relevant TrakCare Window active.,3
    Return

LogEpisode()
{
    FormatTime, timevar , YYYYMMDDHH24MISS, yyyy-MM-dd HH:mm
    FileAppend, %ClipBoard%`,%timevar%`n, %A_ScriptDir%\episode_log.txt
    if ErrorLevel 
        MsgBox, Unable to write to file Documents\ChemHelp\episode_log.txt.  Make sure the directory exists in the "My Documents" folder.
    else
        MsgBox, 1,,Episode %ClipBoard% logged at %timevar% `n at ChemHelp-directory\episode_log.txt,1.8
}    
    
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
sleep, 200 ; changed down from 300
send, {Enter}
sleep, 150 ; changed down from 300
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
sleep, 100
send, {tab}
SendInput, %full_name%
send, {ShiftDown}{tab}{ShiftUp}

return
#IfWinActive

;-------------------------                                                                  Ctrl + Alt + P Enter PHONC

#IfWinActive, Medical Validation :   (Authorise By Episode)
^!p::
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
sleep, 200 ; changed down from 300
send, {Enter}
sleep, 150 ; changed down from 300
MouseClick, Left, 340, 133, 2, 100
sleep, 100

return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                            Ctrl+Alt+R Enter CCOM with Dr. update links


#IfWinActive, Medical Validation :   (Authorise By Episode)
^!r::
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

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                 ALT+W for EPR

ButtonEPR:
send, {AltDown}{Tab}{AltUp}
sleep, 200 

!w::
sleep, 300
IfWinActive, Result Entry - Single - 
{
MRNSingle()
txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
return
}
    
else IfWinActive, Medical Validation :   (Authorise By Episode)
{
MRNMedVal()
txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
return
}

else IfWinActive, Result Verify - Single - 
{
MRNResultVerSingle()
txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
return
}

else IfWinActive, Result Entry (Total
{
    MRNResultEntTotal()
    txt := Clipboard
    url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
    run, %url%
    sleep, 800
    WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
    return
}

else IfWinActive, Result Entry - Verify
{
    MRNResultEntVer()
    txt := Clipboard
    url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
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

;                                                         Alt+F for Form
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
url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
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
url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
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
url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
}
else
IfWinActive, Result Entry - Verify
    {
        EpResEntrVer()
        txt := ClipBoard
        url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
        run, %url%
        sleep, 1500
        WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
        Return
    }
else
IfWinActive, Result Entry (Total Count
    {
        EpResEntrTot()
        txt := ClipBoard
        url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
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
^!f::
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 1000
WinActivate, Test Set Maintenance
WinWaitActive, Test Set Maintenance
sleep, 500
ImageSearch, FoundX, FoundY, 15,86, 523, 379, %A_ScriptDir%\trakcare_icons\cdum.png
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for CDUM. Possibly the file cdum.png is missing in ChemHelp\trakcare_icons folder.
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
sleep, 200
send, {AltDown}o{AltUp}
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                          Button dCDUM
ButtondCDUM:
^!d::
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 200
WinWaitActive, Test Set Maintenance
sleep, 800
ImageSearch, FoundX, FoundY, 15,86, 523, 379, %A_ScriptDir%\trakcare_icons\cdum.png
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for CDUM. `nPossibly the file cdum.png is missing in ChemHelp\trakcare_icons folder.`n If this is the issue, find it on github.com/dietervdwes/chemhelp
else if (ErrorLevel = 1)
    MsgBox CDUM could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x%FoundY%.
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
    send, {AltDown}d{AltUp}
sleep, 300
WinClose, Test Set Maintenance
;MsgBox, Please confirm that CDUM2 has been removed.
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                          Button rFix
ButtonrFix:
send, {AltDown}{Escape}{AltUp}
#IfWinActive Result Entry - Single -
    send, {AltDown}u{AltUp}
    sleep, 1200
    send, {AltDown}2{AltUp}
    sleep, 300
    send, s
    sleep, 300
    WinActivate, Test Set Maintenance
    sleep, 100
    MouseClick, Left, 830, 308 
    sleep, 100
    send, IFE
    sleep, 200
    send, {tab down}{tab up}
    send, {enter}
    sleep, 300
    WinClose, Test Set Maintenance
    sleep, 800
    send, {Alt}
    sleep, 400
    send, 2
    sleep, 400
    send, n
    sleep, 500
    send, {AltDown}1{AltUp}
    sleep, 300
    send, f
    sleep, 300
    WinWaitActive, Refer To Verification Queue
    WinActivate, Refer To Verification Queue
    sleep, 300
    MouseClick, Left, 429, 111
    sleep, 500
    MouseClick, Left, 429, 165
    sleep, 500
    send, {AltDown}o{AltUp}
    sleep, 500
    WinClose, Result Entry - Single
    return
#IfWinActive
MsgBox, Result Entry - Single window is not active.
return
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                               Button Verified to insert staff note "Transcription Verified"  
ButtonVerified:
sleep, 200
send, {AltDown}{Escape}{AltUp}
sleep, 200
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                Alt+V for Staff note in Med Val or Single: Transcription Verified 

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
send, {altdown}o{altup}
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
send, {altdown}o{altup}
sleep, 50
Return
}
else
{
return
}

;;;;;;;;;;;;;;;;;;;                                                                             Alt+/ for Staff note in Med Val or Single: Transcription Verified 
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

;;;;;;;;;;;;;;;;;;;;;;                                                                                     Page Down when Alt+z is pressed.
#IfWinActive Medical Validation :   (Authorise By Episode)
!z::
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgDn down}{PgDn up}
Return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;                                                                                                     Page Up when Alt+q is pressed.
/*
#IfWinActive Medical Validation :   (Authorise By Episode)
!q::
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
Return
#IfWinActive
*/
;;
!q::
   If WinActive("Result Entry - Single - ")
   {
   mouseclick, left, 78,  209, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
   Return
   }
   If WinActive("Cumulative View - ")
   {
    mouseclick, left, 78,  209, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
   sleep, 150
   Return
   }
   Else
   {
   send, {PgUp}
    Return
   }
!z::
   If WinActive("Result Entry - Single - ")
   {
   mouseclick, left, 78,  209, 1
   sleep, 50
   send, {PgDn down}{PgDn up}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgDn down}{PgDn up}
   Return
   }
   If WinActive("Cumulative View - ")
   {
    mouseclick, left, 78,  209, 1
   sleep, 50
   send, {PgDn down}{PgDn up}
   sleep, 150
   Return
   }
   Else
   {
   send, {PgDn}
    Return
   }


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                   Alt+B for Back 
!b::
   If WinActive("Result Entry - Single - ")
   {
   send, {altdown}<{altup}
   sleep, 80
   send, {Enter}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   send, {altdown}<{altup}
   sleep, 500
   Return
   }
   Else
   {
   MsgBox, No TrakCare Window selected.
    Return
   }


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                   Alt+X for Next
!x::
   If WinActive("Result Entry - Single - ")
   {
   send, {altdown}>{altup}
   sleep, 80
   send, {Enter}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   send, {altdown}>{altup}
   sleep, 500
   Return
   }
   Else
   {
   MsgBox, No TrakCare Window selected.
    Return
   }

;                                                                                                                  Alt+I to invalidate a Potassium
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

;                                                                                                                  Alt+C to cancel a PHONC
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
MouseClick, left, 765 , 305, 1 
sleep, 50
Return

WheelUp::
mouseclick, left, 765 , 65, 1
sleep, 50
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

#IfWinActive Cumulative View - 
WheelDown::
mouseclick, left, 993 , 670, 1
Return

WheelUp::
mouseclick, left, 993 , 208, 1
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
url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
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
sleep, 100 ; changed this from 250 (2022-01-12)
send, {altdown}i{altup}
sleep, 100 ; changed this down from 200
send, c
sleep, 50 ;changed this down from 100
WinActivate, Clinical History
WinWaitactive, Clinical History
sleep, 100 ; changed this down from 100
MouseClick, left, 122, 75, 2, 100
clipboard=
sleep, 100 ;changed down from 100
send, ^c
ClipWait
sleep, 100
txt := Clipboard
sleep, 100 ; changed this down from 100
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}


EpResEntrTot()
{
WinActivate, Result Entry (Total
WinWaitActive, Result Entry (Total
sleep, 100 ; changed this from 250 (2022-01-12)
send, {altdown}i{altup}
sleep, 100 ; changed this down from 200
send, c{Enter} 
sleep, 50 ;changed this down from 100
WinActivate, Clinical History
WinWaitactive, Clinical History
sleep, 100 ; changed this down from 100
MouseClick, left, 122, 75, 2, 100
clipboard=
sleep, 100 ;changed down from 100
send, ^c
ClipWait
sleep, 100
txt := Clipboard
sleep, 100 ; changed this down from 100
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

EpResEntrVer()
{
WinActivate, Result Entry - Verify
WinWaitActive, Result Entry - Verify
sleep, 100 ; changed this from 250 (2022-01-12)
send, {altdown}i{altup}
sleep, 100 ; changed this down from 200
send, c{Enter}
sleep, 50 ;changed this down from 100
WinActivate, Clinical History
WinWaitactive, Clinical History
sleep, 100 ; changed this down from 100
MouseClick, left, 122, 75, 2, 100
clipboard=
sleep, 100 ;changed down from 100
send, ^c
ClipWait
sleep, 100
txt := Clipboard
sleep, 100 ; changed this down from 100
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

EpResultSingle()
{
  WinActivate, Result Entry - Single -
WinWaitActive, Result Entry - Single -
sleep, 200
send, {AltDown}3{AltUp}
sleep, 500
;send, 3
sleep, 300
send, c
sleep, 100
WinActivate, Clinical History
WinWaitactive, Clinical History
sleep, 50
mouseclick, left, 122, 75, 2, 100
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 100
txt := Clipboard
sleep, 50
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

EpResultVerSingle()
{
    WinActivate, Result Verify - Single - 
WinWaitActive, Result Verify - Single - 
sleep, 50
send, {AltDown}3{AltUp}
sleep, 300
send, c
sleep, 100
WinActivate, Clinical History
WinWaitActive, Clinical History
sleep, 50
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100
send, ^c
ClipWait
sleep, 100
txt := Clipboard
sleep, 50
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

MRNSingle()
{
    WinActivate, Result Entry - Single - 
WinWaitActive, Result Entry - Single - 
sleep, 500 ; changed from 500
send, {AltDown}3
sleep, 300
send, {AltUp}
sleep, 200 ; changed from 500
send, {enter}
sleep, 100 ; changed from 500
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 100 ; changed from 200
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100 ; changed from 100
send, ^c
ClipWait
sleep, 100 ; changed from 200
txt := Clipboard
sleep, 100 ; changed from 200
WinClose, Patient History
WinWaitClose, Patient History
sleep, 50
}


MRNMedVal()
{
    sleep, 200 ;changed down from 500
send, {altdown}i{altup}
sleep, 200 ; changed down from 500
send, {enter}
sleep, 50 ; changed down from 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 50 ; changed down from 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100 ; changed down from 100
send, ^c
ClipWait
sleep, 50 ; changed down from 200
txt := Clipboard
sleep, 50 ; changed down from 200
WinClose, Patient History
WinWaitClose, Patient History
sleep, 50
}

MRNResultVerSingle()
{
    sleep, 500
send, {AltDown}
sleep, 100
send, 3{AltUp}
sleep, 300
send, {enter}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
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

MRNResultEntTotal()
{
sleep, 200 ;changed down from 500
send, {altdown}i{altup}
sleep, 200 ; changed down from 500
send, {enter}
sleep, 50 ; changed down from 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 50 ; changed down from 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100 ; changed down from 100
send, ^c
ClipWait
sleep, 50 ; changed down from 200
txt := Clipboard
sleep, 50 ; changed down from 200
WinClose, Patient History
WinWaitClose, Patient History
sleep, 50
}

MRNResultEntVer()
{
sleep, 200 ;changed down from 500
send, {altdown}i{altup}
sleep, 200 ; changed down from 500
send, {enter}
sleep, 50 ; changed down from 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 50 ; changed down from 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100 ; changed down from 100
send, ^c
ClipWait
sleep, 50 ; changed down from 200
txt := Clipboard
sleep, 50 ; changed down from 200
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
run node %A_ScriptDir%\SavePatientEPR.js %txt%
if (ErrorLevel = "ERROR")
    MsgBox There was an error, likely Node is not installed.
else
return
}
    
else IfWinActive, Medical Validation :   (Authorise By Episode)
{
MRNMedVal()
txt := Clipboard
run node %A_ScriptDir%\SavePatientEPR.js %txt%
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
run node %A_ScriptDir%\SavePatientEPR.js %txt%
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
Loop, read, %A_ScriptDir%\extract_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{
    LoopNumber := A_Index ;A_Index stores the number of the current look, i.e. in the case of 'loop, read' it will be the line number, as AHK reads files line by line.
    Looplinecontents := A_LoopReadLine ;This code stores the contents of the current line into the variable "Looplinecontents" or %Looplinecontents%.
    LineArray := StrSplit(Looplinecontents, ",")
    ;This portion configures which columns is present in the CSV
    begin_date := LineArray[1]
    end_date := LineArray[2]
    loc_prov := LineArray[3]
    branch := LineArray[4]
    user_site := LineArray[5]
    test_set_test_item := LineArray[6]
    email := LineArray[7]
    non_reportable := LineArray[8]
    sort_direction := LineArray[9]
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
  send, {Tab}
  sleep, 500
  send, %branch%
  ;send, {Down} ;Uncomment this if branch is present
  sleep, 1000
  send, {Tab}
  sleep, 500
  send, %user_site%
  ;send, {Down}  ;Uncomment this if province is present
  sleep, 1000
  ;send, {Tab} ;Uncomment this tab if above user site is present.
  sleep, 500
  loop, 6
    {
        send, {Tab}
        sleep, 200
    }
  send, %test_set_test_item%
  sleep, 200
  send, {Tab}
  sleep, 500
  #IfWinActive, CLNSTATLIST
  sleep, 1000
  send, {Enter}
  WinWaitClose, CLNSTATLIST
  #IfWinActive
  sleep, 1000
  MouseClick, Left, 470, 533
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
  FileAppend, %A_Now%`,%begin_date%`,%end_date%`,%loc_prov%`,%test_set_test_item%`,%email%`,%non_reportable%`,%sort_direction%`n, %A_ScriptDir%\extract_list.txt
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

ButtonAutoA: ;Feature hidden by default. Use this with extreme caution - it will auto-authorize samples. Useful for a batch processing but shouldn't be used if you don't know what you're doing.
MsgBox, Do not use this feature if you don't know what you're doing.  Hit ESC immediately if you don't know what this feature is.
WinActivate, Medical Validation
WinWaitActive, Medical Validation
loop, 100
    {
        Send, {AltDown}a{AltUp}
        sleep, 800
        PixelGetColor, color, 105, 229
        While !(color = 0xD1B499)
                {
                PixelGetColor, color, 105, 229
                ToolTip, Color is %color%
                sleep, 200
                Tooltip
                }
        sleep, 1500
    }
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                   SPE canned text 

::INFLAM::The alpha-1, alpha-2 and beta-2 (complement) regions are elevated. The gamma region measures  g/L (8 - 14 g/L). `n`No monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::INFLAMP::The alpha-1, alpha-2 and beta-2 (complement) regions are elevated and there is a polyclonal hypergammaglobulinaemia at  g/L (8 - 14 g/L). `nNo monoclonal peaks are visible. This pattern suggests a chronic inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::INFLAMS::The alpha-1, -2 and beta-2 (complement) region is elevated and the gamma region measures  g/L (8 - 14 g/L). `nNo monoclonal peaks are visible. This pattern suggests an inflammatory process.  `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::INFLAMOLIGO::Hypoalbuminemia is present.`nThe alpha-1, alpha-2 and beta-2 (complement) regions are elevated and there is hypergammaglobulinemia of _ g/L (8-14 g/L).`nMultiple irregularities (all <1 g/L) are present in the gamma region on the background polyclonal hypergammaglobulinemia.  This pattern suggests a chronic inflammatory process with an oligoclonal response.`nSerum free light chain analysis is recommended, if clinically indicated.`nRepeat serum protein electrophoresis is recommended in 3-6 months or when the inflammatory condition has subsided.
::HYPOA::Hypoalbuminemia is present.
::PREV::The previously typed monoclonal Ig_ persists in _ gamma at _ g/L.  Immunoparesis is _.
::PROM::A prominent peak is present in the mid-gamma region measuring g/L. The remainder of the gamma region measures g/L (8-14 g/L).  Immunotyping will be performed. Please see results below, to follow.
::SURINE::Please send urine (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) for Bence Jones protein electrophoresis and serum for free light chain analysis.
::BGB::Beta-gamma bridging is present. This is consistent with a chronic inflammatory process associated with an IgA response.  Causes may include cirrhosis, and cutaneous or mucosal inflammation.
::SUSP::If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::SUSPPEAK::A suspicious peak is present in the _-gamma region measuring _ g/L. `nThe rest of the gamma region measures _ g/L (8-14 g/L).`nImmunofixation will be performed.  See result below, to follow. 
::REPEATSPE::Repeat serum protein electrophoresis is recommended in 3-6 months or when the inflammatory condition has subsided.
::NORMP::Normal protein electrophoresis pattern.  No monoclonal peaks are visible. The gamma region measures _ g/L (8-14 g/L). `nIf the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::NEPHR::Hypoalbuminaemia is present.  The alpha-2 (macroglobulin) region is significantly increased at _ g/L (5-9 g/L).  The gamma region measures _ g/L (8-14 g/L). No monoclonal peaks are visible. `nThis pattern suggests nephrotic syndrome. If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended. 
::A-1::The alpha-1 peak is biphasic, suggesting alpha-1-antitrypsin heterozygosity.
::IMMFIX::Immunofixation will be performed.  See result below, to follow.
::HYPOG::Hypogammaglobulinemia is present at _ g/L (8-14 g/L).`nImmunofixation will be performed. See results below, to follow.
::ARRCOM::Unable to calculate the aldosterone:renin ratio due to the upper / lower measuring limit of aldosterone / renin, but the ratio is </> _.
::CSFELEC::
(
CSF electrophoresis:
Samples with high total protein concentrations (>16.8 g/L) will not be run due to the increased likelihood of false negative results.
)
::POLYG::Polyclonal hypergammaglobulinemia is present at
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

; Work in progress
;sending a hotstring with a variable https://www.autohotkey.com/board/topic/95718-hotstrings-that-have-a-variable-in-the-command/
;These are the default endchars for a hotkey, pressing any of these keys will trigger the hotstring
;You can add or remove characters as you need
endchars = -()[]{}':;/\,.?`n `t!

;'*B0 = do not auto-backspace, and do not wait for a hotstring endchar to run this hostring
:*B0:hotstring::
;::hotstring::
  ;Wait for additional text
  Input, modifier, V, %endchars%
  ;Extract the endkey pressed from ErrorLevel
  RegexMatch(ErrorLevel, "EndKey:\K.*", EndKey)
  ;Send backspaces to clear out the hotstring text
  SendInput % "{bs " strlen(modifier)+strlen(A_ThisHotkey)-4 "}"

  ;Put the text you want to send here
  ;"modifier" contains the text entered after the hotstring
  ;"EndKey" contains the key that terminated the hotkey
  ;Place EndKey between brackets { } as it is not always a literal character, for example, EndKey might be the text "Space"
  SendInput You pressed these keys: %modifier% after the hotstring{%EndKey%}
return
 

#IfWinNotActive, ahk_class XLMAIN ahk_exe EXCEL.EXE

::STN::Stable negative bias.
::STP::Stable positive bias.
::CTM::Close to mean.  Stable.
::CTT::Close to target. Stable.
::FreeT4comment::Unfortunately we were unable to get hold of the attending clinician for the free-T4 result which has changed after re-analysis.

;If you have a long name like me, you can enter a shortcode like below, without the preceeding ";".  
;Type "d1" then space or tab and the text as coded will be typed automatically.

::D1::
SendInput, %full_name%
return
#IfWinNotActive


;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                        Button Info 

Button_i:
run, https://github.com/dietervdwes/chemhelp/
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
Gui, Add, Button, gMoreScript01, [&01] Open form on Equation
Gui, Add, Button, gMoreScript02, [&02] Open MRN in EPR view
Gui, Add, Button, gMoreScript03, [&03] Locate Episode Storage position
Gui, Add, Button, gMoreScript04, [&04] Open in Patient Entry 
Gui, Add, Button, gMoreScript05, [&05] Extract MRN to ExcelFile
Gui, Add, Button, gMoreScript06, [&06] SPE Comment
Gui, Add, Button, gMoreScript07, [&07] IFE Comment
Gui, Add, Button, gMoreScript08, [&08] Histo Comment
Gui, Add, Button, gMoreScript09, [&09] UOA Comment
Gui, Add, Button, glog_episode, [&12] Log Episode ;see Ctrl+Alt+N or Ctrl+Alt+L ~line 450
Gui, Add, Button, x106 y178 w100 gMoreScript10, [&10] UPE Comment
Gui, Add, Button, x102 y207 w100 gMoreScript11, [&11] UIFE Comment


Gui, Add, Button, x150 y265 w45 h20 gExitApp,  Restart 
;Gui, -Border
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
ypos := A_ScreenHeight-500
xpos := A_ScreenWidth-320
Gui, Margin, 0, 0
Gui, Show, x%xpos% y%ypos% w220 h400, ChemHelp - More
return

MoreScript01:
        Gui, Submit
		url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
		run, %url%
		sleep, 1500
		WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
		Return
		
MoreScript02:
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . Episode . ""
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
		Return
MoreScript03:
		Gui, Submit
		sleep, 100
        Specinfolookup(Episode)
		Return
MoreScript04:
        Gui, Submit
		sleep, 100
        OpenEpInPatientEntry(Episode)
		Return
MoreScript05:
        Gui, Submit
        run node %A_ScriptDir%\SavePatientEPR.js %Episode%
        if (ErrorLevel = "ERROR")
            MsgBox There was an error, likely Node is not installed.
        else
        Return
MoreScript06:
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^C112^1^CF112^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript07:
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^C113^1^CF113^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript08:
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^A030^1^T^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
        Return
MoreScript09: 
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^C440^1^C3200^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript10: 
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "%5EC323%5E1%5EC1245%5E"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript11: 
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "%5EC324%5E1%5EC1260%5E"
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

;Linear Spoon's Screen capture script from https://www.autohotkey.com/board/topic/121619-screencaptureahk-broken-capturescreen-function-win-81-x64/
/* CaptureScreen(aRect, bCursor, sFileTo, nQuality)
1) If the optional parameter bCursor is True, captures the cursor too.
2) If the optional parameter sFileTo is 0, set the image to Clipboard.
   If it is omitted or "", saves to screen.bmp in the script folder,
   otherwise to sFileTo which can be BMP/JPG/PNG/GIF/TIF.
   sFile should be 0 if you want it to go to the clipboard, leaving it blank makes it save to some default filename in the working directory. 
3) The optional parameter nQuality is applicable only when sFileTo is JPG. Set it to the desired quality level of the resulting JPG, an integer between 0 - 100.
4) If aRect is 0/1/2/3, captures the entire desktop/active window/active client area/active monitor.
5) aRect can be comma delimited sequence of coordinates, e.g., "Left, Top, Right, Bottom" or "Left, Top, Right, Bottom, Width_Zoomed, Height_Zoomed".
   In this case, only that portion of the rectangle will be captured. Additionally, in the latter case, zoomed to the new width/height, Width_Zoomed/Height_Zoomed.

Example:
CaptureScreen(0)
CaptureScreen(1)
CaptureScreen(2)
CaptureScreen(3)
CaptureScreen("100, 100, 200, 200")
CaptureScreen("100, 100, 200, 200, 400, 400")   ; Zoomed
*/

/* Convert(sFileFr, sFileTo, nQuality)
Convert("C:\image.bmp", "C:\image.jpg")
Convert("C:\image.bmp", "C:\image.jpg", 95)
Convert(0, "C:\clip.png")   ; Save the bitmap in the clipboard to sFileTo if sFileFr is "" or 0.
*/


CaptureScreen(aRect = 0, bCursor = False, sFile ="", nQuality = "")
{
	If !aRect
	{
		SysGet, nL, 76  ; virtual screen left & top
		SysGet, nT, 77
		SysGet, nW, 78	; virtual screen width and height
		SysGet, nH, 79
	}
	Else If aRect = 1
		WinGetPos, nL, nT, nW, nH, A
	Else If aRect = 2
	{
		WinGet, hWnd, ID, A
		VarSetCapacity(rt, 16, 0)
		DllCall("GetClientRect" , "ptr", hWnd, "ptr", &rt)
		DllCall("ClientToScreen", "ptr", hWnd, "ptr", &rt)
		nL := NumGet(rt, 0, "int")
		nT := NumGet(rt, 4, "int")
		nW := NumGet(rt, 8)
		nH := NumGet(rt,12)
	}
	Else If aRect = 3
	{
		VarSetCapacity(mi, 40, 0)
		DllCall("GetCursorPos", "int64P", pt), NumPut(40,mi,0,"uint")
		DllCall("GetMonitorInfo", "ptr", DllCall("MonitorFromPoint", "int64", pt, "Uint", 2, "ptr"), "ptr", &mi)
		nL := NumGet(mi, 4, "int")
		nT := NumGet(mi, 8, "int")
		nW := NumGet(mi,12, "int") - nL
		nH := NumGet(mi,16, "int") - nT
	}
	Else
	{
		StringSplit, rt, aRect, `,, %A_Space%%A_Tab%
		nL := rt1	; convert the Left,top, right, bottom into left, top, width, height
		nT := rt2
		nW := rt3 - rt1
		nH := rt4 - rt2
		znW := rt5
		znH := rt6
	}

	mDC := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
	hBM := CreateDIBSection(mDC, nW, nH)
	oBM := DllCall("SelectObject", "ptr", mDC, "ptr", hBM, "ptr")
	hDC := DllCall("GetDC", "ptr", 0, "ptr")
	DllCall("BitBlt", "ptr", mDC, "int", 0, "int", 0, "int", nW, "int", nH, "ptr", hDC, "int", nL, "int", nT, "Uint", 0x40CC0020)
	DllCall("ReleaseDC", "ptr", 0, "ptr", hDC)
	If bCursor
		CaptureCursor(mDC, nL, nT)
	DllCall("SelectObject", "ptr", mDC, "ptr", oBM)
	DllCall("DeleteDC", "ptr", mDC)
	If znW && znH
		hBM := Zoomer(hBM, nW, nH, znW, znH)
	If sFile = 0
		SetClipboardData(hBM)
	Else Convert(hBM, sFile, nQuality), DllCall("DeleteObject", "ptr", hBM)
}

CaptureCursor(hDC, nL, nT)
{
	VarSetCapacity(mi, 32, 0), Numput(16+A_PtrSize, mi, 0, "uint")
	DllCall("GetCursorInfo", "ptr", &mi)
	bShow   := NumGet(mi, 4, "uint")
	hCursor := NumGet(mi, 8)
	xCursor := NumGet(mi,8+A_PtrSize, "int")
	yCursor := NumGet(mi,12+A_PtrSize, "int")

	DllCall("GetIconInfo", "ptr", hCursor, "ptr", &mi)
	xHotspot := NumGet(mi, 4, "uint")
	yHotspot := NumGet(mi, 8, "uint")
	hBMMask  := NumGet(mi,8+A_PtrSize)
	hBMColor := NumGet(mi,16+A_PtrSize)

	If bShow
		DllCall("DrawIcon", "ptr", hDC, "int", xCursor - xHotspot - nL, "int", yCursor - yHotspot - nT, "ptr", hCursor)
	If hBMMask
		DllCall("DeleteObject", "ptr", hBMMask)
	If hBMColor
		DllCall("DeleteObject", "ptr", hBMColor)
}

Zoomer(hBM, nW, nH, znW, znH)
{
	mDC1 := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
	mDC2 := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
	zhBM := CreateDIBSection(mDC2, znW, znH)
	oBM1 := DllCall("SelectObject", "ptr", mDC1, "ptr",  hBM, "ptr")
	oBM2 := DllCall("SelectObject", "ptr", mDC2, "ptr", zhBM, "ptr")
	DllCall("SetStretchBltMode", "ptr", mDC2, "int", 4)
	DllCall("StretchBlt", "ptr", mDC2, "int", 0, "int", 0, "int", znW, "int", znH, "ptr", mDC1, "int", 0, "int", 0, "int", nW, "int", nH, "Uint", 0x00CC0020)
	DllCall("SelectObject", "ptr", mDC1, "ptr", oBM1)
	DllCall("SelectObject", "ptr", mDC2, "ptr", oBM2)
	DllCall("DeleteDC", "ptr", mDC1)
	DllCall("DeleteDC", "ptr", mDC2)
	DllCall("DeleteObject", "ptr", hBM)
	Return zhBM
}

Convert(sFileFr = "", sFileTo = "", nQuality = "")
{
	If (sFileTo = "")
		sFileTo := A_ScriptDir . "\screen.bmp"
	SplitPath, sFileTo, , sDirTo, sExtTo, sNameTo
	
	If Not hGdiPlus := DllCall("LoadLibrary", "str", "gdiplus.dll", "ptr")
		Return	sFileFr+0 ? SaveHBITMAPToFile(sFileFr, sDirTo (sDirTo = "" ? "" : "\") sNameTo ".bmp") : ""
	VarSetCapacity(si, 16, 0), si := Chr(1)
	DllCall("gdiplus\GdiplusStartup", "UintP", pToken, "ptr", &si, "ptr", 0)

	If !sFileFr
	{
		DllCall("OpenClipboard", "ptr", 0)
		If	(DllCall("IsClipboardFormatAvailable", "Uint", 2) && (hBM:=DllCall("GetClipboardData", "Uint", 2, "ptr")))
			DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", hBM, "ptr", 0, "ptr*", pImage)
		DllCall("CloseClipboard")
	}
	Else If	sFileFr Is Integer
		DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", sFileFr, "ptr", 0, "ptr*", pImage)
	Else	DllCall("gdiplus\GdipLoadImageFromFile", "wstr", sFileFr, "ptr*", pImage)
	DllCall("gdiplus\GdipGetImageEncodersSize", "UintP", nCount, "UintP", nSize)
	VarSetCapacity(ci,nSize,0)
	DllCall("gdiplus\GdipGetImageEncoders", "Uint", nCount, "Uint", nSize, "ptr", &ci)
	struct_size := 48+7*A_PtrSize, offset := 32 + 3*A_PtrSize, pCodec := &ci - struct_size
	Loop, %	nCount
		If InStr(StrGet(Numget(offset + (pCodec+=struct_size)), "utf-16") , "." . sExtTo)
			break

	If (InStr(".JPG.JPEG.JPE.JFIF", "." . sExtTo) && nQuality<>"" && pImage && pCodec < &ci + nSize)
	{
		DllCall("gdiplus\GdipGetEncoderParameterListSize", "ptr", pImage, "ptr", pCodec, "UintP", nCount)
		VarSetCapacity(pi,nCount,0), struct_size := 24 + A_PtrSize
		DllCall("gdiplus\GdipGetEncoderParameterList", "ptr", pImage, "ptr", pCodec, "Uint", nCount, "ptr", &pi)
		Loop, %	NumGet(pi,0,"uint")
			If (NumGet(pi,struct_size*(A_Index-1)+16+A_PtrSize,"uint")=1 && NumGet(pi,struct_size*(A_Index-1)+20+A_PtrSize,"uint")=6)
			{
				pParam := &pi+struct_size*(A_Index-1)
				NumPut(nQuality,NumGet(NumPut(4,NumPut(1,pParam+0,"uint")+16+A_PtrSize,"uint")),"uint")
				Break
			}
	}

	If pImage
		pCodec < &ci + nSize	? DllCall("gdiplus\GdipSaveImageToFile", "ptr", pImage, "wstr", sFileTo, "ptr", pCodec, "ptr", pParam) : DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "ptr", pImage, "ptr*", hBitmap, "Uint", 0) . SetClipboardData(hBitmap), DllCall("gdiplus\GdipDisposeImage", "ptr", pImage)

	DllCall("gdiplus\GdiplusShutdown" , "Uint", pToken)
	DllCall("FreeLibrary", "ptr", hGdiPlus)
}


CreateDIBSection(hDC, nW, nH, bpp = 32, ByRef pBits = "")
{
	VarSetCapacity(bi, 40, 0)
	NumPut(40, bi, "uint")
	NumPut(nW, bi, 4, "int")
	NumPut(nH, bi, 8, "int")
	NumPut(bpp, NumPut(1, bi, 12, "UShort"), 0, "Ushort")
	Return DllCall("gdi32\CreateDIBSection", "ptr", hDC, "ptr", &bi, "Uint", 0, "UintP", pBits, "ptr", 0, "Uint", 0, "ptr")
}

SaveHBITMAPToFile(hBitmap, sFile)
{
	VarSetCapacity(oi,104,0)
	DllCall("GetObject", "ptr", hBitmap, "int", 64+5*A_PtrSize, "ptr", &oi)
	fObj := FileOpen(sFile, "w")
	fObj.WriteShort(0x4D42)
	fObj.WriteInt(54+NumGet(oi,36+2*A_PtrSize,"uint"))
	fObj.WriteInt64(54<<32)
	fObj.RawWrite(&oi + 16 + 2*A_PtrSize, 40)
	fObj.RawWrite(NumGet(oi, 16+A_PtrSize), NumGet(oi,36+2*A_PtrSize,"uint"))
	fObj.Close()
}

SetClipboardData(hBitmap)
{
	VarSetCapacity(oi,104,0)
	DllCall("GetObject", "ptr", hBitmap, "int", 64+5*A_PtrSize, "ptr", &oi)
	sz := NumGet(oi,36+2*A_PtrSize,"uint")
	hDIB :=	DllCall("GlobalAlloc", "Uint", 2, "Uptr", 40+sz, "ptr")
	pDIB := DllCall("GlobalLock", "ptr", hDIB, "ptr")
	DllCall("RtlMoveMemory", "ptr", pDIB, "ptr", &oi + 16 + 2*A_PtrSize, "Uptr", 40)
	DllCall("RtlMoveMemory", "ptr", pDIB+40, "ptr", NumGet(oi, 16+A_PtrSize), "Uptr", sz)
	DllCall("GlobalUnlock", "ptr", hDIB)
	DllCall("DeleteObject", "ptr", hBitmap)
	DllCall("OpenClipboard", "ptr", 0)
	DllCall("EmptyClipboard")
	DllCall("SetClipboardData", "Uint", 8, "ptr", hDIB)
	DllCall("CloseClipboard")
}

ButtonCapture:
sleep, 200
send, {AltDown}{Tab}{AltUp}
sleep, 300
^!c::
FormatTime, Timevar,, yyyy-MM-dd_HH-mm-ss
CaptureScreen(1,"","" . A_ScriptDir . "\Screenshots\" . Timevar . ".png","90")
if errorlevel
	MsgBox, Error capturing screen to "ChemHelp-directory\Screenshots"
else
MsgBox, 1,,Window screenshot captured at `nChemHelp-directory\Screenshots,2
    sleep,500
Return

RemoveToolTip:
Tooltip
return

Button_t:
^!'::
CoordMode,ToolTip,Screen
Loop % (20){
    Random,posx,0,% A_ScreenWidth
    Random,posy,0,% A_ScreenHeight
    Random,TimeIn,500,3000
    Random,TimeOut,5000,10000
    ShowToolTip("Random tooltip " A_Index,A_Index,TimeOut/1000,TimeIn/1000,posx,posy)
}
MsgBox, Do not click random buttons... `nNow watch the ToolTips.`nHopefully they disappear soon...
sleep, 5000
return

ShowToolTip(Text, Nr=1,Out=5,In=0,X="",Y=""){
    static
    If (text=""){
        If ((T_Text%Nr% && !T_%Nr% && T_In%Nr%<A_TickCount) || (T_%Nr% && T_Out%Nr%<A_TickCount && !((T_Text%Nr%:="") . (T_X%Nr%:="") . (T_Y%Nr%:="") . (T_In%Nr%:="") . (T_Out%Nr%:="") . (T_%Nr%:=""))))
            ToolTip % T_Text%Nr%,% T_X%Nr%,% T_Y%Nr%,% Nr+(T_%Nr%:=1)-1
        Return (!T_%Nr% ? T_In%Nr% : (T_Out%Nr%>A_TickCount ? T_Out%Nr% : ""))
    }
   T_Text%Nr%:=text,T_X%Nr%:= X,T_Y%Nr%:= Y,timer ? "" : (timer:=1),T_In%Nr%:=Round(A_TickCount + In*1000),T_Out%Nr%:=Round(A_TickCount + (In+Out)*1000), NextTimer ? "" : (NextTimer:=10)
    ShowToolTip:
    NextTimer=
    Loop 20
        If (ErrorLevel:=ShowToolTip("",A_Index))
            If (!NextTimer || NextTimer+A_TickCount>ErrorLevel)
                NextTimer:=ErrorLevel-A_TickCount+10
    SetTimer, ShowToolTip,% NextTimer ? (-1*(NextTimer+10)) : "Off"
    Return
}


/*
;-----------------------------   Logging each Authorize
#IfWinActive, Medical Validation :   (Authorise By Episode) - \\Remote
!a::
FileAppend, %A_now%`n, %A_ScriptDir%\chemhelp_log.txt
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


/*
responding to the first post:

close balloon tips on a case by case basis
http://www.autohotkey.com/forum/viewtopic.php?t=55390

Leef_me's old post, rediscovered, yeah!
ControlClick coordinates and Windows Spy question (reads a tooltip in the process)
http://www.autohotkey.com/forum/topic54333.html

Keyboard shortcut to close balloon tips
http://www.autohotkey.com/forum/viewtopic.php?t=15141&highlight=close+balloon
*/

/* This script tried to pull information from a tooltip (aim is to use the tooltip displayed over the VQ and write an API which sends updates as the size of this list increases.
#singleinstance force
CoordMode, ToolTip, screen
CoordMode, Mouse, screen
return

;======== show the information from a single tooltip =========
F1::
ControlGetText, tooltip2,,ahk_class tooltips_class32
;tooltip, ******************`n%tooltip2%`n******************, 1300, 300
;sleep, 1000
;tooltip
msgbox, %tooltip2%
return


;======== show the information from a multiple tooltips =========
F2::
WinGet, ID, LIST,ahk_class tooltips_class32
tt_text=
Loop, %id%
{
  this_id := id%A_Index%
  ControlGetText, tooltip2,,ahk_id %this_id%
  tt_text.= "******************`n" tooltip2 "`n******************`n"
}

;tooltip, %tt_text%, 300, 300
msgbox, %tt_text%
;sleep, 3000
;tooltip

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
