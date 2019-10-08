# ChemHelp
ChemHelp is an AutoHotKey based script to help signing of blood results in Chemical Pathology on the TrakCare LIS.

# Introduction
Written by Dieter van der Westhuizen, 2018 - current
Inspired by Chad Centner - TrakHelper, 2015-2018.

ChemHelp was developed to assist reviewers in Chemical Pathology to eliminate unnecessary mouse use and clicking by automating some of the tasks while using TrakCare for reviewing blood results.
NB: Please report any ideas or errors to dieter.vdwesthuizen@nhls.ac.za.  That way I can try improving the app.

# Installation
1.	Install AutoHotKey by downloading from this page: https://www.autohotkey.com/
If you want to edit the script for use in your lab, then you will need an applicable text editor, see below.

2.	Copy "ChemHelp.ahk" to local desktop or other convenient location on the local workstation

3.	(Optional when wanting to use the Login button) 
Copy the contents of "Copy_contents_to_My_Documents" to your My Documents folder and replace the contents of each text file with your login details.  Also replace the "tabs_citrix.txt" contents with the amount of {TAB} buttons it takes to navigate to the blue TrakCare icon after having logged into Citrix. (This is usually around 16-20 tabs, depending on how many bookmarks etc. you have and you will need to test this after manually logging into Citrix and counting how many times you need to press tab until the blue TrakCare icon highlights.  Save this number in the "tabs_citrix.txt" file in your Documents folder.
It is important that the contents (i.e. individual files and NOT the folder), be copied, to your "Documents" root folder.

4.	Double click "ChemHelp" which you’ve copied to your Desktop to launch ChemHelp.

5.	(Optional) Although editing of the script in Notepad is possible, the following editor is recommended: Scite4AutoHotkey, downloadable from: http://fincs.ahk4.net/scite4ahk/.


When ChemHelp window is running, the following shortcuts are active:

 
# ChemHelp Shortcuts

Alt + A: authorizes the episode.

Alt + Q: Page up  

Alt + Z: Page down.  (An automated mouse click on the first line accommodates the latter two functions)

Alt + F (or “Form” button on the app window): Opens the scanned Form on Equation Document viewer via the default web browser and resizes the window to the right half of the screen if Internet Explorer is your default browser.

Alt + W (or “EPR” button on the app window): Opens EPR in default web browser and maximizes the window. (Similar to clicking.)

Alt + V (or “Verified” button on the app window): Opens Staff Notes and enter the text: "Transcription Verified."

Alt + /: Opens Staff Notes and enter the text: "Not scanned yet."

Alt + B: Back one episode

Alt + X: NeXt Episode

Alt + N: Copy Episode Number to the clipboard.

Alt + E (or “Episode” button on the app window): Opens a new text box to enter an Episode number to open up the scanned form on Equation document viewer.

Alt + P: From within “Medical Validation”, copies the episode number and navigates to Result Entry window to enter or add the PHONC.  This is also useful to enter to “Result Entry” window from within “Medical Validation”

Alt + F4: Closes any active window.  

Alt + Tab: Switches active windows.  (Latter two shortcuts are universal on Windows computers.)

FPSA button on ChemHelp to add a Free PSA to the current episode.

Login button: Runs the login script 
The source code can be seen by right-clicking on ChemHelp and opening with a Text editor.  Please don't change on the S-drive unless discussed.  Of course you can customize your own version, should you wish, after copying the file (not the shortcut) to your PC.
# Known Errors / Improvements to be made

ChemHelp was originally developed for use at Groote Schuur Hospital Chemical Pathology, but most functions should also be useable by other laboratories within the NHLS.
Most functions, eg. the Form button (or Ctrl + F), can be adapted to work at other laboratories with their respective file- / pdf servers.
The ECM and interaction with it has not yet been configured, as we do not use the ECM server in the Western Cape.
I am however keen to visit another lab (or have an ECHO session) to learn how this works and try to implement it to work on this system as well.


# Feedback
Please remember to give feedback as an error can be resolved and may only lead to improvement: dieter.vdwesthuizen@nhls.ac.za 
