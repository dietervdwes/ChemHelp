# ChemHelp
ChemHelp is an AutoHotKey based script to help signing of blood results in Chemical Pathology on the TrakCare LIS.

# Introduction
Written by Dieter van der Westhuizen, 2018 - current; 
Inspired by Chad Centner - TrakHelper, 2015-2018.

ChemHelp was developed to assist reviewers in Chemical Pathology to eliminate unnecessary mouse use and clicking by automating some of the tasks while using TrakCare for reviewing blood results.
ChemHelpMobile was made as an interface to sign results via a remote desktop viewer and just enlarges the buttons to be able to click the buttons more easily on a mobile phone.  It includes all the features of ChemHelp, and also needs ChemHelp to be in the same folder.
NB: Please report any ideas or errors to dieter.vdwesthuizen@nhls.ac.za.  That way I can try improving the app.

# Installation
1.	Install AutoHotKey by downloading from this page: https://www.autohotkey.com/
If you want to edit the script for use in your lab, then you will need an applicable text editor, see below.

2.	Copy "ChemHelp.ahk" and "ChemHelpMobile.ahk" to local desktop or other convenient location on the local workstation

3.	(Optional when wanting to use the Login button) 
Copy the file chemhelp_settings.txt to your My Documents folder and replace the contents of each field in the file with your login details.  Replace the "18" in the file (last field) with the amount of {TAB} buttons it takes to navigate to the blue TrakCare icon after having logged into Citrix. (This is usually around 16-20 tabs, depending on how many bookmarks etc. you have and you will need to test this after manually logging into Citrix and counting how many times you need to press tab until the blue TrakCare icon highlights.  It is important that the chemhelp_settings.txt file (i.e. individual file and NOT the folder), be copied, to your "Documents" root folder.
Important: ChemHelpMobile needs to be in the same folder as ChemHelp if you want to use the Mobile interface (like Teamviewer or Anydesk).

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

Alt + P: From within “Medical Validation”, copies the episode number and navigates to Result Entry window to enter or add the PHONC.  This is also useful to easily navigate to “Result Entry” window from within “Medical Validation”, even when a PHONC is not due for entry.

Ctrl + Alt + P: From within “Medical Validation”, copies the episode number and navigates to Result Entry window to add a CCOM and put a standard comment with a link for Dr. to update details.

Alt + C: Cancels phoning if the test item is highlighted in Results Entry window.

Alt + I: Invalidates due to NSS (for Potassiums from peripheral clinics).

Alt + F4: Closes any active window.  

Alt + Tab: Switches active windows.

Double click on a text Episode eg. from an email then Alt + Right Click opens a Menu to open features with the episode.

FPSA button on ChemHelp to add a Free PSA to the current episode.

Login button: Runs the login script 

Ex button runs a data extraction script

The source code can be seen by right-clicking on ChemHelp and opening with a Text editor.  Of course you can customize your own version, should you wish, after copying the file to your PC.

# Mouse Scroll Wheel
Mouse scroll wheel has been setup to work in some windows.

# Hotstrings for reporting protein electrophoreses
Note that this list can be edited by editing ChemHelp.ahk with a text editor and entering your own text in a similar format.  To learn how to edit these click here: https://www.autohotkey.com/docs/Hotstrings.htm

`::INFLAM::The alpha-1, -2 and beta-2 (complement) regions are elevated and there is a polyclonal hypergammaglobulinaemia at  g/L (8 - 14 g/L). No monoclonal peaks are visible. This pattern suggests an inflammatory process.  If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.

::INFLAMS::The alpha-1, -2 and beta-2 (complement) region is elevated and the gamma region measures _ g/L (8 - 14 g/L). No monoclonal peaks are visible. This pattern suggests an inflammatory process. If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.

::PREV::The previously typed monoclonal Ig_ persists in _ gamma at _ g/L.  Immunoparesis is _.

::PROM::A prominent peak is present in the mid-gamma region measuring g/L. The remainder of the gamma region measures g/L (8-14 g/L).  Immunotyping will be performed. Please see results below.

::SURINE::Please send urine (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) for Bence Jones protein electrophoresis.

::BGB::Beta-gamma bridging is present. This is consistent with a chronic inflammatory process associated with an IgA response.  Causes may include cirrhosis, and cutaneous or mucosal inflammation.

::SUSP::If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.

::REPEATSPE::Repeat SPE is recommended in 3-6 months or when the inflammatory condition has subsided.

::NORMP::Normal protein electrophoresis pattern.  No monoclonal peaks are present. The gamma region measures _ g/L (8-14 g/L). If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.

::NEPHR::Hypoalbuminaemia is present.  The alpha-2 (macroglobulin) region is significantly increased at _ g/L (5-9 g/L).  The gamma region measures _ g/L (8-14 g/L). No monoclonal peaks are visible. This pattern suggests nephrotic syndrome. If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended. 

::A-1::The alpha-1 peak is biphasic, suggesting alpha-1-antitrypsin heterozygosity.

::text1::
(
Any text
Newline preserved in this way
)
`
# Known Errors / Improvements to be made

ChemHelp was originally developed for use at Groote Schuur Hospital Chemical Pathology, but most functions should also be useable by other laboratories within the NHLS.
Most functions, eg. the Form button (or Ctrl + F), can be adapted to work at other laboratories with their respective file- / pdf servers.
The ECM and interaction with it has not yet been configured, as we do not use the ECM server in the Western Cape.
I am however keen to visit another lab (or have an ECHO session) to learn how this works and try to implement it to work on this system as well.

ChemHelp-Chrome.ahk was a work in progress to open the scanned forms as well as the EPR in Chrome browser rather than the default Internet Explorer.  I am not working on this now anymore, as Internet Explorer seems to be working fine.


# Feedback
Please remember to give feedback as an error can be resolved and may only lead to improvement: dieter.vdwesthuizen@nhls.ac.za 
