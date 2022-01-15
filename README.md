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

Ctrl + Alt + F: Adds a Free-PSA and removes the CDUM.
Ctrl + Alt + D: Removes a CDUM.

Alt + W (or “EPR” button on the app window): Opens EPR in default web browser and maximizes the window. (Similar to clicking.)

Alt + Shift + W: Saves the EPR to an Excel file in a folder "MRN_archives" which should be created in the root directory of ChemHelp.ahk (see EPR to Excel section below).

Alt + V (or “Verified” button on the app window): Opens Staff Notes and enter the text: "Transcription Verified."

Alt + /: Opens Staff Notes and enter the text: "Not scanned yet."

Alt + B: Back one episode

Alt + X: NeXt Episode

Alt + N: Copy Episode Number to the clipboard. This only works from within Result Entry / Verify windows with the episode active.

Ctrl + Alt + N / Ctrl + Alt + L - saves the episode number and the time in ChemHelp-directory\episode_log.txt

Alt + P: From within “Medical Validation”, copies the episode number and navigates to Result Entry window to enter or add the PHONC.  This is also useful to easily navigate to “Result Entry” window from within “Medical Validation”, even when a PHONC is not due for entry.

Alt + R: From within “Medical Validation”, copies the episode number and navigates to Result Entry window.

Ctrl + Alt + C: Screen Captures (screenshots) the active window into ChemHelp-directory\Screenshots\

Ctrl + Alt + P: From within “Medical Validation”, copies the episode number and navigates to Result Entry window to add a CCOM and put a standard comment with a link for clinician to update details on tinyurl.com/nhls-update.

Shift + Alt + P: Enable Proxy (via Registry)

Shift + Alt + R: Disable Proxy (via Registry, for example to use wifi on an NHLS computer)

Alt + C: Cancels phoning if the test item is highlighted in Results Entry window.

Alt + I: Invalidates due to NSS (for Potassiums from peripheral clinics).

Alt + F4: Closes any active window.  

Alt + Tab: Switches active windows.

Double click on a text Episode eg. from an email (to highlight the episode, then Alt + Right Click opens a Menu to open features with the episode.

# Buttons
FPSA button on ChemHelp to add a Free PSA to the current episode and remove the CDUM2. (Requires the cdum.png file to be present in root\trakcare_icons folder)
dCDUM button removes the CDUM2 only.
rFix button (request immunofixation) - if in Result Entry - Single window, it will 1. Place the episode on hold, 2. Add an immunofixation, 3. Refer the episode to the Chem - Registrar VQ list to make it visible to clinicians to prevent duplicate sample sending.
Log-on button: Runs the login script 
Mobile button: Opens the ChemHelpMobile.ahk script from the root directory. This is useful when authorizing results from a mobile device
Ex button runs a data extraction script
Verified button inserts the staff note: "Transcription Verified."
KeepOpen runs a script to refresh the VQ list in Result Verify window so TrakCare doesn't automatically log out.
More button opens a few additional options (see "More" section below).
The source code can be seen by right-clicking on ChemHelp and opening with a Text editor.  Of course you can customize your own version, should you wish, after copying the file to your PC.

# Mouse Scroll Wheel
Mouse scroll wheel has been setup to work in some windows.

# Hotstrings
Note that this list can be edited by editing ChemHelp.ahk with a text editor and entering your own text in a similar format.  Follow this format: https://www.autohotkey.com/docs/Hotstrings.htm

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

d1::[sends your full_name as specified in the chemhelp_settings.txt file in the ChemHelp root folder.
`
# EPR to Excel
This function uses a JavaScript webscraping technique via Node.js and Puppeteer to open the EPR in the background and save the patient's ePR to an Excel file.
Requirements:
1. Node.js must be installed on the PC.
2. SavePatientEPR.js file must be present in ChemHelp root directory
3. package.json and package-lock.json must be present in ChemHelp root directory
4. node_modules must be present in ChemHelp root directory with Puppeteer initialised via "npm install puppeteer" command via command prompt in ChemHelp root directory. (This is not possible from within the NHLS network and your PC will need a connection other than NHLS-intranet when doing this - due to proxy servers blocking this update.)
5. MRN_archives folder must be created in the root folder of ChemHelp.ahk.  This is where the new mrn.xlsx files will be created.
Call this function by hitting Ctrl + Shift + W while a TrakCare window is open with a patient's result.  AHK will copy the MRN to the clipboard and pass this variable to the JavaScript function (SavePatientEPR.js) to be scraped.  One can continue with other work or continue signing once the Node window has opened.
See the Puppeteer documentation online for more info on the web scraping technique or see my other data extraction script as this is basically a repurposed script from that project.

# More button
This opens a small window which includes an edit box which can take an Episode or MRN and perform various functions such as:
1. Open the form of an episode on equation document viewer (needs the configuration file "chemhelp_settings.txt" correctly formatted in the root directory with the Equation server IP of your lab).
2. Open MRN number in EPR view (needs the MRN prefix to the MRN number)
3. Locate an episode's storage position.
4. Open the episode in Patient Entry - Single window (Single window needs to be closed for this to work)
5. Extract a whole patient's cumulative results to an Excel sheet - needs Node and Puppeteer installed with the "SavePatientEPR.js" script in your ChemHelp root folder.  Speak to me if you need more info on this or follow the steps in above section "EPR to Excel" - the same function is used for this button.
6. Open various "free-text / F6 comments" such as SPE, IFE, HISTO, UOA (others coming soon).

There is no field validation option in the edit box, yet, thus it’s your responsibility to ensure you enter the correct type of identifier (episode or MRN) into the edit box.
In version 1.11 I have updated to add a VPN button which should open the Shrewsoft VPN app, use your saved Citrix username and Password and log in – I haven’t tested it on other PC’s other than my PC at home, so use with caution, and report issues please.


# Known Errors / Improvements to be made

ChemHelp was originally developed for use at Groote Schuur Hospital Chemical Pathology, but most functions should also be useable by other laboratories within the NHLS.
Most functions, eg. the Form button (or Ctrl + F), can be adapted to work at other laboratories with their respective file- / pdf servers, such as the ECM. Currently the IP address of the Equation file server should be entered into the chemhelp_settings.txt file. 172.22.4.40 (GSH), 172.22.8.50 (RXH), 172.22.12.40 (TBH), 172.22.16.63 (?Green Point Lab)
The ECM and interaction with it has not yet been configured, as we do not use the ECM server in the Western Cape.
I am however keen to visit another lab (or have a Zoom session) to learn how this works and try to implement it to work on this system as well.

The EPR to Excel Function must be coded to extract Free Text (F6) comments too.  This is quite a task which will likely take some time and planning to complete one day.

ChemHelp-Chrome.ahk was a work in progress to open the scanned forms as well as the EPR in Chrome browser rather than the default Internet Explorer.  I am not working on this now anymore, as Internet Explorer seems to be working fine.


# Feedback
Please remember to give feedback as an error can be resolved and may only lead to improvement: dieter.vdwesthuizen@nhls.ac.za 
