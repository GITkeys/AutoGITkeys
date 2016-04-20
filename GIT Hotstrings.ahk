#SingleInstance force  ;Only allow one script with this name to run
#Persistent ;Keep this program running unless the user closes it
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
;SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory. [Enable if needed]

	;Tray icon happens here
	Menu, Tray, Click, 1
	Menu, Tray, NoStandard
	Menu, Tray, Add, Click for a list of commands, HotstringList ;See HotstringList to see how this works
	Menu, Tray, Default, Click for a list of commands
	Menu, Tray, Add, Exit
	Return

;---------------------------------------------------------------------------------------------

Exit:
ExitApp

;---------------------------------------------------------------------------------------------
SendThis: ;Save the current clipboard contents and send VarTemplateText

If VarTemplateText =
msgbox Could not find the template, it may have been removed or renamed
else
OldClipboard = %clipboard%
clipboard = %VarTemplateText%
Sleep 750
If RegExMatch(clipboard , "[xX]{8,30}|<licnum>") {
If RegExMatch(OldClipboard , "([a-zA-Z0-9]{6}-){5}[a-zA-Z0-9]{6}|([a-zA-Z0-9]{5}-){5}[a-zA-Z0-9]{5}") {
clipboard := RegExReplace(clipboard, "[xX]{8,30}|<licnum>", OldClipboard) 
}
Else
TrayTip, AHK, No L/N on clipboard - Please replace XXXXXXXXXXXXX with correct L/N,,1
}
Send ^v
Sleep 100
clipboard = %OldClipboard%
Return

;---------------------------------------------------------------------------------------------
 ;Email sign-offs
;---------------------------------------------------------------------------------------------

::tfe::Thank you for your email.
::tfc::Thank you for your call.
::lladw::It looks like you have some adware [Advertising software] installed on your computer. `n `n
::ihth::I hope this helps. If you have further questions or issues, please do not hesitate to contact us.
::iyh::If you have further questions or issues, please do not hesitate to contact us.
::tgn::That is great news to hear that the issue is now resolved. {Enter} {Enter}As always, if you have any further questions or comments regarding our service then please do not hesitate to contact us.
::gth::Glad to hear that the issue has been resolved. {Enter} {Enter}As always, if you have any further questions or comments regarding our service please do not hesitate to contact us.
::atb::All the best,
::bre::Best Regards,
::tyfyc::Thank you for your cooperation, we appreciate it. If you have further questions or issues, please do not hesitate to contact us.
::sfs::
FileRead, VarTemplateText, H:\GITAutohotkeysUpdater\Signature.txt
Goto, SendThis

;---------------------------------------------------------------------------------------------
;Common CMS notes
;---------------------------------------------------------------------------------------------

::-_::_____________________________________________________________________________________
::afs::-Asking for scan logs
::jmd::-Join.me'd
::sadw::- Sending adwcleaner
::nac::No active components
::srri::-Sending remover {+} reinstall info
::smui::-Sending manual update info
::ssff::-Sending specified file fix
::sce::-Sending courtesy email
::sra::-Sending reset access
::wfcr::- Marked as waiting for customer reply
::gsmsr::- Getting safe mode scan results

;---------------------------------------------------------------------------------------------
;Master templates
;---------------------------------------------------------------------------------------------

;---------------------------------------------------------------------------------------------
;[Fix][Thing]
;---------------------------------------------------------------------------------------------
 
::fixaccess:: ;Reset Access
::resetacc::
FileRead, VarTemplateText, Q:\Installation and Download\install - reset access.txt
Goto, SendThis

::fixfile::
::fixsfnf:: ;Fix Specified file not found
::sfnf::
FileRead, VarTemplateText, Q:\Installation and Download\install error - specified file not found.txt
Goto, SendThis 

::fixpath::
::fixspnf:: ;Fix Specified path not found
::spnf::
FileRead, VarTemplateText, Q:\Installation and Download\error - specified path not found.TXT
Goto, SendThis

::fixchkdsk:: ;Send chkdsk and defrag
::fixdefrag:: ;Send chkdsk and defrag
FileRead, VarTemplateText, Q:\Hardware Issues\chkdsk +defrag.txt
Goto, SendThis

;---------------------------------------------------------------------------------------------
;[Get][Thing]
;--------------------------------------------------------------------------------------------

::getrevo:: ;Get Revo uninstaller
FileRead, VarTemplateText, Q:\Short Templates (how to's)\uninstallation - Revo Uninstaller.txt
Goto, SendThis

::getfwdiag:: ;Get fwdiag
FileRead, VarTemplateText, Q:\Firewall\firewall - fwdiag (installation issue).TXT
Goto, SendThis

::getscrnsht:: ;Just "scrnsht" works too
::scrnsht::
FileRead, VarTemplateText, Q:\Short Templates (how to's)\screenshot (how to).txt
Goto, SendThis

::getgmer:: ;Get a GMER scan  
::gmerdir::
FileRead, VarTemplateText, Q:\Infections\gmer scan.txt
Goto, SendThis

::getautgmer:: ;Get an Autoruns and GMER scan
::autgmer::
FileRead, VarTemplateText, Q:\Infections\gmer and autoruns.txt
Goto, SendThis

::getotl:: ;Get an OTL scan
::otldir::
FileRead, VarTemplateText, Q:\Infections\OTL.txt
Goto, SendThis
 
::getrs:: ;Get rs detection and scan logs
FileRead, VarTemplateText, Q:\Scanning\scan - rs detection and scan log.txt
Goto, SendThis

::getrsms:: ;Get rs detection, scan logs and msinfo
FileRead, VarTemplateText, Q:\Scanning\scan - rs detection and scan log + msinfo.txt
Goto, SendThis 

::getsms:: ;Get them to do a scan in safe mode
FileRead, VarTemplateText, Q:\Infections\safe mode scan.txt
Goto, SendThis 

::getsmsgmer::
::getgmersms:: ;Get them to do a scan in safe mode
FileRead, VarTemplateText, Q:\Infections\gmer + safe mode scan.txt
Goto, SendThis 

::getsfc:: ;Get them to do an sfc scan
FileRead, VarTemplateText, Q:\Infections\Infection - sfc for whitelisted files.txt
Goto, SendThis 

::getmsinfo:: 
::getms:: ;Get msinfo
FileRead, VarTemplateText, Q:\Logs and diagnostics\msinfo32.txt
;FileRead, VarTemplateText, Q:\Short Templates (how to's)\msinfo32.txt ;why are there two of these?
Goto, SendThis

::getftp:: ;Get rs detection, scan logs and msinfo
FileRead, VarTemplateText, Q:\Short Templates (how to's)\ftp upload.txt
Goto, SendThis 

::getrevo:: ;Get them to run Revo 
FileRead, VarTemplateText, Q:\Short Templates (how to's)\uninstallation - Revo Uninstaller.txt
Goto, SendThis 

::getlmi:: ;Get them to call us for an LMI
FileRead, VarTemplateText, Q:\LMI\LMI session - offer (support staff).txt
Goto, SendThis 

;---------------------------------------------------------------------------------------------
;[Inst]all[Thing]
;---------------------------------------------------------------------------------------------
 
::instavg:: ;Install AVG
FileRead, VarTemplateText, Q:\Installation and Download\install - AVG 2014 (home).txt
Goto, SendThis 
  
;---------------------------------------------------------------------------------------------
;[Rem]ove[Thing]
;---------------------------------------------------------------------------------------------

::getadw::
::remads:: ;Send AdwCleaner
::adwc::
FileRead, VarTemplateText, Q:\Infections\adwCleaner.txt
Goto, SendThis

::remqvo6:: ;Remove Qvo6
FileRead, VarTemplateText, Q:\Infections\qvo6.txt
Goto, SendThis

::remxp:: ;Send template about XP Support
FileRead, VarTemplateText, Q:\Misc Issues\No Windows XP Support.txt
Goto, SendThis

::remunos:: ;Uninstall any program template
FileRead, VarTemplateText, Q:\Short Templates (how to's)\uninstall a program - unknown OS.txt
Goto, SendThis

::rempct:: ;Remove PCT 
FileRead, VarTemplateText, Q:\PC TuneUp\PC Tuneup - Remover.txt
Goto, SendThis 

::remrein:: ;Remove AVG
FileRead, VarTemplateText, Q:\Installation and Download\AVG remover.txt
Goto, SendThis

::remmca:: ;Remove McAfee
FileRead, VarTemplateText, Q:\Conflicts with other AV software\mcafee removal.txt
Goto, SendThis

::remtrend:: ;Remove TrendMicro
FileRead, VarTemplateText, Q:\Conflicts with other AV software\TrendMicro.txt
Goto, SendThis

::remnort:: ;Remove Norton
FileRead, VarTemplateText, Q:\Conflicts with other AV software\Norton remover linked.txt
Goto, SendThis

;--------------------------------------------------------------------------------------------
;Autoruns 

::getautoruns:: ;Get autoruns
::autdir::
FileRead, VarTemplateText, Q:\Infections\autoruns - direct link.txt
Goto, SendThis

::getautms:: ;Get autoruns + msinfo
FileRead, VarTemplateText, Q:\Logs and diagnostics\Autoruns + MSINFO.txt
Goto, SendThis  

::sentexe:: ;Client sent EXE instead of ARN
::autexe::
FileRead, VarTemplateText, Q:\Infections\autoruns - sent exe file instead of arn.txt
Goto, SendThis 

::sendautrem:: ;Send Autoruns Removal template
::autrem::
FileRead, VarTemplateText, Q:\Infections\autoruns - removal template.txt
Goto, SendThis

;--------------------------------------------------------------------------------------------
 ;Autoruns removal template line breaks
;--------------------------------------------------------------------------------------------

::utlt:: ;Under the Logon tab: 
Send, --------------------------------------------------------------------------------------------------- {Enter} Under the Logon tab: {Enter}{Enter}
Return

::utiet:: ;Under the Internet Explorer tab: 
Send, --------------------------------------------------------------------------------------------------- {Enter} Under the Internet Explorer tab: {Enter}{Enter}
Return

::utstt:: ;Under the Scheduled Tasks tab: 
Send, --------------------------------------------------------------------------------------------------- {Enter} Under the Scheduled Tasks tab: {Enter}{Enter}
Return

::utst:: ;Under the Services tab: 
Send, --------------------------------------------------------------------------------------------------- {Enter} Under the Services tab: {Enter}{Enter}
Return

::utait:: ;Under the AppInit tab: 
Send, --------------------------------------------------------------------------------------------------- {Enter} Under the AppInit tab: {Enter}{Enter}
Return 


;--------------------------------------------------------------------------------------------
;Misc Templates
;--------------------------------------------------------------------------------------------

::youspud::
::unclem:: ;Unclear Email
::unclearem::
FileRead, VarTemplateText, Q:\Short Templates (how to's)\unclear email.txt
Goto, SendThis 

;--------------------------------------------------------------------------------------------
;This section automatically documents hotstrings with comments and spits them out into a user readable list when clicking on the tray icon
;--------------------------------------------------------------------------------------------

HotstringList: ;Uses RegEx to find all hotstrings in this file

FileRead, ScriptContents, %A_ScriptFullPath%
if not ErrorLevel
{
  AllHotstrings := ""

  Loop, parse, ScriptContents, `n, `r
  {
    if (RegExMatch(A_LoopField, ":[a-z]*?:(.+?)::+(.+)", SubPat)) ;Not an actual Hotstring, Just gets detected as one and automatically added to this list
    {
      AllHotstrings := AllHotstrings . SubPat . "`n`n"
    }
  }
  ;MsgBox 0, Hotstrings, %AllHotstrings%
}


;--------------------------------------------------------------------------------------------

AboutBox: ;Makes the HotstringList readable for users

about_box_title = GIT Hotstrings
Gui, 2:+Resize  -Maximizebox +minsize450x250 ; Make the window resizable.
Gui 2:Add, edit, vEdit2about r3 w400 readonly -E0x200,
Gui, 2:Add, Button, Default, OK
Gui, 2:Show ,, %about_box_title%


aboutdata =
           ( Join 
List of all hotstrings:`n
`n
 %AllHotstrings%
                 )

Guicontrol,2:, edit2about, %aboutdata%
return

;--------------------------------------------------------------------------------------------

2ButtonOK:  ; This section is used by the "AboutBox" above.
2GuiClose:
2GuiEscape:
Gui, 1:-Disabled  ; Re-enable the main window (must be done prior to the next step).
Gui Destroy  ; Destroy the about box.
return

2GuiSize:
  If A_EventInfo = 1  ; The window has been minimized.  No action needed.
    Return
  ; Otherwise, the window has been resized or maximized. Resize the controls to match.
  GuiControl 2:Move, edit2about, % "H" . (A_GuiHeight-40) . " W" . (A_GuiWidth-20)
  GuiControl 2:Move, OK, % "y" (A_GuiHeight-25) . "x" (A_guiwidth/2)  
  GuiControl 2:Move, OK2, % "y" (A_GuiHeight-5) . "x" (A_guiwidth/4)
Return 



;--------------------------------------------------------------------------------------------
;This makes Windows Key + Capslock change case
;--------------------------------------------------------------------------------------------

cycleNumber := 1

#Capslock::
If (cycleNumber==1)
{
ConvertUpper()
cycleNumber:= 2
}
Else If (cycleNumber==2)
{
ConvertLower()
cycleNumber:= 3
}
Else
{
ConvertSentence()
cycleNumber:= 1
}
Return





ConvertUpper()
{
	clipSave := Clipboard
	Clipboard = ; Empty the clipboard so that ClipWait has something to detect
	SendInput, ^c ;copies selected text
	ClipWait
	StringReplace, Clipboard, Clipboard, `r`n, `n, All ;Fix for SendInput sending Windows linebreaks 
	StringUpper, Clipboard, Clipboard ;Convert case
	Len:= Strlen(Clipboard) ;Set number of characters
    Send %Clipboard% ;Send new string
	Send +{left %Len%} ;Re-select text
    VarSetCapacity(OutputText, 0) ;free memory
	Clipboard := clipSave ;Restore previous clipboard
}

ConvertLower()
{
	clipSave := Clipboard
	Clipboard = ; Empty the clipboard so that ClipWait has something to detect
	SendInput, ^c ;copies selected text
	ClipWait
	StringReplace, Clipboard, Clipboard, `r`n, `n, All ;Fix for SendInput sending Windows linebreaks 
	StringLower, Clipboard, Clipboard ;Convert case
	Len:= Strlen(Clipboard) ;Set number of characters
    Send %Clipboard% ;Send new string
	Send +{left %Len%} ;Re-select text
    VarSetCapacity(OutputText, 0) ;free memory
	Clipboard := clipSave ;Restore previous clipboard
}

ConvertSentence()
{
	clipSave := Clipboard
	Clipboard = ; Empty the clipboard so that ClipWait has something to detect
	SendInput, ^c ;copies selected text
	ClipWait
	StringReplace, Clipboard, Clipboard, `r`n, `n, All ;Fix for SendInput sending Windows linebreaks 
	StringLower, Clipboard, Clipboard ;Convert case
	Clipboard := RegExReplace(Clipboard, "(((^|([.!?]+\s+))[a-z])| i | i')", "$u1")
	Len:= Strlen(Clipboard) ;Set number of characters
    Send %Clipboard% ;Send new string
	Send +{left %Len%} ;Re-select text
    VarSetCapacity(OutputText, 0) ;free memory
	Clipboard := clipSave ;Restore previous clipboard
}

;--------------------------------------------------------------------------------------------
;This makes Windows Key + Insert - paste text only; formatted version remains in clipboard
;--------------------------------------------------------------------------------------------

#Insert::
	OriginalClipboardContents = %ClipBoardAll%
	ClipBoard = %ClipBoard%                             ; Convert to text
	Send ^v 						
	Sleep 100                                           ; Don't change clipboard while it is pasted
	ClipBoard := OriginalClipboardContents              ; Restore original clipboard contents
	OriginalClipboardContents =                         ; Free memory
	Return 	
