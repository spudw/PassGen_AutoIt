#AutoIt3Wrapper_Icon = "Icon.ico"
#AutoIt3Wrapper_Compression = 4
Const $sVersion = "1.3b"

#pragma compile(FileDescription, Password Generator Tool)
#pragma compile(ProductName, PassGen)
#pragma compile(ProductVersion, 1.3 Beta)
#pragma compile(FileVersion, 0.1.3.1) ; The last parameter is optional.

;#AutoIt3Wrapper_Res_File_Add

; Version History
;
; Version 1.3 - 2021/01/25
;		New Features & Misc Code Cleanup/Optimization
;		[+] Added KeyArchive / KeyList mechanisms
;			[+] Added Key Archive functions
;				- Select, Add, Remove, Modify, Date-Change, Import, Export, Clear
;			[*] Altered Key Change button - Integrated Key Archive management functions
;				- Dynamic Context Menu with Key Archive management function controls
;			[*] Changed Registry key storage logic to accommodate KeyArchive
;		[*]Key and Password become masked on Window Focus Lost or Minimize
;		[*] Tray Icon Hidden when Not Minimized to Tray
;		[+] Added Auto Purge from Clipboard after timer
;			[*] Periodically checks clipboard and clears copied password or passphrase from clipboard
;			    and Copied to labels when clipboard no longer contains password or passphrase
;		[+] Added Key Complexity Requirement
;		[+] Added Legacy Key logic - to migrate RegKey from prior version.
;
; Version 1.2.1.1 - 2020/08/24
;		Minor Tray Tweaks
;		[*] Made system tray icon always shown
;		[*] Minor Tray Menu changes
;
; Version 1.2.1 - 2020/08/24
;		Functional Changes & Minor Code Cleanup
;		[+] Restored Minimize button and added Taskbar Peek Preview GUI hide workaround
;		[+] Added Close to Tray function option
;		[*] Minor Code Cleanup
;
; Version 1.2 - 2020/08/21
;		Code Cleanup & Feature Additions
;		[*] Removed unnecessary code
;		[+] Clear Clipboard on Exit
;		[+] Added Menubar with functions
;			File > Exit
;			Options > Automatic Startup on Login
;		[+] Added Tray Menu with functions
;			Open, Quit
;		[*] Changed Tray Icon behavior - Hide when GUI visible, Show when GUI hidden
;		[*] Changed Tray Icon click behavior
;			Single Left-Click restores GUI
;			Single Right-Click open Tray Menu
;		[-] Removed Minimize button
;
; Version 1.1 - 2020/08/20
;		Feature Additions
;		[+] Added Password Masking functionality
;		[+] Added Minimize to Tray functionality
;
; Version 1.0.1 - 2020/02/25
;		Minor Customization
;		[*] Changed EXE Icom
;
; Version 1.0 - 2020/05/20
;		Original Full Release

#Region - Includes and Variables
#include <Array.au3>
#include <Crypt.au3>
#include <Date.au3>
#include <Misc.au3>
#include <WinAPI.au3>
#include <WinAPICom.au3>
#include <GuiComboBox.au3>
#include <GuiDateTimePicker.au3>
#include <GuiMenu.au3>
#include <ButtonConstants.au3>
#include <ColorConstants.au3>
#include <EditConstants.au3>
#include <FileConstants.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <TrayConstants.au3>
#include <WindowsConstants.au3>

If _Singleton("PassGenA", 1) = 0 Then
	$sRunningProcessPath = _WinAPI_GetProcessFileName(ProcessExists("PassGen.exe"))
	If $sRunningProcessPath = @ScriptFullPath Then Exit
	If _VersionCompare(FileGetVersion(@ScriptFullPath), FileGetVersion($sRunningProcessPath)) = 1 Then
		ProcessClose("PassGen.exe")
	Else
		Exit
	EndIf
EndIf

Opt("GUIEventOptions", 1)
Opt("GUIOnEventMode", 1)
Opt("TrayMenuMode", 1 + 2)
Opt("TrayOnEventMode", 1)
TraySetClick(8)

Const $CHARACTERLIST = "ABCEFGHKLMNPQRSTUVWXYZ0987654321abdefghjmnqrtuwy"
Const $CHARACTERLISTLEN = StringLen($CHARACTERLIST)
Const $REGKEYPATH = "HKCU\Software\PassGen\test"
Const $REGKEYCURRENT = "CurrentKey"
Const $STARTUPLINK = @StartupDir & "\PassGen.exe.lnk"
Const $PROGRAMPATH = @ProgramsDir & "\PassGen\PassGen.exe"

Const $tagDATA_BLOB = "DWORD cbData;ptr pbData;"
Const $tagCRYPTPROTECT_PROMPTSTRUCT = "DWORD cbSize;DWORD dwPromptFlags;HWND hwndApp;ptr szPrompt;"

Enum $KEYARCHIVEACTION_SELECT = 1000, $KEYARCHIVEACTION_ADD, $KEYARCHIVEACTION_REMOVE, $KEYARCHIVEACTION_MODIFY, $KEYARCHIVEACTION_DATECHANGE, _
		$KEYARCHIVEACTION_IMPORT, $KEYARCHIVEACTION_EXPORT, $KEYARCHIVEACTION_CLEAR
Enum Step *2 $COMPLEXITY_UPPER, $COMPLEXITY_LOWER, $COMPLEXITY_NUMBER, $COMPLEXITY_SYMBOL

Const $BINARYFORMAT_SOH = Binary("0x01")
Const $BINARYFORMAT_STX = Binary("0x02")
Const $BINARYFORMAT_ETX = Binary("0x03")
Const $BINARYFORMAT_EOM = Binary("0x19")
Const $BINARYFORMAT_RS = Binary("0x1E")
Const $BINARYFORMAT_US = Binary("0x1F")
Const $BINARYFORMAT_SUB = Binary("0x1A")
Const $BINARYFORMAT_FS = Binary("0x1C")
Const $EXPORTFILEHEADER = $BINARYFORMAT_SOH & StringToBinary("PGEF") & $BINARYFORMAT_STX ;[Start  of Heading] MagicNumber [Start of text]
Const $EXPORTFILEEOF = $BINARYFORMAT_ETX & StringToBinary("PGEF") & $BINARYFORMAT_EOM ;[End of Text]  MagicNumber [End of Medium]
Const $EXPORTFILEKEYENTRYHEADER_KEYDATE = $BINARYFORMAT_RS & StringToBinary("KD") & $BINARYFORMAT_US ;[Record Separator] MagicNumber [Unit Separator]
Const $EXPORTFILEKEYENTRYHEADER_KEYVALUE = $BINARYFORMAT_US & StringToBinary("KV") & $BINARYFORMAT_US ;[Unit Separator] MagicNumber [Unit Separator]
Const $EXPORTFILEKEYENTRYHEADER_KEYLENGTH = $BINARYFORMAT_US & StringToBinary("KL") & $BINARYFORMAT_US ;[Unit Separator] MagicNumber [Unit Separator]
Const $EXPORTFILEENCRYPTEDHEADER = $BINARYFORMAT_SUB & StringToBinary("PGENCEXP") & $BINARYFORMAT_FS ;[Substitution] MagicNumber [File Separator]
Const $EXPORTFILEHEADERLEN = BinaryLen($EXPORTFILEHEADER)
Const $EXPORTFILEEOFLEN = BinaryLen($EXPORTFILEEOF)
Const $EXPORTFILEKEYENTRYHEADER_KEYDATELEN = BinaryLen($EXPORTFILEKEYENTRYHEADER_KEYDATE)
Const $EXPORTFILEKEYENTRYHEADER_KEYVALUELEN = BinaryLen($EXPORTFILEKEYENTRYHEADER_KEYVALUE)
Const $EXPORTFILEKEYENTRYHEADER_KEYLENGTHLEN = BinaryLen($EXPORTFILEKEYENTRYHEADER_KEYLENGTH)
Const $EXPORTFILEENCRYPTEDHEADERLEN = BinaryLen($EXPORTFILEENCRYPTEDHEADER)
Const $DATEFORMATBYTELEN = StringLen("YYYY/MM/DD")

Const $AUTOPURGETIME = 1 ;minutes
Global $g_sActiveKeyGUID = "", $g_aKeyArchive, $g_iKeyOperation, $hTimer
#EndRegion - Includes and Variables

#Region - UI Creation
Dim $aGUI[1] = ["hwnd|id"]
Enum $hGUI = 1, $idMnuFile, $idMnuFileQuit, $idMnuOptions, $idMnuOptionsAutoStart, $idMnuOptionsCloseToTray, $idTrayOpen, $idTrayQuit, $idBtnRevealKey, $idLblKey, $idTxtKey, $idCmbKeyList, _
		$idDateKeyDatePicker, $idBtnKey, $idLblPassphrase, $idLblPassphraseUse, $idTxtPassphrase, $idBtnPassphrase, $idLblPassphraseMsg, $idLblPassword, $idLblPasswordUse, $idTxtPassword, _
		$idBtnPassword, $idLblPasswordMsg, $iGUILast
ReDim $aGUI[$iGUILast]

$aGUI[$hGUI] = GUICreate("PassGen v" & $sVersion, 508, 230, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_SYSMENU))
$aGUI[$idMnuFile] = GUICtrlCreateMenu("&File")
$aGUI[$idMnuFileQuit] = GUICtrlCreateMenuItem("&Quit", $aGUI[$idMnuFile])
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idMnuOptions] = GUICtrlCreateMenu("&Options")
$aGUI[$idMnuOptionsAutoStart] = GUICtrlCreateMenuItem("&Automatic Start on Login", $aGUI[$idMnuOptions])
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idMnuOptionsCloseToTray] = GUICtrlCreateMenuItem("&Enable Close to Tray", $aGUI[$idMnuOptions])
GUICtrlSetOnEvent(-1, "GUIEvents")
TrayCreateItem("PassGen v" & $sVersion)
TrayItemSetState(-1, $TRAY_DISABLE)
TrayCreateItem("")
$aGUI[$idTrayOpen] = TrayCreateItem("Open")
TrayItemSetOnEvent(-1, "TrayEvents")
TrayItemSetState(-1, $TRAY_DEFAULT)
TrayCreateItem("")
$aGUI[$idTrayQuit] = TrayCreateItem("Quit")
TrayItemSetOnEvent(-1, "TrayEvents")
$aGUI[$idBtnRevealKey] = GUICtrlCreateCheckbox("&Show", 12, 12, 44, 28, BitOR($BS_PUSHLIKE, $BS_AUTOCHECKBOX))
GUICtrlSetState(-1, $GUI_UNCHECKED)
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idLblKey] = GUICtrlCreateLabel("Key:", 64, 18, 40, 20)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTUNDER)
$aGUI[$idTxtKey] = GUICtrlCreateInput("", 104, 10, 330, 34, $ES_PASSWORD)
Const $ES_PASSWORDCHAR = GUICtrlSendMsg(-1, $EM_GETPASSWORDCHAR, 0, 0)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetFont(-1, 18, $FW_BOLD, Default, "Consolas")
$aGUI[$idCmbKeyList] = GUICtrlCreateCombo("", 104, 10, 330, 34, $CBS_DROPDOWNLIST)
_GUICtrlComboBox_SetCueBanner($aGUI[$idCmbKeyList], "Select a Key, or hit Cancel")
GUICtrlSetState(-1, $GUI_HIDE)
GUICtrlSetFont(-1, 14, $FW_BOLD, Default, "Consolas")
$aGUI[$idDateKeyDatePicker] = GUICtrlCreateDate("", 104, 10, 330, 34, BitOR($DTS_SHORTDATEFORMAT, $DTS_APPCANPARSE))
_GUICtrlDTP_SetFormat(ControlGetHandle($aGUI[$hGUI], "", $aGUI[$idDateKeyDatePicker]), "yyyy/MM/dd")
GUICtrlSetState(-1, $GUI_HIDE)
GUICtrlSetFont(-1, 18, $FW_BOLD, Default, "Consolas")
$aGUI[$idBtnKey] = GUICtrlCreateButton("&Change", 442, 10, 58, 34, $BS_DEFPUSHBUTTON)
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idLblPassphrase] = GUICtrlCreateLabel("Passphrase:", 16, 68, 100, 20)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTUNDER)
$aGUI[$idLblPassphraseUse] = GUICtrlCreateLabel("Send with Email", 15, 88, 100, 20)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 9, $FW_NORMAL, $GUI_FONTITALIC, "Times New Roman")
$aGUI[$idTxtPassphrase] = GUICtrlCreateInput("", 104, 62, 330, 34)
GUICtrlSetState(-1, $GUI_FOCUS)
GUICtrlSetFont(-1, 18, $FW_BOLD, Default, "Consolas")
$aGUI[$idBtnPassphrase] = GUICtrlCreateButton("C&opy", 451, 62, 40, 34)
GUICtrlSetOnEvent(-1, "GUIEvents")
GUICtrlSetState(-1, $GUI_DISABLE)
$aGUI[$idLblPassphraseMsg] = GUICtrlCreateLabel("", 104, 96, 330, 40, $SS_CENTER)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTITALIC)
$aGUI[$idLblPassword] = GUICtrlCreateLabel("Password:", 30, 148, 80, 20)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTUNDER)
$aGUI[$idLblPasswordUse] = GUICtrlCreateLabel("Use to Encrypt", 24, 167, 100, 20)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 9, $FW_NORMAL, $GUI_FONTITALIC, "Times New Roman")
$aGUI[$idTxtPassword] = GUICtrlCreateInput("", 104, 142, 330, 34, BitOR($ES_READONLY, $SS_CENTER, $ES_PASSWORD))
;~ GUICtrlSetBkColor(-1, 0xFFFFFF)
idTxtPassword_Enable(False)
GUICtrlSetFont(-1, 18, $FW_BOLD, Default, "Consolas")
$aGUI[$idBtnPassword] = GUICtrlCreateButton("Co&py", 451, 142, 40, 34)
GUICtrlSetOnEvent(-1, "GUIEvents")
GUICtrlSetState(-1, $GUI_DISABLE)
$aGUI[$idLblPasswordMsg] = GUICtrlCreateLabel("", 104, 176, 330, 40, $SS_CENTER)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTITALIC)

GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_ACTIVATE, "WM_ACTIVATE")
GUISetOnEvent($GUI_EVENT_CLOSE, "GUIEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "GUIEvents")
GUISetOnEvent($GUI_EVENT_RESTORE, "GUIEvents")
TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "GUIShow")
#EndRegion - UI Creation

#Region - Main
;~ UpdatePassGen()

_Crypt_Startup()

_LegacyKeyConvert()

UILock()

KeyReadFromReg(RegistryKeyGetCurrent())
If KeyIsValid() Then
	UILock(False)
	idtxtPassphrase_Focus()
EndIf

;~ If AutoStartIsEnabled() Then GUICtrlSetState($aGUI[$idMnuOptionsAutoStart], $GUI_CHECKED)
If CloseToTrayIsEnabled() Then GUICtrlSetState($aGUI[$idMnuOptionsCloseToTray], $GUI_CHECKED)

If $CmdLineRaw = "/silent" Then
	GUIHide()
Else
	GUIShow()
EndIf

While 1
	Sleep(10)
WEnd
#EndRegion - Main

#Region - UI Event Functions
Func _Exit()
	_Crypt_Shutdown()
	$hTimer = 0
	ClipboardClear()
	Exit
EndFunc   ;==>_Exit

Func GUICtrl_Enable($iCtrl, $bFlag = True)
	Return GUICtrl_SetState($iCtrl, $bFlag, $GUI_ENABLE, $GUI_DISABLE)
EndFunc   ;==>GUICtrl_Enable

Func GUICtrl_GetState($iCtrl, $iState)
	Return BitAND(GUICtrlGetState($iCtrl), $iState)
EndFunc   ;==>GUICtrl_GetState

Func GUICtrl_MenuItemToggle($iCtrl)
	If BitAND(GUICtrl_Read($iCtrl), $GUI_CHECKED) Then
		Return GUICtrl_SetChecked($iCtrl, False)
	Else
		Return GUICtrl_SetChecked($iCtrl)
	EndIf
EndFunc   ;==>GUICtrl_MenuItemToggle

Func GUICtrl_Read($iCtrl)
	Return GUICtrlRead($iCtrl)
EndFunc   ;==>GUICtrl_Read

Func GUICtrl_SetChecked($iCtrl, $bFlag = True)
	Return GUICtrl_SetState($iCtrl, $bFlag, $GUI_CHECKED, $GUI_UNCHECKED)
EndFunc   ;==>GUICtrl_SetChecked

Func GUICtrl_SetData($iCtrl, $sValue)
	Return GUICtrlSetData($iCtrl, $sValue)
EndFunc   ;==>GUICtrl_SetData

Func GUICtrl_SetFocus($iCtrl, $bFlag = True)
	Return GUICtrl_SetState($iCtrl, $bFlag, $GUI_FOCUS)
EndFunc   ;==>GUICtrl_SetFocus

Func GUICtrl_SetState($iCtrl, $bEval, $iTrueValue, $iFalseValue = 0)
	If Not IsBool($bEval) Then Return SetError(1, 0, 0)
	Local $iState = ($bEval = True) ? $iTrueValue : $iFalseValue
	Return SetExtended(GUICtrlSetState($iCtrl, $iState), $iState)
EndFunc   ;==>GUICtrl_SetState

Func GUICtrl_Show($iCtrl, $bFlag = True)
	Return GUICtrl_SetState($iCtrl, $bFlag, $GUI_SHOW, $GUI_HIDE)
EndFunc   ;==>GUICtrl_Show

Func GUIEvents()
	$iCtrl = @GUI_CtrlId
	Switch $iCtrl
		Case $GUI_EVENT_CLOSE
			Return (CloseToTrayIsEnabled()) ? GUIHide() : _Exit()
		Case $GUI_EVENT_MINIMIZE
			GUIMinimize()
		Case $GUI_EVENT_RESTORE
			GUIRestore()
		Case $aGUI[$idMnuFileQuit]
			_Exit()
		Case $aGUI[$idMnuOptionsAutoStart]
			idMnuOptionsAutoStart_Click()
		Case $aGUI[$idMnuOptionsCloseToTray]
			idMnuOptionsCloseToTray_Click()
		Case $aGUI[$idBtnRevealKey]
			idBtnRevealKey_Click()
		Case $aGUI[$idBtnKey]
			idBtnKey_Click()
		Case $aGUI[$idBtnPassphrase]
			idBtnPassphrase_Click()
		Case $aGUI[$idBtnPassword]
			idBtnPassword_Click()
	EndSwitch
EndFunc   ;==>GUIEvents

Func GUIHide()
	GUIMinimize()
	TraySetState($TRAY_ICONSTATE_SHOW)
	GUISetState(@SW_HIDE)
	GUISetState(@SW_DISABLE)
EndFunc   ;==>GUIHide

Func GUIMinimize()
	idtxtPassphrase_Focus()
	KeyHide()
	GUISetState(@SW_MINIMIZE)
EndFunc   ;==>GUIMinimize

Func GUIRestore()
	GUISetState(@SW_RESTORE)
EndFunc   ;==>GUIRestore

Func GUIShow()
	GUIRestore()
	TraySetState($TRAY_ICONSTATE_HIDE)
	GUISetState(@SW_ENABLE)
	GUISetState(@SW_SHOW)
	WinActivate($aGUI[$hGUI])
EndFunc   ;==>GUIShow

Func TrayEvents()
	$iCtrl = @TRAY_ID
	Switch $iCtrl
		Case $aGUI[$idTrayOpen]
			GUIShow()
		Case $aGUI[$idTrayQuit]
			_Exit()
	EndSwitch
EndFunc   ;==>TrayEvents

Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	Local $iIDFrom = BitAND($wParam, 0xFFFF) ; LoWord - this gives the control which sent the message
	Local $iCode = BitShift($wParam, 16) ; HiWord - this gives the message that was sent
	Switch $iCode
		Case $EN_CHANGE ; If we have the correct message
			Switch $iIDFrom ; See if it comes from one of the inputs
				Case $aGUI[$idTxtKey]
					Return idTxtKey_OnChange()
				Case $aGUI[$idTxtPassphrase]
					Return idTxtPassphrase_OnChange()
			EndSwitch
		Case $EN_SETFOCUS
			Switch $iIDFrom
				Case $aGUI[$idTxtPassword]
					Return PasswordShow()
			EndSwitch
		Case $EN_KILLFOCUS
			Switch $iIDFrom
				Case $aGUI[$idTxtPassword]
					Return PasswordHide()
			EndSwitch
		Case $CBN_SELENDCANCEL ;Key Archive Combobox selection not made / canceled
			Switch $iIDFrom
				Case $aGUI[$idCmbKeyList]
					UILock(False)
			EndSwitch
		Case $CBN_SELENDOK ;Key Archive Combobox selection made
			Switch $iIDFrom
				Case $aGUI[$idCmbKeyList]
					$g_sActiveKeyGUID = KeyArchiveGetGUID(_GUICtrlComboBox_GetCurSel($aGUI[$idCmbKeyList]))
					idBtnKey_SetCaption("&Save")
					idCmbKeyList_Visible(False)
					KeyArchiveClearFromMem()
					idTxtKey_SetData(KeyArchiveParseValue($g_sActiveKeyGUID))
					RegistryKeySelect($g_sActiveKeyGUID)
					UILock(False)
					Return $g_sActiveKeyGUID
			EndSwitch
	EndSwitch
	Switch $wParam
		Case $KEYARCHIVEACTION_SELECT To $KEYARCHIVEACTION_CLEAR
			Return KeyArchiveOperation($wParam)
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func WM_ACTIVATE($hWnd, $iMsg, $wParam, $lParam)
	Local $iCode = BitAND($wParam, 0xFFFF)
	Switch $hWnd
		Case $aGUI[$hGUI]
			Switch $iCode
				Case 0 ; WA_INACTIVE
					If Not WinActive("PassGen") Then KeyHide()
					PasswordHide()
			EndSwitch
	EndSwitch
EndFunc   ;==>WM_ACTIVATE
#EndRegion - UI Event Functions

#Region - UI Control Functions
Func idBtnKey_Click()
	Switch idBtnKey_GetCaption()
		Case "&Change"
			KeyArchiveContextMenuCreate()
		Case "&Save"
			Switch $g_iKeyOperation
				Case $KEYARCHIVEACTION_ADD
					KeySave()
				Case $KEYARCHIVEACTION_DATECHANGE
					KeyDateSave()
				Case $KEYARCHIVEACTION_MODIFY
					KeySave()
			EndSwitch
		Case "&Cancel"
			Switch $g_iKeyOperation
				Case $KEYARCHIVEACTION_ADD
					$sCurrentKey = RegistryKeyGetCurrent()
					If Not @error Then
						KeyReadFromReg($sCurrentKey)
						UILock(False)
					Else
						idTxtKey_SetData("")
						idTxtKey_Enable(False)
						UILock()
						idBtnKey_SetCaption("&Change")
					EndIf
			EndSwitch
;~ 			UILock(False)
	EndSwitch
EndFunc   ;==>idBtnKey_Click

Func idBtnKey_Focus()
	GUICtrl_SetFocus($aGUI[$idBtnKey])
EndFunc   ;==>idBtnKey_Focus

Func idBtnKey_GetCaption()
	Return GUICtrl_Read($aGUI[$idBtnKey])
EndFunc   ;==>idBtnKey_GetCaption

Func idBtnKey_SetCaption($sCaption)
	GUICtrl_SetData($aGUI[$idBtnKey], $sCaption)
EndFunc   ;==>idBtnKey_SetCaption

Func idBtnPassphrase_Click()
	ClipboardCopyData(GUICtrl_Read($aGUI[$idTxtPassphrase]))
	idLblPasswordMsg_SetMessage()
	idLblPassphraseMsg_SetMessage("Copied to Clipboard")
	AutoPurgeTimer()
EndFunc   ;==>idBtnPassphrase_Click

Func idBtnPassphrase_Enabled($bFlag = True)
	GUICtrl_Enable($aGUI[$idBtnPassphrase], $bFlag)
EndFunc   ;==>idBtnPassphrase_Enabled

Func idBtnPassword_Click()
	ClipboardCopyData(GUICtrl_Read($aGUI[$idTxtPassword]))
	idLblPassphraseMsg_SetMessage()
	idLblPasswordMsg_SetMessage("Copied to Clipboard")
	AutoPurgeTimer()
EndFunc   ;==>idBtnPassword_Click

Func idBtnPassword_Enabled($bFlag = True)
	GUICtrl_Enable($aGUI[$idBtnPassword], $bFlag)
EndFunc   ;==>idBtnPassword_Enabled

Func idBtnRevealKey_AcceleratorOption($iFlag = 0)
	GUICtrl_SetData($aGUI[$idBtnRevealKey], ($iFlag = 1) ? "Sho&w" : "&Show")
EndFunc   ;==>idBtnRevealKey_AcceleratorOption

Func idBtnRevealKey_Depressed($bFlag = True)
	GUICtrl_SetChecked($aGUI[$idBtnRevealKey], $bFlag)
EndFunc   ;==>idBtnRevealKey_Depressed

Func idBtnRevealKey_Enabled($bFlag = True)
	GUICtrl_Enable($aGUI[$idBtnRevealKey], $bFlag)
EndFunc   ;==>idBtnRevealKey_Enabled

Func idBtnRevealKey_Click()
	Local $iState = GUICtrl_Read($aGUI[$idBtnRevealKey])
	If $iState = $GUI_UNCHECKED Then
		KeyHide()
	Else
		KeyShow()
	EndIf
EndFunc   ;==>idBtnRevealKey_Click

Func idCmbKeyList_Clear()
	_GUICtrlComboBox_ResetContent($aGUI[$idCmbKeyList])
EndFunc   ;==>idCmbKeyList_Clear

Func idCmbKeyList_Visible($bFlag = True)
	GUICtrl_Show($aGUI[$idCmbKeyList], $bFlag)
EndFunc   ;==>idCmbKeyList_Visible

Func idCmbKeyList_Expand($bFlag = True)
	_GUICtrlComboBox_ShowDropDown($aGUI[$idCmbKeyList], $bFlag)
	If $bFlag = True Then idCmbKeyList_Focus()
EndFunc   ;==>idCmbKeyList_Expand

Func idCmbKeyList_Focus()
	GUICtrl_SetFocus($aGUI[$idCmbKeyList])
EndFunc   ;==>idCmbKeyList_Focus

Func idDateKeyDatePicker_Read()
	Return GUICtrl_Read($aGUI[$idDateKeyDatePicker])
EndFunc   ;==>idDateKeyDatePicker_Read

Func idDateKeyDatePicker_SetData($sValue)
	GUICtrl_SetData($aGUI[$idDateKeyDatePicker], $sValue)
EndFunc   ;==>idDateKeyDatePicker_SetData

Func idDateKeyDatePicker_Visible($bFlag = True)
	GUICtrl_Show($aGUI[$idDateKeyDatePicker], $bFlag)
EndFunc   ;==>idDateKeyDatePicker_Visible

Func idLblPassphraseMsg_SetMessage($sMsg = "")
	GUICtrl_SetData($aGUI[$idLblPassphraseMsg], $sMsg)
EndFunc   ;==>idLblPassphraseMsg_SetMessage

Func idLblPasswordMsg_SetMessage($sMsg = "")
	GUICtrl_SetData($aGUI[$idLblPasswordMsg], $sMsg)
EndFunc   ;==>idLblPasswordMsg_SetMessage

Func idMnuOptionsAutoStart_Click()
	Local $iState = GUICtrl_MenuItemToggle($aGUI[$idMnuOptionsAutoStart])
;~ 	AutoStart(($iState == $GUI_CHECKED) ? True : False)
EndFunc   ;==>idMnuOptionsAutoStart_Click

Func idMnuOptionsCloseToTray_Click()
	Local $iState = GUICtrl_MenuItemToggle($aGUI[$idMnuOptionsCloseToTray])
	CloseToTraySetting(($iState == $GUI_CHECKED) ? True : False)
EndFunc   ;==>idMnuOptionsCloseToTray_Click

Func idTxtKey_Enable($bFlag = True)
	GUICtrl_Enable($aGUI[$idTxtKey], $bFlag)
EndFunc   ;==>idTxtKey_Enable

Func idTxtKey_Focus()
	GUICtrl_SetFocus($aGUI[$idTxtKey])
EndFunc   ;==>idTxtKey_Focus

Func idTxtKey_IsEnabled()
	Return (GUICtrl_GetState($aGUI[$idTxtKey], $GUI_ENABLE)) ? True : False
EndFunc   ;==>idTxtKey_IsEnabled

Func idTxtKey_OnChange()
	AutoPurgeTimer(False)
	KeyIsValid()
EndFunc   ;==>idTxtKey_OnChange

Func idTxtKey_Read()
	Return GUICtrl_Read($aGUI[$idTxtKey])
EndFunc   ;==>idTxtKey_Read

Func idTxtKey_SetData($sValue)
	GUICtrl_SetData($aGUI[$idTxtKey], $sValue)
EndFunc   ;==>idTxtKey_SetData

Func idTxtKey_Visible($bFlag = True)
	GUICtrl_Show($aGUI[$idTxtKey], $bFlag)
EndFunc   ;==>idTxtKey_Visible

Func idTxtPassphrase_Enable($bFlag = True)
	GUICtrl_Enable($aGUI[$idTxtPassphrase], $bFlag)
EndFunc   ;==>idTxtPassphrase_Enable

Func idtxtPassphrase_Focus()
	GUICtrl_SetFocus($aGUI[$idTxtPassphrase])
EndFunc   ;==>idtxtPassphrase_Focus

Func idTxtPassphrase_OnChange()
	AutoPurgeTimer(False)
	If PassphraseIsValid() Then
		GUICtrl_Enable($aGUI[$idBtnPassphrase])
		GUICtrl_Enable($aGUI[$idBtnPassword])
		GeneratePassword()
	Else
		GUICtrl_Enable($aGUI[$idBtnPassphrase], False)
		GUICtrl_Enable($aGUI[$idBtnPassword], False)
	EndIf
EndFunc   ;==>idTxtPassphrase_OnChange

Func idTxtPassphrase_Read()
	Return GUICtrl_Read($aGUI[$idTxtPassphrase])
EndFunc   ;==>idTxtPassphrase_Read

Func idTxtPassword_Enable($bFlag = True)
	GUICtrlSetBkColor($aGUI[$idTxtPassword], ($bFlag = True) ? 0xFFFFFF : 0xF0F0F0)
	GUICtrl_Enable($aGUI[$idTxtPassword], $bFlag)
EndFunc   ;==>idTxtPassword_Enable

Func idTxtPassword_Read()
	Return GUICtrl_Read($aGUI[$idTxtPassword])
EndFunc   ;==>idTxtPassword_Read

Func idTxtPassword_SetData($sValue)
	GUICtrl_SetData($aGUI[$idTxtPassword], $sValue)
EndFunc   ;==>idTxtPassword_SetData
#EndRegion - UI Control Functions

#Region - Additonal Functions
Func AutoPurgeClipboard()
	If (TimerDiff($hTimer) >= (1000 * 60 * $AUTOPURGETIME)) Or (ClipboardContainsPass() = False) Then
		$hTimer = TimerInit
		idLblPassphraseMsg_SetMessage("")
		idLblPasswordMsg_SetMessage("")
		If ClipboardContainsPass() Then
			ClipboardClear()
		EndIf
	EndIf
EndFunc   ;==>AutoPurgeClipboard

Func AutoPurgeTimer($bFlag = True)
	Switch $bFlag
		Case True
			$hTimer = TimerInit()
			AdlibRegister("AutoPurgeClipboard", 10000)
		Case False
			$hTimer = 0
			AdlibUnRegister("AutoPurgeClipboard")
	EndSwitch
EndFunc   ;==>AutoPurgeTimer

Func AutoStart($bEnable = True)
	Switch $bEnable
		Case True
			FileCopy(@ScriptFullPath, $PROGRAMPATH, $FC_CREATEPATH + $FC_OVERWRITE)
			FileCreateShortcut($PROGRAMPATH, $STARTUPLINK, "", "/silent")
		Case False
			If FileExists($STARTUPLINK) Then Return FileDelete($STARTUPLINK)
	EndSwitch
EndFunc   ;==>AutoStart

Func AutoStartIsEnabled()
	Local $aShortcut = FileGetShortcut($STARTUPLINK)
	If @error Then Return False
	Return (FileExists($aShortcut[0])) ? True : False
EndFunc   ;==>AutoStartIsEnabled

Func ClipboardClear()
	ClipPut("")
EndFunc   ;==>ClipboardClear

Func ClipboardContainsPass()
	$sClipboardText = ClipboardGet()
	If $sClipboardText == idTxtPassphrase_Read() Then Return SetError(0, 1, True)
	If $sClipboardText == idTxtPassword_Read() Then Return SetError(0, 2, True)
	Return False
EndFunc   ;==>ClipboardContainsPass

Func ClipboardCopyData($vData)
	ClipPut($vData)
EndFunc   ;==>ClipboardCopyData

Func ClipboardGet()
	Return ClipGet()
EndFunc   ;==>ClipboardGet

Func CloseToTrayIsEnabled()
	RegRead($REGKEYPATH, "NoCloseToTray")
	Return (@error <> 0) ? True : False
EndFunc   ;==>CloseToTrayIsEnabled

Func CloseToTraySetting($bEnable = True)
	Switch $bEnable
		Case True
			RegDelete($REGKEYPATH, "NoCloseToTray")
		Case False
			RegWrite($REGKEYPATH, "NoCloseToTray", "REG_DWORD", 0)
	EndSwitch
EndFunc   ;==>CloseToTraySetting

Func GeneratePassword() ;Create Password based on Key & Passphrase hash
	$sKey = KeyGetValue()
	$sPassphrase = PassphraseGetValue()
	$dHash = _Crypt_HashData($sKey & "-" & $sPassphrase, $CALG_SHA1)
	$sPassword = ""
	For $iX = 1 To 20
		$hByte = BinaryMid($dHash, $iX, 1)
		$iChar = Mod(Int(Dec(StringRight($hByte, 2))), $CHARACTERLISTLEN)
		$sPassword &= StringMid($CHARACTERLIST, $iChar + 1, 1)
	Next
	$sPassword = StringMid($sPassword, 1, 4) & "-" & StringMid($sPassword, 5, 4) & "-" & StringMid($sPassword, 9, 4) & "-" & StringMid($sPassword, 13, 4) & "-" & StringMid($sPassword, 17, 4)
	idTxtPassword_SetData($sPassword)
	idBtnPassword_Click()
EndFunc   ;==>GeneratePassword

Func InputboxMask($iCtrl, $bMask = True)
	Switch $bMask
		Case False
			GUICtrlSendMsg($iCtrl, $EM_SETPASSWORDCHAR, 0, 0)
		Case True
			GUICtrlSendMsg($iCtrl, $EM_SETPASSWORDCHAR, $ES_PASSWORDCHAR, 0)
	EndSwitch
	Local $aRes = DllCall("user32.dll", "int", "RedrawWindow", "hwnd", GUICtrlGetHandle($iCtrl), "ptr", 0, "ptr", 0, "dword", 5)
EndFunc   ;==>InputboxMask

Func IsStringComplex($sText, $iMinLength = 8, $iMinReq = 4, $iReqFlags = Default)
	;Input Parameter Error Checking
	$sText = String($sText)
;~ 	If Not StringLen($sText) Then Return SetError(1, 0, Null) ;$sText length is less than 1
	If Not IsInt($iMinLength) Or $iMinLength < 0 Then Return SetError(2, 0, Null) ;$iMinLen is not an integer Or is less then 0
	If Not IsInt($iMinReq) Then Return SetError(3, 0, Null) ;$iMinReq is not an intger
	If $iMinReq < 0 Or $iMinReq > 4 Then Return SetError(3, 1, Null) ;$iMinReq integer is outside acceptable range
	If $iReqFlags = Default Then $iReqFlags = BitOR($COMPLEXITY_UPPER, _
			$COMPLEXITY_LOWER, $COMPLEXITY_NUMBER, $COMPLEXITY_SYMBOL)
	If Not IsInt($iReqFlags) Then Return SetError(4, 0, Null) ;$iReqFlags is not an integer
	If $iReqFlags < 0 Or $iReqFlags > 15 Then Return SetError(4, 1, Null) ;$iReqFlags integer is outside acceptable range
	Local $iReqChecksCount = 0
	If BitAND($iReqFlags, $COMPLEXITY_UPPER) Then $iReqChecksCount += 1
	If BitAND($iReqFlags, $COMPLEXITY_LOWER) Then $iReqChecksCount += 1
	If BitAND($iReqFlags, $COMPLEXITY_UPPER) Then $iReqChecksCount += 1
	If BitAND($iReqFlags, $COMPLEXITY_UPPER) Then $iReqChecksCount += 1
	If $iReqChecksCount > $iMinReq Then Return SetError(5, 0, Null) ;$iReqFlags is not compatible with $iMinReq - More Required Flags set than Minimum Required

	;First String Complexity Test - String Length
	If Not (StringLen($sText) >= $iMinLength) Then Return SetError(1, 0, False) ;Failed due to string length being less than $iMinLength

	;Conduct Remaining String Complexity Test
	Local $iTestPassCount = 0
	Local $iFailedReq = 0
	Local $bTest
	$bTest = (StringRegExp($sText, "[[:upper:]]")) ? True : False
	If $bTest Then $iTestPassCount += 1
	If BitAND($iReqFlags, $COMPLEXITY_UPPER) And $bTest = False Then $iFailedReq += $COMPLEXITY_UPPER
	$bTest = (StringRegExp($sText, "[[:lower:]]")) ? True : False
	If $bTest Then $iTestPassCount += 1
	If BitAND($iReqFlags, $COMPLEXITY_LOWER) And $bTest = False Then $iFailedReq += $COMPLEXITY_LOWER
	$bTest = (StringRegExp($sText, "[[:digit:]]")) ? True : False
	If $bTest Then $iTestPassCount += 1
	If BitAND($iReqFlags, $COMPLEXITY_NUMBER) And $bTest = False Then $iFailedReq += $COMPLEXITY_NUMBER
	$bTest = (StringRegExp($sText, "[[:punct:]]")) ? True : False
	If $bTest Then $iTestPassCount += 1
	If BitAND($iReqFlags, $COMPLEXITY_SYMBOL) And $bTest = False Then $iFailedReq += $COMPLEXITY_SYMBOL

	;Evaluate Test Results
	If $iTestPassCount < $iMinReq Then Return SetError(2, 0, False) ;Failed due to less than $iMinReq
	If $iFailedReq Then Return SetError(3, $iFailedReq, False) ;Failed due to one or more $iRegFlags failing

	Return True
EndFunc   ;==>IsStringComplex

Func KeyAdd()
	idTxtKey_SetData("")
	KeyChange()
	KeyIsValid()
EndFunc   ;==>KeyAdd

Func KeyArchiveClear($bForce = False)
	If $bForce = False Then
		If MsgBox(BitOR($MB_ICONWARNING, $MB_YESNO, $MB_DEFBUTTON2), "Clear Key Archive", "Are you absolutely sure you want to clear the Key Archive?" & @CRLF & @CRLF & "This action can not be undone!") <> $IDYES Then Return 0
	EndIf
	KeyArchiveGet()
	For $sKey In $g_aKeyArchive
		RegDelete($REGKEYPATH, $sKey)
	Next
	RegDelete($REGKEYPATH, $REGKEYCURRENT)
	idTxtKey_SetData("")
	idTxtKey_OnChange()
	UILock()
	KeyArchiveClearFromMem()
EndFunc   ;==>KeyArchiveClear

Func KeyArchiveClearFromMem()
	$g_aKeyArchive = 0
EndFunc   ;==>KeyArchiveClearFromMem

Func KeyArchiveExportCanceled($sCustomMsg = "")
	If $sCustomMsg <> "" Then
		$sCustomMsg &= @CRLF & @CRLF
	EndIf
	$sCustomMsg &= "Export Canceled"
	MsgBox(0, "", $sCustomMsg)
	Return SetError(WinActivate($aGUI[$hGUI]), 0, 0)
EndFunc   ;==>KeyArchiveExportCanceled

Func KeyArchiveExportKey($sDate, $sKey)
	If Not _DateIsValid($sDate) Then Return SetError(1, 0, 0)
	If Not IsString($sKey) Then Return SetError(2, 0, 0)
	$dKeyBytes = StringToBinary($sKey, 4)
	$iKeyByteLen = StringToBinary(BinaryLen($dKeyBytes))
	Return $EXPORTFILEKEYENTRYHEADER_KEYDATE & StringToBinary($sDate) & $EXPORTFILEKEYENTRYHEADER_KEYVALUE & $dKeyBytes
;~ 	Return $EXPORTFILEKEYENTRYHEADER_KEYDATE & StringToBinary($sDate) & $EXPORTFILEKEYENTRYHEADER_KEYLENGTH & $iKeyByteLen & $EXPORTFILEKEYENTRYHEADER_KEYVALUE & $dKeyBytes
EndFunc   ;==>KeyArchiveExportKey

Func KeyArchiveExportRoutine()
	Local $sPassGenExportFilePath = FileSaveDialog("PassGen Key Archive Export", @MyDocumentsDir & "\", "PassGen Key Archive Export(*.pge)", $FD_PROMPTOVERWRITE, "PassGenKeyArchive.pge")
	If @error Then Return KeyArchiveExportCanceled()
	Local $sPGEPassword = ""

	Do
		$sPGEPassword = InputBox("Enter Export Password", "Enter a password to protect the export." & @CRLF & @CRLF & _
				"The password must have at least 8 characters and contain at least 3 of the 4 requirements:" & @CRLF & @CRLF & "Upper case, Lower case, Number, or Symbol.", $sPGEPassword)
		If @error = 1 Then
			Return KeyArchiveExportCanceled()
		EndIf
		Local $bPasswordIsValid = IsStringComplex($sPGEPassword, 8, 3, 0)
		If Not $bPasswordIsValid Then MsgBox(BitOR($MB_ICONERROR, $MB_TASKMODAL), "Password Error", "Password does not meet complexity requirements.")
	Until $bPasswordIsValid

	$sBIN = $EXPORTFILEHEADER

	KeyArchiveGet()
	For $iKey In $g_aKeyArchive
		$sBIN = $sBIN & KeyArchiveExportKey(KeyArchiveParseDate($iKey), KeyArchiveParseValue($iKey))
	Next
	$sBIN = $sBIN & $EXPORTFILEEOF

	Local $dEncryptionKey = _Crypt_DeriveKey(StringToBinary($sPGEPassword), $CALG_AES_256)
	$sPGEPassword = ""
	If @error Then Return KeyArchiveExportCanceled("An unexpected error occured while preparing the Encryption key")

	$hExportFile = FileOpen($sPassGenExportFilePath, BitOR($FO_OVERWRITE, $FO_CREATEPATH, $FO_BINARY))
	If $hExportFile = -1 Then Return KeyArchiveExportCanceled("An unexpected error occured while attempting to create the export file")
	FileWrite($hExportFile, $EXPORTFILEENCRYPTEDHEADER)
	FileWrite($hExportFile, _Crypt_EncryptData($sBIN, $dEncryptionKey, $CALG_USERKEY))
	_Crypt_DestroyKey($dEncryptionKey) ; Destroy the cryptographic key.
	FileClose($hExportFile)
	MsgBox(0, "Export Complete", "PassGen Key Archive exported successfully", 5)
EndFunc   ;==>KeyArchiveExportRoutine

Func KeyArchiveGet()
	Local $iRegValIndex = 1
	Local $aKeyArchive[0][2]
	While 1
		Local $sRegValName = RegEnumVal($REGKEYPATH, $iRegValIndex)
		If @error Then ExitLoop
		If StringLeft($sRegValName, 1) == "{" And StringRight($sRegValName, 1) == "}" Then
			Local $sRegValue = KeyUnprotect(RegRead($REGKEYPATH, $sRegValName))
			_ArrayAdd($aKeyArchive, $sRegValName & "|" & $sRegValue)
		EndIf
		$iRegValIndex += 1
	WEnd
	$g_aKeyArchive = $aKeyArchive
	_ArraySort($g_aKeyArchive, 1, 0, 0, 1)
	_ArrayColDelete($g_aKeyArchive, 1, True)
	Return 1
EndFunc   ;==>KeyArchiveGet

Func KeyArchiveGetCount()
	If Not IsArray($g_aKeyArchive) Then KeyArchiveGet()
	If IsArray($g_aKeyArchive) Then
		Return UBound($g_aKeyArchive)
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc   ;==>KeyArchiveGetCount

Func KeyArchiveGetGUID($iKeyIndex)
	If Not IsArray($g_aKeyArchive) Then Return SetError(1, 0, "")
	Return $g_aKeyArchive[$iKeyIndex]
EndFunc   ;==>KeyArchiveGetGUID

Func KeyArchiveListClear()
	idCmbKeyList_Clear()
EndFunc   ;==>KeyArchiveListClear

Func KeyArchiveListDisplay()
	UILock()
	idTxtKey_Visible(False)
	idBtnKey_SetCaption("&Cancel")
	KeyArchiveListClear()
	If KeyArchiveGet() Then
		For $iX = 0 To KeyArchiveGetCount() -1
			_GUICtrlComboBox_AddString($aGUI[$idCmbKeyList], KeyArchiveParseDate($g_aKeyArchive[$iX]) & " - " & KeyArchiveParseValue($g_aKeyArchive[$iX]))
		Next
	EndIf
	idCmbKeyList_Visible()
	idCmbKeyList_Expand()
EndFunc   ;==>KeyArchiveListDisplay

Func KeyArchiveContextMenuCreate() ;Dynamic Key Archive context menu
	Local $idContextMenu = _GUICtrlMenu_CreatePopup($MNS_AUTODISMISS)
	Local $iKeyArchiveCount = KeyArchiveGetCount()
	If $iKeyArchiveCount > 1 Then
		_GUICtrlMenu_AddMenuItem($idContextMenu, "&Select Key from Archive", $KEYARCHIVEACTION_SELECT)
	EndIf
	_GUICtrlMenu_AddMenuItem($idContextMenu, "&Add Key to Archive", $KEYARCHIVEACTION_ADD)
	If $iKeyArchiveCount = 1 Then
		_GUICtrlMenu_AddMenuItem($idContextMenu, "&Remove Key", $KEYARCHIVEACTION_REMOVE)
		_GUICtrlMenu_AddMenuItem($idContextMenu, "&Modify Key", $KEYARCHIVEACTION_MODIFY)
	ElseIf $iKeyArchiveCount > 1 Then
		_GUICtrlMenu_AddMenuItem($idContextMenu, "&Remove Key from Archive", $KEYARCHIVEACTION_REMOVE)
		_GUICtrlMenu_AddMenuItem($idContextMenu, "&Modify Key in Archive", $KEYARCHIVEACTION_MODIFY)
		_GUICtrlMenu_AddMenuItem($idContextMenu, "&Change Key Date", $KEYARCHIVEACTION_DATECHANGE)
	EndIf
	_GUICtrlMenu_AddMenuItem($idContextMenu, "")
	_GUICtrlMenu_AddMenuItem($idContextMenu, "&Import Key Archive", $KEYARCHIVEACTION_IMPORT)
	If $iKeyArchiveCount > 0 Then
		_GUICtrlMenu_AddMenuItem($idContextMenu, "&Export Key Archive", $KEYARCHIVEACTION_EXPORT)
	EndIf
	If $iKeyArchiveCount > 1 Then
		_GUICtrlMenu_AddMenuItem($idContextMenu, "Clear Key Archive", $KEYARCHIVEACTION_CLEAR)
	EndIf

	$aCtrlPos = ControlGetPos($aGUI[$hGUI], "", $aGUI[$idBtnKey])
	$aWinPos = WinGetPos($aGUI[$hGUI])
	_GUICtrlMenu_TrackPopupMenu($idContextMenu, $aGUI[$hGUI], $aWinPos[0] + $aCtrlPos[0], $aWinPos[1] + $aCtrlPos[1] + 40)
	_GUICtrlMenu_DestroyMenu($idContextMenu)
EndFunc   ;==>KeyArchiveContextMenuCreate

Func KeyArchiveImportCanceled($sCustomMsg = "")
	If $sCustomMsg <> "" Then
		$sCustomMsg &= @CRLF & @CRLF
	EndIf
	$sCustomMsg &= "Import Canceled"
	MsgBox(0, "", $sCustomMsg)
	Return SetError(WinActivate($aGUI[$hGUI]), 0, 0)
EndFunc   ;==>KeyArchiveImportCanceled

Func KeyArchiveImportFileParse(ByRef $dExportFileData) ;Routine to Parse Binary Key Archive File [decrypted]
	Local $iExportFileLen = BinaryLen($dExportFileData)
	Local $iBinaryLocation = $EXPORTFILEHEADERLEN + 1
	Local $iKeyDateLocation, $iKeyValueStartLocation, $iKeyValueEndLocation
	If Not _BinaryCheckRecordHeader($dExportFileData, $iBinaryLocation, $EXPORTFILEKEYENTRYHEADER_KEYDATELEN, $EXPORTFILEKEYENTRYHEADER_KEYDATE) Then Return SetError(1, 0, 0) ;Record header not found where expected.
	Local $bErr = False
	Local $bEOF = False
	Local $dKeys
	Do
		$iBinaryLocation += $EXPORTFILEKEYENTRYHEADER_KEYDATELEN
		$iKeyDateLocation = $iBinaryLocation
		$iBinaryLocation += $DATEFORMATBYTELEN
		If Not _BinaryCheckValueHeader($dExportFileData, $iBinaryLocation, $EXPORTFILEKEYENTRYHEADER_KEYVALUELEN, $EXPORTFILEKEYENTRYHEADER_KEYVALUE) Then
			$bErr = True
		Else
			$iBinaryLocation += $EXPORTFILEKEYENTRYHEADER_KEYVALUELEN
			$iKeyValueStartLocation = $iBinaryLocation
			Local $dByte = 0
			Do
				$iBinaryLocation += 1
				$dByte = _BinaryExtract($dExportFileData, $iBinaryLocation, 1)
				If _BinaryCompare($dByte, $BINARYFORMAT_ETX) Then
					If _BinaryCheckEOF($dExportFileData, $iBinaryLocation, $EXPORTFILEEOFLEN, $EXPORTFILEEOF) Then
						$bEOF = True
					EndIf
				ElseIf $iBinaryLocation > $iExportFileLen - $EXPORTFILEEOFLEN + 1 Then
					$bErr = True
					ExitLoop
				EndIf
			Until _BinaryCompare($dByte, $BINARYFORMAT_RS) Or _BinaryCompare($dByte, $BINARYFORMAT_ETX)
			$iKeyValueEndLocation = $iBinaryLocation
			If Not $bErr Then
				$dRecord = _BinaryExtract($dExportFileData, $iKeyDateLocation, $DATEFORMATBYTELEN) & _
						_BinaryExtract($dExportFileData, $iKeyValueStartLocation, $iKeyValueEndLocation - $iKeyValueStartLocation)
				If Not _BinaryCheckEOF($dExportFileData, $iBinaryLocation, $EXPORTFILEEOFLEN, $EXPORTFILEEOF) Then $dRecord &= StringToBinary("|")
				$dKeys = Binary($dKeys) & Binary($dRecord)
			EndIf
		EndIf
	Until $bEOF = True Or $bErr = True
	Return ($bEOF == True And $bErr == False) ? $dKeys : SetError(2, $bErr, -1) ;Failed to reach EOF and/or an error occured.
EndFunc   ;==>KeyArchiveImportFileParse

Func KeyArchiveImportFileValidate($dExportFile) ;Routine to Validate Key Archive file format [decrypted]
	Local $iExportFileLen = BinaryLen($dExportFile)
	Return (_BinaryCheckFileHeader($dExportFile, 1, $EXPORTFILEHEADERLEN, $EXPORTFILEHEADER) And _
			_BinaryCheckEOF($dExportFile, $iExportFileLen - ($EXPORTFILEEOFLEN - 1), $EXPORTFILEEOFLEN, $EXPORTFILEEOF)) ? True : False
EndFunc   ;==>KeyArchiveImportFileValidate

Func KeyArchiveImportRoutine()
	Local $sPassGenExportFilePath = FileOpenDialog("PassGen Key Archive Export", @MyDocumentsDir & "\", "PassGen Key Archive Export(*.pge)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST))
	If @error Then Return KeyArchiveImportCanceled()
	If Not FileExists($sPassGenExportFilePath) Then Return KeyArchiveImportCanceled("Unable to locate export file")
	$hExportFile = FileOpen($sPassGenExportFilePath, $FO_BINARY)
	$dExportFileData = FileRead($hExportFile)
	FileClose($hExportFile)

	$bExportEncryptedFileHeaderValid = _BinaryCheckEncryptedFileHeader($dExportFileData, 1, $EXPORTFILEENCRYPTEDHEADERLEN, $EXPORTFILEENCRYPTEDHEADER)
	If Not $bExportEncryptedFileHeaderValid Then
		Return KeyArchiveImportCanceled("File does not appear to be a valid PassGen Key Archive Export")
	EndIf

	Local $sPGEPassword = InputBox("Enter Export Password", "Enter the password that was originally to protect the key archive export.")
	Local $dDecryptionKey = _Crypt_DeriveKey(StringToBinary($sPGEPassword), $CALG_AES_256)
	$sPGEPassword = ""
	If @error Then Return KeyArchiveImportCanceled("An unexpected error occured while preparing the Decryption key")

	$dExportFileData = _BinaryExtract($dExportFileData, $EXPORTFILEENCRYPTEDHEADERLEN + 1, BinaryLen($dExportFileData) - $EXPORTFILEENCRYPTEDHEADERLEN)

	$dExportFileData = _Crypt_DecryptData($dExportFileData, $dDecryptionKey, $CALG_USERKEY)
	_Crypt_DestroyKey($dDecryptionKey) ; Destroy the cryptographic key.

	$bExportFileValid = KeyArchiveImportFileValidate($dExportFileData)
	If $bExportFileValid Then
		$dKeys = KeyArchiveImportFileParse($dExportFileData)
		If @error Then KeyArchiveImportCanceled("An unexpected error occured while processing the PassGen Key Archive")
	Else
		$hExportFile = 0
		$dExportFileData = 0
		Return KeyArchiveImportCanceled("Failed to decrypt PassGen Key Archive with the provided password")
	EndIf
	$hExportFile = 0
	If KeyArchiveGetCount() Then
		$iImportOverwrite = MsgBox(BitOR($MB_YESNOCANCEL, $MB_ICONQUESTION), "Import Key Archive", _
				"Do you want to overwrite the existing Key Archive?" & @CRLF & @CRLF & "Yes = Overwrite, No = Append")
		If $iImportOverwrite = $IDYES Then KeyArchiveClear(True)
		If $iImportOverwrite = $IDCANCEL Then
			$dKeys = 0
			Return KeyArchiveImportCanceled()
		EndIf
	EndIf
	Local $aKeys = StringSplit(BinaryToString($dKeys), "|", $STR_NOCOUNT)
	For $sKey In $aKeys
		Local $sKeyGUID = _WinAPI_CreateGUID()
		Local $sDate = StringLeft($sKey, $DATEFORMATBYTELEN)
		Local $sValue = StringRight($sKey, StringLen($sKey) - $DATEFORMATBYTELEN)
		KeySaveToReg($sKeyGUID, StringRight($sKey, StringLen($sKey) - $DATEFORMATBYTELEN), StringLeft($sKey, $DATEFORMATBYTELEN))
	Next
	KeyArchiveGet()
	RegistryKeySelect($g_aKeyArchive[0])
	KeyReadFromReg(RegistryKeyGetCurrent())
	KeyArchiveClearFromMem()
	idBtnKey_SetCaption("&Change")
	idBtnRevealKey_AcceleratorOption()
	idBtnRevealKey_Depressed(False)
	KeyHide()
	UILock(False)
	idtxtPassphrase_Focus()
	idTxtPassphrase_OnChange()
	MsgBox(0, "Import Complete", "PassGen Key Archive imported successfully", 5)
EndFunc   ;==>KeyArchiveImportRoutine

Func KeyArchiveOperation($iKeyOperation = 0)
	$g_iKeyOperation = $iKeyOperation
	Switch $iKeyOperation
		Case $KEYARCHIVEACTION_ADD
			KeyAdd()
		Case $KEYARCHIVEACTION_CLEAR
			KeyArchiveClear()
		Case $KEYARCHIVEACTION_DATECHANGE
			KeyDateChange()
		Case $KEYARCHIVEACTION_EXPORT
			KeyArchiveExportRoutine()
		Case $KEYARCHIVEACTION_IMPORT
			KeyArchiveImportRoutine()
		Case $KEYARCHIVEACTION_MODIFY
			KeyChange()
		Case $KEYARCHIVEACTION_REMOVE
			KeyRemove()
		Case $KEYARCHIVEACTION_SELECT
			KeyArchiveListDisplay()
	EndSwitch
EndFunc   ;==>KeyArchiveOperation

Func KeyArchiveParseDate($sKeyGUID)
	Local $sKey = RegistryKeyRead($sKeyGUID)
	If @error Then Return SetError(@error, 0, "")
	$sKey = KeyUnprotect($sKey)
	Return StringLeft($sKey, $DATEFORMATBYTELEN)
EndFunc   ;==>KeyArchiveParseDate

Func KeyArchiveParseValue($sKeyGUID)
	Local $sKey = RegistryKeyRead($sKeyGUID)
	If @error Then Return SetError(@error, 0, "")
	$sKey = KeyUnprotect($sKey)
	Return StringRight($sKey, StringLen($sKey) - $DATEFORMATBYTELEN)
EndFunc   ;==>KeyArchiveParseValue

Func KeyArchiveSelectionMade($iKeyOperation)
	Switch $iKeyOperation
		Case $KEYARCHIVEACTION_SELECT
			idTxtKey_SetData(KeyArchiveParseValue($g_sActiveKeyGUID))
			RegistryKeySelect($g_sActiveKeyGUID)
			UILock(False)
	EndSwitch
EndFunc   ;==>KeyArchiveSelectionMade

Func KeyChange()
	UILock()
	idTxtKey_Visible()
	idTxtKey_Enable()
	idTxtKey_Focus()
	If StringLen(idTxtKey_Read()) Then
		idBtnKey_SetCaption("&Save")
	Else
		idBtnKey_SetCaption("&Cancel")
	EndIf
	idBtnRevealKey_AcceleratorOption(1)
	idBtnRevealKey_Depressed()
	KeyShow()
EndFunc   ;==>KeyChange

Func KeyDateChange()
	UILock()
	$g_iKeyOperation = $KEYARCHIVEACTION_DATECHANGE
	idTxtKey_Visible(False)
	idDateKeyDatePicker_SetData(KeyArchiveParseDate($g_sActiveKeyGUID))
	idDateKeyDatePicker_Visible(True)
	idBtnKey_SetCaption("&Save")
EndFunc   ;==>KeyDateChange

Func KeyDateSave()
	Local $sKeyGUID = RegistryKeyGetCurrent()
	Local $sKeyValue = idTxtKey_Read()
	Local $sKeyDate = idDateKeyDatePicker_Read()
	KeySaveToReg($sKeyGUID, $sKeyValue, $sKeyDate)
	KeyArchiveClearFromMem()
	idDateKeyDatePicker_Visible(False)
	idTxtKey_Visible()
	UILock(False)
EndFunc   ;==>KeyDateSave

Func KeyGetValue()
	Return StringStripWS(idTxtKey_Read(), $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>KeyGetValue

Func KeyHide()
	idBtnRevealKey_Depressed(False)
	InputboxMask($aGUI[$idTxtKey])
EndFunc   ;==>KeyHide

Func KeyIsValid()
	If Not idTxtKey_IsEnabled() And Not idTxtKey_Read() Then Return 0
	idLblPassphraseMsg_SetMessage()
	idLblPasswordMsg_SetMessage()
	Local $sKey = KeyGetValue()
	If StringLen($sKey) = 0 Then Return 0
	$bIsKeyComplex = IsStringComplex($sKey, 8, 3, 0)
	If $bIsKeyComplex = False Then
		Switch @error
			Case 1
				idLblPassphraseMsg_SetMessage("Key is too short." & @CRLF & "Must contain at least 8 characters.")
			Case 2
				idLblPassphraseMsg_SetMessage("Key must contain 3 of the 4 requirements:" & @CRLF & "Upper case, Lower case, Number, or Symbol.")
		EndSwitch
		idBtnKey_SetCaption("&Cancel")
	Else
		idBtnKey_SetCaption("&Save")
	EndIf
	If $bIsKeyComplex == -1 Then Return False
	Return $bIsKeyComplex
EndFunc   ;==>KeyIsValid

Func KeyModify()
	KeyChange()
EndFunc   ;==>KeyModify

Func KeyProtect($sValue)
	Return _CryptProtectData($sValue)
EndFunc   ;==>KeyProtect

Func KeyReadFromReg($sKeyGUID)
	$hProtectedKey = RegistryKeyRead($sKeyGUID)
	If @error Then Return 0
	$sKey = KeyUnprotect($hProtectedKey)
	idTxtKey_SetData(StringRight($sKey, StringLen($sKey) - $DATEFORMATBYTELEN))
	Return 1
EndFunc   ;==>KeyReadFromReg

Func KeyRemove()
	UILock()
	Local $sKeyGUID = RegistryKeyGetCurrent()
	Local $sWarningMsg = "Are you sure you want to remove the selected key?"
	$sWarningMsg &= @CRLF & @CRLF & "This action can not be undone."
	idTxtKey_Visible()
	KeyShow()
	If MsgBox(BitOR($MB_YESNO, $MB_ICONWARNING, $MB_DEFBUTTON2, $MB_TASKMODAL, $MB_SETFOREGROUND), _
			"PassGen Key Removal Confirmation", $sWarningMsg) = $IDNO Then
		UILock(False)
		KeyHide()
		Return 0
	EndIf
	RegDelete($REGKEYPATH, $sKeyGUID)
	KeyArchiveClearFromMem()
	idTxtKey_SetData("")
	If KeyArchiveGetCount() = 0 Then
		RegDelete($REGKEYPATH, $REGKEYCURRENT)
		KeyIsValid()
	Else
		KeyHide()
		RegistryKeySelect($g_aKeyArchive[0])
		KeyReadFromReg(RegistryKeyGetCurrent())
		UILock(False)
	EndIf
	KeyArchiveClearFromMem()
EndFunc   ;==>KeyRemove

Func KeySave()
	If Not KeyIsValid() Then Return 0
	Local $sKeyGUID = ""
	Switch $g_iKeyOperation
		Case $KEYARCHIVEACTION_MODIFY
			$sKeyGUID = RegistryKeyGetCurrent()
			Local $sKeyDate = KeyArchiveParseDate($sKeyGUID)
			KeySaveToReg($sKeyGUID, idTxtKey_Read(), $sKeyDate)
		Case Else ;$KEYARCHIVEACTION_ADD
			$sKeyGUID = _WinAPI_CreateGUID()
			KeySaveToReg($sKeyGUID, idTxtKey_Read())
			RegistryKeySelect($sKeyGUID)
	EndSwitch
	KeyArchiveClearFromMem()
	idBtnKey_SetCaption("&Change")
	idBtnRevealKey_AcceleratorOption()
	idBtnRevealKey_Depressed(False)
	KeyHide()
	UILock(False)
	idtxtPassphrase_Focus()
	idTxtPassphrase_OnChange()
EndFunc   ;==>KeySave

Func KeySaveToReg($sGUID, $sValue, $sDate = "")
	If $sDate = "" Then $sDate = @YEAR & "/" & @MON & "/" & @MDAY
	$hProtectedKey = KeyProtect($sDate & $sValue)
	RegistryKeyWriteBinary($sGUID, $hProtectedKey)
EndFunc   ;==>KeySaveToReg

Func KeyShow()
	idBtnRevealKey_Depressed()
	InputboxMask($aGUI[$idTxtKey], False)
EndFunc   ;==>KeyShow

Func KeyUnprotect($hProtectedKey)
	Return _CryptUnprotectData($hProtectedKey)
EndFunc   ;==>KeyUnprotect

Func PassphraseGetValue()
	Return StringStripWS(idTxtPassphrase_Read(), $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>PassphraseGetValue

Func PassphraseIsValid()
	idLblPassphraseMsg_SetMessage()
	idLblPasswordMsg_SetMessage()
	Local $sPassphrase = PassphraseGetValue()
	If StringLen($sPassphrase) < 1 Then Return False
	$bIsPassphraseComplex = IsStringComplex($sPassphrase, 8, 0, 0)
	If $bIsPassphraseComplex = False Then
		Switch @error
			Case 1
				idLblPassphraseMsg_SetMessage("Passphrase is too short." & @CRLF & "Must contain at least 8 characters.")
				idTxtPassword_SetData("")
		EndSwitch
	EndIf
	If $bIsPassphraseComplex == -1 Then Return False
	Return $bIsPassphraseComplex
EndFunc   ;==>PassphraseIsValid

Func PasswordClear()
	ClipboardClear()
	idTxtPassword_SetData("")
EndFunc   ;==>PasswordClear

Func PasswordHide()
	InputboxMask($aGUI[$idTxtPassword])
EndFunc   ;==>PasswordHide

Func PasswordShow()
	InputboxMask($aGUI[$idTxtPassword], False)
EndFunc   ;==>PasswordShow

Func RegistryKeyGetCurrent()
	Local $sRegValue = RegRead($REGKEYPATH, $REGKEYCURRENT)
	Return SetError(@error, 0, $sRegValue)
EndFunc   ;==>RegistryKeyGetCurrent

Func RegistryKeyRead($sKeyGUID)
	Return RegRead($REGKEYPATH, $sKeyGUID)
EndFunc   ;==>RegistryKeyRead

Func RegistryKeySelect($sKeyGUID = "")
	If $sKeyGUID == "" Then $sKeyGUID = RegistryKeyGetCurrent()
	RegWrite($REGKEYPATH, "CurrentKey", "REG_SZ", $sKeyGUID)
EndFunc   ;==>RegistryKeySelect

Func RegistryKeyWriteBinary($sKeyGUID, $hValue)
	RegWrite($REGKEYPATH, $sKeyGUID, "REG_BINARY", $hValue)
EndFunc   ;==>RegistryKeyWriteBinary

Func UILock($bFlag = True)
	AutoPurgeTimer(False)
	Switch $bFlag
		Case True
			idBtnKey_Focus()
			idBtnPassphrase_Enabled(False)
			idBtnPassword_Enabled(False)
			idBtnRevealKey_Enabled(False)
			idTxtPassphrase_Enable(False)
			idTxtPassword_Enable(False)
			idTxtPassword_SetData("")
			idLblPassphraseMsg_SetMessage()
			idLblPasswordMsg_SetMessage()
		Case False
			idBtnRevealKey_Enabled()
			idTxtPassphrase_Enable()
			idTxtPassword_Enable()
			idCmbKeyList_Expand(False)
			idCmbKeyList_Visible(False)
			idTxtKey_Enable(False)
			idTxtKey_Visible()
			idBtnKey_SetCaption("&Change")
			idTxtPassphrase_OnChange()
	EndSwitch
EndFunc   ;==>UILock

Func UpdatePassGen()
	If AutoStartIsEnabled() And @ScriptFullPath <> $PROGRAMPATH Then AutoStart(1)
EndFunc   ;==>UpdatePassGen
#EndRegion - Additonal Functions

#Region - Internal Functions
Func _BinaryCheckEncryptedFileHeader($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
	Return _BinaryCheckBytes($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
EndFunc   ;==>_BinaryCheckEncryptedFileHeader

Func _BinaryCheckFileHeader($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
	Return _BinaryCheckBytes($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
EndFunc   ;==>_BinaryCheckFileHeader

Func _BinaryCheckEOF($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
	Return _BinaryCheckBytes($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
EndFunc   ;==>_BinaryCheckEOF

Func _BinaryCheckValueHeader($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
	Return _BinaryCheckBytes($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
EndFunc   ;==>_BinaryCheckValueHeader

Func _BinaryCheckRecordHeader($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
	Return _BinaryCheckBytes($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
EndFunc   ;==>_BinaryCheckRecordHeader

Func _BinaryCheckBytes($dData, $iBinaryLocation, $iBinaryLen, $dBytes)
	Return _BinaryCompare(_BinaryExtract($dData, $iBinaryLocation, $iBinaryLen), $dBytes)
EndFunc   ;==>_BinaryCheckBytes

Func _BinaryExtract(ByRef $bData, $iStartByte, $iByteLen) ;Internal Function to Extract bytes from Binary Data
	Return BinaryMid($bData, $iStartByte, $iByteLen)
EndFunc   ;==>_BinaryExtract

Func _BinaryCompare($bData, $bCompareData) ;Internal Function for Binary Data Comparrison
	Return ($bData == $bCompareData) ? True : False
EndFunc   ;==>_BinaryCompare

Func _CryptProtectData($sString, $sDesc = "", $sPwd = "", $iFlag = 0, $pPrompt = 0)
	$hDLL_CryptProtect = DllOpen("crypt32.dll")

	;funkey 2014.08.11th
	Local $aRet, $iError, $tEntropy, $tDesc, $pEntropy = 0, $pDesc = 0
	Local $tDataIn = _DataToBlob($sString)
	If $sPwd <> "" Then
		$tEntropy = _DataToBlob($sPwd)
		$pEntropy = DllStructGetPtr($tEntropy)
	EndIf

	If $sDesc <> "" Then
		$tDesc = DllStructCreate("wchar desc[" & StringLen($sDesc) + 1 & "]")
		DllStructSetData($tDesc, "desc", $sDesc)
		$pDesc = DllStructGetPtr($tDesc)
	EndIf

	Local $tDataBuf = DllStructCreate($tagDATA_BLOB)

	$aRet = DllCall($hDLL_CryptProtect, "BOOL", "CryptProtectData", "struct*", $tDataIn, "ptr", $pDesc, "ptr", $pEntropy, "ptr", 0, "ptr", $pPrompt, "DWORD", $iFlag, "struct*", $tDataBuf)
	$iError = @error

	_WinAPI_LocalFree(DllStructGetData($tDataIn, "pbData"))

	If $sPwd <> "" Then _WinAPI_LocalFree(DllStructGetData($tEntropy, "pbData"))
	If $iError Then Return SetError(1, 0, "")
	If $aRet[0] = 0 Then Return SetError(2, _WinAPI_GetLastError(), "")

	Local $tDataOut = DllStructCreate("byte data[" & DllStructGetData($tDataBuf, "cbData") & "]", DllStructGetData($tDataBuf, "pbData"))
	Local $bData = DllStructGetData($tDataOut, "data")

	_WinAPI_LocalFree(DllStructGetData($tDataBuf, "pbData"))

	DllClose($hDLL_CryptProtect)

	Return $bData
EndFunc   ;==>_CryptProtectData

;http://msdn.microsoft.com/en-us/library/aa380882(v=vs.85).aspx
Func _CryptUnprotectData($bData, $sDesc = "", $sPwd = "", $iFlag = 0, $pPrompt = 0)
	$hDLL_CryptProtect = DllOpen("crypt32.dll")

	;funkey 2014.08.11th
	Local $aRet, $iError, $tEntropy, $pEntropy = 0
	Local $tDataIn = _DataToBlob($bData)
	$sDesc = ""

	If $sPwd <> "" Then
		$tEntropy = _DataToBlob($sPwd)
		$pEntropy = DllStructGetPtr($tEntropy)
	EndIf

	Local $tDataBuf = DllStructCreate($tagDATA_BLOB)
	Local $tDesc = DllStructCreate("ptr desc")
	Local $pDesc = DllStructGetPtr($tDesc)

	$aRet = DllCall($hDLL_CryptProtect, "BOOL", "CryptUnprotectData", "struct*", $tDataIn, "ptr*", $pDesc, "ptr", $pEntropy, "ptr", 0, "ptr", $pPrompt, "DWORD", $iFlag, "struct*", $tDataBuf)
	$iError = @error

	_WinAPI_LocalFree(DllStructGetData($tDataIn, "pbData"))

	If $sPwd <> "" Then _WinAPI_LocalFree(DllStructGetData($tEntropy, "pbData"))
	If $iError Then Return SetError(1, 0, "")
	If $aRet[0] = 0 Then Return SetError(2, _WinAPI_GetLastError(), "")

	Local $tDataOut = DllStructCreate("char data[" & DllStructGetData($tDataBuf, "cbData") & "]", DllStructGetData($tDataBuf, "pbData"))
	Local $sData = DllStructGetData($tDataOut, "data")

	Local $aLen = DllCall("msvcrt.dll", "UINT:cdecl", "wcslen", "ptr", $aRet[2])
	Local $tDesc = DllStructCreate("wchar desc[" & $aLen[0] + 1 & "]", $aRet[2])
	$sDesc = DllStructGetData($tDesc, "desc")

	_WinAPI_LocalFree($aRet[2])
	_WinAPI_LocalFree(DllStructGetData($tDataBuf, "pbData"))

	DllClose($hDLL_CryptProtect)

	Return $sData
EndFunc   ;==>_CryptUnprotectData

;Creates a DATA_BLOB structure where the function stores the decrypted data.
;When you have finished using the DATA_BLOB structure, free its pbData member by calling the _WinAPI_LocalFree function.
Func _DataToBlob($data)
	;funkey 2014.08.11th
	Local $iLen, $tDataIn, $tData, $aMem
	Local Const $LMEM_ZEROINIT = 0x40
	Select
		Case IsString($data)
			$iLen = StringLen($data)
		Case IsBinary($data)
			$iLen = BinaryLen($data)
		Case Else
			Return SetError(1, 0, 0)
	EndSelect

	$tDataIn = DllStructCreate($tagDATA_BLOB)
	$aMem = DllCall("Kernel32.dll", "handle", "LocalAlloc", "UINT", $LMEM_ZEROINIT, "UINT", $iLen)
	$tData = DllStructCreate("byte[" & $iLen & "]", $aMem[0])

	DllStructSetData($tData, 1, $data)
	DllStructSetData($tDataIn, "cbData", $iLen)
	DllStructSetData($tDataIn, "pbData", DllStructGetPtr($tData))

	Return $tDataIn
EndFunc   ;==>_DataToBlob

Func _LegacyKeyConvert()
	RegRead($REGKEYPATH, "Key")
	If @error = -1 Then Return 0
	If @error = 0 Then
		$iConvertResponse = MsgBox(BitOR($MB_YESNO, $MB_ICONQUESTION), "PassGen - Old version key detected", "A PassGen key was detected that was produced using an older version of the PassGen Tool. " & _
				"This version is no longer compatible with 1.3+ version of PassGen." & @CRLF & @CRLF & "Do you want to convert the older key to the new format?" & @CRLF & "Yes = Convert old key, No = Remove old key" & _
				@CRLF & @CRLF & "You will need to change the key date as appropriate.")
		If $iConvertResponse = $IDYES Then
			$sKeyGUID = _WinAPI_CreateGUID()
			$sOldKey = KeyUnprotect(RegRead($REGKEYPATH, "Key"))
			$sDate = _LegacyKeyRegGetTimeStamp()
			If Not _DateIsValid($sDate) Then $sDate = ""
			KeySaveToReg($sKeyGUID, $sOldKey, $sDate)
			RegistryKeySelect($sKeyGUID)
		EndIf
		RegDelete($REGKEYPATH, "Key")
	EndIf
EndFunc   ;==>_LegacyKeyConvert

;tailored from https://www.autoitscript.com/forum/topic/76544-registry-timestamp/?do=findComment&comment=556271
Func _LegacyKeyRegGetTimeStamp($iPC = "\\" & @ComputerName, $iRegHive = 0x80000001, $sRegKey = "Software\PassGen")
    Local $sRes='', $aRet, $hReg = DllStructCreate("int")
    Local $hRemoteReg = DllStructCreate("int")
    Local $FILETIME = DllStructCreate("dword;dword")
    Local $SYSTEMTIME1 = DllStructCreate("ushort;ushort;ushort;ushort;ushort;ushort;ushort;ushort")
    Local $SYSTEMTIME2 = DllStructCreate("ushort;ushort;ushort;ushort;ushort;ushort;ushort;ushort")
    Local $hAdvAPI=DllOpen('advapi32.dll'), $hKernel=DllOpen('kernel32.dll')
    If $hAdvAPI=-1 Or $hKernel=-1 Then Return SetError(1, $aRet[0], 'DLL Open Error!')

    $connect = DllCall("advapi32.dll", "int", "RegConnectRegistry", _
        "str", $iPC , _
        "int", $iRegHive, _
        "ptr", DllStructGetPtr($hRemoteReg))

    $aRet = DllCall("advapi32.dll", "int", "RegOpenKeyEx", _
        "int", DllStructGetData($hRemoteReg,1), _
        "str", $sRegKey, _
        "int", 0, _
        "int", 0x20019, _
        "ptr", DllStructGetPtr($hReg))
    If $aRet[0] Then Return SetError(2, $aRet[0], 'Registry Key Open Error!')

    $aRet = DllCall("advapi32.dll", "int", "RegQueryInfoKey", _
        "int", DllStructGetData($hReg,1), _
        "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, _
        "ptr", DllStructGetPtr($FILETIME))
    If $aRet[0] Then Return SetError(3, $aRet[0], 'Registry Key Query Error!')


    $aRet = DllCall("advapi32.dll", "int", "RegCloseKey", _
        "int", DllStructGetData($hReg,1))
    If $aRet[0] Then Return SetError(4, $aRet[0], 'Registry Key Close Error!')


    $aRet = DllCall("kernel32.dll", "int", "FileTimeToSystemTime", _
        "ptr", DllStructGetPtr($FILETIME), _
        "ptr", DllStructGetPtr($SYSTEMTIME1))
    If $aRet[0]=0 Then Return SetError(5, 0, 'Time Convert Error!')


    $aRet = DllCall("kernel32.dll", "int", "SystemTimeToTzSpecificLocalTime", _
        "ptr", 0, _
        "ptr", DllStructGetPtr($SYSTEMTIME1), _
        "ptr", DllStructGetPtr($SYSTEMTIME2))
    If $aRet[0]=0 Then Return SetError(5, 0, 'Time Convert Error!')

    $sRes &= StringFormat("%.2d",DllStructGetData($SYSTEMTIME2,1)) & "/"
	$sRes &= StringFormat("%.2d",DllStructGetData($SYSTEMTIME2,2)) & "/"
	$sRes &= StringFormat("%.2d",DllStructGetData($SYSTEMTIME2,4))

    Return $sRes
EndFunc
#EndRegion - Internal Functions