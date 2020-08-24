#AutoIt3Wrapper_Icon = "Icon.ico"
#AutoIt3Wrapper_Compression = 4
#AutoIt3Wrapper_Res_FileVersion = 1.2.1
#AutoIt3Wrapper_Res_ProductName = PassGen

; Version History
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

#include <WinAPI.au3>
#include <Crypt.au3>
#include <Misc.au3>
#include <WinAPISysWin.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ColorConstants.au3>
#include <FontConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <TrayConstants.au3>
#include <WinAPIGdi.au3>

If _Singleton("PassGen", 1) = 0 Then
	$sRunningProcessPath = _WinAPI_GetProcessFileName(ProcessExists("PassGen.exe"))
	If $sRunningProcessPath = @ScriptFullPath Then Exit
	If _VersionCompare(FileGetVersion(@ScriptFullPath), FileGetVersion($sRunningProcessPath)) = 1 Then
		ProcessClose("PassGen.exe")
	Else
		Exit
	EndIf
EndIf

Opt("GUIOnEventMode", 1)
Opt("TrayMenuMode", 1 + 2)
Opt("TrayOnEventMode", 1)
TraySetClick(16)

Const $sCharacterList = "ABCEFGHKLMNPQRSTUVWXYZ0987654321abdefghjmnqrtuwy"
Const $iCharacterListLen = StringLen($sCharacterList)
Const $sRegKeyPath = "HKCU\Software\PassGen"
Const $sStartupLink = @StartupDir & "\PassGen.exe.lnk"
Const $sProgramPath = @ProgramsDir & "\PassGen\PassGen.exe"

UpdatePassGen()

Const $tagDATA_BLOB = "DWORD cbData;ptr pbData;"
Const $tagCRYPTPROTECT_PROMPTSTRUCT = "DWORD cbSize;DWORD dwPromptFlags;HWND hwndApp;ptr szPrompt;"

Dim $aGUI[1] = ["hwnd|id"]
Enum $hGUI = 1, $idMnuFile, $idMnuFileQuit, $idMnuOptions, $idMnuOptionsAutoStart, $idMnuOptionsCloseToTray, $idTrayOpen, $idTrayQuit, $idBtnRevealKey, $idLblKey, $idTxtKey, $idBtnKey, _
		$idLblPassphrase, $idLblPassphraseUse, $idTxtPassphrase, $idBtnPassphrase, $idLblPassphraseMsg, $idLblPassword, $idLblPasswordUse, $idTxtPassword, $idBtnPassword, $idLblPasswordMsg, $iGUILast
ReDim $aGUI[$iGUILast]

#Region - UI Creation
$aGUI[$hGUI] = GUICreate("PassGenTool", 508, 230, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_SYSMENU))
$aGUI[$idMnuFile] = GUICtrlCreateMenu("&File")
$aGUI[$idMnuFileQuit] = GUICtrlCreateMenuItem("&Quit", $aGUI[$idMnuFile])
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idMnuOptions] = GUICtrlCreateMenu("&Options")
$aGUI[$idMnuOptionsAutoStart] = GUICtrlCreateMenuItem("&Automatic Start on Login", $aGUI[$idMnuOptions])
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idMnuOptionsCloseToTray] = GUICtrlCreateMenuItem("&Enable Close to Tray", $aGUI[$idMnuOptions])
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idTrayOpen] = TrayCreateItem("Open")
TrayItemSetOnEvent(-1, "TrayEvents")
TrayCreateItem("")
$aGUI[$idTrayQuit] = TrayCreateItem("Quit")
TrayItemSetOnEvent(-1, "TrayEvents")
$aGUI[$idBtnRevealKey] = GUICtrlCreateCheckbox("&Show", 12, 12, 44, 28, BitOR($BS_PUSHLIKE, $BS_AUTOCHECKBOX))
GUICtrlSetState(-1, $GUI_UNCHECKED)
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idLblKey] = GUICtrlCreateLabel("Key:", 64, 18, 40, 20)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTUNDER)
GUICtrlSetColor(-1, $COLOR_RED)
$aGUI[$idTxtKey] = GUICtrlCreateInput("", 104, 10, 330, 34, $ES_PASSWORD)
Const $ES_PASSWORDCHAR = GUICtrlSendMsg(-1, $EM_GETPASSWORDCHAR, 0, 0)
GUICtrlSetState(-1, $GUI_DISABLE)
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
GUICtrlSetBkColor(-1, 0xFFFFFF)
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
TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "GUIRestore")
_WinAPI_DwmSetWindowAttribute($aGUI[$hGUI], $DWMWA_FORCE_ICONIC_REPRESENTATION, 1)
#EndRegion - UI Creation

KeyReadFromReg()
If Not KeyIsPresent() Then KeyChange()
If AutoStartIsEnabled() Then GUICtrlSetState($aGUI[$idMnuOptionsAutoStart], $GUI_CHECKED)
If CloseToTrayIsEnabled() Then GUICtrlSetState($aGUI[$idMnuOptionsCloseToTray], $GUI_CHECKED)

If $CmdLineRaw = "/silent" Then
	GUIHide()
Else
	GUIRestore()
EndIf

While 1
	Sleep(10)
WEnd

#Region - UI Event Functions
Func _Exit()
	ClipboardClear()
	Exit
EndFunc   ;==>_Exit

Func GUIEvents()
	$iCtrl = @GUI_CtrlId
	Switch $iCtrl
		Case $GUI_EVENT_CLOSE
			Return (CloseToTrayIsEnabled()) ? GUIHide() : _Exit()
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
	PasswordHide()
	GUISetState(@SW_HIDE)
	GUISetState(@SW_DISABLE)
	TraySetState($TRAY_ICONSTATE_SHOW)
EndFunc   ;==>GUIHide

Func GUIRestore()
	TraySetState($TRAY_ICONSTATE_HIDE)
	GUISetState(@SW_ENABLE)
	GUISetState(@SW_SHOW)
	WinActivate($aGUI[$hGUI])
EndFunc   ;==>GUIRestore

Func idBtnRevealKey_Click()
	Local $iCtrl = $aGUI[$idBtnRevealKey]
	Local $iState = GUICtrlRead($iCtrl)
	If $iState = $GUI_UNCHECKED Then
		KeyHide()
	Else
		KeyShow()
	EndIf
EndFunc   ;==>idBtnRevealKey_Click

Func idBtnKey_Click()
	If GUICtrlRead($aGUI[$idBtnKey]) = "&Change" Then
		KeyChange()
	Else
		KeySave()
	EndIf
EndFunc   ;==>idBtnKey_Click

Func idBtnPassphrase_Click()
	ClipboardCopyData(GUICtrlRead($aGUI[$idTxtPassphrase]))
	GUICtrlSetData($aGUI[$idLblPassphraseMsg], "Copied to Clipboard")
	GUICtrlSetData($aGUI[$idLblPasswordMsg], "")
EndFunc   ;==>idBtnPassphrase_Click

Func idBtnPassword_Click()
	ClipboardCopyData(GUICtrlRead($aGUI[$idTxtPassword]))
	GUICtrlSetData($aGUI[$idLblPassphraseMsg], "")
	GUICtrlSetData($aGUI[$idLblPasswordMsg], "Copied to Clipboard")
EndFunc   ;==>idBtnPassword_Click

Func idMnuOptionsAutoStart_Click()
	Local $iState = MenuItemToggle($aGUI[$idMnuOptionsAutoStart])
	AutoStart($iState)
EndFunc   ;==>idMnuOptionsAutoStart_Click

Func idMnuOptionsCloseToTray_Click()
	Local $iState = MenuItemToggle($aGUI[$idMnuOptionsCloseToTray])
	CloseToTraySetting($iState)
EndFunc   ;==>idMnuOptionsCloseToTray_Click

Func idTxtKey_OnChange()
	Return KeyIsPresent()
EndFunc   ;==>idTxtKey_OnChange

Func idTxtKey_Read()
	Return GUICtrlRead($aGUI[$idTxtKey])
EndFunc   ;==>idTxtKey_Read

Func idTxtPassphrase_OnChange()
	If PassphraseIsValid() Then
		GUICtrlSetState($aGUI[$idBtnPassphrase], $GUI_ENABLE)
		GUICtrlSetState($aGUI[$idBtnPassword], $GUI_ENABLE)
		GeneratePassword()
	Else
		GUICtrlSetState($aGUI[$idBtnPassphrase], $GUI_DISABLE)
		GUICtrlSetState($aGUI[$idBtnPassword], $GUI_DISABLE)
	EndIf
EndFunc   ;==>idTxtPassphrase_OnChange

Func idTxtPassphrase_Read()
	Return GUICtrlRead($aGUI[$idTxtPassphrase])
EndFunc   ;==>idTxtPassphrase_Read

Func idTxtPassword_SetData($sValue)
	GUICtrlSetData($aGUI[$idTxtPassword], $sValue)
EndFunc   ;==>idTxtPassword_SetData

Func MenuItemToggle($id)
	If BitAND(GUICtrlRead($id), $GUI_CHECKED) Then
		GUICtrlSetState($id, $GUI_UNCHECKED)
		Return 0
	Else
		GUICtrlSetState($id, $GUI_CHECKED)
		Return 1
	EndIf
EndFunc   ;==>MenuItemToggle

Func TrayEvents()
	$iCtrl = @TRAY_ID
	Switch $iCtrl
		Case $aGUI[$idTrayOpen]
			GUIRestore()
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
					idTxtKey_OnChange()
				Case $aGUI[$idTxtPassphrase]
					idTxtPassphrase_OnChange()
			EndSwitch
		Case $EN_SETFOCUS
			Switch $iIDFrom
				Case $aGUI[$idTxtPassword]
					PasswordShow()
			EndSwitch
		Case $EN_KILLFOCUS
			Switch $iIDFrom
				Case $aGUI[$idTxtPassword]
					PasswordHide()
			EndSwitch
	EndSwitch
EndFunc   ;==>WM_COMMAND

Func WM_ACTIVATE($hWnd, $iMsg, $wParam, $lParam)
	Local $iCode = BitAND($wParam, 0xFFFF)
	Switch $hWnd
		Case $aGUI[$hGUI]
			Switch $iCode
				Case 0 ; WA_INACTIVE
					PasswordHide()
			EndSwitch
	EndSwitch
EndFunc   ;==>WM_ACTIVATE
#EndRegion - UI Event Functions

#Region - Additonal Functions
Func AutoStart($bEnable = 1)
	If $bEnable Then
		FileCopy(@ScriptFullPath, $sProgramPath, $FC_CREATEPATH + $FC_OVERWRITE)
		FileCreateShortcut($sProgramPath, $sStartupLink, "", "/silent")
	Else
		If FileExists($sStartupLink) Then FileDelete($sStartupLink)
	EndIf
EndFunc   ;==>AutoStart

Func AutoStartIsEnabled()
	Local $aShortcut = FileGetShortcut($sStartupLink)
	If @error Then Return False
	Return (FileExists($aShortcut[0])) ? True : False
EndFunc   ;==>AutoStartIsEnabled

Func ClipboardClear()
	ClipPut("")
EndFunc   ;==>ClipboardClear

Func ClipboardCopyData($vData)
	ClipPut($vData)
EndFunc   ;==>ClipboardCopyData

Func CloseToTrayIsEnabled()
	RegRead($sRegKeyPath, "NoCloseToTray")
	Return (@error <> 0) ? True : False
EndFunc   ;==>CloseToTrayIsEnabled

Func CloseToTraySetting($bEnable = 1)
	If $bEnable Then
		RegDelete($sRegKeyPath, "NoCloseToTray")
	Else
		RegWrite($sRegKeyPath, "NoCloseToTray", "REG_DWORD", 0)
	EndIf
EndFunc   ;==>CloseToTraySetting

Func GeneratePassword()
	$sKey = KeyGetValue()
	$sPassphrase = PassphraseGetValue()
	$dHash = _Crypt_HashData($sKey & "-" & $sPassphrase, $CALG_SHA1)
	$sPassword = ""
	For $iX = 1 To 20
		$hByte = BinaryMid($dHash, $iX, 1)
		$iChar = Mod(Int(Dec(StringRight($hByte, 2))), $iCharacterListLen)
		$sPassword &= StringMid($sCharacterList, $iChar + 1, 1)
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

Func KeyChange()
	PasswordClear()
	GUICtrlSetData($aGUI[$idLblPassphraseMsg], "")
	GUICtrlSetData($aGUI[$idLblPasswordMsg], "")
	GUICtrlSetState($aGUI[$idTxtPassphrase], $GUI_DISABLE)
	GUICtrlSetState($aGUI[$idTxtPassword], $GUI_DISABLE)
	GUICtrlSetBkColor($aGUI[$idTxtPassword], 0xF0F0F0)
	GUICtrlSetData($aGUI[$idBtnKey], "&Save")
	GUICtrlSetData($aGUI[$idBtnRevealKey], "Sho&w")
	GUICtrlSetState($aGUI[$idBtnRevealKey], $GUI_CHECKED)
	KeyShow()
	GUICtrlSetState($aGUI[$idTxtKey], $GUI_ENABLE)
	GUICtrlSetState($aGUI[$idTxtKey], $GUI_FOCUS)
EndFunc   ;==>KeyChange

Func KeyIsPresent()
	Local $iCtrl = $aGUI[$idLblKey]
	Local $iMsgCtrl = $aGUI[$idLblPasswordMsg]
	Local $bKeyIsPresent = (KeyGetValue() <> "") ? True : False
	Switch $bKeyIsPresent
		Case True
			GUICtrlSetColor($iCtrl, 0x000000)
			GUICtrlSetData($iMsgCtrl, "")
		Case False
			GUICtrlSetColor($iCtrl, $COLOR_RED)
			GUICtrlSetData($iMsgCtrl, "Missing Key")
	EndSwitch
	Return $bKeyIsPresent
EndFunc   ;==>KeyIsPresent

Func KeyGetValue()
	Return StringStripWS(idTxtKey_Read(), $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>KeyGetValue

Func KeyHide()
	InputboxMask($aGUI[$idTxtKey])
EndFunc   ;==>KeyHide

Func KeyProtect($sValue)
	Return _CryptProtectData($sValue)
EndFunc   ;==>KeyProtect

Func KeyReadFromReg()
	$hProtectedKey = RegistryKeyRead()
	If @error Then Return ""
	$sKey = KeyUnprotect($hProtectedKey)
	GUICtrlSetData($aGUI[$idTxtKey], $sKey)
EndFunc   ;==>KeyReadFromReg

Func KeySave()
	KeySaveToReg()
	$sKey = KeyGetValue()
	If $sKey = "" Then Return -1
	GUICtrlSetData($aGUI[$idTxtKey], $sKey)
	GUICtrlSetState($aGUI[$idTxtKey], $GUI_DISABLE)
	GUICtrlSetState($aGUI[$idTxtPassphrase], $GUI_ENABLE)
	GUICtrlSetState($aGUI[$idTxtPassword], $GUI_ENABLE)
	GUICtrlSetBkColor($aGUI[$idTxtPassword], 0xFFFFFF)
	GUICtrlSetData($aGUI[$idBtnKey], "&Change")
	GUICtrlSetData($aGUI[$idBtnRevealKey], "&Show")
	GUICtrlSetState($aGUI[$idBtnRevealKey], $GUI_UNCHECKED)
	KeyHide()
	GUICtrlSetState($aGUI[$idTxtPassphrase], $GUI_FOCUS)
	idTxtPassphrase_OnChange()
EndFunc   ;==>KeySave

Func KeySaveToReg()
	$hProtectedKey = KeyProtect(KeyGetValue())
	RegistryKeyWriteBinary($hProtectedKey)
EndFunc   ;==>KeySaveToReg

Func KeyShow()
	InputboxMask($aGUI[$idTxtKey], False)
EndFunc   ;==>KeyShow

Func KeyUnprotect($hProtectedKey)
	Return _CryptUnprotectData($hProtectedKey)
EndFunc   ;==>KeyUnprotect

Func PassphraseGetValue()
	Return StringStripWS(idTxtPassphrase_Read(), $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>PassphraseGetValue

Func PassphraseIsValid()
	Local $iCtrl = $aGUI[$idTxtPassphrase]
	Local $iMsgCtrl = $aGUI[$idLblPassphraseMsg]
	Local $sPassphrase = PassphraseGetValue()
	Switch StringLen($sPassphrase)
		Case 0
			GUICtrlSetData($aGUI[$idLblPassphraseMsg], "")
			GUICtrlSetData($aGUI[$idLblPasswordMsg], "")
			PasswordClear()
		Case 1 To 7
			GUICtrlSetData($iMsgCtrl, "Passphrase is too short" & @CRLF & "8 characters minimum")
			GUICtrlSetData($aGUI[$idLblPasswordMsg], "")
			PasswordClear()
		Case Else
			GUICtrlSetData($iMsgCtrl, "")
			Return 1
	EndSwitch
	Return 0
EndFunc   ;==>PassphraseIsValid

Func PasswordClear()
	ClipboardClear()
	idTxtPassword_SetData("")
EndFunc   ;==>PasswordClear

Func PasswordHide()
	InputboxMask($aGUI[$idTxtPassword])
	GUICtrlSetState($aGUI[$idTxtPassphrase], $GUI_FOCUS)
EndFunc   ;==>PasswordHide

Func PasswordShow()
	InputboxMask($aGUI[$idTxtPassword], False)
EndFunc   ;==>PasswordShow

Func RegistryKeyWriteBinary($hValue)
	Return RegWrite($sRegKeyPath, "Key", "REG_BINARY", $hValue)
EndFunc   ;==>RegistryKeyWriteBinary

Func RegistryKeyRead()
	Return RegRead($sRegKeyPath, "Key")
EndFunc   ;==>RegistryKeyRead

Func UpdatePassGen()
	If AutoStartIsEnabled() And @ScriptFullPath <> $sProgramPath Then AutoStart(1)
EndFunc   ;==>UpdatePassGen
#EndRegion - Additonal Functions

#Region - Internal Functions
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
#EndRegion - Internal Functions
