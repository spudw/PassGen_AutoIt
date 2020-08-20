#AutoIt3Wrapper_Icon = "Icon.ico"
#AutoIt3Wrapper_Compression = 4
#AutoIt3Wrapper_Res_FileVersion = 1.1

; Version History
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
#include <WinAPISysWin.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ColorConstants.au3>
#include <FontConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <StringConstants.au3>
#include <StaticConstants.au3>
#include <TrayConstants.au3>

#include <APISysConstants.au3>
#include <WinAPISysWin.au3>


Opt("GUIOnEventMode", 1)
Opt("TrayMenuMode", 1)
Opt("TrayOnEventMode", 1)
TraySetOnEvent($TRAY_EVENT_PRIMARYDOUBLE, "GUIRestore")
;Opt("TrayIconHide", 1)

Const $sCharacterList = "ABCEFGHKLMNPQRSTUVWXYZ0987654321abdefghjmnqrtuwy"
Const $iCharacterListLen = StringLen($sCharacterList)
Const $sRegKeyPath = "HKCU\Software\PassGen"

Const $tagDATA_BLOB = "DWORD cbData;ptr pbData;"
Const $tagCRYPTPROTECT_PROMPTSTRUCT = "DWORD cbSize;DWORD dwPromptFlags;HWND hwndApp;ptr szPrompt;"

Dim $aGUI[1] = ["hwnd|id"]
Enum $hGUI = 1, $idBtnRevealKey, $idLblKey, $idTxtKey, $idBtnKey, $idLblPassphrase, $idLblPassphraseUse, $idTxtPassphrase, $idBtnPassphrase, $idLblPassphraseMsg, _
		$idLblPassword, $idLblPasswordUse, $idTxtPassword, $idBtnPassword, $idLblPasswordMsg, $iGUILast
ReDim $aGUI[$iGUILast]

#Region - UI Creation
$aGUI[$hGUI] = GUICreate("PassGenTool", 508, 230, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU, $WS_MINIMIZEBOX))
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
GUICtrlSendMsg(-1, $EM_GETPASSWORDCHAR, $ES_PASSWORDCHAR, 0)
GUICtrlSetBkColor(-1, 0xFFFFFF)
GUICtrlSetFont(-1, 18, $FW_BOLD, Default, "Consolas")
$aGUI[$idBtnPassword] = GUICtrlCreateButton("Co&py", 451, 142, 40, 34)
GUICtrlSetOnEvent(-1, "GUIEvents")
GUICtrlSetState(-1, $GUI_DISABLE)
$aGUI[$idLblPasswordMsg] = GUICtrlCreateLabel("", 104, 176, 330, 40, $SS_CENTER)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTITALIC)

GUISetOnEvent($GUI_EVENT_CLOSE, "GUIEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "GUIEvents")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_ACTIVATE, "WM_ACTIVATE")
GUIRegisterMsg($WM_SIZE, "WM_SIZE")

GUISetState(@SW_SHOW)
#EndRegion - UI Creation

KeyReadFromReg()
If Not KeyIsPresent() Then KeyChange()

While 1
	Sleep(10)
WEnd

#Region - UI Event Functions
Func GUIEvents()
	$iCtrl = @GUI_CtrlId
	Switch $iCtrl
		Case $GUI_EVENT_CLOSE
			Exit
		Case $GUI_EVENT_MINIMIZE
			PasswordHide()
			GUIHide()
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
	GUISetState(@SW_HIDE)
EndFunc   ;==>GUIHide

Func GUIRestore()
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

Func idTxtKey_OnChange()
	If Not KeyIsPresent() Then Return 0
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

Func WM_SIZE($hWnd, $iMsg, $wParam, $lParam)
	Switch $hWnd
		Case $aGUI[$hGUI]
			Switch $wParam
				Case 1 ; WA_INACTIVE
					PasswordHide()
			EndSwitch
	EndSwitch
EndFunc   ;==>WM_SIZE
#EndRegion - UI Event Functions

#Region - Additonal Functions
Func ClipboardClear()
	ClipPut("")
EndFunc   ;==>ClipboardClear

Func ClipboardCopyData($vData)
	ClipPut($vData)
EndFunc   ;==>ClipboardCopyData

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

Func KeyMask($bMask = True)
	Switch $bMask
		Case False
			GUICtrlSendMsg($aGUI[$idTxtKey], $EM_SETPASSWORDCHAR, 0, 0)
		Case True
			GUICtrlSendMsg($aGUI[$idTxtKey], $EM_SETPASSWORDCHAR, $ES_PASSWORDCHAR, 0)
	EndSwitch
	Local $aRes = DllCall("user32.dll", "int", "RedrawWindow", "hwnd", GUICtrlGetHandle($aGUI[$idTxtKey]), "ptr", 0, "ptr", 0, "dword", 5)
EndFunc   ;==>KeyMask

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
