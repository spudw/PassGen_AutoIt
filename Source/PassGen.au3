#AutoIt3Wrapper_Icon = "Icon.ico"
#AutoIt3Wrapper_Compression = 4
Const $sVersion = "2.0.4"

#pragma compile(FileDescription, Password Generator Tool)
#pragma compile(ProductName, PassGen)
#pragma compile(ProductVersion, 2.0.4)
#pragma compile(FileVersion, 2.0.4.0)
;#AutoIt3Wrapper_Res_File_Add


; Version History
;
; Version 2.0.4 - 2021/02/19
;		Bug Fixes
;		[*] Tooltip not disappearing when cancel edit
;
; Version 2.0.3 - 2021/02/19
;		Minor Cosmetic Changes
;		[*] Changed GUI size and layout
;			Spaced controls mor uniformly
;		[*] Show Key button
;			Toggled state text change
;		[*] Passphrase field operation
;			Greyed out Password field when no Password present
;			"Type here" message with empty, non-focused Passphrase field
;
; Version 2.0.2 - 2021/02/18
;		Bug fixes & AutoUpdate function tweak
;		[*] Import operation
;			Select most recent dated Key as Active
;		[*] HotKeyManager Issues
;			KeyManager Activate
;			Import operation
;			Unsaved changed queue stacking
;		[*] Empty Key Archive
;			Export operation halt
;			New Key doesn't register as change
;		[*] PasswordIsValid
;			Password cleared if Passphrase is cleared logic
;		[*] AutoUpdate
;			Added update message and re-execute from local path logic
;
; Version 2.0.1 - 2021/02/17
;		Minor Bug fixes & cosmetic changes/corrections
;		[*] Changed Key Change Button text
;		[+] Added File Menu options on main GUI
;			Open Key Manager, Import, Export
;
; Version 2.0 - 2021/02/17
;		New Major Version - New Features & Code Cleanup
;		[+] Added dedicated Key Archive Manager UI
;			Change button on Main UI opens new Key Archive Manager (KM) UI
;			KM is equipped with editable listview which enables electing an active key, key date and key values
;			Basic UI functions included (Menubar, context menu, buttons) for performing key archive management functions (add keys, modify keys)
;		[*] Migrated Key Archive functions from ver 1.3 into new Key Archive Manager
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
#include <GuiEdit.au3>
#include <GuiListView.au3>
#include <GuiMenu.au3>
#include <ButtonConstants.au3>
#include <ColorConstants.au3>
#include <EditConstants.au3>
#include <FileConstants.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <TrayConstants.au3>
#include <WinAPIvkeysConstants.au3>
#include <WindowsConstants.au3>

If _Singleton("PassGen", 1) = 0 Then
	$sRunningProcessPath = _WinAPI_GetProcessFileName(ProcessExists("PassGen.exe"))
	If $sRunningProcessPath = @ScriptFullPath Then Exit
	If _VersionCompare(FileGetVersion(@ScriptFullPath), FileGetVersion($sRunningProcessPath)) = 1 Then
		ProcessClose("PassGen.exe")
	Else
		Exit
	EndIf
EndIf

Opt("GUICloseOnESC", 0) ;Don't send the $GUI_EVENT_CLOSE message when ESC is pressed.
Opt("GUIEventOptions", 1) ;suppress windows behavior on minimize, restore or maximize click button or window resize
Opt("GUIOnEventMode", 1) ;enable
Opt("TrayMenuMode", 1 + 2) ; no default menu & items will not automatically check/uncheck when clicked
Opt("TrayOnEventMode", 1) ;enable
TraySetClick(8)

Const $CHARACTERLIST = "ABCEFGHKLMNPQRSTUVWXYZ0987654321abdefghjmnqrtuwy"
Const $CHARACTERLISTLEN = StringLen($CHARACTERLIST)
Const $REGKEYPATH = "HKCU\Software\PassGen"
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

Dim $aGUI[1] = ["hwnd|id"]
Enum $hGUI = 1, $idMnuFile, $idMnuFileKeyMgr, $idMnuFileImport, $idMnuFileExport, $idMnuFileQuit, $idMnuOptions, $idMnuOptionsAutoStart, $idMnuOptionsCloseToTray, $idTrayOpen, $idTrayQuit, $idBtnRevealKey, $idLblKey, $idTxtKey, $idCmbKeyList, _
		$idDateKeyDatePicker, $idBtnKey, $idLblPassphrase, $idLblPassphraseUse, $idTxtPassphrase, $idBtnPassphrase, $idLblPassphraseMsg, $idLblPassword, $idLblPasswordUse, $idTxtPassword, _
		$idBtnPassword, $idLblPasswordMsg, $iGUILast
ReDim $aGUI[$iGUILast]

Dim $aKeyManagerGUI[1] = ["hwnd|id"]
Enum $hKeyManagerGUI = 1, $idKMMnuFile, $idKMMnuFileNew, $idKMMnuFileExport, $idKMMnuFileImport, $idKMMnuFileClose, _
		$idKMMnuEdit, $idKMMnuEditRemove, $idKMMnuEditSelectAll, $idKMMnuEditDeselectAll, _
		$idKMListView, $hKMListView, $idKMBtnCancel, $idKMBtnSave, $idKMListViewDummy, $iKeyManagerGUILast
ReDim $aKeyManagerGUI[$iKeyManagerGUILast]

Enum $e_KMActivate = 1000, $e_KMEditDate, $e_KMEditValue, $e_KMEditRemove
Enum Step *2 $e_HotKeyESC, $e_HotKeyDEL, $e_HotKeyEnter

Global $aItemPos, $idTempEditCtrl, $hTempEditCtrl, $aCurrentListViewItem[2], $bChangesMade = False, $bChangesPending = False, $bListViewSortDirection = False, $g_iEditing = False
Global $g_iCurCol = -1, $g_iSortDir = 1, $g_bSet = False, $g_iCol = -1
Global $hTimer, $g_aKeyArchive, $g_bKeyManagerBusy = False
Const $sEmptyPassphraseMsg = "Type Passphrase Here", $AUTOPURGETIME = 3 ;minutes
#EndRegion - Includes and Variables

#Region - UI Creation
$aGUI[$hGUI] = GUICreate("PassGen v" & $sVersion, 510, 230, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_SYSMENU))
$aGUI[$idMnuFile] = GUICtrlCreateMenu("&File")
$aGUI[$idMnuFileKeyMgr] = GUICtrlCreateMenuItem("Open Key &Manager" & @TAB & "Ctrl + M", $aGUI[$idMnuFile])
GUICtrlSetOnEvent(-1, "GUIEvents")
GUICtrlCreateMenuItem("", $aGUI[$idMnuFile])
$aGUI[$idMnuFileImport] = GUICtrlCreateMenuItem("&Import Key Archive" & @TAB & "Ctrl + I", $aGUI[$idMnuFile])
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idMnuFileExport] = GUICtrlCreateMenuItem("&Export Key Archive" & @TAB & "Ctrl + E", $aGUI[$idMnuFile])
GUICtrlSetOnEvent(-1, "GUIEvents")
GUICtrlCreateMenuItem("", $aGUI[$idMnuFile])
$aGUI[$idMnuFileQuit] = GUICtrlCreateMenuItem("&Quit" & @TAB & "Ctrl + Q", $aGUI[$idMnuFile])
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
$aGUI[$idBtnRevealKey] = GUICtrlCreateCheckbox("&Show", 12, 10, 44, 34, BitOR($BS_PUSHLIKE, $BS_AUTOCHECKBOX))
GUICtrlSetState(-1, $GUI_UNCHECKED)
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idLblKey] = GUICtrlCreateLabel("Key:", 64, 18, 40, 20)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTUNDER)
$aGUI[$idTxtKey] = GUICtrlCreateInput("", 104, 10, 326, 34, $ES_PASSWORD)
Const $ES_PASSWORDCHAR = GUICtrlSendMsg(-1, $EM_GETPASSWORDCHAR, 0, 0)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetFont(-1, 18, $FW_BOLD, Default, "Consolas")
$aGUI[$idBtnKey] = GUICtrlCreateButton("Key &Mgr", 442, 10, 58, 34, $BS_DEFPUSHBUTTON)
GUICtrlSetOnEvent(-1, "GUIEvents")
$aGUI[$idLblPassphrase] = GUICtrlCreateLabel("Passphrase:", 16, 86, 100, 20)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTUNDER)
$aGUI[$idLblPassphraseUse] = GUICtrlCreateLabel("Send with Email", 15, 106, 100, 20)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 9, $FW_NORMAL, $GUI_FONTITALIC, "Times New Roman")
$aGUI[$idTxtPassphrase] = GUICtrlCreateInput("", 104, 80, 326, 34)
GUICtrlSetState(-1, $GUI_FOCUS)
idTxtPassphrase_SetStyle(1)
$aGUI[$idBtnPassphrase] = GUICtrlCreateButton("C&opy", 442, 80, 58, 34)
GUICtrlSetOnEvent(-1, "GUIEvents")
GUICtrlSetState(-1, $GUI_DISABLE)
$aGUI[$idLblPassphraseMsg] = GUICtrlCreateLabel("", 104, 114, 330, 40, $SS_CENTER)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTITALIC)
$aGUI[$idLblPassword] = GUICtrlCreateLabel("Password:", 30, 156, 80, 20)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTUNDER)
$aGUI[$idLblPasswordUse] = GUICtrlCreateLabel("Use to Encrypt", 24, 175, 100, 20)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 9, $FW_NORMAL, $GUI_FONTITALIC, "Times New Roman")
$aGUI[$idTxtPassword] = GUICtrlCreateInput("", 104, 150, 326, 34, BitOR($ES_READONLY, $SS_CENTER, $ES_PASSWORD))
idTxtPassword_Enable(False)
GUICtrlSetFont(-1, 18, $FW_BOLD, Default, "Consolas")
$aGUI[$idBtnPassword] = GUICtrlCreateButton("Co&py", 442, 150, 58, 34)
GUICtrlSetOnEvent(-1, "GUIEvents")
GUICtrlSetState(-1, $GUI_DISABLE)
$aGUI[$idLblPasswordMsg] = GUICtrlCreateLabel("", 104, 184, 330, 20, $SS_CENTER)
GUICtrlSetColor(-1, $COLOR_RED)
GUICtrlSetFont(-1, 10, $FW_BOLD, $GUI_FONTITALIC)

GUIRegisterMsg($WM_ACTIVATE, "WM_ACTIVATE")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUISetOnEvent($GUI_EVENT_CLOSE, "GUIEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "GUIEvents")
GUISetOnEvent($GUI_EVENT_RESTORE, "GUIEvents")
TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "GUIShow")

Local $aAccelKeys[][2] = [["^m", $aGUI[$idMnuFileKeyMgr]], ["^i", $aGUI[$idMnuFileImport]], ["^e", $aGUI[$idMnuFileExport]], ["^q", $aGUI[$idMnuFileQuit]]]
GUISetAccelerators($aAccelKeys)

$aKeyManagerGUI[$hKeyManagerGUI] = GUICreate("Key Archive Manager", 318, 400, -1, -1, $WS_SIZEBOX)

$aKeyManagerGUI[$idKMMnuFile] = GUICtrlCreateMenu("&File")
$aKeyManagerGUI[$idKMMnuFileNew] = GUICtrlCreateMenuItem("Add &New Key" & @TAB & "Ctrl + N", $aKeyManagerGUI[$idKMMnuFile])
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
GUICtrlCreateMenuItem("", $aKeyManagerGUI[$idKMMnuFile])
$aKeyManagerGUI[$idKMMnuFileImport] = GUICtrlCreateMenuItem("&Import Key Archive" & @TAB & "Ctrl + I", $aKeyManagerGUI[$idKMMnuFile])
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
$aKeyManagerGUI[$idKMMnuFileExport] = GUICtrlCreateMenuItem("&Export Key Archive" & @TAB & "Ctrl + E", $aKeyManagerGUI[$idKMMnuFile])
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
GUICtrlCreateMenuItem("", $aKeyManagerGUI[$idKMMnuFile])
$aKeyManagerGUI[$idKMMnuFileClose] = GUICtrlCreateMenuItem("&Close" & @TAB & "Ctrl + C", $aKeyManagerGUI[$idKMMnuFile])
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
$aKeyManagerGUI[$idKMMnuEdit] = GUICtrlCreateMenu("&Edit")
$aKeyManagerGUI[$idKMMnuEditRemove] = GUICtrlCreateMenuItem("&Remove Selected Key(s)" & @TAB & "DEL", $aKeyManagerGUI[$idKMMnuEdit])
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
GUICtrlCreateMenuItem("", $aKeyManagerGUI[$idKMMnuEdit])
$aKeyManagerGUI[$idKMMnuEditSelectAll] = GUICtrlCreateMenuItem("Select &All" & @TAB & "Ctrl + A", $aKeyManagerGUI[$idKMMnuEdit])
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
$aKeyManagerGUI[$idKMMnuEditDeselectAll] = GUICtrlCreateMenuItem("&Deselect All" & @TAB & "Ctrl + D", $aKeyManagerGUI[$idKMMnuEdit])
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
$aKeyManagerGUI[$idKMListView] = GUICtrlCreateListView("Active|Key Date|Key Value|GUID", 4, 4, 308, 310, BitOR($LVS_SORTDESCENDING, $WS_BORDER, $LVS_SHOWSELALWAYS), BitOR($LVS_EX_CHECKBOXES, $LVS_EX_FULLROWSELECT))
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 50)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 100)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 158)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 0)
$aKeyManagerGUI[$hKMListView] = GUICtrlGetHandle($aKeyManagerGUI[$idKMListView])
$aKeyManagerGUI[$idKMListViewDummy] = GUICtrlCreateDummy()
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
$aKeyManagerGUI[$idKMBtnCancel] = GUICtrlCreateButton("&Cancel", 10, 320, 140, 30)
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
$aKeyManagerGUI[$idKMBtnSave] = GUICtrlCreateButton("&Save", 166, 320, 140, 30, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
GUICtrlSetOnEvent(-1, "KeyManager_GUIEvents")
GUICtrlSetState(-1, $GUI_DISABLE)
_GUICtrlListView_RegisterSortCallBack($aKeyManagerGUI[$hKMListView], False, False)

GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
GUIRegisterMsg($WM_GETMINMAXINFO, "WM_GETMINMAXINFO")
GUISetOnEvent($GUI_EVENT_CLOSE, "KeyManager_GUIEvents")

Local $aAccelKeys[][2] = [["^n", $aKeyManagerGUI[$idKMMnuFileNew]], ["^i", $aKeyManagerGUI[$idKMMnuFileImport]], ["^e", $aKeyManagerGUI[$idKMMnuFileExport]], ["^c", $aKeyManagerGUI[$idKMMnuFileClose]], ["^a", $aKeyManagerGUI[$idKMMnuEditSelectAll]], ["^d", $aKeyManagerGUI[$idKMMnuEditDeselectAll]]]
GUISetAccelerators($aAccelKeys)
#EndRegion - UI Creation

#Region - Main
UpdatePassGen()

_Crypt_Startup()

_LegacyKeyConvert()

UILock()

KeyReadFromReg(RegistryKeyGetCurrent())
If idTxtKey_Read() Then UILock(False)

If AutoStartIsEnabled() Then GUICtrlSetState($aGUI[$idMnuOptionsAutoStart], $GUI_CHECKED)
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
	If $bChangesMade Then
		If Not KeyManager_Close("Quit") Then Return 0
	EndIf
	_GUICtrlListView_UnRegisterSortCallBack($aKeyManagerGUI[$hKMListView])
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
			If @GUI_WinHandle = $aGUI[$hGUI] Then
				Return (CloseToTrayIsEnabled()) ? GUIHide() : _Exit()
			Else
				KeyManager_Close()
			EndIf
		Case $GUI_EVENT_MINIMIZE
			GUIMinimize()
		Case $GUI_EVENT_RESTORE
			GUIRestore()
		Case $aGUI[$idMnuFileKeyMgr]
			KeyManager_OpenGUI()
		Case $aGUI[$idMnuFileImport]
			KeyManager_OpenGUI()
			KeyArchiveImportRoutine()
		Case $aGUI[$idMnuFileExport]
			KeyArchiveExportRoutine()
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
	GUISetState(@SW_HIDE, $aGUI[$hGUI])
	GUISetState(@SW_DISABLE, $aGUI[$hGUI])
EndFunc   ;==>GUIHide

Func GUIMinimize()
	KeyHide()
	GUISetState(@SW_MINIMIZE)
EndFunc   ;==>GUIMinimize

Func GUIRestore()
	GUISetState(@SW_RESTORE)
EndFunc   ;==>GUIRestore

Func GUIShow()
	GUIRestore()
	TraySetState($TRAY_ICONSTATE_HIDE)
	GUISetState(@SW_ENABLE, $aGUI[$hGUI])
	GUISetState(@SW_SHOW, $aGUI[$hGUI])
	WinActivate($aGUI[$hGUI])
EndFunc   ;==>GUIShow

Func KeyManager_GUIEvents()
	$iCtrl = @GUI_CtrlId
	Switch $iCtrl
		Case $GUI_EVENT_CLOSE, $aKeyManagerGUI[$idKMMnuFileClose]
			KeyManager_Close()
		Case $aKeyManagerGUI[$idKMMnuFileNew]
			KeyManager_AddKey()
			If idKMListView_GetCount() = 1 Then idKMListView_CheckItem(0)
		Case $aKeyManagerGUI[$idKMMnuEditRemove]
			KeyManager_RemoveSelected()
		Case $aKeyManagerGUI[$idKMMnuEditSelectAll]
			KeyManager_SelectAll()
		Case $aKeyManagerGUI[$idKMMnuEditDeselectAll]
			KeyManager_SelectAll(False)
		Case $aKeyManagerGUI[$idKMListViewDummy]
			KeyManager_KeyActivateEvent()
		Case $aKeyManagerGUI[$idKMListView]
			KeyManager_Sort()
		Case $aKeyManagerGUI[$idKMBtnCancel]
			KeyManager_Cancel()
		Case $aKeyManagerGUI[$idKMBtnSave]
			KeyManager_Save()
		Case $aKeyManagerGUI[$idKMMnuFileImport]
			KeyArchiveImportRoutine()
		Case $aKeyManagerGUI[$idKMMnuFileExport]
			KeyArchiveExportRoutine()
	EndSwitch
EndFunc   ;==>KeyManager_GUIEvents

Func TrayEvents()
	$iCtrl = @TRAY_ID
	Switch $iCtrl
		Case $aGUI[$idTrayOpen]
			GUIShow()
		Case $aGUI[$idTrayQuit]
			_Exit()
	EndSwitch
EndFunc   ;==>TrayEvents

Func WM_ACTIVATE($hWnd, $iMsg, $wParam, $lParam)
	Local $iCode = BitAND($wParam, 0xFFFF)
	Switch $hWnd
		Case $aGUI[$hGUI]
			Switch $iCode
				Case 0 ;WA_INACTIVE
					KeyHide()
					PasswordHide()
				Case 1 To 2 ;WA_ACTIVE & WA_CLICKACTIVE
					HotKeyManager()
			EndSwitch
		Case $aKeyManagerGUI[$hKeyManagerGUI]
			Switch $iCode
				Case 0 ;WA_INACTIVE
					If Not $g_bKeyManagerBusy Then
						If $g_iEditing Then KeyManager_DeleteTempEditControl()
						KeyManager_Hide()
					EndIf
				Case 1 To 2 ;WA_ACTIVE & WA_CLICKACTIVE
					HotKeyManager($e_HotKeyDEL + $e_HotKeyESC)
			EndSwitch
	EndSwitch
EndFunc   ;==>WM_ACTIVATE

Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	Local $iIDFrom = BitAND($wParam, 0xFFFF) ; LoWord - this gives the control which sent the message
	Local $iCode = BitShift($wParam, 16) ; HiWord - this gives the message that was sent

	Switch $wParam
		Case $e_KMActivate
			KeyManager_SetActive()
			Return $GUI_RUNDEFMSG
		Case $e_KMEditDate
			KeyManager_ActivatedItem($aCurrentListViewItem[0], 1)
			KeyManager_EditItem()
			Return $GUI_RUNDEFMSG
		Case $e_KMEditValue
			KeyManager_ActivatedItem($aCurrentListViewItem[0], 2)
			KeyManager_EditItem()
			Return $GUI_RUNDEFMSG
		Case $e_KMEditRemove
			KeyManager_RemoveSelected()
			Return $GUI_RUNDEFMSG
	EndSwitch

	Switch $iCode
		Case $EN_CHANGE ; If we have the correct message
			Switch $iIDFrom ; See if it comes from one of the inputs
				Case $aGUI[$idTxtPassphrase]
					Return idTxtPassphrase_OnChange()
				Case $idTempEditCtrl
					TempEditControlValidKeyValue()
					Return $GUI_RUNDEFMSG
			EndSwitch
		Case $EN_SETFOCUS
			Switch $iIDFrom
				Case $aGUI[$idTxtPassphrase]
					idTxtPassphrase_SetStyle(0)
					If idTxtPassphrase_Read() = $sEmptyPassphraseMsg Then idTxtPassphrase_SetData("")
				Case $aGUI[$idTxtPassword]
					Return PasswordShow()
			EndSwitch
		Case $EN_KILLFOCUS
			Switch $iIDFrom
				Case $aGUI[$idTxtPassphrase]
					If Not StringLen(idTxtPassphrase_Read()) Then
						idTxtPassphrase_SetStyle(1)
						idTxtPassphrase_SetData($sEmptyPassphraseMsg)
					EndIf
				Case $aGUI[$idTxtPassword]
					Return PasswordHide()
				Case $idTempEditCtrl

			EndSwitch
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func WM_GETMINMAXINFO($hWnd, $Msg, $wParam, $lParam)
	If $hWnd = $aKeyManagerGUI[$hKeyManagerGUI] Then
		$minmaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
		DllStructSetData($minmaxinfo, 7, 332) ; min X
		DllStructSetData($minmaxinfo, 8, 250) ; min Y
		DllStructSetData($minmaxinfo, 9, 332) ; max X
		DllStructSetData($minmaxinfo, 10, 700) ; max Y
		Return 0
	EndIf
EndFunc   ;==>WM_GETMINMAXINFO

Func WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
	$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "HwndFrom"))
	$iCode = DllStructGetData($tNMHDR, "Code")

	Switch $hWndFrom
		Case $hTempEditCtrl
			Switch $iCode
				Case $DTN_CLOSEUP
					KeyManager_EditDate()
			EndSwitch
			Return $GUI_RUNDEFMSG

		Case $aKeyManagerGUI[$hKMListView]
			Local $tInfo = DllStructCreate($tagNMLISTVIEW, $lParam)
			Local $iItem = DllStructGetData($tInfo, "Item")
			Local $iSubItem = DllStructGetData($tInfo, "SubItem")
			If $iItem < 0 Then Return $GUI_RUNDEFMSG
			Switch $iCode
				Case $NM_CLICK, $NM_DBLCLK
					If $iSubItem == 0 Then
						KeyManager_ActivatedItem($iItem, $iSubItem)
						GUICtrlSendToDummy($aKeyManagerGUI[$idKMListViewDummy])
					EndIf
				Case $NM_RCLICK
					KeyManager_ActivatedItem($iItem, $iSubItem)
					KeyManager_SelectAll(False)
					idKMListView_SelectItem($iItem)
					KeyManager_ContextMenu()
				Case $LVN_ITEMACTIVATE
					If $iSubItem == 0 Then Return $GUI_RUNDEFMSG
					KeyManager_ActivatedItem($iItem, $iSubItem)
					KeyManager_EditItem()
			EndSwitch
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY
#EndRegion - UI Event Functions

#Region - UI Control Functions
Func idBtnKey_Click()
	KeyManager_OpenGUI()
EndFunc   ;==>idBtnKey_Click

Func idBtnKey_Enabled($bFlag = True)
	GUICtrl_Enable($aGUI[$idBtnKey], $bFlag)
EndFunc   ;==>idBtnKey_Enabled

Func idBtnKey_Focus()
	GUICtrl_SetFocus($aGUI[$idBtnKey])
EndFunc   ;==>idBtnKey_Focus

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
	GUICtrl_SetData($aGUI[$idBtnRevealKey], ($iFlag = 1) ? "&Hide" : "&Show")
EndFunc   ;==>idBtnRevealKey_AcceleratorOption

Func idBtnRevealKey_Click()
	Local $iState = GUICtrl_Read($aGUI[$idBtnRevealKey])
	If $iState = $GUI_UNCHECKED Then
		KeyHide()
	Else
		KeyShow()
	EndIf
EndFunc   ;==>idBtnRevealKey_Click

Func idBtnRevealKey_Depressed($bFlag = True)
	GUICtrl_SetChecked($aGUI[$idBtnRevealKey], $bFlag)
EndFunc   ;==>idBtnRevealKey_Depressed

Func idBtnRevealKey_Enabled($bFlag = True)
	GUICtrl_Enable($aGUI[$idBtnRevealKey], $bFlag)
EndFunc   ;==>idBtnRevealKey_Enabled

Func idKMBtnSave_Enable($bFlag = True)
	GUICtrl_Enable($aKeyManagerGUI[$idKMBtnSave], $bFlag)
EndFunc   ;==>idKMBtnSave_Enable

Func idKMListView_CheckItem($iItem)
	Return _GUICtrlListView_SetItemChecked($aKeyManagerGUI[$hKMListView], $iItem)
EndFunc   ;==>idKMListView_CheckItem

Func idKMListView_GetCount()
	Return _GUICtrlListView_GetItemCount($aKeyManagerGUI[$hKMListView])
EndFunc   ;==>idKMListView_GetCount

Func idKMListView_GetItem($iItem, $iSubItem)
	Return _GUICtrlListView_GetItemText($aKeyManagerGUI[$hKMListView], $iItem, $iSubItem)
EndFunc   ;==>idKMListView_GetItem

Func idKMListView_SelectItem($iItem)
	Return _GUICtrlListView_SetItemSelected($aKeyManagerGUI[$hKMListView], $iItem)
EndFunc   ;==>idKMListView_SelectItem

Func idLblPassphraseMsg_SetMessage($sMsg = "")
	GUICtrl_SetData($aGUI[$idLblPassphraseMsg], $sMsg)
EndFunc   ;==>idLblPassphraseMsg_SetMessage

Func idLblPasswordMsg_SetMessage($sMsg = "")
	GUICtrl_SetData($aGUI[$idLblPasswordMsg], $sMsg)
EndFunc   ;==>idLblPasswordMsg_SetMessage

Func idMnuOptionsAutoStart_Click()
	Local $iState = GUICtrl_MenuItemToggle($aGUI[$idMnuOptionsAutoStart])
	AutoStart(($iState == $GUI_CHECKED) ? True : False)
EndFunc   ;==>idMnuOptionsAutoStart_Click

Func idMnuOptionsCloseToTray_Click()
	Local $iState = GUICtrl_MenuItemToggle($aGUI[$idMnuOptionsCloseToTray])
	CloseToTraySetting(($iState == $GUI_CHECKED) ? True : False)
EndFunc   ;==>idMnuOptionsCloseToTray_Click

Func idTempEditCtrl_Read()
	Return GUICtrl_Read($idTempEditCtrl)
EndFunc   ;==>idTempEditCtrl_Read

Func idTxtKey_Enable($bFlag = True)
	GUICtrl_Enable($aGUI[$idTxtKey], $bFlag)
EndFunc   ;==>idTxtKey_Enable

Func idTxtKey_IsEnabled()
	Return (GUICtrl_GetState($aGUI[$idTxtKey], $GUI_ENABLE)) ? True : False
EndFunc   ;==>idTxtKey_IsEnabled

Func idTxtKey_OnChange()
	AutoPurgeTimer(False)
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

Func idTxtPassphrase_OnChange()
	AutoPurgeTimer(False)
	If idTxtPassphrase_Read() <> $sEmptyPassphraseMsg Then
		If PassphraseIsValid() Then
			GUICtrl_Enable($aGUI[$idBtnPassphrase])
			GUICtrl_Enable($aGUI[$idBtnPassword])
			idTxtPassword_Enable()
			GeneratePassword()
		Else
			GUICtrl_Enable($aGUI[$idBtnPassphrase], False)
			GUICtrl_Enable($aGUI[$idBtnPassword], False)
			idTxtPassword_Enable(False)
		EndIf
	Else
		idTxtPassword_Enable(False)
	EndIf
EndFunc   ;==>idTxtPassphrase_OnChange

Func idTxtPassphrase_Read()
	Return GUICtrl_Read($aGUI[$idTxtPassphrase])
EndFunc   ;==>idTxtPassphrase_Read

Func idTxtPassphrase_SetData($sValue)
	GUICtrl_SetData($aGUI[$idTxtPassphrase], $sValue)
EndFunc   ;==>idTxtPassphrase_SetData

Func idTxtPassphrase_SetStyle($iFlag = 0)
	Local $idCtrl = $aGUI[$idTxtPassphrase]
	Local $iFontAttr, $iColor, $iStyle, $iSize

	Switch $iFlag
		Case 0
			$iFontAttr = Default
			$iColor = 0x000000
			$iStyle = $ES_LEFT
			$iSize = 18
		Case 1
			$iFontAttr = $GUI_FONTITALIC
			$iColor = 0xA0A0A0
			$iStyle = $SS_CENTER
			$iSize = 16
	EndSwitch
	GUICtrlSetFont($idCtrl, $iSize, $FW_BOLD, $iFontAttr, "Consolas")
	GUICtrlSetColor($idCtrl, $iColor)
	GUICtrlSetStyle($idCtrl, $iStyle)
EndFunc   ;==>idTxtPassphrase_SetStyle

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
			If @ScriptFullPath <> $PROGRAMPATH Then
				SplashTextOn("PassGen Updating...", "Please wait a few moments for" & @CRLF & "PassGen to update and restart", 400, 120)
				ShellExecute($PROGRAMPATH)
				_Exit()
			EndIf
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

Func HotKeyManager($iHotKeys = 0)
	HotKeySet("{ESC}")
	HotKeySet("{DEL}")
	HotKeySet("{ENTER}")
	If BitAND($iHotKeys, $e_HotKeyDEL) = $e_HotKeyDEL Then
		HotKeySet("{DEL}", "KeyManager_RemoveSelected")
	EndIf
	If BitAND($iHotKeys, $e_HotKeyESC) = $e_HotKeyESC Then
		HotKeySet("{ESC}", "KeyManager_Cancel")
	EndIf
	If BitAND($iHotKeys, $e_HotKeyEnter) = $e_HotKeyEnter Then
		HotKeySet("{ENTER}", "KeyManager_Save")
	EndIf
EndFunc   ;==>HotKeyManager

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
	$sText = StringStripWS(String($sText), $STR_STRIPLEADING + $STR_STRIPTRAILING)
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

Func KeyArchiveClear($bForce = False)
	If $bForce = False Then
		If MsgBox(BitOR($MB_ICONWARNING, $MB_YESNO, $MB_DEFBUTTON2), "Clear Key Archive", "Are you absolutely sure you want to clear the Key Archive?" & _
				@CRLF & @CRLF & "This action can not be undone!") <> $IDYES Then Return 0
	EndIf
	KeyArchiveGet()
	For $sKey In $g_aKeyArchive
		RegDelete($REGKEYPATH, $sKey)
	Next
	RegDelete($REGKEYPATH, $REGKEYCURRENT)
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
	Return KeyManager_Busy(False)
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
	HotKeyManager()
	KeyManager_Busy()
	If Not _GUICtrlListView_GetItemCount($aKeyManagerGUI[$hKMListView]) Then Return KeyArchiveExportCanceled("Nothing to Export")
	If $bChangesMade Then
		$iRet = MsgBox(BitOR($MB_YESNO, $MB_ICONWARNING, $MB_DEFBUTTON2), "Unsaved Changes", "You have unsaved changes." & _
				@CRLF & "Unsaved changes may not be exported." & @CRLF & @CRLF & "Do you want to continue with exporting?")
		If $iRet = $IDNO Then
			KeyArchiveExportCanceled()
			Return 0
		ElseIf $iRet = $IDCANCEL Then
			Return KeyManager_Show()
		EndIf
	EndIf

	Local $sPassGenExportFilePath = FileSaveDialog("PassGen Key Archive Export", @MyDocumentsDir & "\", "PassGen Key Archive (*.pge)", $FD_PROMPTOVERWRITE, "PassGenKeyArchive.pge")
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
	KeyManager_Busy(False)
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
	_ArraySort($aKeyArchive, 1, 0, 0, 1)
	$g_aKeyArchive = $aKeyArchive
	_ArrayColDelete($g_aKeyArchive, 1, True)
	Return $aKeyArchive
EndFunc   ;==>KeyArchiveGet

Func KeyArchiveGetCount()
	If Not IsArray($g_aKeyArchive) Then KeyArchiveGet()
	If IsArray($g_aKeyArchive) Then
		Return UBound($g_aKeyArchive)
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc   ;==>KeyArchiveGetCount

Func KeyArchiveImportCanceled($sCustomMsg = "")
	If $sCustomMsg <> "" Then
		$sCustomMsg &= @CRLF & @CRLF
	EndIf
	$sCustomMsg &= "Import Canceled"
	MsgBox(0, "", $sCustomMsg)
	Return KeyManager_Busy(False)
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
	HotKeyManager()
	KeyManager_Busy()
	Local $sPassGenExportFilePath = FileOpenDialog("PassGen Key Archive Import", @MyDocumentsDir & "\", "PassGen Key Archive (*.pge)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST))
	If @error Then Return KeyArchiveImportCanceled()
	If Not FileExists($sPassGenExportFilePath) Then Return KeyArchiveImportCanceled("Unable to locate PassGen export file")
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
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	If KeyArchiveGetCount() Then
		$iImportOverwrite = MsgBox(BitOR($MB_YESNOCANCEL, $MB_ICONQUESTION), "Import Key Archive", _
				"Do you want to overwrite the existing Key Archive?" & @CRLF & @CRLF & "Yes = Overwrite, No = Append")
		If $iImportOverwrite = $IDYES Then _GUICtrlListView_DeleteAllItems($hCtrl)
		If $iImportOverwrite = $IDCANCEL Then
			$dKeys = 0
			Return KeyArchiveImportCanceled()
		EndIf
	EndIf
	_GUICtrlListView_BeginUpdate($hCtrl)
	Local $aKeys = StringSplit(BinaryToString($dKeys), "|", $STR_NOCOUNT)
	For $sKey In $aKeys
		Local $sKeyGUID = _WinAPI_CreateGUID()
		Local $sDate = StringLeft($sKey, $DATEFORMATBYTELEN)
		Local $sValue = StringRight($sKey, StringLen($sKey) - $DATEFORMATBYTELEN)
		KeyManager_AddKey($sKeyGUID, $sDate, $sValue)
	Next
	If KeyManager_GetActive() = -1 Then
		Local $iCount = _GUICtrlListView_GetItemCount($hCtrl)
		If $iCount > 1 Then
			_GUICtrlListView_SortItems($hCtrl, 1)
			Local $sFirstDate = idKMListView_GetItem(0, 1)
			Local $sLastDate = idKMListView_GetItem($iCount - 1, 1)
			Local $iDateDiff = _DateDiff("D", $sFirstDate, $sLastDate)
			If $iDateDiff >= 0 Then _GUICtrlListView_SortItems($hCtrl, 1)
			idKMListView_CheckItem(0)
		Else
			idKMListView_CheckItem(0)
		EndIf
	EndIf
	_GUICtrlListView_EndUpdate($hCtrl)
	KeyArchiveClearFromMem()
	KeyManager_ChangeMade()
	KeyManager_Busy(False)
EndFunc   ;==>KeyArchiveImportRoutine

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

Func KeyGetValue()
	Return StringStripWS(idTxtKey_Read(), $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>KeyGetValue

Func KeyHide()
	idBtnRevealKey_Depressed(False)
	idBtnRevealKey_AcceleratorOption()
	InputboxMask($aGUI[$idTxtKey])
EndFunc   ;==>KeyHide

Func KeyManager_ActivatedItem($iItem, $iSubItem)
	$aCurrentListViewItem[0] = $iItem
	$aCurrentListViewItem[1] = $iSubItem
EndFunc   ;==>KeyManager_ActivatedItem

Func KeyManager_ActivatedItemPosition($hCtrl, $iItem, $iSubItem)
	Local $aListViewPos = ControlGetPos($aKeyManagerGUI[$hKeyManagerGUI], "", $hCtrl)
	Local $aSubItemRect = _GUICtrlListView_GetSubItemRect($hCtrl, $iItem, $iSubItem)
	Local $iLeft = $aListViewPos[0] + $aSubItemRect[0] + 1
	Local $iTop = $aListViewPos[1] + $aSubItemRect[1] + 2
	Local $iWidth = $aSubItemRect[2] - $aSubItemRect[0] - 2
	Local $iHeight = $aSubItemRect[3] - $aSubItemRect[1] - 1
	Local $aListViewItemPos[4] = [$iLeft, $iTop, $iWidth, $iHeight]
	Return $aListViewItemPos
EndFunc   ;==>KeyManager_ActivatedItemPosition

Func KeyManager_AddKey($sGUID = "", $sDate = _NowCalcDate(), $sValue = "New Key!", $bActive = False)
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	If $sGUID = "" Then
		$sGUID = _WinAPI_CreateGUID()
		KeyManager_ChangeMade()
	EndIf
	_GUICtrlListView_BeginUpdate($hCtrl)
	Local $iNewItem = _GUICtrlListView_AddItem($hCtrl, "")
	_GUICtrlListView_AddSubItem($hCtrl, $iNewItem, $sDate, 1)
	_GUICtrlListView_AddSubItem($hCtrl, $iNewItem, $sValue, 2)
	_GUICtrlListView_AddSubItem($hCtrl, $iNewItem, $sGUID, 3)
	If $bActive Then _GUICtrlListView_SetItemChecked($hCtrl, $iNewItem)
	_GUICtrlListView_EndUpdate($hCtrl)
EndFunc   ;==>KeyManager_AddKey

Func KeyManager_Busy($bFlag = True)
	$g_bKeyManagerBusy = $bFlag
EndFunc   ;==>KeyManager_Busy

Func KeyManager_Cancel()
	ToolTip("")
	If $g_iEditing Then
		Switch $g_iEditing
			Case 1

			Case 2
				KeyManager_DeleteTempEditControl()
				If $bChangesMade Then KeyManager_ChangeMade()
		EndSwitch
		Return 0
	EndIf
	KeyManager_Close("Cancel")
EndFunc   ;==>KeyManager_Cancel

Func KeyManager_ChangeMade($bFlag = True)
	$bChangesMade = $bFlag
	Switch $bFlag
		Case True
			idKMBtnSave_Enable()
		Case False
			idKMBtnSave_Enable(False)
	EndSwitch
EndFunc   ;==>KeyManager_ChangeMade

Func KeyManager_Close($sMsg = "Close", $bForce = False)
	Local $bQuit = False
	If $bChangesMade And $bForce = False Then
		KeyManager_Busy()
		HotKeyManager()
		$iRet = MsgBox(BitOR($MB_YESNO, $MB_ICONWARNING, $MB_DEFBUTTON2), "Unsaved Changes", "You have unsaved changes." & _
				@CRLF & "Any unsaved changes will be lost." & @CRLF & @CRLF & "Are you sure you want to " & $sMsg & "?")
		If $iRet = $IDNO Then
			Return KeyManager_Busy(False)
		ElseIf $iRet = $IDYES Then
			$bChangesMade = False
			$bChangesPending = False
			$bQuit = True
			UILock(False)
		EndIf
	EndIf
	KeyManager_Busy(False)
	KeyManager_Hide()
	Return $bQuit
EndFunc   ;==>KeyManager_Close

Func KeyManager_ContextMenu()
	If Not _GUICtrlListView_GetSelectedCount($aKeyManagerGUI[$hKMListView]) Then Return True
	Local $hMenu = _GUICtrlMenu_CreatePopup($MNS_AUTODISMISS)
	_GUICtrlMenu_InsertMenuItem($hMenu, 0, "Set as Active Key", $e_KMActivate)
	_GUICtrlMenu_InsertMenuItem($hMenu, 1, "")
	_GUICtrlMenu_InsertMenuItem($hMenu, 2, "Change Key Date", $e_KMEditDate)
	_GUICtrlMenu_InsertMenuItem($hMenu, 3, "Modify Key Value", $e_KMEditValue)
	_GUICtrlMenu_InsertMenuItem($hMenu, 4, "")
	_GUICtrlMenu_InsertMenuItem($hMenu, 5, "Remove Key", $e_KMEditRemove)
	_GUICtrlMenu_TrackPopupMenu($hMenu, $aKeyManagerGUI[$hKeyManagerGUI])
	_GUICtrlMenu_DestroyMenu($hMenu)
	Return True
EndFunc   ;==>KeyManager_ContextMenu

Func KeyManager_DeleteTempEditControl()
	$g_iEditing = False
	GUICtrlDelete($idTempEditCtrl)
	$idTempEditCtrl = -1
	$hTempEditCtrl = -1
	GUICtrlSetState($aKeyManagerGUI[$idKMListView], BitOR($GUI_ENABLE, $GUI_FOCUS))
EndFunc   ;==>KeyManager_DeleteTempEditControl

Func KeyManager_EditDate()
	Local $sNewDTPValue = GUICtrlRead($idTempEditCtrl)
	Local $sCurrentDTPValue = idKMListView_GetItem($aCurrentListViewItem[0], $aCurrentListViewItem[1])
	If $sNewDTPValue <> $sCurrentDTPValue Then
		KeyManager_SetValue($sNewDTPValue)
	Else
		KeyManager_DeleteTempEditControl()
		If $bChangesMade Then KeyManager_ChangeMade()
	EndIf
EndFunc   ;==>KeyManager_EditDate

Func KeyManager_EditItem()
	GUICtrlSetState($aKeyManagerGUI[$idKMListView], $GUI_DISABLE)
	Local $hWndFrom = $aKeyManagerGUI[$hKMListView]
	Local $iItem = $aCurrentListViewItem[0]
	Local $iSubItem = $aCurrentListViewItem[1]
	Local $aItemPos = KeyManager_ActivatedItemPosition($hWndFrom, $iItem, $iSubItem)
	_GUICtrlListView_SetItemSelected($hWndFrom, -1, False, False)
	Local $sText = _GUICtrlListView_GetItemText($hWndFrom, $iItem, $iSubItem)
	Switch $iSubItem
		Case 1
			$idTempEditCtrl = GUICtrlCreateDate($sText, $aItemPos[0], $aItemPos[1], $aItemPos[2], $aItemPos[3], $DTS_SHORTDATEFORMAT)
			GUICtrlSendMsg(-1, $DTM_SETFORMATW, 0, "yyyy/MM/dd")
		Case 2
			$idTempEditCtrl = GUICtrlCreateInput($sText, $aItemPos[0], $aItemPos[1], $aItemPos[2], $aItemPos[3])
	EndSwitch
	$hTempEditCtrl = GUICtrlGetHandle($idTempEditCtrl)
	Sleep(10)
	_WinAPI_SetFocus($hTempEditCtrl)
	$g_iEditing = $iSubItem
	Switch $iSubItem
		Case 1
			_SendMessage($hTempEditCtrl, $WM_LBUTTONDOWN, 1, $aItemPos[2] - 10)
			idKMBtnSave_Enable(False)
		Case 2
			KeyManager_EditValue()
	EndSwitch
EndFunc   ;==>KeyManager_EditItem

Func KeyManager_EditValue()
	$g_iEditing = 2
	idKMBtnSave_Enable(False)
	GUICtrlSendMsg($idTempEditCtrl, $EM_SETSEL, 0, -1)
	HotKeyManager($e_HotKeyESC + $e_HotKeyEnter)
	TempEditControlValidKeyValue()
EndFunc   ;==>KeyManager_EditValue

Func KeyManager_EditValueSave()
	KeyManager_Busy(False)
	HotKeyManager($e_HotKeyDEL)
	Local $bValidKeyValue = TempEditControlValidKeyValue()
	If Not $bValidKeyValue Then
		Return KeyManager_EditValue()
	EndIf
	Local $sText = GUICtrlRead($idTempEditCtrl)
	Local $sCurrentText = idKMListView_GetItem($aCurrentListViewItem[0], $aCurrentListViewItem[1])
	If $sText <> $sCurrentText Then
		KeyManager_SetValue($sText)
	Else
		KeyManager_DeleteTempEditControl()
		If $bChangesMade Then KeyManager_ChangeMade()
	EndIf
	ToolTip("")
EndFunc   ;==>KeyManager_EditValueSave

Func KeyManager_GetActive()
	Local $iActiveItem = -1
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	For $iX = 0 To _GUICtrlListView_GetItemCount($hCtrl)
		If _GUICtrlListView_GetItemChecked($hCtrl, $iX) Then $iActiveItem = $iX
	Next
	Return $iActiveItem
EndFunc   ;==>KeyManager_GetActive

Func KeyManager_Hide()
	ToolTip("")
	HotKeyManager()
	If $bChangesMade Then $bChangesPending = True
	UILock(False)
	GUISetState(@SW_HIDE, $aKeyManagerGUI[$hKeyManagerGUI])
	GUISwitch($aGUI[$hGUI])
	KeyHide()
EndFunc   ;==>KeyManager_Hide

Func KeyManager_KeyActivateEvent()
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	_GUICtrlListView_BeginUpdate($hCtrl)
	Local $iActivatedItem = $aCurrentListViewItem[0]
	Local $iActivatedItemIsChecked = _GUICtrlListView_GetItemChecked($hCtrl, $iActivatedItem)
	If $iActivatedItemIsChecked Then
		For $iX = 0 To _GUICtrlListView_GetItemCount($hCtrl) - 1
			If $iX <> $iActivatedItem Then
				_GUICtrlListView_SetItemChecked($hCtrl, $iX, False)
			EndIf
		Next
	Else
		If KeyManager_GetActive() = -1 Then
			_GUICtrlListView_SetItemChecked($hCtrl, $iActivatedItem)
		EndIf
	EndIf
	KeyManager_KeySetActive(KeyManager_GetActive())
	_GUICtrlListView_EndUpdate($hCtrl)
EndFunc   ;==>KeyManager_KeyActivateEvent

Func KeyManager_KeySetActive($iIndex)
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	Local $sGUID = _GUICtrlListView_GetItemText($hCtrl, $iIndex, 3)
	Local $bActive = _GUICtrlListView_GetItemChecked($hCtrl, $iIndex)
	If $bActive Then RegistryKeySelect($sGUID)
	idTxtKey_SetData(KeyArchiveParseValue(RegistryKeyGetCurrent()))
EndFunc   ;==>KeyManager_KeySetActive

Func KeyManager_OpenGUI()
	If $bChangesPending = True Then
		$bChangesPending = False
	Else
		KeyManager_ChangeMade(False)
		KeyManager_Populate()
	EndIf
	KeyManager_Show()
EndFunc   ;==>KeyManager_OpenGUI

Func KeyManager_Populate()
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	_GUICtrlListView_BeginUpdate($hCtrl)
	_GUICtrlListView_DeleteAllItems($hCtrl)
	Local $aKeyArchive = KeyArchiveGet()
	Local $iCount = UBound($aKeyArchive)
	If Not $iCount Then Return 0
	Local $sCurrentKey = RegistryKeyGetCurrent()
	For $iX = $iCount - 1 To 0 Step -1
		Local $sDate = StringLeft($aKeyArchive[$iX][1], $DATEFORMATBYTELEN)
		Local $sValue = StringRight($aKeyArchive[$iX][1], StringLen($aKeyArchive[$iX][1]) - $DATEFORMATBYTELEN)
		Local $sGUID = $aKeyArchive[$iX][0]
		Local $bActive = ($sGUID = $sCurrentKey) ? True : False
		KeyManager_AddKey($sGUID, $sDate, $sValue, $bActive)
	Next
	_GUICtrlListView_EndUpdate($hCtrl)
EndFunc   ;==>KeyManager_Populate

Func KeyManager_RemoveSelected()
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	If Not _GUICtrlListView_GetSelectedCount($hCtrl) Then Return 0
	_GUICtrlListView_BeginUpdate($hCtrl)
	_GUICtrlListView_DeleteItemsSelected($hCtrl)
	If KeyManager_GetActive() = -1 Then
		If _GUICtrlListView_GetItemCount($hCtrl) Then _GUICtrlListView_SetItemChecked($hCtrl, 0)
	EndIf
	_GUICtrlListView_EndUpdate($hCtrl)
	KeyManager_ChangeMade()
EndFunc   ;==>KeyManager_RemoveSelected

Func KeyManager_Save()
	If $g_iEditing Then
		Switch $g_iEditing
			Case 1

			Case 2
				KeyManager_EditValueSave()
		EndSwitch
		Return 0
	EndIf
	KeyArchiveClear(True)
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	For $iX = 0 To _GUICtrlListView_GetItemCount($hCtrl) - 1
		Local $sDate = _GUICtrlListView_GetItemText($hCtrl, $iX, 1)
		Local $sValue = _GUICtrlListView_GetItemText($hCtrl, $iX, 2)
		Local $sGUID = _GUICtrlListView_GetItemText($hCtrl, $iX, 3)
		Local $bActive = _GUICtrlListView_GetItemChecked($hCtrl, $iX)
		If $bActive Then RegistryKeySelect($sGUID)
		KeySaveToReg($sGUID, $sValue, $sDate)
	Next
	idTxtKey_SetData(KeyArchiveParseValue(RegistryKeyGetCurrent()))
	$bChangesMade = False
	idKMBtnSave_Enable(False)
EndFunc   ;==>KeyManager_Save

Func KeyManager_SelectAll($bFlag = True)
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	_GUICtrlListView_BeginUpdate($hCtrl)
	For $iX = 0 To _GUICtrlListView_GetItemCount($hCtrl) - 1
		_GUICtrlListView_SetItemSelected($hCtrl, $iX, $bFlag)
	Next
	_GUICtrlListView_EndUpdate($hCtrl)
EndFunc   ;==>KeyManager_SelectAll

Func KeyManager_SetActive()
	idKMListView_CheckItem($aCurrentListViewItem[0])
	KeyManager_KeyActivateEvent()
EndFunc   ;==>KeyManager_SetActive

Func KeyManager_SetValue($sValue)
	_GUICtrlListView_SetItemText($aKeyManagerGUI[$hKMListView], $aCurrentListViewItem[0], $sValue, $aCurrentListViewItem[1])
	KeyManager_ChangeMade()
	KeyManager_DeleteTempEditControl()
EndFunc   ;==>KeyManager_SetValue

Func KeyManager_Show()
	UILock()
	KeyManager_Busy(False)
	GUISwitch($aKeyManagerGUI[$hKeyManagerGUI])
	GUISetState(@SW_SHOW)
	KeyShow()
EndFunc   ;==>KeyManager_Show

Func KeyManager_Sort($iCol = -1)
	Local $hCtrl = $aKeyManagerGUI[$hKMListView]
	If $iCol = -1 Then $iCol = GUICtrlGetState($aKeyManagerGUI[$idKMListView])
	If $iCol = 0 Then Return 0
	_GUICtrlListView_BeginUpdate($hCtrl)
	_GUICtrlListView_SortItems($hCtrl, $iCol)
	_GUICtrlListView_EndUpdate($hCtrl)
EndFunc   ;==>KeyManager_Sort

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

Func KeySaveToReg($sGUID, $sValue, $sDate = "")
	If $sDate = "" Then $sDate = @YEAR & "/" & @MON & "/" & @MDAY
	$hProtectedKey = KeyProtect($sDate & $sValue)
	RegistryKeyWriteBinary($sGUID, $hProtectedKey)
EndFunc   ;==>KeySaveToReg

Func KeyShow()
	idBtnRevealKey_Depressed()
	idBtnRevealKey_AcceleratorOption(1)
	InputboxMask($aGUI[$idTxtKey], False)
EndFunc   ;==>KeyShow

Func KeyUnprotect($hProtectedKey)
	Return _CryptUnprotectData($hProtectedKey)
EndFunc   ;==>KeyUnprotect

Func PassphraseGetValue()
	Return StringStripWS(idTxtPassphrase_Read(), $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>PassphraseGetValue

Func PassphraseIsValid()
	idTxtPassword_SetData("")
	idLblPassphraseMsg_SetMessage()
	idLblPasswordMsg_SetMessage()
	Local $sPassphrase = PassphraseGetValue()
	If StringLen($sPassphrase) < 1 Then Return False
	$bIsPassphraseComplex = IsStringComplex($sPassphrase, 8, 0, 0)
	If $bIsPassphraseComplex = False Then
		Switch @error
			Case 1
				idLblPassphraseMsg_SetMessage("Passphrase is too short." & @CRLF & "Must contain at least 8 characters.")
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

Func TempEditControlValidKeyValue()
	Local $sValue = GUICtrlRead($idTempEditCtrl)
	Local $sMsg = ""
	Local $aWinPos = WinGetPos($aKeyManagerGUI[$hKeyManagerGUI])
	Local $aEditCtrlPos = ControlGetPos($aKeyManagerGUI[$hKMListView], "", $hTempEditCtrl)
	$bIsKeyComplex = IsStringComplex($sValue, 8, 3, 0)
	If Not $bIsKeyComplex Then
		Local $iError = @error
		Switch $iError
			Case 1
				$sMsg = "Key is too short." & @CRLF & @CRLF & "Must contain at least 8 characters."
			Case 2
				$sMsg = "Key does not meet complexity requirements." & @CRLF & @CRLF & "Must contain at least 3 of the 4 following requirements:" & @CRLF & "Uppercase, Lowercase, Number, Symbol."
		EndSwitch
	Else
		$sMsg = "Press the Enter key to finish editing," & @CRLF & "or press the Esc key to cancel."
	EndIf
	ToolTip($sMsg, ($aWinPos[0] + $aEditCtrlPos[0] + ($aEditCtrlPos[2] / 2) + 8), ($aWinPos[1] + $aEditCtrlPos[1] + ($aEditCtrlPos[3] / 2) + 50), "", 0, BitOR($TIP_BALLOON, $TIP_CENTER))
	Return $bIsKeyComplex
EndFunc   ;==>TempEditControlValidKeyValue

Func UILock($bFlag = True)
	AutoPurgeTimer(False)
	Switch $bFlag
		Case True
			If ClipboardContainsPass() Then ClipboardClear()
			idBtnKey_Focus()
			idBtnPassphrase_Enabled(False)
			idBtnPassword_Enabled(False)
			idBtnRevealKey_Enabled(False)
			idTxtPassphrase_Enable(False)
			idTxtPassword_Enable(False)
			idTxtPassword_SetData("")
			idLblPassphraseMsg_SetMessage()
			idLblPasswordMsg_SetMessage()
			If idTxtPassphrase_Read() = $sEmptyPassphraseMsg Then idTxtPassphrase_SetData("")
		Case False
			If idTxtKey_Read() = "" Then Return 0
			idBtnKey_Enabled()
			idBtnRevealKey_Enabled()
			idTxtPassphrase_Enable()
			idTxtPassword_Enable()
			idTxtKey_Visible()
			idBtnPassphrase_Enabled(False)
			idTxtPassphrase_OnChange()
			If Not StringLen(idTxtPassphrase_Read()) Then idTxtPassphrase_SetData($sEmptyPassphraseMsg)
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
	Local $sRes = '', $aRet, $hReg = DllStructCreate("int")
	Local $hRemoteReg = DllStructCreate("int")
	Local $FILETIME = DllStructCreate("dword;dword")
	Local $SYSTEMTIME1 = DllStructCreate("ushort;ushort;ushort;ushort;ushort;ushort;ushort;ushort")
	Local $SYSTEMTIME2 = DllStructCreate("ushort;ushort;ushort;ushort;ushort;ushort;ushort;ushort")
	Local $hAdvAPI = DllOpen('advapi32.dll'), $hKernel = DllOpen('kernel32.dll')
	If $hAdvAPI = -1 Or $hKernel = -1 Then Return SetError(1, $aRet[0], 'DLL Open Error!')

	$connect = DllCall("advapi32.dll", "int", "RegConnectRegistry", _
			"str", $iPC, _
			"int", $iRegHive, _
			"ptr", DllStructGetPtr($hRemoteReg))

	$aRet = DllCall("advapi32.dll", "int", "RegOpenKeyEx", _
			"int", DllStructGetData($hRemoteReg, 1), _
			"str", $sRegKey, _
			"int", 0, _
			"int", 0x20019, _
			"ptr", DllStructGetPtr($hReg))
	If $aRet[0] Then Return SetError(2, $aRet[0], 'Registry Key Open Error!')

	$aRet = DllCall("advapi32.dll", "int", "RegQueryInfoKey", _
			"int", DllStructGetData($hReg, 1), _
			"ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, "ptr", 0, _
			"ptr", DllStructGetPtr($FILETIME))
	If $aRet[0] Then Return SetError(3, $aRet[0], 'Registry Key Query Error!')


	$aRet = DllCall("advapi32.dll", "int", "RegCloseKey", _
			"int", DllStructGetData($hReg, 1))
	If $aRet[0] Then Return SetError(4, $aRet[0], 'Registry Key Close Error!')


	$aRet = DllCall("kernel32.dll", "int", "FileTimeToSystemTime", _
			"ptr", DllStructGetPtr($FILETIME), _
			"ptr", DllStructGetPtr($SYSTEMTIME1))
	If $aRet[0] = 0 Then Return SetError(5, 0, 'Time Convert Error!')


	$aRet = DllCall("kernel32.dll", "int", "SystemTimeToTzSpecificLocalTime", _
			"ptr", 0, _
			"ptr", DllStructGetPtr($SYSTEMTIME1), _
			"ptr", DllStructGetPtr($SYSTEMTIME2))
	If $aRet[0] = 0 Then Return SetError(5, 0, 'Time Convert Error!')

	$sRes &= StringFormat("%.2d", DllStructGetData($SYSTEMTIME2, 1)) & "/"
	$sRes &= StringFormat("%.2d", DllStructGetData($SYSTEMTIME2, 2)) & "/"
	$sRes &= StringFormat("%.2d", DllStructGetData($SYSTEMTIME2, 4))

	Return $sRes
EndFunc   ;==>_LegacyKeyRegGetTimeStamp
#EndRegion - Internal Functions
