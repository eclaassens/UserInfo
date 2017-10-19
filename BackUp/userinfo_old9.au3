#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#AutoIt3Wrapper_run_obfuscator=y
;#Obfuscator_parameters=/striponly /om

Opt("WinTitleMatchMode", 2)
Opt("TrayIconDebug", 0)
Opt("TrayIconHide", 1)
Opt("GUICloseOnESC", 0)

If Not @Compiled Then
	Opt("TrayIconDebug", 1)
	Opt("TrayIconHide", 0)
EndIf

#include <Array.au3>
#include <Date.au3>
#include <GUIConstants.au3>
#include <ButtonConstants.au3>
#include <IE.au3>
#include <GuiStatusBar.au3>
#include <File.au3>
#include <Math.au3>; importing the max command
#include <Misc.au3>; importing _singleton
#include <ComboConstants.au3>
#include <GuiListView.au3>
#include <INet.au3>
#include <String.au3>
#include <localization.au3>; load translations

#Region  ; Define AD Constants
Global Const $ADS_GROUP_TYPE_GLOBAL_GROUP = 0x2
Global Const $ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 0x4
Global Const $ADS_GROUP_TYPE_UNIVERSAL_GROUP = 0x8
Global Const $ADS_GROUP_TYPE_SECURITY_ENABLED = 0x80000000
Global Const $ADS_GROUP_TYPE_GLOBAL_SECURITY = BitOR($ADS_GROUP_TYPE_GLOBAL_GROUP, $ADS_GROUP_TYPE_SECURITY_ENABLED)
Global Const $ADS_GROUP_TYPE_UNIVERSAL_SECURITY = BitOR($ADS_GROUP_TYPE_UNIVERSAL_GROUP, $ADS_GROUP_TYPE_SECURITY_ENABLED)
Global Const $ADS_GROUP_TYPE_DOMAIN_LOCAL_SECURITY = BitOR($ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP, $ADS_GROUP_TYPE_SECURITY_ENABLED)

Global Const $ADS_UF_PASSWD_NOTREQD = 0x0020
Global Const $ADS_UF_WORKSTATION_TRUST_ACCOUNT = 0x1000
Global Const $ADS_ACETYPE_ACCESS_ALLOWED = 0x0
Global Const $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = 0x5
Global Const $ADS_ACETYPE_ACCESS_DENIED_OBJECT = 0x6
Global Const $ADS_FLAG_OBJECT_TYPE_PRESENT = 0x1
Global Const $ADS_RIGHT_GENERIC_READ = 0x80000000
Global Const $ADS_RIGHT_DS_SELF = 0x8
Global Const $ADS_RIGHT_DS_WRITE_PROP = 0x20
Global Const $ADS_RIGHT_DS_CONTROL_ACCESS = 0x100
Global Const $ADS_UF_ACCOUNTDISABLE = 2
Global Const $ADS_OPTION_SECURITY_MASK = 3
Global Const $ADS_SECURITY_INFO_DACL = 4
Global Const $ADS_UF_DONT_EXPIRE_PASSWORD = 0x10000

Global Const $ALLOWED_TO_AUTHENTICATE = "{68B1D179-0D15-4d4f-AB71-46152E79A7BC}"
Global Const $RECEIVE_AS = "{AB721A56-1E2f-11D0-9819-00AA0040529B}"
Global Const $SEND_AS = "{AB721A54-1E2f-11D0-9819-00AA0040529B}"
Global Const $USER_CHANGE_PASSWORD = "{AB721A53-1E2f-11D0-9819-00AA0040529b}"
Global Const $USER_FORCE_CHANGE_PASSWORD = "{00299570-246D-11D0-A768-00AA006E0529}"
Global Const $USER_ACCOUNT_RESTRICTIONS = "{4C164200-20C0-11D0-A768-00AA006E0529}"
Global Const $VALIDATED_DNS_HOST_NAME = "{72E39547-7B18-11D1-ADEF-00C04FD8D5CD}"
Global Const $VALIDATED_SPN = "{F3A64788-5306-11D1-A9C5-0000F80367C1}"
Const $Member_SchemaIDGuid = "{BF9679C0-0DE6-11D0-A285-00AA003049E2}"
#EndRegion  ; Define AD Constants


Global $objConnection = ObjCreate("ADODB.Connection") ; Create COM object to AD
If @error Then
	MsgBox(16, "Error", "Could not create ADODB Connection")
	Exit
EndIf
$objConnection.ConnectionString = "Provider=ADsDSOObject"
$objConnection.Open("Active Directory Provider") ; Open connection to AD

Global $ADSystemInfo = ObjCreate("ADSystemInfo")

#Region  ; User Interface
$titel = "Userinfo v3.0"
If @AutoItX64 Then
	$titel &= " [64 Bit]"
Else
	$titel &= " [32 Bit]"
EndIf

If @Compiled And _Singleton("userinfo", 1) = 0 Then ; Only start one instance of the application
	WinActivate($titel)
	Exit
EndIf

;frmMain - Create main window
$labelWidth = 90 ; default label width
$labelHeight = 17 ; default label height
$valueWidth = 370 ; default value width
$valueHeight = 17 ; default value height
$cellPadding = 5 ; default distance between

$col1LabelLeft = 16 ; Start first label
$col1ValueLeft = $col1LabelLeft + $labelWidth + $cellPadding ; Start first value
$col2LabelLeft = $col1ValueLeft + $valueWidth + $cellPadding ; Start  second label
$col2ValueLeft = $col2LabelLeft + $labelWidth + $cellPadding; Start  second value
$breite1 = 370 ; Width  first value
$breite2 = 180 ; Width  second value
$breite3 = 60
$breite4 = 220

Global Const $buttonLeft = 825
Global Const $buttonWidth = 140
Global Const $buttonBigHeight = 40
Global Const $buttonNormalHeight = 20
Global Const $buttonPadding = 5

$frmMain = GUICreate($titel, 980, 800, -1, -1)
GUISetBkColor(0xeeeeee)
GUICtrlSetFont($frmMain, 6)

GUICtrlCreateGroup("Functions", 820, $buttonPadding, $buttonWidth + (2 * $buttonPadding), 770) ;Create Functions group

; for some reason, the "flashing" of labels (when clicked) only works when lables are created after creating the buttons...
#Region ; Create all buttons
$top = 20
$btnRefreshQuery = GUICtrlCreateButton("Refresh", $buttonLeft, $top, $buttonWidth, $buttonBigHeight)
GUICtrlSetImage(-1, "shell32.dll", 255)

$top += 40
$btnNewQuery = GUICtrlCreateButton("New Search", $buttonLeft, $top, $buttonWidth, $buttonBigHeight)
GUICtrlSetImage(-1, "shell32.dll", 23)

$top += 60
$btnCountDeptUsers = GUICtrlCreateButton("Count Users of Department", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)

$top += 20
$btnDepartment = GUICtrlCreateButton("Show Users of Department", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)

$top += 70
$btnCompareGroups = GUICtrlCreateButton("Compare Groups", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)

$top += 20
$btnGroups = GUICtrlCreateButton("Copy Group List", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)

$top += 20
$btnGroupMembers = GUICtrlCreateButton("Group-Info", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)

$top += 60
$btnExit = GUICtrlCreateButton("Exit", $buttonLeft, $top, $buttonWidth, $buttonBigHeight)
GUICtrlSetImage(-1, "shell32.dll", 28)

GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group
#EndRegion ; Create all buttons

#Region ; Create all info lables
$top = 10
GUICtrlCreateLabel("Name:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblName = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("UserID:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblUserId = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Title:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_anrede = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("OU:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_ou = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Telephone:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_telefon = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Department:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblDepartment = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Domain:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblDomain = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Email:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_email = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Location:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_ort = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Address:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_adresse = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Room:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_raum = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$breite4 = 330
$top += 20
GUICtrlCreateLabel("Last update:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblLastUpdate = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Valid till:", $col1LabelLeft + 220, $top, $labelWidth / 2, $labelHeight)
$lbl_expiration = GUICtrlCreateLabel("", $col1LabelLeft + 280, $top, 130, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("PW Required:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblPwRequired = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("PW Change:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_change = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Locked/Deact:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblLocked = GUICtrlCreateLabel("", $col1ValueLeft, $top, 30, $labelHeight)
GUICtrlSetFont(-1, -1, 800)
GUICtrlCreateLabel("/", $col1ValueLeft + 30, $top, 5, $labelHeight)
$lblDeactivated = GUICtrlCreateLabel("-", $col1ValueLeft + 40, $top, 50, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$bn_unlock = GUICtrlCreateButton("Unlock", 200, $top - 4, $labelWidth, 20)

GUICtrlCreateLabel("Created on:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_erzeugt = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Failed count:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_errorcount = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth / 4, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Last logon:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblLastLogon = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Homedirectory:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_homeverz = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Homedrive:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_homelauf = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Loginscript:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblLogonScript = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Profile Path:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lbl_profil = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidth, $labelHeight)
GUICtrlSetFont(-1, -1, 800)
#EndRegion ; Create all info lables

$top += 25
$listGroups = GUICtrlCreateListView("Name of Group|samAccountName|Group type|Member of|Scope", 8, $top, 800, 770 - $top)
_GUICtrlListView_SetColumnWidth($listGroups, 0, 240)
_GUICtrlListView_SetColumnWidth($listGroups, 1, 200)

$sbMain = _GUICtrlStatusBar_Create($frmMain)


; frmSearchAD - Create window for entering the user name and domain
$frmSearchAD = GUICreate($titel, 370, 140, -1, -1)
GUICtrlCreateLabel("UserID/Name/Tel.:", 5, 8, 100, 17)
$edit = GUICtrlCreateInput(@UserName, 110, 5, 250, 20)
GUICtrlCreateLabel("Wildcards:", 5, 35, 100, 17)
$radio1 = GUICtrlCreateRadio("UserID", 110, 35, 70, 20)
$radio2 = GUICtrlCreateRadio("Surname", 190, 35, 80, 20)
$radio3 = GUICtrlCreateRadio("Christian name", 270, 35, 120, 20)
GUICtrlSetState($radio2, $GUI_CHECKED)
$Enter_key = GUICtrlCreateDummy()
Dim $a_AccelKeys[1][2] = [["{ENTER}", $Enter_key]] ; Hotkey array for evaluating the Enter button on frmSearchAD
GUISetAccelerators($a_AccelKeys, $frmSearchAD)


GUICtrlCreateLabel("Domain:", 5, 65, 100, 17)
$cmbDomain = GUICtrlCreateCombo("", 110, 62, 250, 250, $CBS_DROPDOWNLIST)
$btnOk = GUICtrlCreateButton("&Ok", 110, 90, 120, 25)
$btnCancel = GUICtrlCreateButton("E&xit", 240, 90, 120, 25)
$sbSearchAD = _GUICtrlStatusBar_Create($frmSearchAD)

#EndRegion  ; User Interface

Global $oMyError = ""
$oMyError = ObjEvent("AutoIt.Error", "_ADDoError") ; Install a custom error handler

Dim $groups
Global $lw_m = ""
Global $lw_y = ""
Global $pc = ""
Global $user = ""
Global $objRootDSE
Global $strDNSDomain
Global $strHostServer
Global $strConfiguration
Global $intUAC
Global $oUsr
Global $objRecordSet
Global $time ; Zeitstempel f�rs Stoppen des Auslesens der Userdaten
Global $pc_status = 0 ; PC online?
Global $cancel ; Flag zum Beenden der Dom�nensuche
Global $loopflag ; dient zur Kontrolle der Schleife bei der Userauswahl
Global $MaxPasswordAge ; wie lang ein PW in der Dom�ne g�ltig ist

; ==================================================================================================================
;        Main program
; ==================================================================================================================
listDomains()
setUser()

; ==================================================================================================================
;        Functions
; ==================================================================================================================

Func listDomains() ; discover all Domains in forest and select ours
	$oDS = ObjGet("WinNT:")
	$y = ""
	For $x In $oDS
		$y = $y & "|" & $x.Name
	Next
	GUICtrlSetData($cmbDomain, $y, "NL")
EndFunc   ;==>listDomains

Func setUser($userid_tmp = "")
	ControlFocus($titel, "", $btnNewQuery)
	_GUICtrlStatusBar_SetText($sbSearchAD, "Ready")
	$pc = ""
	$lw_m = ""
	GUISetState(@SW_HIDE, $frmMain)
	GUISetState(@SW_SHOW, $frmSearchAD)

	$loopflag = 0
	If $userid_tmp <> "" Then
		GUICtrlSetData($edit, $userid_tmp)
		$loopflag = domainSearch()
	EndIf

	If $loopflag = 0 Then
		ControlClick($titel, "", $edit)
		Send("+{home}")
		$loopflag = 1
		While $loopflag
			$msg = GUIGetMsg()
			Switch $msg
				Case $Enter_key
					domainSearch()
				Case $GUI_EVENT_CLOSE
					Exit
				Case $btnOk
					domainSearch()
				Case $btnCancel
					Exit
			EndSwitch
		WEnd
	EndIf
	GUISetState(@SW_HIDE, $frmSearchAD)

	; if we get here, useraccount has been found. -> get our values
	$ldap_entry = $objRecordSet.fields(0).value

	$oUsr = ObjGet($ldap_entry) ; Retrieve the COM Object for the logged on user
	$user = $oUsr.samAccountName ; get SamAccountname

	GUISetState(@SW_SHOW, $frmMain)
	GUISetCursor(15, 1, $frmMain)
	_GUICtrlStatusBar_SetText($sbSearchAD, "Reading user data...")

	queryAD()
	GUISetCursor(2)
	GUILoop()
EndFunc   ;==>setUser

Func GUILoop()
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $btnExit
				Exit
			Case $btnGroups
				For $i = 0 To UBound($groups) - 1
					$tmp = StringSplit($groups[$i], ",")
					$tmp[1] = StringReplace($tmp[1], "CN=", "")
					$groups[$i] = $tmp[1]
				Next
				_ArrayToClip($groups, 1)
			Case $btnGroupMembers
				$x = GUICtrlRead($listGroups)
				If $x <> "" Then ; nur wenn auch ein User gew�hlt wurde...
					$y = GUICtrlRead($x) ; das ListviewItem auslesen
					$tmp = StringSplit($y, "|") ; aufteilen und nur die UserID behalten
					_ADGetGroupMembers($tmp[2])
				EndIf
			Case $bn_unlock
				unlock($user)
			Case $btnCompareGroups
				compareGroup()
			Case $btnRefreshQuery
				$time = TimerInit() ; reset timer
				queryAD()
			Case $btnNewQuery
				setUser()
			Case $edit ; otherwise the text from the search edit field is copied to the clipboard when a hit on Wildcard search was carried out
				;
			Case $btnDepartment
				searchDepartment()
			Case $btnCountDeptUsers
				countUsers()
			Case -100 To 0
				;
			Case Else
				$tmp = GUICtrlRead($nMsg)
				If $tmp <> "" Then ; only if a text has been detected
					While StringLeft($tmp, 1) = "|" ; strip leading hyphens
						$tmp = StringMid($tmp, 2)
					WEnd
					$tmp = StringReplace($tmp, "|", @CRLF)
					$tmp = StringReplace($tmp, @CRLF & @CRLF, "")
					ClipPut($tmp) ; copy to clipboard
					_GUICtrlStatusBar_SetText($sbSearchAD, $tmp & " has been copied to the clipboard") ; update status bar
					$x = ControlGetFocus($titel) ; find out what was clicked
					If Not StringInStr($x, "SysListView32") Then ; If not something was clicked in the group list
						GUICtrlSetBkColor($nMsg, 0xff0000) ; color the entry red
						Sleep(250) ; wait a quarter of a second
						GUICtrlSetBkColor($nMsg, -1) ; Revert color of text
					EndIf
				EndIf
		EndSwitch
		Sleep(50)
	WEnd
EndFunc   ;==>GUILoop

Func domainSearch($user = "", $domain = "")
	If $user = "" Then $user = GUICtrlRead($edit)
	If $domain = "" Then $domain = GUICtrlRead($cmbDomain)

	If $user <> "" Then
		$time = TimerInit()
		$user = StringStripWS($user, 3)

		$objRootDSE = ObjGet("LDAP://" & $domain & "/RootDSE")
		If @error Then
			_GUICtrlStatusBar_SetText($sbSearchAD, "Domain " & $domain & " could not be reached.")
			$loopflag = 1
			Return 0
		EndIf
		$strDNSDomain = $objRootDSE.Get("defaultNamingContext") ; Retrieve the current AD domain name
		$strHostServer = $objRootDSE.Get("dnsHostName") ; Retrieve the name of the connected DC
		$strConfiguration = $objRootDSE.Get("ConfigurationNamingContext") ; Retrieve the Configuration naming context

		Switch _ADObjectExists($user)
			Case 1 ; User is found
				$loopflag = 0
				$x = ObjGet("WinNT://" & $domain)
				$MaxPasswordAge = $x.MaxPasswordAge
				Return 1
			Case -1 ; Query has multiple results
				ControlClick($titel, "", $btnOk)
				Return 0 ; und brav wieder das Fensterchen anzeigen, nun aber mit der eindeutigen Userkennung
			Case 0 ; User not found
				_GUICtrlStatusBar_SetText($sbSearchAD, "The account " & $user & " does not exist in this domain.")
				$loopflag = 1
				Return 0
		EndSwitch
	EndIf
EndFunc   ;==>domainSearch

Func queryAD()
	$pc = ""
	$lastlogon = "-"
	GUICtrlSetState($bn_unlock, $GUI_DISABLE)
	GUICtrlSetData($lbl_homelauf, "")
	GUICtrlSetData($lbl_homeverz, "")
	GUICtrlSetData($lblLogonScript, "")
	_GUICtrlListView_DeleteAllItems($listGroups)

	_ADGetUserData($user)
	_ADGetUserGroups($groups, $user)
	_ADGetAccount($user)

	$tmp = $oUsr.scriptpath

	If $lastlogon = "-" Then ; Dom�nen LastLogin auswerten
		$tmp = $oUsr.LastLogin
		If $tmp <> "" Then
			$lastlogon = Zeit($oUsr.LastLogin)
		Else
			$lastlogon = "Never"
		EndIf
	EndIf
	GUICtrlSetData($lblLastLogon, $lastlogon)


		 	$tmp = TimerDiff($time) / 1000
		 	$tmp1 = StringFormat("%.2f", $tmp)
	#CS
		; 	;Empirum() ; Daten aus Empirum auslesen
		; 	;NC_Daten()
		;
		;
	#CE
	$tmp = TimerDiff($time) / 1000
	$tmp2 = StringFormat("%.2f", $tmp)

	_GUICtrlStatusBar_SetText($sbSearchAD, "Userdata read (in " & $tmp2 & " seconds, with " & $tmp1 & " from ADS), Domaincontroller: " & $strHostServer)
EndFunc   ;==>queryAD


Func _ADDoError()
	$HexNumber = Hex($oMyError.number, 8)

	If $HexNumber = 80020009 Then
		SetError(3)
		Return
	EndIf
	If $HexNumber = "8007203A" Then
		SetError(4)
		Return
	EndIf


	MsgBox(262144, "", "We intercepted a COM Error !" & @CRLF & _
			"Number is: " & @TAB & $HexNumber & @CRLF & _
			"Windescription is: " & @TAB & $oMyError.windescription & @CRLF & _
			"err.description is: " & @TAB & $oMyError.description & @CRLF & _
			"err.source is: " & @TAB & $oMyError.source & @CRLF & _
			"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
			"err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF & _
			"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
			"err.retcode is: " & @TAB & $oMyError.retcode & @CRLF & _
			"Script Line number is: " & @TAB & $oMyError.scriptline)
	Select
		Case $oMyError.windescription = "Access is denied"
			$objConnection.Close("Active Directory Provider")
			$objConnection.Open("Active Directory Provider")
			SetError(2)
		Case 1
			SetError(1)
	EndSelect
EndFunc   ;==>_ADDoError


Func _ADObjectExists($object)
	$flag = 0 ; f�r Auswertung der Radiobuttons
	$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sAMAccountName=" & $object & "));ADsPath;subtree"
	$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	Switch $objRecordSet.RecordCount
		Case 0; User wurde nicht UserID gefunden
			; nix tun, weitere Suchen ablaufen lassen
		Case 1 ; User wurde eindeutig identifiziert anhand des Namens
			GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
			Return 1
		Case Else ; mehrere User wurden �ber UserID gefunden
			If GUICtrlRead($radio1) = $GUI_CHECKED Then $flag = 1 ; Ergebnis nur behalten wenn suche nach UserID aktiviert
	EndSwitch

	If $flag = 0 Then ; wenn bei der UserID Suche nichts gefunden wurde bzw. Suche nach Name/Vorname gew�nscht
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sn=" & $object & "));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
		Switch $objRecordSet.RecordCount
			Case 0; User wurde nicht per Nachname gefunden
				; nix tun, weiter pr�fen
			Case 1 ; User wurde eindeutig identifiziert anhand des Namens
				GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
				Return 1
			Case Else ; mehrere User wurden �ber UserID gefunden -> Userliste anzeigen lassen
				If GUICtrlRead($radio2) = $GUI_CHECKED Then $flag = 1 ; Ergebnis nur behalten wenn suche nach Nachname aktiviert
		EndSwitch
	EndIf

	If ($flag = 0) And (GUICtrlRead($radio3) = $GUI_CHECKED) Then ; Wenn immer noch nichts gefunden, suche nach dem Vornamen
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(givenName=" & $object & "));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery)
	EndIf

	If $objRecordSet.RecordCount = 0 Then ; So far nothing has been found purely -> So we are looking for a phone number
		If StringIsDigit($object) Then $object = "*" & $object
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(telephoneNumber=" & $object & "));ADsPath;subtree" ; also suche nach Telefonnummer
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	EndIf

	#CS 	If $objRecordSet.RecordCount = 0 Then ; Bisher wurde rein gar nichts gefunden -> also suchen wir nach einer Postkorbnummer
		; 		$tmp2 = suche_postkorb($object)
		; 		If $tmp2 <> "" Then
		; 			GUICtrlSetData($edit, $tmp2) ; die Userkennung aus dem Ergebnis der Postkorbsuche wird im Suchfenster eingetragen
		; 			Return -1 ; Zur�ck zum Aufrufer mit Info, dass neue UserID in Suchmaske eingetragen wurde
		; 		EndIf
		; 	EndIf
	#CE


	Switch $objRecordSet.RecordCount
		Case 0 ; dieser User wurde nicht gefunden
			Return 0
		Case 1 ; User wurde eindeutig identifiziert anhand des Namens
			GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
			Return 1
		Case Else ; User existiert, aber mehrere Treffer (Suchbegriff: Meier ^^)
			Dim $treffer_arry[1]
			$z = ""
			Do
				$y = $objRecordSet.Fields(0).Value ; FQDN-Name des Users
				If Not StringInStr($y, "ou=Empf�nger") Then ; skip all mail-only accounts (ou=benutzer,ou=empf�nger)
					$oUsr = ObjGet($objRecordSet.Fields(0).Value) ; Retrieve the COM Object for the logged on user
					_ArrayAdd($treffer_arry, $oUsr.sn & "," & $oUsr.givenName & "|" & $oUsr.samAccountName)
				EndIf
				$objRecordSet.MoveNext
			Until $objRecordSet.EOF
			If UBound($treffer_arry) = 2 Then ; Array hat nur 1 Element => die anderen Treffer waren reine Mailempf�nger und wurden beim �bertragen ins Array �bergangen
				$x = StringInStr($treffer_arry[1], "|")
				GUICtrlSetData($edit, StringMid($treffer_arry[1], $x + 1)) ; die Userkennung wird im Suchfenster eingetragen
				GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
				Return -1
			Else ; aha, es gibt tats�chlich mehrere User die in Frage kommen
				GUISetState(@SW_HIDE, $frmSearchAD)
				$Form3 = GUICreate($texts[$language][69], 350, 355)
				$Enter_key2 = GUICtrlCreateDummy()
				Dim $b_AccelKeys[1][2] = [["{ENTER}", $Enter_key2]] ; Hotkey-Array f�r das Auswerten der Enter-Taste in Form3
				GUISetAccelerators($b_AccelKeys, $Form3)
				$liste = GUICtrlCreateListView($texts[$language][70], 5, 40, 340, 280)
				_GUICtrlListView_SetColumnWidth($liste, 0, 250)
				$btn_userwahl = GUICtrlCreateButton("Ok", 5, 325, 340, 25)
				GUICtrlCreateLabel(UBound($treffer_arry) - 1 & " users found. Please select one:", 5, 5, 290, 30)
				_ArrayDelete($treffer_arry, 0) ; das leere erste Feld l�schen
				_ArraySort($treffer_arry) ; Treffer sortieren

				For $i = 0 To UBound($treffer_arry) - 1 ; und alle Ergebnisse in Listview kopieren
					GUICtrlCreateListViewItem($treffer_arry[$i], $liste)
				Next
				GUISetState(@SW_SHOW, $Form3)

				While 1
					$msg = GUIGetMsg()
					If $msg = $GUI_EVENT_CLOSE Then Exit
					If ($msg = $btn_userwahl) Or ($msg = $Enter_key2) Then
						$x = GUICtrlRead($liste)
						If $x <> "" Then ; nur wenn auch ein User gew�hlt wurde...
							$y = GUICtrlRead($liste) ; die Control-ID des markierten Listview-Items holen
							$y = GUICtrlRead($y) ; das ListviewItem auslesen
							$tmp = StringSplit($y, "|") ; aufteilen und nur die UserID behalten
							$x = GUICtrlSetData($edit, $tmp[2]) ; die Userkennung wird im Suchfenster eingetragen
							ExitLoop
						EndIf
					EndIf
				WEnd
				GUIDelete($Form3) ; Clear selection GUI
				GUISetState(@SW_SHOW, $frmSearchAD) ; Display window for user selection again
				Return -1 ; User was added in form
			EndIf
	EndSwitch
EndFunc   ;==>_ADObjectExists

Func _ADGetUserData($user)
	$oUsr.GetInfo() ; ADS-Cache der Attribute neu auslesen sonst kriegt das Programm evtl. keine Gruppen�nderungen mit
	$tmp = $oUsr.DisplayName
	If $tmp <> "" Then
		GUICtrlSetData($lblName, $oUsr.DisplayName)
	Else
		GUICtrlSetData($lblName, $oUsr.sn & ", " & $oUsr.givenName)
	EndIf

	GUICtrlSetData($lbl_anrede, $oUsr.title)

	; evaluation of the "Released on" value. Case: If a "" in the name of the user himself is in it, he is as \ masked.
	; That bothers StringSplit but unfortunately not. Therefore we convert the thing before StringSplit to and then back.
	$tmp = $oUsr.distinguishedName
	$tmp = StringReplace($tmp, "\,", "��")
	$tmp = StringSplit($tmp, ",")

	$tmp2 = ""
	For $i = 1 To $tmp[0] ; Abtrennen der DC= Klamotten
		$tmp[$i] = StringReplace($tmp[$i], "��", ",")
		$tmp[$i] = StringMid($tmp[$i], 4)
	Next ; die hinteren 3 mit einem . trennen
	$max = $tmp[0]
	$tmp2 = $ADSystemInfo.DomainShortName & " (" & $tmp[$max - 2] & "." & $tmp[$max - 1] & "." & $tmp[$max] & ")"

	GUICtrlSetData($lblUserId, $user)
	GUICtrlSetData($lblDomain, $tmp2)

	$tmp2 = $tmp2 & "/"
	For $i = $tmp[0] - 3 To 2 Step -1
		$tmp2 = $tmp2 & $tmp[$i] & "/"
	Next
	$tmp2 = StringReplace($tmp2, "/\", "/")
	$tmp2 = StringLeft($tmp2, StringLen($tmp2) - 1)

	GUICtrlSetData($lbl_ou, $tmp2)

	GUICtrlSetData($lbl_email, $oUsr.mail)
	GUICtrlSetData($lbl_telefon, $oUsr.telephoneNumber)
	GUICtrlSetData($lblDepartment, StringReplace($oUsr.department, "&", "&&"))
	GUICtrlSetData($lbl_profil, $oUsr.profilePath)

	$lw_m = $oUsr.homeDirectory
	GUICtrlSetData($lbl_homelauf, $oUsr.homeDrive)
	GUICtrlSetColor($lbl_homeverz, -1)
	GUICtrlSetColor($lbl_homelauf, -1)
	If $lw_m <> "" Then
		GUICtrlSetData($lbl_homeverz, $lw_m)
	EndIf

	GUICtrlSetData($lbl_ort, StringStripWS($oUsr.postalCode & " " & $oUsr.l, 3))
	GUICtrlSetData($lbl_raum, $oUsr.physicalDeliveryOfficeName)
	GUICtrlSetData($lbl_adresse, $oUsr.streetAddress)
EndFunc   ;==>_ADGetUserData

Func _ADGetUserGroups(ByRef $usergroups, $user = @UserName)
	$usergroups = $oUsr.GetEx("memberof")
	$count = UBound($usergroups)
	If $count = 0 Then
		GUICtrlSetData($btnGroups, $texts[$language][12]) ; reset label of button
		Return ; catch empty list
	EndIf
	_ArrayInsert($usergroups, 0, $count)
	_ArraySort($usergroups, 0, 1)
	For $i = 1 To $count
		$tmp = $usergroups[$i]
		$tmp = StringReplace($tmp, "/", "\/")
		$objGroup = ObjGet("LDAP://" & $tmp)
		If BitAND($objGroup.groupType, $ADS_GROUP_TYPE_SECURITY_ENABLED) Then
			$tmp2 = "Security"
		Else
			$tmp2 = "Distribution"
		EndIf
		$tmp3 = $objGroup.memberOf
		$tmp3 = StringMid($tmp3, 4)
		$tmp3 = StringSplit($tmp3, ",")

		Switch $objGroup.InstanceType
			Case 2
				$tmp4 = "Global"
			Case 4
				$tmp4 = "Local"
			Case 8
				$tmp4 = "Universal"
		EndSwitch
		$tmp = StringReplace($objGroup.name, "\/", "/") ; Maskierung wieder r�ckg�ngig machen f�r gescheite Anzeige
		$tmp = StringReplace($tmp, "CN=", "") & "|" & $objGroup.samAccountName & "|" & $tmp2 & "|" & $tmp3[1] & "|" & $tmp4
		GUICtrlCreateListViewItem($tmp, $listGroups)
	Next
	GUICtrlSetData($btnGroups, $texts[$language][12] & " (" & $count & ")")
EndFunc   ;==>_ADGetUserGroups

Func _ADGetAccount($user)
	GUICtrlSetData($lblLastUpdate, Zeit($oUsr.PasswordLastChanged))
	If $oUsr.IsAccountLocked Then
		GUICtrlSetData($lblLocked, $texts[$language][74])
		GUICtrlSetColor($lblLocked, 0xff0000)
		GUICtrlSetState($bn_unlock, $GUI_ENABLE)
	Else
		GUICtrlSetData($lblLocked, $texts[$language][75])
		GUICtrlSetColor($lblLocked, 0x00aa00)
		GUICtrlSetState($bn_unlock, $GUI_DISABLE)
	EndIf

	GUICtrlSetData($lbl_errorcount, $oUsr.BadLoginCount)

	$intUAC = $oUsr.Get("userAccountControl")

	; Pr�fen, ob der User sein Passwort �ndern darf. Da keine direkte Abfrage m�glich ist, erfolgt dies �ber den Security Descriptor
	$oSecDesc = $oUsr.Get("ntSecurityDescriptor")
	$oACL = $oSecDesc.DiscretionaryACL
	For $oACE In $oACL
		If ($oACE.ObjectType = $USER_CHANGE_PASSWORD) And (($oACE.Trustee = "Everyone") Or ($oACE.Trustee = "Jeder")) Then
			If ($oACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT) Then
				GUICtrlSetData($lbl_change, $texts[$language][76])
				GUICtrlSetColor($lbl_change, 0x00aa00)
			Else
				GUICtrlSetData($lbl_change, $texts[$language][77])
				GUICtrlSetColor($lbl_change, 0xff0000)
			EndIf
		EndIf
	Next

	If BitAND(0x00020, $intUAC) Then ; PW not required flag ist gesetzt
		GUICtrlSetData($lblPwRequired, $texts[$language][75])
	Else ; nicht gesetzt -> Passwort n�tig
		GUICtrlSetData($lblPwRequired, $texts[$language][74])
	EndIf

	GUICtrlSetData($lbl_erzeugt, Zeit($oUsr.whenCreated))
	GUICtrlSetData($lblLogonScript, $oUsr.scriptPath)

	If BitAND($intUAC, $ADS_UF_ACCOUNTDISABLE) Then ; Disabled flag ist gesetzt
		GUICtrlSetData($lblDeactivated, $texts[$language][74])
		GUICtrlSetColor($lblDeactivated, 0xff0000)
	Else
		GUICtrlSetData($lblDeactivated, $texts[$language][75])
		GUICtrlSetColor($lblDeactivated, 0x00aa00)
	EndIf


	GUICtrlSetColor($lbl_expiration, 0x000000)
	$dummy = $oUsr.AccountExpirationDate
	$tmp = Zeit2($dummy) ; Umgewandelte Zeit f�r DateDiff
	$dummy = Zeit($dummy) ; Umgewandelte Zeit f�r die Ausgabe

	$tmp2 = _DateDiff("D", _NowCalcDate(), $tmp)
	If ($tmp2 < 1) And ($tmp2 > -148883) Then GUICtrlSetColor($lbl_expiration, 0xff0000) ; Passwort ist keinen Tag mehr g�ltig

	If ($dummy = "01.01.1601 02:00") Or ($dummy = "01.01.1970 00:00") Or ($dummy = "01.01.1601 01:00") Then
		$dummy = $texts[$language][78]
		GUICtrlSetColor($lbl_expiration, 0x000000)
	EndIf
	GUICtrlSetData($lbl_expiration, $dummy)

EndFunc   ;==>_ADGetAccount

Func unlock($user)
	If $oUsr.IsAccountLocked Then
		$oUsr.IsAccountLocked = False
		$oUsr.SetInfo
		Sleep(500)
		$oUsr.GetInfo() ; Refresh ADS-Cache
		If Not $oUsr.IsAccountLocked Then
			GUICtrlSetData($lblLocked, "No")
			GUICtrlSetColor($lblLocked, 0x00aa00);Green
			GUICtrlSetState($bn_unlock, $GUI_DISABLE)
		Else
			GUICtrlSetData($lblLocked, "Yes")
			GUICtrlSetColor($lblLocked, 0xff0000) ; Red
			GUICtrlSetState($bn_unlock, $GUI_ENABLE)
		EndIf
	EndIf
EndFunc   ;==>unlock

Func passwort_req($user)
	$oUsr.Put("userAccountControl", BitXOR($intUAC, 0x00020))
	$oUsr.SetInfo

	$oUsr.GetInfo
	$intUAC = $oUsr.Get("userAccountControl")
	If BitAND(0x00020, $intUAC) Then ; PW not required flag is set
		GUICtrlSetData($lblPwRequired, "No")
		GUICtrlSetColor($lblPwRequired, 0xff0000)
	Else ; not set -> Password required
		GUICtrlSetData($lblPwRequired, "Yes")
		GUICtrlSetColor($lblPwRequired, 0x00aa00)
	EndIf
EndFunc   ;==>passwort_req

Func Zeit($zeit) ; Umwandeln der AD-Zeit ins deutsche Format
	If $zeit = "" Then Return "---" ; User wurde grade frisch zur�ckgesetzt
	$tmp = StringMid($zeit, 7, 2) & "." & StringMid($zeit, 5, 2) & "." & StringLeft($zeit, 4) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2)
	Return $tmp
EndFunc   ;==>Zeit

Func Zeit2($zeit) ; Umwandeln der AD-Zeit ins englische Format f�r _DateAdd
	$tmp = StringLeft($zeit, 4) & "/" & StringMid($zeit, 5, 2) & "/" & StringMid($zeit, 7, 2) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2) & ":" & StringMid($zeit, 13, 2)
	Return $tmp
EndFunc   ;==>Zeit2

Func Zeit3($zeit) ; Umwandeln vom englischen Datumsformat YYYYMMDD ins deutsche
	If $zeit = "0" Then Return "---" ; User wurde grade frisch zur�ckgesetzt
	$tmp = StringSplit($zeit, "/")
	$tmp2 = StringLeft($tmp[3], 2) & "." & $tmp[2] & "." & $tmp[1] & " " & StringMid($tmp[3], 4, 5)
	Return ($tmp2)
EndFunc   ;==>Zeit3

Func Zeit4($zeit) ; umwandeln von MM/DD/YYYY ins deutsche Format z.B.: 9/5/2008 4:57:39 PM
	$tmp = StringSplit($zeit, "/")
	If Not IsArray($tmp) Or ($tmp[0] < 3) Then Return "-"
	If StringLen($tmp[1]) = 1 Then $tmp[1] = "0" & $tmp[1] ; notfalls den Monat zweistellig machen
	If StringLen($tmp[2]) = 1 Then $tmp[2] = "0" & $tmp[2] ; notfalls den Tag zweistellig machen
	$tmp3 = StringLeft($tmp[3], 4) ; das Jahr
	$tmp4 = StringStripWS(StringMid($tmp[3], 5), 3) ; die Uhrzeit
	If StringRight($tmp4, 2) = "PM" Then ; wenn Nachmittags: 12 Stunden draufpacken
		$doppelpunkt = StringInStr($tmp4, ":")
		$tmp4 = StringLeft($tmp4, $doppelpunkt - 1) + 12 & StringMid($tmp4, $doppelpunkt, 6)
	Else ; wenn vormittags: nur das " AM" entfernen
		$tmp4 = StringLeft($tmp4, StringLen($tmp4) - 3)
	EndIf

	If StringInStr($tmp4, ":") = 2 Then $tmp4 = "0" & $tmp4

	$tmp2 = $tmp[2] & "." & $tmp[1] & "." & $tmp3 & " " & StringLeft($tmp4, 5) ; Datum zusammensetzen
	Return $tmp2
EndFunc   ;==>Zeit4

Func Deaktivieren($user)
	$intUAC = $oUsr.Get("userAccountControl")
	If BitAND($intUAC, $ADS_UF_ACCOUNTDISABLE) Then ; Disabled flag ist gesetzt
		GUICtrlSetData($lblDeactivated, $texts[$language][75])
		GUICtrlSetColor($lblDeactivated, 0x00aa00)
	Else
		GUICtrlSetData($lblDeactivated, $texts[$language][74])
		GUICtrlSetColor($lblDeactivated, 0xff0000)
	EndIf
	$oUsr.Put("userAccountControl", BitXOR($intUAC, $ADS_UF_ACCOUNTDISABLE)) ; Flag �ndern
	$oUsr.SetInfo
EndFunc   ;==>Deaktivieren

Func _ADGetGroupMembers($gruppe)
	If $gruppe = "" Then Return

	$groupdn = _ADSamAccountNameToFQDN($gruppe)
	If $groupdn = "" Then
		MsgBox(16, "Problem", $texts[$language][79])
		Return
	EndIf
	$objGroup = ObjGet("LDAP://" & $groupdn)

	Dim $userarry[50000][2]
	$i = 0
	For $objmember In $objGroup.members
		$userarry[$i][0] = StringReplace($objmember.displayname, " (LBV)", "") ; remove useless additional string from our Actice Directory
		$userarry[$i][1] = $objmember.samAccountName
		$i = $i + 1
	Next
	ReDim $userarry[$i][2]

	_ArraySort($userarry) ; sortiere nach Displayname
	$counter = 0
	$result1 = ""
	For $strMember = 0 To UBound($userarry) - 1
		$result1 = $result1 & "<td bgcolor=66ff66>" & $userarry[$counter][0] & "<br>(" & $userarry[$counter][1] & ")</td>"
		$counter = $counter + 1
		If Mod($counter, 5) = 0 Then $result1 = $result1 & "</tr><tr><th>&nbsp;</th>"
	Next
	$result1 = StringReplace($result1, "\", "")

	_ArraySort($userarry, 0, 0, 0, 1) ; sortiere nach UserID
	$counter = 0
	$result2 = ""
	For $strMember = 0 To UBound($userarry) - 1
		$tmp = StringSplit($strMember, ",")
		$result2 = $result2 & "<td bgcolor=66ff66>" & $userarry[$counter][1] & "<br>(" & $userarry[$counter][0] & ")</td>"
		$counter = $counter + 1
		If Mod($counter, 5) = 0 Then $result2 = $result2 & "</tr><tr><th>&nbsp;</th>"
	Next
	$result2 = StringReplace($result2, "\", "")

	$objGroup.GetInfo

	$intGroupType = $objGroup.GroupType

	If $intGroupType And $ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP Then
		$x = "Domain local"
	ElseIf $intGroupType And $ADS_GROUP_TYPE_GLOBAL_GROUP Then
		$x = "Global"
	ElseIf $intGroupType And $ADS_GROUP_TYPE_UNIVERSAL_GROUP Then
		$x = "Universal"
	Else
		$x = "Unknown"
	EndIf

	If $intGroupType And $ADS_GROUP_TYPE_SECURITY_ENABLED Then
		$z = "Security group"
	Else
		$z = "Distribution group"
	EndIf

	$erzeugt = $objGroup.whenCreated
	$erzeugt = StringLeft($erzeugt, 4) & "-" & StringMid($erzeugt, 5, 2) & "-" & StringMid($erzeugt, 7, 2) & " " & StringMid($erzeugt, 9, 2) & ":" & StringMid($erzeugt, 11, 2) & ":" & StringMid($erzeugt, 13, 2)

	$geaendert = $objGroup.whenChanged
	$geaendert = StringLeft($geaendert, 4) & "-" & StringMid($geaendert, 5, 2) & "-" & StringMid($geaendert, 7, 2) & " " & StringMid($geaendert, 9, 2) & ":" & StringMid($geaendert, 11, 2) & ":" & StringMid($geaendert, 13, 2)

	$datei = FileOpen(@TempDir & "\userinfo.htm", 2)
	FileWriteLine($datei, "<table>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][80] & ":</th><td bgcolor=66ff66 colspan=10>" & $gruppe & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][81] & ":</th><td bgcolor=66ff66 colspan=10>" & $erzeugt & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][82] & ":</th><td bgcolor=66ff66 colspan=10>" & $geaendert & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][83] & ":</th><td bgcolor=66ff66 colspan=10>" & $objGroup.mail & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][84] & ":</th><td bgcolor=66ff66 colspan=10>" & $objGroup.info & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][85] & ":</th><td bgcolor=66ff66 colspan=10>" & $objGroup.description & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][86] & ":</th><td bgcolor=66ff66 colspan=10>" & $x & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][87] & ":</th><td bgcolor=66ff66 colspan=10>" & $z & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][88] & ":</th><td bgcolor=66ff66 colspan=10>" & UBound($userarry) - 1 & "</td></tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][89] & ":</th>" & $result1 & "</tr>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][90] & ":</th>" & $result2 & "</tr>")
	FileWriteLine($datei, "</table>")
	FileClose($datei)
	_IECreate(@TempDir & "\userinfo.htm")

	_GUICtrlStatusBar_SetText($sbSearchAD, $texts[$language][91])
	Return
EndFunc   ;==>_ADGetGroupMembers

Func _ADSamAccountNameToFQDN($samname)
	$strQuery2 = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(sAMAccountName=" & $samname & ");distinguishedName;subtree"
	$objRecordSet2 = $objConnection.Execute($strQuery2)

	If $objRecordSet2.RecordCount = 1 Then
		$fqdn = $objRecordSet2.fields(0).value
		$objRecordSet2 = 0
		Return StringReplace($fqdn, "/", "\/")
	Else
		$objRecordSet2 = 0
		Return ""
	EndIf
EndFunc   ;==>_ADSamAccountNameToFQDN

Func compareGroup($compareUser = "")
	Local $strQuery2, $objRecordSet2, $ldap_entry2
	Dim $vergleichsgruppen

	If $compareUser = "" Then
		$compareUser = InputBox($texts[$language][92], $texts[$language][93], "", "", 300, 120)
		If @error Then Return
	EndIf

	$strQuery2 = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(sAMAccountName=" & $compareUser & ");ADsPath;subtree"
	$objRecordSet2 = $objConnection.Execute($strQuery2) ; Retrieve the FQDN for the logged on user
	$ldap_entry2 = $objRecordSet2.fields(0).value
	$oUsr2 = ObjGet($ldap_entry2) ; Retrieve the COM Object for the logged on user
	$vergleichsgruppen = $oUsr2.GetEx("memberof")
	$oUsr2 = 0
	$count = UBound($vergleichsgruppen)
	If $count = 0 Then Return ; fange leere Gruppenlisten ab
	_ArrayInsert($vergleichsgruppen, 0, $count)
	For $i = 0 To $count
		$tmp = StringSplit($vergleichsgruppen[$i], ",")
		$tmp[1] = StringReplace($tmp[1], "CN=", "")
		$vergleichsgruppen[$i] = $tmp[1]
	Next
	_ArraySort($vergleichsgruppen, 0, 1)
	$missing1 = ""
	$missing2 = ""
	$commonGroups = ""

	For $i = 0 To _GUICtrlListView_GetItemCount($listGroups) - 1 ; durchlaufe alle Gruppen des ge�ffneten Users
		$flag = 0 ; per default hat der User die Gruppe, der Vergleichsuser nicht
		$name1 = _GUICtrlListView_GetItem($listGroups, $i, 0)
		$name1 = $name1[3]
		For $name2 In $vergleichsgruppen ; durchlaufe alle Gruppen des Vergleichs-Users
			If $name1 = $name2 Then ; wenn der Name der aktuellen Gruppe vom User auch beim Vergleichsuser vorhanden ist
				$flag = 1 ; setze das Flag
				$commonGroups = $commonGroups & $name1 & @CRLF ; und f�ge den Gruppennamen zur Liste der gemeinsamen Gruppen hinzu
				ExitLoop ; verlasse den Durchlauf der Gruppen des Vergleichs-Users
			EndIf
		Next
		; wenn Gruppen des Vergleichs-Users durchsucht wurden ohne Treffer, f�ge die Gruppe des Users zu Liste der Gruppen hinzu, die nur der User hat
		If $flag = 0 Then $missing1 = $missing1 & $name1 & @CRLF
	Next

	For $name2 In $vergleichsgruppen ; durchlaufe alle Gruppen des Vergleichsusers
		$flag = 0
		For $i = 0 To _GUICtrlListView_GetItemCount($listGroups) - 1 ; step through the groups of open Users
			$name1 = _GUICtrlListView_GetItem($listGroups, $i, 0) ; get the advertised (!) group name from the first column
			$name1 = $name1[3]
			If $name2 = $name1 Then
				$flag = 1
				ExitLoop
			EndIf
		Next
		If $flag = 0 Then $missing2 = $missing2 & $name2 & @CRLF
	Next

	$missing1 = StringSplit($missing1, @CRLF)
	$missing2 = StringSplit($missing2, @CRLF)
	$commonGroups = StringSplit($commonGroups, @CRLF)

	; k�rze bei den beiden Listen der Gruppen die nur 1 User hat, die ersten 2 Zeilen (Anzahl der Elemente und ein Leerfeld) weg
	;	_ArrayDelete($missing1, 0)
	_ArrayDelete($missing2, 0)
	;	_ArrayDelete($missing1, 0)
	_ArrayDelete($missing2, 0)

	; ermittle die maximale Anzahl der Eintr�ge aller 3 Spalten
	$max = 0
	$m1 = UBound($missing1)
	$m2 = UBound($missing2)
	$m3 = UBound($commonGroups)
	$max = _Max($max, $m1)
	$max = _Max($max, $m2)
	$max = _Max($max, $m3)

	$datei = FileOpen(@TempDir & "\userinfo.htm", 2)
	FileWriteLine($datei, "<table>")
	FileWriteLine($datei, "<tr><th bgcolor=ff6666>" & $texts[$language][94] & $user & $texts[$language][95] & "</th><th bgcolor=66ff66>" & $texts[$language][96] & "</th><th bgcolor=ff6666>" & $texts[$language][94] & $compareUser & $texts[$language][95] & "</th></tr>")
	For $i = 1 To $max
		$s1 = ""
		$s2 = ""
		$s3 = ""
		If $i < $m1 Then $s1 = $missing1[$i]
		If $i < $m2 Then $s2 = $missing2[$i]
		If $i < $m3 Then $s3 = $commonGroups[$i]
		If $s1 = "" And $s2 = "" And $s3 = "" Then ContinueLoop
		FileWriteLine($datei, "<tr><td bgcolor=ff9999>" & $s1 & "</td><td bgcolor=99ff99>" & $s3 & "</td><td bgcolor=ff9999>" & $s2 & "</td></tr>")
	Next
	FileWriteLine($datei, "</table>")
	FileClose($datei)
	_IECreate(@TempDir & "\userinfo.htm")
	_GUICtrlStatusBar_SetText($sbSearchAD, $texts[$language][97])
EndFunc   ;==>compareGroup

Func countUsers()
	$z = ""
	$summe = 0
	For $i = 1 To 6
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(department=*0" & $i & "*));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
		$z = $z & "Abteilung " & $i & ": " & $objRecordSet.RecordCount & @CRLF
		$summe += $objRecordSet.RecordCount
	Next
	$z &= "===========================" & @CRLF & "Summe: " & $summe
	MsgBox(0, $texts[$language][124], $z)
EndFunc   ;==>countUsers

Func searchDepartment()
	$room = $oUsr.physicalDeliveryOfficeName
	$dezernat = $oUsr.department

	; Create GUI
	$Form4 = GUICreate($texts[$language][114], 550, 510)
	$liste = GUICtrlCreateListView($texts[$language][115], 5, 10, 540, 400)
	_GUICtrlListView_SetColumnWidth($liste, 0, 70)
	_GUICtrlListView_SetColumnWidth($liste, 1, 250)
	$btnNeighbours_vergleich = GUICtrlCreateButton($texts[$language][116], 5, 415, 540, 20)
	$btnNeighbours_wechseln = GUICtrlCreateButton($texts[$language][117], 5, 440, 540, 20)
	$btnNeighbours_ok = GUICtrlCreateButton("Ok", 5, 465, 540, 35)

	; suche nach Nachbarn auf dieser Etage
	Dim $treffer_array[1]
	$z = ""
	$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(department=" & $dezernat & "));ADsPath;subtree"
	$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	Do
		$y = $objRecordSet.Fields(0).Value ; FQDN-Name des Users
		If Not StringInStr($y, "ou=Empf�nger") Then ; pass over all that exists only as a mail receiver (which are namely ou = users, ou = receiver
			$o_temp_Usr = ObjGet($objRecordSet.Fields(0).Value) ; Retrieve the COM Object for the logged on user
			_ArrayAdd($treffer_array, $o_temp_Usr.physicalDeliveryOfficeName & "|" & $o_temp_Usr.sn & "," & $o_temp_Usr.givenName & "|" & $o_temp_Usr.samAccountName & "|" & $o_temp_Usr.telephoneNumber)
		EndIf
		$objRecordSet.MoveNext
	Until $objRecordSet.EOF
	_ArrayDelete($treffer_array, 0) ; das leere erste Feld l�schen
	_ArraySort($treffer_array) ; Treffer sortieren

	$room = $oUsr.physicalDeliveryOfficeName
	For $i = 0 To UBound($treffer_array) - 1 ; und alle Ergebnisse in Listview kopieren
		$x = GUICtrlCreateListViewItem($treffer_array[$i], $liste)
		$wert = StringSplit($treffer_array[$i], "|")
		If $wert[1] = $room Then GUICtrlSetBkColor($x, 0xffff00) ; die Raumnummer des Users einf�rben
	Next

	GUISetState(@SW_SHOW, $Form4)

	While 1
		$msg = GUIGetMsg()
		If ($msg = $btnNeighbours_wechseln) Or ($msg = $btnNeighbours_vergleich) Then
			$x = GUICtrlRead($liste)
			If $x <> "" Then ; nur wenn auch ein User gew�hlt wurde...
				$y = GUICtrlRead($liste) ; die Control-ID des markierten Listview-Items holen
				$y = GUICtrlRead($y) ; das ListviewItem auslesen
				$tmp = StringSplit($y, "|") ; aufteilen in einzelne Felder
				If $msg = $btnNeighbours_wechseln Then
					GUIDelete($Form4) ; Auswahl-GUI l�schen
					setUser($tmp[3])
				Else
					GUIDelete($Form4) ; Auswahl-GUI l�schen
					compareGroup($tmp[3])
				EndIf
				ExitLoop
			EndIf
		EndIf

		If ($msg = $GUI_EVENT_CLOSE) Or ($msg = $btnNeighbours_ok) Then
			GUISetState(@SW_SHOW, $frmMain)
			GUIDelete($Form4) ; Auswahl-GUI l�schen
			ExitLoop ; und raus aus der Schleife
		EndIf
	WEnd
EndFunc   ;==>searchDepartment
