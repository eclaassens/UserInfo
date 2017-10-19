#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
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
#include <Math.au3> ; importing the max command
#include <Misc.au3>  ; importing _singleton
#include <ComboConstants.au3>
#include <GuiListView.au3>
#include <INet.au3>
#include <String.au3>
#include <localization.au3> ; load translations

#region  ; Define AD Constants
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
#endregion  ; Define AD Constants


Global $objConnection = ObjCreate("ADODB.Connection") ; Create COM object to AD
If @error Then
	MsgBox(16, "Problem", $texts[$language][0])
	Exit
EndIf
$objConnection.ConnectionString = "Provider=ADsDSOObject"
$objConnection.Open("Active Directory Provider") ; Open connection to AD

;Global $connEmpirum = ObjCreate("ADODB.Connection") ; empirum verbindung
;Global $connNCDaten = ObjCreate("ADODB.Connection") ; NC Daten Verbindung
;Global $connSBDAT = ObjCreate("ADODB.Connection") ; Sachbearbeiter-Datei Verbindung
;Global $rs = ObjCreate("ADODB.Recordset") ; Recordset für Suchergebnisse bei den Datenbanken
Global $ADSystemInfo = ObjCreate("ADSystemInfo")
Global $anzahl_pcs
Global $anzahl_ncs

#region  ; User Interface
$titel = "Userinfo v2.79 (w)"
If @AutoItX64 Then
	$titel &= " [64 Bit]"
Else
	$titel &= " [32 Bit]"
EndIf

If @Compiled And _Singleton("userinfo", 1) = 0 Then ; Only start one instance of the application
	WinActivate($titel)
	Exit
EndIf

$Form1 = GUICreate($titel, 980, 800, -1, -1)
GUISetBkColor(0xeeeeee)
GUICtrlSetFont($Form1, 6)
$reihe1 = 16 ; Start  1. description
$reihe2 = 110 ; Start  1. value
$reihe3 = 410 ; Start  2. description
$reihe4 = 480 ; Start  2. value
$breite1 = 370 ; Width  1. value
$breite2 = 180 ; Width  2. value
$breite3 = 60
$breite4 = 220

Global Const $buttonreihe = 825
Global Const $buttonbreite = 140

GUICtrlCreateGroup($texts[$language][1], 820, 5, 150, 770)
$y = 20
; for some reason, the "flashing" of labels (when clicked) only works when labes are created after creating the buttons...
$bn_refresh = GUICtrlCreateButton($texts[$language][2], $buttonreihe, $y, $buttonbreite, 40)
GUICtrlSetImage(-1, "shell32.dll", 255)
$y += 40

$btn_nachbarn = GUICtrlCreateButton($texts[$language][3], $buttonreihe, $y, $buttonbreite, 40)
GUICtrlSetImage(-1, @SystemDir & "\shell32.dll", 269)
$y += 40

$btnChangeUser = GUICtrlCreateButton($texts[$language][4], $buttonreihe, $y, $buttonbreite, 40)
GUICtrlSetImage(-1, "shell32.dll", 23)
$y += 60

$bn_overview_copy = GUICtrlCreateButton($texts[$language][5], $buttonreihe, $y, $buttonbreite, 40)
GUICtrlSetImage(-1, "shell32.dll", 223)
$y += 40

$bn_overview_print = GUICtrlCreateButton($texts[$language][6], $buttonreihe, $y, $buttonbreite, 40)
GUICtrlSetImage(-1, "shell32.dll", 17)
$y += 80

$bn_vnc_nc = GUICtrlCreateButton($texts[$language][7], $buttonreihe, $y, $buttonbreite, 20)
$y += 20

$bn_citrix = GUICtrlCreateButton($texts[$language][8], $buttonreihe, $y, $buttonbreite, 20)
$y += 30

$bn_userzaehler = GUICtrlCreateButton($texts[$language][9], $buttonreihe, $y, $buttonbreite, 20)
$y += 20
$btnDepartment = GUICtrlCreateButton($texts[$language][10], $buttonreihe, $y, $buttonbreite, 20)
$y += 70

$btnCompareGroups = GUICtrlCreateButton($texts[$language][11], $buttonreihe, $y, $buttonbreite, 20)
$y += 20

$btnGroups = GUICtrlCreateButton($texts[$language][12], $buttonreihe, $y, $buttonbreite, 20)
$y += 20

$bn_gruppenmitglieder = GUICtrlCreateButton($texts[$language][13], $buttonreihe, $y, $buttonbreite, 20)
$y += 60

$bn_ende = GUICtrlCreateButton($texts[$language][19], $buttonreihe, $y, $buttonbreite, 40)
GUICtrlSetImage(-1, "shell32.dll", 28)
$y += 40

GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group


$y = 10
GUICtrlCreateLabel($texts[$language][20], $reihe1, $y, 90, 17)
$lbl_name = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)
GUICtrlCreateLabel($texts[$language][21], $reihe3, $y, $breite3, 17)
$lblUserId = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][22], $reihe1, $y, 90, 17)
$lbl_anrede = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)
;GUICtrlCreateLabel($texts[$language][23], $reihe3, $y, $breite3, 17)
;$lbl_fnummer = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
;GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][24], $reihe1, $y, 90, 17)
$lbl_ou = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][25], $reihe3, $y, $breite3, 17)
$lbl_telefon = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][26], $reihe1, $y, 69, 17)
$lblDepartment = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][27], $reihe3, $y, $breite3, 17)
$lblDomain = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][28], $reihe1, $y, 28, 17)
$lbl_email = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][29], $reihe3, $y, 69, 17)
$lbl_ort = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][30], $reihe1, $y, 90, 17)
$lbl_adresse = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][31], $reihe3, $y, 69, 17)
$lbl_raum = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
; Tabelle zur Anzeige der PC/NC Namen
$computer = GUICtrlCreateListView($texts[$language][32], 15, $y, 800, 85)
_GUICtrlListView_SetColumnWidth($computer, 0, 180)
_GUICtrlListView_SetColumnWidth($computer, 1, 90)
_GUICtrlListView_SetColumnWidth($computer, 2, 90)
_GUICtrlListView_SetColumnWidth($computer, 3, 28)
_GUICtrlListView_SetColumnWidth($computer, 4, 180) ; Name des NCs
_GUICtrlListView_SetColumnWidth($computer, 5, 90)
_GUICtrlListView_SetColumnWidth($computer, 6, 90)
_GUICtrlListView_SetColumnWidth($computer, 7, 28)


$breite4 = 330
$y += 90
GUICtrlCreateLabel($texts[$language][33], $reihe1, $y, 90, 17)
$lbl_aenderung = GUICtrlCreateLabel("", $reihe2, $y, 130, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][34], $reihe1 + 220, $y, 55, 17)
$lbl_expiration = GUICtrlCreateLabel("", $reihe1 + 280, $y, 130, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][35], $reihe3, $y, $breite3, 17)
$lbl_pwreq = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20

GUICtrlCreateLabel($texts[$language][36], $reihe3, $y, $breite3, 17)
$lbl_change = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][37], $reihe1, $y, 85, 17)
$lbl_gesperrt = GUICtrlCreateLabel("", $reihe2, $y, 30, 17)
GUICtrlSetFont(-1, -1, 800)
GUICtrlCreateLabel("/", $reihe2 + 30, $y, 5, 17)
$lbl_deakt = GUICtrlCreateLabel("-", $reihe2 + 40, $y, 50, 17)
GUICtrlSetFont(-1, -1, 800)

$bn_unlock = GUICtrlCreateButton($texts[$language][38], 200, $y - 4, 90, 20)

GUICtrlCreateLabel($texts[$language][39], $reihe3, $y, $breite3, 17)
$lbl_erzeugt = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][40], $reihe1, $y, 85, 17)
$lbl_errorcount = GUICtrlCreateLabel("", $reihe2, $y, 85, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][41], $reihe3, $y, $breite3, 17)
$lbl_login = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][42], $reihe1, $y, 89, 17)
$lbl_homeverz = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][43], $reihe3, $y, $breite3, 17)
$lbl_homelauf = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][44], $reihe1, $y, 55, 17)
$lbl_script = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][45], $reihe3, $y, $breite3, 17)
$lbl_profil = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][46], $reihe1, $y, 85, 17)
$lbl_geboren = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

;GUICtrlCreateLabel($texts[$language][47], $reihe3, $y, 55, 17)
;$lbl_TSProfil = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
;GUICtrlSetFont(-1, -1, 800)

;$y += 20
;GUICtrlCreateLabel($texts[$language][48], $reihe1, $y, 85, 17)
;$lbl_postkorb = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
;GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][49], $reihe3, $y, 55, 17)
$lbl_AmtsBez = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][50], $reihe1, $y, 85, 17)
$lbl_pwchange = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel($texts[$language][51], $reihe3, $y, 55, 17)
$lbl_geschlecht = GUICtrlCreateLabel("", $reihe4, $y, $breite4, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 20
GUICtrlCreateLabel($texts[$language][52], $reihe1, $y, 85, 17)
$lbl_SAP_PNR = GUICtrlCreateLabel("", $reihe2, $y, $breite1, 17)
GUICtrlSetFont(-1, -1, 800)

$y += 25
$listGroups = GUICtrlCreateListView($texts[$language][53], 8, $y, 800, 770 - $y)
_GUICtrlListView_SetColumnWidth($listGroups, 0, 240)
_GUICtrlListView_SetColumnWidth($listGroups, 1, 200)

$bar = _GUICtrlStatusBar_Create($Form1)


; Construction of the window for entering the user name and domain
$Form2 = GUICreate($titel, 370, 140, -1, -1)
GUICtrlCreateLabel("UserID/Name/Tel.:", 5, 8, 100, 17)
$edit = GUICtrlCreateInput(@UserName, 110, 5, 250, 20)
GUICtrlCreateLabel("Wildcards:", 5, 35, 100, 17)
$radio1 = GUICtrlCreateRadio("UserID", 110, 35, 70, 20)
$radio2 = GUICtrlCreateRadio($texts[$language][54], 190, 35, 80, 20)
$radio3 = GUICtrlCreateRadio($texts[$language][55], 270, 35, 120, 20)
GUICtrlSetState($radio2, $GUI_CHECKED)
$Enter_key = GUICtrlCreateDummy()
Dim $a_AccelKeys[1][2] = [["{ENTER}", $Enter_key]] ; Hotkey array for evaluating the Enter button on Form2
GUISetAccelerators($a_AccelKeys, $Form2)


GUICtrlCreateLabel("Domain:", 5, 65, 100, 17)
$cmbDomain = GUICtrlCreateCombo("", 110, 62, 250, 250, $CBS_DROPDOWNLIST)
$btnOk = GUICtrlCreateButton("Ok", 110, 90, 120, 25)
$btn_cancel = GUICtrlCreateButton($texts[$language][56], 240, 90, 120, 25)
$statusbar = _GUICtrlStatusBar_Create($Form2)

#endregion  ; User Interface

Global $oMyError = ""
$oMyError = ObjEvent("AutoIt.Error", "_ADDoError") ; Install a custom error handler

Dim $gruppen
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
Global $time ; Zeitstempel fürs Stoppen des Auslesens der Userdaten
Global $pc_status = 0 ; PC online?
Global $cancel ; Flag zum Beenden der Domänensuche
Global $loopflag ; dient zur Kontrolle der Schleife bei der Userauswahl
Global $passwortdauer ; wie lang ein PW in der Domäne gültig ist

; ==================================================================================================================
;        Main program ;)
; ==================================================================================================================
listDomains()
establishUser()

; ==================================================================================================================
;        Functions
; ==================================================================================================================

Func listDomains() ; get all known Domains - and, in our case, select ours
	$oDS = ObjGet("WinNT:")
	$y = ""
	For $x In $oDS
		$y = $y & "|" & $x.Name
	Next
	GUICtrlSetData($cmbDomain, $y, "NL")
EndFunc   ;==>listDomains

Func establishUser($userid_tmp = "")
	ControlFocus($titel, "", $btnChangeUser)
	_GUICtrlStatusBar_SetText($statusbar, $texts[$language][58])
	$pc = ""
	$lw_m = ""
	GUISetState(@SW_HIDE, $Form1)
	GUISetState(@SW_SHOW, $Form2)

	$loopflag = 0
	If $userid_tmp <> "" Then
		GUICtrlSetData($edit, $userid_tmp)
		$loopflag = finde_domaene()
	EndIf

	If $loopflag = 0 Then
		ControlClick($titel, "", $edit)
		Send("+{home}")
		$loopflag = 1
		While $loopflag
			$msg = GUIGetMsg()
			Switch $msg
				Case $Enter_key
					finde_domaene()
				Case $GUI_EVENT_CLOSE
;					$connEmpirum.Close
;					$objConnection.Close
;					$connNCDaten.Close
					Exit
				Case $btnOk
					finde_domaene()
				Case $btn_cancel
					Exit
			EndSwitch
		WEnd
	EndIf
	GUISetState(@SW_HIDE, $Form2)

	; if we get here, useraccount has been found. -> get our values
	$ldap_entry = $objRecordSet.fields(0).value

	$oUsr = ObjGet($ldap_entry) ; Retrieve the COM Object for the logged on user
	$user = $oUsr.samAccountName ; get SamAccountname

	GUISetState(@SW_SHOW, $Form1)
	GUISetCursor(15, 1, $Form1)
	_GUICtrlStatusBar_SetText($bar, $texts[$language][59])

	queryAD()
	GUISetCursor(2)
	GUILoop()
EndFunc   ;==>establishUser

Func GUILoop()
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
;				$connEmpirum.Close
;				$objConnection.Close
;				$connNCDaten.Close
				Exit
			Case $bn_ende
;				$connEmpirum.Close
;				$objConnection.Close
;				$connNCDaten.Close
				Exit
			Case $btnGroups
				For $i = 0 To UBound($gruppen) - 1
					$tmp = StringSplit($gruppen[$i], ",")
					$tmp[1] = StringReplace($tmp[1], "CN=", "")
					$gruppen[$i] = $tmp[1]
				Next
				_ArrayToClip($gruppen, 1)
			Case $bn_gruppenmitglieder
				$x = GUICtrlRead($listGroups)
				If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
					$y = GUICtrlRead($x) ; das ListviewItem auslesen
					$tmp = StringSplit($y, "|") ; aufteilen und nur die UserID behalten
					_ADGetGroupMembers($tmp[2])
				EndIf
			Case $bn_unlock
				Entsperren($user)
			Case $btnCompareGroups
				compareGroup()
			Case $bn_refresh
				$time = TimerInit() ; reset timer
				queryAD()
			;Case $bn_vnc_nc ; NC mit VNC öffnen
			;	vnc("nc")
			Case $btnChangeUser
				establishUser()
			Case $edit ; sonst wird der Text aus dem Edit-Feld der Suchmaske ins Clipboard kopiert wenn ein Treffer über Wildcardsuche erfolgte
				;
			Case $computer ; Sonst wird der String aus der Computerliste beim Anklicken kopiert
			;Case $bn_overview_copy
;				drucken(0)
			;Case $bn_overview_print
;				drucken(1)
			;Case $bn_citrix
			;	Citrixprofile()
			Case $btn_nachbarn
				suche_nachbarn()
			Case $btnDepartment
				suche_dezernat()
			Case $bn_userzaehler
				userzaehlen()
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
					_GUICtrlStatusBar_SetText($bar, $tmp & $texts[$language][60]) ; update status bar
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

Func finde_domaene($user = "", $domain = "")
	If $user = "" Then $user = GUICtrlRead($edit)
	If $domain = "" Then $domain = GUICtrlRead($cmbDomain)

	If $user <> "" Then
		$time = TimerInit()
		$user = StringStripWS($user, 3)

		$objRootDSE = ObjGet("LDAP://" & $domain & "/RootDSE")
		If @error Then
			_GUICtrlStatusBar_SetText($statusbar, $texts[$language][61] & $domain & $texts[$language][62])
			$loopflag = 1
			Return 0
		EndIf
		$strDNSDomain = $objRootDSE.Get("defaultNamingContext") ; Retrieve the current AD domain name
		$strHostServer = $objRootDSE.Get("dnsHostName") ; Retrieve the name of the connected DC
		$strConfiguration = $objRootDSE.Get("ConfigurationNamingContext") ; Retrieve the Configuration naming context

		Switch _ADObjectExists($user)
			Case 1 ; User wurde erfolgreich gefunden
				$loopflag = 0
				$x = ObjGet("WinNT://" & $domain)
				$passwortdauer = $x.MaxPasswordAge
				Return 1
			Case -1 ; Suchbegriff lieferte mehrere Treffer, daher wurde die UserID ins Formular eingetragen
				ControlClick($titel, "", $btnOk)
				Return 0 ; und brav wieder das Fensterchen anzeigen, nun aber mit der eindeutigen Userkennung
			Case 0 ; User wurde nicht gefunden
				_GUICtrlStatusBar_SetText($statusbar, $texts[$language][63] & $user & $texts[$language][64])
				$loopflag = 1
				Return 0
		EndSwitch
	EndIf
EndFunc   ;==>finde_domaene

Func queryAD()
	$pc = ""
	$lastlogon = "-"
	GUICtrlSetState($bn_unlock, $GUI_DISABLE)
	GUICtrlSetState($bn_vnc_nc, $GUI_DISABLE)
	;GUICtrlSetState($bn_eventvwr, $GUI_DISABLE)
	;GUICtrlSetState($bn_remote, $GUI_DISABLE)
	;GUICtrlSetState($bn_laufwerk_c, $GUI_DISABLE)
	;GUICtrlSetState($bn_versionsinfo, $GUI_DISABLE)
	;GUICtrlSetState($bn_software, $GUI_DISABLE)
	GUICtrlSetData($lbl_homelauf, "")
	GUICtrlSetData($lbl_homeverz, "")
	GUICtrlSetData($lbl_script, "")
;	GUICtrlSetData($lbl_TSProfil, "")
;	GUICtrlSetData($lbl_postkorb, "")
	GUICtrlSetData($lbl_AmtsBez, "")
	GUICtrlSetData($lbl_pwchange, "")
	GUICtrlSetData($lbl_geschlecht, "")
	GUICtrlSetData($lbl_SAP_PNR, "")
	_GUICtrlListView_DeleteAllItems($listGroups)
	_GUICtrlListView_DeleteAllItems($computer)

	$anzahl_pcs = 0
	$anzahl_ncs = 0

	_ADGetUserData($user)
	_ADGetUserGroups($gruppen, $user)
	_ADGetAccount($user)
	;DB_Abfrage($user)

	$tmp = $oUsr.scriptpath

	If $lastlogon = "-" Then ; Domänen LastLogin auswerten
		$tmp = $oUsr.LastLogin
		If $tmp <> "" Then
			$lastlogon = Zeit($oUsr.LastLogin)
		Else
			$lastlogon = $texts[$language][65]
		EndIf
	EndIf
	GUICtrlSetData($lbl_login, $lastlogon)

	$tmp = TimerDiff($time) / 1000
	$tmp1 = StringFormat("%.2f", $tmp)

	;Empirum() ; Daten aus Empirum auslesen
	;NC_Daten()

	$tmp = TimerDiff($time) / 1000
	$tmp2 = StringFormat("%.2f", $tmp)
	_GUICtrlStatusBar_SetText($bar, $texts[$language][66] & $tmp2 & $texts[$language][67] & $tmp1 & $texts[$language][68] & $strHostServer)
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
		Case $oMyError.windescription = "Access is denied."
			$objConnection.Close("Active Directory Provider")
			$objConnection.Open("Active Directory Provider")
			SetError(2)
		Case 1
			SetError(1)
	EndSelect
EndFunc   ;==>_ADDoError


Func _ADObjectExists($object)
	$flag = 0 ; für Auswertung der Radiobuttons
	$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sAMAccountName=" & $object & "));ADsPath;subtree"
	$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	Switch $objRecordSet.RecordCount
		Case 0; User wurde nicht UserID gefunden
			; nix tun, weitere Suchen ablaufen lassen
		Case 1 ; User wurde eindeutig identifiziert anhand des Namens
			GUISetState(@SW_SHOW, $Form2) ; Fenster zur Benutzerwahl wieder einblenden
			Return 1
		Case Else ; mehrere User wurden über UserID gefunden
			If GUICtrlRead($radio1) = $GUI_CHECKED Then $flag = 1 ; Ergebnis nur behalten wenn suche nach UserID aktiviert
	EndSwitch

	If $flag = 0 Then ; wenn bei der UserID Suche nichts gefunden wurde bzw. Suche nach Name/Vorname gewünscht
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sn=" & $object & "));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
		Switch $objRecordSet.RecordCount
			Case 0; User wurde nicht per Nachname gefunden
				; nix tun, weiter prüfen
			Case 1 ; User wurde eindeutig identifiziert anhand des Namens
				GUISetState(@SW_SHOW, $Form2) ; Fenster zur Benutzerwahl wieder einblenden
				Return 1
			Case Else ; mehrere User wurden über UserID gefunden -> Userliste anzeigen lassen
				If GUICtrlRead($radio2) = $GUI_CHECKED Then $flag = 1 ; Ergebnis nur behalten wenn suche nach Nachname aktiviert
		EndSwitch
	EndIf

	If ($flag = 0) And (GUICtrlRead($radio3) = $GUI_CHECKED) Then ; Wenn immer noch nichts gefunden, suche nach dem Vornamen
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(givenName=" & $object & "));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery)
	EndIf

	If $objRecordSet.RecordCount = 0 Then ; Bisher wurde rein gar nichts gefunden -> also suchen wir nach einer Telefonnummer
		If StringIsDigit($object) Then $object = "*" & $object
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(telephoneNumber=" & $object & "));ADsPath;subtree" ; also suche nach Telefonnummer
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	EndIf

#CS 	If $objRecordSet.RecordCount = 0 Then ; Bisher wurde rein gar nichts gefunden -> also suchen wir nach einer Postkorbnummer
; 		$tmp2 = suche_postkorb($object)
; 		If $tmp2 <> "" Then
; 			GUICtrlSetData($edit, $tmp2) ; die Userkennung aus dem Ergebnis der Postkorbsuche wird im Suchfenster eingetragen
; 			Return -1 ; Zurück zum Aufrufer mit Info, dass neue UserID in Suchmaske eingetragen wurde
; 		EndIf
; 	EndIf
 #CE


	Switch $objRecordSet.RecordCount
		Case 0 ; dieser User wurde nicht gefunden
			Return 0
		Case 1 ; User wurde eindeutig identifiziert anhand des Namens
			GUISetState(@SW_SHOW, $Form2) ; Fenster zur Benutzerwahl wieder einblenden
			Return 1
		Case Else ; User existiert, aber mehrere Treffer (Suchbegriff: Meier ^^)
			Dim $treffer_arry[1]
			$z = ""
			Do
				$y = $objRecordSet.Fields(0).Value ; FQDN-Name des Users
				If Not StringInStr($y, "ou=Empfänger") Then ; skip all mail-only accounts (ou=benutzer,ou=empfänger)
					$oUsr = ObjGet($objRecordSet.Fields(0).Value) ; Retrieve the COM Object for the logged on user
					_ArrayAdd($treffer_arry, $oUsr.sn & "," & $oUsr.givenName & "|" & $oUsr.samAccountName)
				EndIf
				$objRecordSet.MoveNext
			Until $objRecordSet.EOF
			If UBound($treffer_arry) = 2 Then ; Array hat nur 1 Element => die anderen Treffer waren reine Mailempfänger und wurden beim Übertragen ins Array übergangen
				$x = StringInStr($treffer_arry[1], "|")
				GUICtrlSetData($edit, StringMid($treffer_arry[1], $x + 1)) ; die Userkennung wird im Suchfenster eingetragen
				GUISetState(@SW_SHOW, $Form2) ; Fenster zur Benutzerwahl wieder einblenden
				Return -1
			Else ; aha, es gibt tatsächlich mehrere User die in Frage kommen
				GUISetState(@SW_HIDE, $Form2)
				$Form3 = GUICreate($texts[$language][69], 350, 355)
				$Enter_key2 = GUICtrlCreateDummy()
				Dim $b_AccelKeys[1][2] = [["{ENTER}", $Enter_key2]] ; Hotkey-Array für das Auswerten der Enter-Taste in Form3
				GUISetAccelerators($b_AccelKeys, $Form3)
				$liste = GUICtrlCreateListView($texts[$language][70], 5, 40, 340, 280)
				_GUICtrlListView_SetColumnWidth($liste, 0, 250)
				$btn_userwahl = GUICtrlCreateButton("Ok", 5, 325, 340, 25)
				GUICtrlCreateLabel($texts[$language][71] & UBound($treffer_arry) - 1 & $texts[$language][72], 5, 5, 290, 30)
				_ArrayDelete($treffer_arry, 0) ; das leere erste Feld löschen
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
						If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
							$y = GUICtrlRead($liste) ; die Control-ID des markierten Listview-Items holen
							$y = GUICtrlRead($y) ; das ListviewItem auslesen
							$tmp = StringSplit($y, "|") ; aufteilen und nur die UserID behalten
							$x = GUICtrlSetData($edit, $tmp[2]) ; die Userkennung wird im Suchfenster eingetragen
							ExitLoop
						EndIf
					EndIf
				WEnd
				GUIDelete($Form3) ; Auswahl-GUI löschen
				GUISetState(@SW_SHOW, $Form2) ; Fenster zur Benutzerwahl wieder einblenden
				Return -1 ; User wurde im Formuar eingetragen
			EndIf
	EndSwitch
EndFunc   ;==>_ADObjectExists

Func _ADGetUserData($user)

	$oUsr.GetInfo() ; ADS-Cache der Attribute neu auslesen sonst kriegt das Programm evtl. keine Gruppenänderungen mit
	$tmp = $oUsr.DisplayName
	If $tmp <> "" Then
		GUICtrlSetData($lbl_name, $oUsr.DisplayName)
	Else
		GUICtrlSetData($lbl_name, $oUsr.sn & ", " & $oUsr.givenName)
	EndIf

	GUICtrlSetData($lbl_anrede, $oUsr.title)

	; evaluation of the "Released on" value. Case: If a "" in the name of the user himself is in it, he is as \ masked.
	; That bothers StringSplit but unfortunately not. Therefore we convert the thing before StringSplit to and then back.
	$tmp = $oUsr.distinguishedName
	$tmp = StringReplace($tmp, "\,", "õõ")
	$tmp = StringSplit($tmp, ",")

	$tmp2 = ""
	For $i = 1 To $tmp[0] ; Abtrennen der DC= Klamotten
		$tmp[$i] = StringReplace($tmp[$i], "õõ", ",")
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

#CS 	If (@OSArch = "X86") Or (@OSArch = "X64" And @AutoItX64 = 1) Then
; 		GUICtrlSetData($lbl_TSProfil, $oUsr.TerminalServicesProfilePath)
; 	Else
; 		GUICtrlSetData($lbl_TSProfil, $texts[$language][73])
; 	EndIf
 #CE
EndFunc   ;==>_ADGetUserData

#CS Func DB_Abfrage($user)
; 	Return
; 	$connSBDAT.Open("Driver={SQL Server};Server=xxx.de;Database=X500;") ; mit der Datenbank verbinden
; 	If @error Then
; 		SplashTextOn("", $texts[$language][125], 250, 100)
; 		Sleep(2000)
; 		SplashOff()
; 		Return
; 	EndIf
;
; 	$suche = "SELECT EBK, GebDatum,Geblist, PWChange FROM xxxxxxxxxxx WHERE UserID LIKE '" & $user & "'"
; 	$rs.Open($suche, $connSBDAT)
; 	$gebdatum = StringLeft($rs.Fields("GebDatum").Value, 8)
;
; 	;	$gebdatum = StringRight($gebdatum, 2) & "." & StringMid($gebdatum, 5, 2) & "." &
;
; 	$gebjahr = StringLeft($gebdatum, 4)
; 	$gebmon = StringMid($gebdatum, 5, 2)
; 	$gebtag = StringRight($gebdatum, 2)
; 	$gebdatum = $gebtag & "." & $gebmon & "." & $gebjahr & " (" & _DateDiff("y", $gebjahr & "/" & $gebmon & "/" & $gebtag, _NowCalc()) & ")"
; 	$gebliste = $rs.Fields("GebList").Value
;
; 	If $gebliste <> "X" Then
; 		GUICtrlSetData($lbl_geboren, "geheim")
; 	Else
; 		GUICtrlSetData($lbl_geboren, $gebdatum)
; 	EndIf
;
;
; 	GUICtrlSetData($lbl_fnummer, $rs.Fields("EBK").Value)
; 	GUICtrlSetData($lbl_pwchange, $rs.Fields("PWChange").Value)
; 	$rs.Close
;
; 	$suche = "SELECT Postkorb, AmtsBez, Geschlecht, SAP_PNR FROM T_DB2 WHERE UserID LIKE '" & $user & "'"
; 	$rs.Open($suche, $connSBDAT)
; 	GUICtrlSetData($lbl_postkorb, $rs.Fields("Postkorb").Value)
; 	GUICtrlSetData($lbl_AmtsBez, $rs.Fields("AmtsBez").Value)
; 	GUICtrlSetData($lbl_geschlecht, $rs.Fields("Geschlecht").Value)
; 	GUICtrlSetData($lbl_SAP_PNR, $rs.Fields("SAP_PNR").Value)
;
; 	$rs.Close
; 	$connSBDAT.Close
;
; EndFunc   ;==>DB_Abfrage
 #CE

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
			$tmp2 = "Sicherheit"
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
				$tmp4 = "Lokal"
			Case 8
				$tmp4 = "Universell"
		EndSwitch
		$tmp = StringReplace($objGroup.name, "\/", "/") ; Maskierung wieder rückgängig machen für gescheite Anzeige
		$tmp = StringReplace($tmp, "CN=", "") & "|" & $objGroup.samAccountName & "|" & $tmp2 & "|" & $tmp3[1] & "|" & $tmp4
		GUICtrlCreateListViewItem($tmp, $listGroups)
	Next
	GUICtrlSetData($btnGroups, $texts[$language][12] & " (" & $count & ")")
EndFunc   ;==>_ADGetUserGroups

Func _ADGetAccount($user)
	GUICtrlSetData($lbl_aenderung, Zeit($oUsr.PasswordLastChanged))
	If $oUsr.IsAccountLocked Then
		GUICtrlSetData($lbl_gesperrt, $texts[$language][74])
		GUICtrlSetColor($lbl_gesperrt, 0xff0000)
		GUICtrlSetState($bn_unlock, $GUI_ENABLE)
	Else
		GUICtrlSetData($lbl_gesperrt, $texts[$language][75])
		GUICtrlSetColor($lbl_gesperrt, 0x00aa00)
		GUICtrlSetState($bn_unlock, $GUI_DISABLE)
	EndIf

	GUICtrlSetData($lbl_errorcount, $oUsr.BadLoginCount)

	$intUAC = $oUsr.Get("userAccountControl")

	; Prüfen, ob der User sein Passwort ändern darf. Da keine direkte Abfrage möglich ist, erfolgt dies über den Security Descriptor
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
		GUICtrlSetData($lbl_pwreq, $texts[$language][75])
	Else ; nicht gesetzt -> Passwort nötig
		GUICtrlSetData($lbl_pwreq, $texts[$language][74])
	EndIf

	GUICtrlSetData($lbl_erzeugt, Zeit($oUsr.whenCreated))
	GUICtrlSetData($lbl_script, $oUsr.scriptPath)

	If BitAND($intUAC, $ADS_UF_ACCOUNTDISABLE) Then ; Disabled flag ist gesetzt
		GUICtrlSetData($lbl_deakt, $texts[$language][74])
		GUICtrlSetColor($lbl_deakt, 0xff0000)
	Else
		GUICtrlSetData($lbl_deakt, $texts[$language][75])
		GUICtrlSetColor($lbl_deakt, 0x00aa00)
	EndIf


	GUICtrlSetColor($lbl_expiration, 0x000000)
	$dummy = $oUsr.AccountExpirationDate
	$tmp = Zeit2($dummy) ; Umgewandelte Zeit für DateDiff
	$dummy = Zeit($dummy) ; Umgewandelte Zeit für die Ausgabe

	$tmp2 = _DateDiff("D", _NowCalcDate(), $tmp)
	If ($tmp2 < 1) And ($tmp2 > -148883) Then GUICtrlSetColor($lbl_expiration, 0xff0000) ; Passwort ist keinen Tag mehr gültig

	If ($dummy = "01.01.1601 02:00") Or ($dummy = "01.01.1970 00:00") Or ($dummy = "01.01.1601 01:00") Then
		$dummy = $texts[$language][78]
		GUICtrlSetColor($lbl_expiration, 0x000000)
	EndIf
	GUICtrlSetData($lbl_expiration, $dummy)

EndFunc   ;==>_ADGetAccount

Func Entsperren($user)
	If $oUsr.IsAccountLocked Then
		$oUsr.IsAccountLocked = False
		$oUsr.SetInfo
		Sleep(500)
		$oUsr.GetInfo() ; ADS-Cache der Attribute neu auslesen sonst kriegt das Programm evtl. keine Gruppenänderungen mit
		If Not $oUsr.IsAccountLocked Then
			GUICtrlSetData($lbl_gesperrt, $texts[$language][75])
			GUICtrlSetColor($lbl_gesperrt, 0x00aa00);Green
			GUICtrlSetState($bn_unlock, $GUI_DISABLE)
		Else
			GUICtrlSetData($lbl_gesperrt, $texts[$language][74])
			GUICtrlSetColor($lbl_gesperrt, 0xff0000) ; Red
			GUICtrlSetState($bn_unlock, $GUI_ENABLE)
		EndIf
	EndIf
EndFunc   ;==>Entsperren

Func passwort_req($user)
	$oUsr.Put("userAccountControl", BitXOR($intUAC, 0x00020))
	$oUsr.SetInfo

	$oUsr.GetInfo
	$intUAC = $oUsr.Get("userAccountControl")
	If BitAND(0x00020, $intUAC) Then ; PW not required flag ist gesetzt
		GUICtrlSetData($lbl_pwreq, $texts[$language][75])
		GUICtrlSetColor($lbl_pwreq, 0xff0000)
	Else ; nicht gesetzt -> Passwort nötig
		GUICtrlSetData($lbl_pwreq, $texts[$language][74])
		GUICtrlSetColor($lbl_pwreq, 0x00aa00)
	EndIf
EndFunc   ;==>passwort_req

Func Zeit($zeit) ; Umwandeln der AD-Zeit ins deutsche Format
	If $zeit = "" Then Return "---" ; User wurde grade frisch zurückgesetzt
	$tmp = StringMid($zeit, 7, 2) & "." & StringMid($zeit, 5, 2) & "." & StringLeft($zeit, 4) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2)
	Return $tmp
EndFunc   ;==>Zeit

Func Zeit2($zeit) ; Umwandeln der AD-Zeit ins englische Format für _DateAdd
	$tmp = StringLeft($zeit, 4) & "/" & StringMid($zeit, 5, 2) & "/" & StringMid($zeit, 7, 2) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2) & ":" & StringMid($zeit, 13, 2)
	Return $tmp
EndFunc   ;==>Zeit2

Func Zeit3($zeit) ; Umwandeln vom englischen Datumsformat YYYYMMDD ins deutsche
	If $zeit = "0" Then Return "---" ; User wurde grade frisch zurückgesetzt
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

#CS Func besuchen($lw)
; 	$besuch_flag = 0
; 	$besuch_backup = ClipGet()
; 	If Not @error Then $besuch_flag = 1
; 	If WinExists("Altap Salamander") Then
; 		ClipPut($lw)
; 		WinActivate("Altap Salamander")
; 		WinWaitActive("Altap Salamander")
; 		Send("^v")
; 	ElseIf WinExists("Total Commander") Then
; 		WinActivate("Total Commander")
; 		WinWaitActive("Total Commander")
; 		Send("{right}")
; 		ClipPut("cd " & $lw)
; 		Send("^v")
; 		Sleep(100)
; 		Send("{ENTER}")
; 	Else
; 		Run("explorer.exe " & $lw)
; 	EndIf
; 	If $besuch_flag Then
; 		Sleep(500)
; 		ClipPut($besuch_backup)
; 	EndIf
; EndFunc   ;==>besuchen
 #CE

#CS Func Homelaufwerk_setzen()
; 	$tmp = InputBox("Pfad des Laufwerks eingeben", "Bitte den neuen Pfad für das Verzeichnis eingeben", "", "", 300, 120)
; 	If $tmp <> "" Then
; 		If $oUsr.homeDirectory <> "" Then ; wenn schon was eingetragen ist, lieber nachfragen
; 			$dummy = MsgBox(17, "Wirklich ersetzen?", "Soll der Eintrag" & @CRLF & $oUsr.homeDirectory & @CRLF & "wirklich durch" & @CRLF & $tmp & @CRLF & "ersetzt werden?")
; 			If $dummy <> 1 Then Return
; 		EndIf
; 		$oUsr.homeDirectory = $tmp
; 		$oUsr.homeDrive = "M:"
; 		$oUsr.SetInfo
; 		GUICtrlSetData($lbl_homelauf, "M:")
; 		GUICtrlSetColor($lbl_homelauf, 0x000000)
; 		GUICtrlSetData($lbl_homeverz, $tmp)
; 		GUICtrlSetColor($lbl_homeverz, 0x000000)
; 	EndIf
; EndFunc   ;==>Homelaufwerk_setzen
 #CE

Func Deaktivieren($user)
	$intUAC = $oUsr.Get("userAccountControl")
	If BitAND($intUAC, $ADS_UF_ACCOUNTDISABLE) Then ; Disabled flag ist gesetzt
		GUICtrlSetData($lbl_deakt, $texts[$language][75])
		GUICtrlSetColor($lbl_deakt, 0x00aa00)
	Else
		GUICtrlSetData($lbl_deakt, $texts[$language][74])
		GUICtrlSetColor($lbl_deakt, 0xff0000)
	EndIf
	$oUsr.Put("userAccountControl", BitXOR($intUAC, $ADS_UF_ACCOUNTDISABLE)) ; Flag ändern
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

	_GUICtrlStatusBar_SetText($bar, $texts[$language][91])
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

	For $i = 0 To _GUICtrlListView_GetItemCount($listGroups) - 1 ; durchlaufe alle Gruppen des geöffneten Users
		$flag = 0 ; per default hat der User die Gruppe, der Vergleichsuser nicht
		$name1 = _GUICtrlListView_GetItem($listGroups, $i, 0)
		$name1 = $name1[3]
		For $name2 In $vergleichsgruppen ; durchlaufe alle Gruppen des Vergleichs-Users
			If $name1 = $name2 Then ; wenn der Name der aktuellen Gruppe vom User auch beim Vergleichsuser vorhanden ist
				$flag = 1 ; setze das Flag
				$commonGroups = $commonGroups & $name1 & @CRLF ; und füge den Gruppennamen zur Liste der gemeinsamen Gruppen hinzu
				ExitLoop ; verlasse den Durchlauf der Gruppen des Vergleichs-Users
			EndIf
		Next
		; wenn Gruppen des Vergleichs-Users durchsucht wurden ohne Treffer, füge die Gruppe des Users zu Liste der Gruppen hinzu, die nur der User hat
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

	; kürze bei den beiden Listen der Gruppen die nur 1 User hat, die ersten 2 Zeilen (Anzahl der Elemente und ein Leerfeld) weg
	;	_ArrayDelete($missing1, 0)
	_ArrayDelete($missing2, 0)
	;	_ArrayDelete($missing1, 0)
	_ArrayDelete($missing2, 0)

	; ermittle die maximale Anzahl der Einträge aller 3 Spalten
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
	_GUICtrlStatusBar_SetText($bar, $texts[$language][97])
EndFunc   ;==>compareGroup

#CS Func drucken($modus)
;
; 	$pcs = ""
; 	$ncs = ""
; 	For $row = 0 To _GUICtrlListView_GetItemCount($computer)
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 0) ; PC-Namen auslesen und auf 35 Zeichen erweitern
; 		$tmp2 = $tmp1[3]
; 		$tmp2 = $tmp2 & _StringRepeat(" ", 25 - StringLen($tmp2))
; 		$pcs = $pcs & $tmp2
;
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 1) ; IP auslesen und auf 16 Zeichen erweitern
; 		$tmp2 = $tmp1[3]
; 		$tmp2 = $tmp2 & _StringRepeat(" ", 16 - StringLen($tmp2))
; 		$pcs = $pcs & $tmp2
;
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 2) ; Mac-Adresse auslesen, Länge ist konstant
; 		$tmp2 = $tmp1[3]
; 		$pcs = $pcs & $tmp2
;
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 3) ; Onlinestatus auslesen
; 		$tmp2 = $tmp1[3]
; 		$pcs = $pcs & " " & $tmp2
;
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 4) ; NC-Namen auslesen und auf 35 Zeichen erweitern
; 		$tmp2 = $tmp1[3]
; 		$tmp2 = $tmp2 & _StringRepeat(" ", 25 - StringLen($tmp2))
; 		$ncs = $ncs & $tmp2
;
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 5) ; IP auslesen und auf 16 Zeichen erweitern
; 		$tmp2 = $tmp1[3]
; 		$tmp2 = $tmp2 & _StringRepeat(" ", 16 - StringLen($tmp2))
; 		$ncs = $ncs & $tmp2
;
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 6) ; Mac-Adresse auslesen, Länge ist konstant
; 		$tmp2 = $tmp1[3]
; 		$ncs = $ncs & $tmp2
;
; 		$tmp1 = _GUICtrlListView_GetItem($computer, $row, 7) ; Onlinestatus auslesen
; 		$tmp2 = $tmp1[3]
; 		$ncs = $ncs & " " & $tmp2
;
;
; 		$pcs = $pcs & @CRLF & _StringRepeat(" ", 21)
; 		$ncs = $ncs & @CRLF & _StringRepeat(" ", 21)
; 	Next
; 	$pcs = StringLeft($pcs, StringLen($pcs) - 23)
; 	$ncs = StringLeft($ncs, StringLen($ncs) - 23)
;
;
; 	$res = @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][20]) & ": " & GUICtrlRead($lbl_name) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][21]) & ": " & GUICtrlRead($lblUserId) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][51]) & ": " & GUICtrlRead($lbl_geschlecht) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][46]) & ": " & GUICtrlRead($lbl_geboren) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][49]) & ": " & GUICtrlRead($lbl_AmtsBez) & @CRLF
; 	;$res = $res & StringFormat("[%-22s]", $texts[$language][23]) & ": " & GUICtrlRead($lbl_fnummer) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][52]) & ": " & GUICtrlRead($lbl_SAP_PNR) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][27]) & ": " & GUICtrlRead($lblDomain) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][24]) & ": " & GUICtrlRead($lbl_ou) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][28]) & ": " & GUICtrlRead($lbl_email) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][48]) & ": " & GUICtrlRead($lbl_postkorb) & @CRLF
; 	$res = $res & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][26]) & ": " & GUICtrlRead($lblDepartment) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][29]) & ": " & GUICtrlRead($lbl_ort) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][30]) & ": " & GUICtrlRead($lbl_adresse) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][31]) & ": " & GUICtrlRead($lbl_raum) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][25]) & ": " & GUICtrlRead($lbl_telefon) & @CRLF
; 	$res = $res & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][33]) & ": " & GUICtrlRead($lbl_aenderung) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][39]) & ": " & GUICtrlRead($lbl_erzeugt) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][98]) & ": " & GUICtrlRead($lbl_gesperrt) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][99]) & ": " & GUICtrlRead($lbl_deakt) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][40]) & ": " & GUICtrlRead($lbl_errorcount) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][41]) & ": " & GUICtrlRead($lbl_login) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][36]) & ": " & GUICtrlRead($lbl_pwchange) & @CRLF
; 	$res = $res & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][42]) & ": " & GUICtrlRead($lbl_homeverz) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][43]) & ": " & GUICtrlRead($lbl_homelauf) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][44]) & ": " & GUICtrlRead($lbl_script) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][45]) & ": " & GUICtrlRead($lbl_profil) & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][47]) & ": " & GUICtrlRead($lbl_TSProfil) & @CRLF
; 	$res = $res & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][100]) & ": " & $strHostServer & @CRLF
; 	$res = $res & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][101]) & ": " & $pcs & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][102]) & ": " & $ncs & @CRLF
; 	$res = $res & StringFormat("[%-22s]", $texts[$language][103]) & ": " & UBound($gruppen) - 1
; 	$res = $res & @CRLF & "==================================================" & @CRLF
;
; 	For $i = 1 To UBound($gruppen) - 1
; 		$tmp = StringSplit($gruppen[$i], ",")
; 		$tmp[1] = StringReplace($tmp[1], "CN=", "")
; 		$res = $res & $tmp[1] & @CRLF
; 	Next
;
;
;
; 	If $modus Then ; drucken
; 		$datei = FileOpen(@TempDir & "\userinfo.txt", 2)
; 		FileWriteLine($datei, $res)
; 		FileClose($datei)
; 		_FilePrint(@TempDir & "\userinfo.txt")
; 	Else
; 		ClipPut($res)
; 	EndIf
; EndFunc   ;==>drucken
 #CE

#CS Func Empirum() ; Auslesen der PCs über die Empirum Datenbank
; 	Return
; 	$connEmpirum.Open("Driver={SQL Server};Server=xxxxxxx;Database=xxx;UiD=xxx;Pwd=xxx"); mit der Datenbank verbinden als READONLY
; 	If @error Then
; 		SplashTextOn("", $texts[$language][104], 250, 100)
; 		Sleep(2000)
; 		SplashOff()
; 		Return
; 	EndIf
;
; 	$row = 0
; 	$Text = ""
;
; 	; erst mal nach oslogin suchen
; 	$suche = "SELECT distinct Computername, ipaddress, macaddress FROM dmisystem AS dmi INNER JOIN InvComputer AS cl ON dmi.client_id=cl.client_id WHERE oslogin LIKE '" & $user & "' ORDER BY Computername"
; 	;$rs.Open($suche, $connEmpirum)
; 	While Not $rs.EOF
; 		$anzahl_pcs += 1
; 		GUICtrlCreateListViewItem("", $computer)
; 		$x = StringStripWS($rs.Fields("Computername").Value, 3)
; 		_GUICtrlListView_AddSubItem($computer, $row, $x, 0)
; 		_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("ipaddress").Value, 3), 1)
; 		_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("macaddress").Value, 3), 2)
; 		If Not Ping($x, 50) Then
; 			_GUICtrlListView_AddSubItem($computer, $row, "off", 3)
; 		Else
; 			_GUICtrlListView_AddSubItem($computer, $row, "on", 3)
; 		EndIf
; 		$row += 1
; 		If $row > 9 Then ExitLoop
; 		$rs.MoveNext
; 	WEnd
; 	$rs.Close
;
; 	; anschliessend nach PC-Namen suchen die wie der User heissen unter Berücksichtigung von Umlauten
; 	$temp = StringReplace($oUsr.sn, "ä", "ae")
; 	$temp = StringReplace($temp, "ö", "oe")
; 	$temp = StringReplace($temp, "ü", "ue")
; 	$temp = StringReplace($temp, "ß", "ss")
; 	$temp = StringReplace($temp, "Ä", "Ae")
; 	$temp = StringReplace($temp, "Ö", "Oe")
; 	$temp = StringReplace($temp, "Ü", "Ue")
; 	$temp = $temp & "%"
;
; 	$suche = "SELECT distinct Computername, ipaddress, macaddress FROM dmisystem AS dmi INNER JOIN InvComputer AS cl ON dmi.client_id=cl.client_id WHERE Computername LIKE '" & $temp & "' ORDER BY Computername"
; 	If $Text <> "" Then $Text = $Text & @CRLF ; wenn bereits was gefunden, Zeilenumbruch hinzufügen
; 	$rs.Open($suche, $connEmpirum)
;
; 	$tmp = ""
; 	$tInfo = DllStructCreate($tagLVFINDINFO)
; 	DllStructSetData($tInfo, "Flags", $LVFI_STRING)
;
; 	While Not $rs.EOF
; 		$tmp = StringStripWS($rs.Fields("Computername").Value, 3)
; 		If _GUICtrlListView_FindItem($computer, -1, $tInfo, $tmp) = -1 Then ; nur wenn pc noch nicht drin steht
; 			$anzahl_pcs += 1
; 			$x = StringStripWS($rs.Fields("Computername").Value, 3)
; 			GUICtrlCreateListViewItem("", $computer)
; 			_GUICtrlListView_AddSubItem($computer, $row, $tmp, 0)
; 			_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("ipaddress").Value, 3), 1)
; 			_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("macaddress").Value, 3), 2)
; 			If Not Ping($x, 50) Then
; 				_GUICtrlListView_AddSubItem($computer, $row, "off", 3)
; 			Else
; 				_GUICtrlListView_AddSubItem($computer, $row, "on", 3)
; 			EndIf
; 			$row += 1
; 			If $row > 9 Then ExitLoop
; 		EndIf
; 		$rs.MoveNext
; 	WEnd
;
; 	$rs.Close
; 	$connEmpirum.Close
;
; 	If $anzahl_pcs > 0 Then
; 		;GUICtrlSetState($bn_eventvwr, $GUI_ENABLE)
; 		;GUICtrlSetState($bn_laufwerk_c, $GUI_ENABLE)
; 		;GUICtrlSetState($bn_versionsinfo, $GUI_ENABLE)
; 		;GUICtrlSetState($bn_software, $GUI_ENABLE)
; 		;GUICtrlSetState($bn_remote, $GUI_ENABLE)
; 	EndIf
;
;
 #CE
;EndFunc   ;==>Empirum

#CS Func NC_Daten() ; Auslesen der NCs über die NC-Datenbank
; 	Return
; 	If @AutoItX64 Then
; 		$connNCDaten.Open("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=\\ScoutNGServ\ScoutDB\Scout.mdb")
; 	Else
; 		$connNCDaten.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\ScoutNGServ\ScoutDB\Scout.mdb")
; 	EndIf
;
; 	If @error Then
; 		SplashTextOn("", $texts[$language][105], 500, 100)
; 		; MsgBox(0,Hex(@error),"Peng")
; 		Sleep(1000)
; 		SplashOff()
; 		Return
; 	EndIf
;
; 	$tmp = StringReplace(GUICtrlRead($lbl_name), " (LBV)", "")
;
; 	; hier wird der Vorname verarbeitet - z.B. "Schmidt, Christel-Maria" heisst in der NC-Datenbank "Schmidt, Christel"
; 	$tmp2 = StringInStr($tmp, ",") ; finde das Komma zwischen Nach- und Vornamen
; 	$tmp3 = StringInStr($tmp, " ", 0, -1) ; finde das LETZTE Leerzeichen
; 	If $tmp3 < $tmp2 + 2 Then $tmp3 = StringInStr($tmp, "-", 0, -1) ; wenn kein Leerzeichen im Vornamen, dann finde das LETZTE Minuszeichen
; 	If $tmp3 > $tmp2 + 2 Then ; das Minus/Leerzeichen liegt nach dem Komma und es ist mindestens ein Zeichen noch dazwischen
; 		$tmp = StringLeft($tmp, $tmp3 - 1) ; dann nimm als Suchbegriff alles bis zum Trennzeichen -> zweiten Vornamen abschneiden
; 	EndIf
; 	$tmp = $tmp & "%"
; 	$tmp = StringReplace($tmp, "ä", "ae")
; 	$tmp = StringReplace($tmp, "ö", "oe")
; 	$tmp = StringReplace($tmp, "ü", "ue")
; 	$tmp = StringReplace($tmp, "Ä", "Ae")
; 	$tmp = StringReplace($tmp, "Ö", "Oe")
; 	$tmp = StringReplace($tmp, "Ü", "Ue")
;
; 	$suche = "SELECT Name, ip_address, mac_address FROM device WHERE info3 LIKE '" & $tmp & "' ORDER BY Name"
; 	$rs.Open($suche, $connNCDaten)
; 	$Text = ""
; 	$row = 0
;
; 	While Not $rs.EOF
; 		GUICtrlSetState($bn_vnc_nc, $GUI_ENABLE)
; 		$anzahl_ncs += 1
; 		$x = StringStripWS($rs.Fields("name").Value, 3)
; 		If $row > _GUICtrlListView_GetItemCount($computer) - 1 Then GUICtrlCreateListViewItem("", $computer)
; 		_GUICtrlListView_AddSubItem($computer, $row, $x, 4)
; 		_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("ip_address").Value, 3), 5)
; 		_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("mac_address").Value, 3), 6)
; 		If Not Ping($x, 50) Then
; 			_GUICtrlListView_AddSubItem($computer, $row, "off", 7)
; 		Else
; 			_GUICtrlListView_AddSubItem($computer, $row, "on", 7)
; 		EndIf
; 		$row += 1
; 		If $row > 9 Then ExitLoop
; 		$rs.MoveNext
; 	WEnd
; 	$rs.Close
;
; 	; zusätzlich suchen wir noch die NCs, die mit dem gleichen Namen anfangen wie der Nachname des Users
; 	$tmp = StringLeft($tmp, StringInStr($tmp, ",") - 1) & "%"
; 	$suche = "SELECT Name, ip_address, mac_address FROM device WHERE Name LIKE '" & $tmp & "' ORDER BY Name"
; 	$rs.Open($suche, $connNCDaten)
; 	$Text = ""
;
; 	$tmp = ""
; 	$tInfo = DllStructCreate($tagLVFINDINFO)
; 	DllStructSetData($tInfo, "Flags", $LVFI_STRING)
;
; 	While Not $rs.EOF
; 		$tmp = StringStripWS($rs.Fields("name").Value, 3)
;
; 		$bool = 0
; 		For $i = 0 To $anzahl_ncs
; 			$x = _GUICtrlListView_GetItem($computer, $i, 4)
; 			$x = $x[3] ; Text rausziehen
; 			If $x = $tmp Then $bool = 1
; 		Next
; 		If $bool = 0 Then
; 			GUICtrlSetState($bn_vnc_nc, $GUI_ENABLE)
; 			$anzahl_ncs += 1
; 			$x = StringStripWS($rs.Fields("name").Value, 3)
; 			If $row > _GUICtrlListView_GetItemCount($computer) - 1 Then GUICtrlCreateListViewItem("", $computer)
; 			_GUICtrlListView_AddSubItem($computer, $row, $tmp, 4)
; 			_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("ip_address").Value, 3), 5)
; 			_GUICtrlListView_AddSubItem($computer, $row, StringStripWS($rs.Fields("mac_address").Value, 3), 6)
; 			If Not Ping($x, 50) Then
; 				_GUICtrlListView_AddSubItem($computer, $row, "off", 7)
; 			Else
; 				_GUICtrlListView_AddSubItem($computer, $row, "on", 7)
; 			EndIf
; 			$row += 1
; 			If $row > 9 Then ExitLoop
; 		EndIf
; 		$rs.MoveNext
; 	WEnd
; 	$rs.Close
;
; 	$connNCDaten.Close
; EndFunc   ;==>NC_Daten
 #CE

#CS Func vnc($art)
; 	$tmp = ""
; 	If $anzahl_ncs = 1 Then
; 		$tmp = _GUICtrlListView_GetItem($computer, 0, 4)
; 		$tmp = $tmp[3]
; 	EndIf
;
; 	; prüfen, ob bereits eine Zeile im Listview angeklickt war = versuche, dieses Gerät zu nehmen
; 	If $tmp = "" Then
; 		$y = GUICtrlRead($computer)
; 		If $y <> 0 Then
; 			$y = GUICtrlRead($y) ; das ListviewItem auslesen
; 			$tmp = StringSplit($y, "|") ; aufteilen und...
; 			$tmp = $tmp[5];  nur die IP des NCs behalten
; 		EndIf
; 	EndIf
;
; 	If $tmp = "" Then ; es gibt mehr als 1 potentielles Gerät zur Auswahl und es war keines angeklickt -> frage nach
; 		$Form4 = GUICreate($texts[$language][106], 350, 250)
; 		$liste = GUICtrlCreateListView($texts[$language][107], 5, 5, 340, 210)
; 		_GUICtrlListView_SetColumnWidth($liste, 0, 230)
; 		$btn_pcwahl = GUICtrlCreateButton("Ok", 5, 220, 340, 25)
;
; 		For $i = 0 To _GUICtrlListView_GetItemCount($computer)
; 			$tmp2 = _GUICtrlListView_GetItem($computer, $i, 4) ; Name des PCs
; 			$tmp3 = _GUICtrlListView_GetItem($computer, $i, 5) ; IP
; 			$temp = $tmp2[3] & "|" & $tmp3[3]
; 			GUICtrlCreateListViewItem($temp, $liste)
; 		Next
;
; 		GUISetState(@SW_SHOW, $Form4)
;
; 		While 1
; 			$msg = GUIGetMsg()
; 			If $msg = $GUI_EVENT_CLOSE Then Exit
; 			If $msg = $btn_pcwahl Then
; 				$x = GUICtrlRead($liste)
; 				If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
; 					$y = GUICtrlRead($liste) ; die Control-ID des markierten Listview-Items holen
; 					$y = GUICtrlRead($y) ; das ListviewItem auslesen
; 					$tmp = StringSplit($y, "|") ; aufteilen und...
; 					$tmp = $tmp[2];  nur die IPbehalten
; 					ExitLoop
; 				EndIf
; 			EndIf
; 		WEnd
; 		GUIDelete($Form4) ; Auswahl-GUI löschen
; 		GUISetState(@SW_SHOW, $Form1) ; Zurück zum Hauptfenster
; 	EndIf
;
; 	; nachdem nun ein eindeutiges Gerät ausgewählt wurde...
; 	$x = 0
; 	If FileExists("C:\Programme\RealVNC\VNC4\vncviewer.exe") Then
; 		$x = Run("C:\Programme\RealVNC\VNC4\vncviewer.exe", @SystemDir)
; 	Else
; 		MsgBox(16, $texts[$language][108], $texts[$language][109])
; 	EndIf
; 	If $x Then ; nur wenn der Prozess erfolgreich gestartet wurde
; 		WinWaitActive("VNC Viewer")
; 		ControlFocus("VNC Viewer", "", "Edit1")
; 		Send($tmp)
; 		Sleep(100)
; 		ControlClick("VNC Viewer", "OK", "Button3")
; 		Sleep(200)
; 		If WinExists("VNC Viewer : Authentication") Then
; 			ControlFocus("VNC Viewer : Authentication", "", "Edit2")
; 			Send("xxxxxx")
; 			Sleep(50)
; 			Send("{ENTER}")
; 		EndIf
; 	EndIf
; EndFunc   ;==>vnc
 #CE

#CS Func Citrixprofile()
; 	$ergebnis = ""
; 	For $i = 1 To 28
;
; 		If StringLen($i) = 1 Then
; 			$j = "0" & $i
; 		Else
; 			$j = $i
; 		EndIf
;
; 		$pfad = "\\cts" & $j & "\c$\Dokumente und Einstellungen\" & $user
; 		If FileExists($pfad & "\.") Then
; 			$x = DirMove($pfad, $pfad & "." & @YEAR & "-" & @MON & "-" & @MDAY, 1)
; 			If $x = 1 Then
; 				$ergebnis = $ergebnis & $pfad & $texts[$language][110] & @CRLF
; 			Else
; 				$ergebnis = $ergebnis & $pfad & $texts[$language][111] & @CRLF
; 			EndIf
; 		Else
; 			;
; 		EndIf
; 	Next
; 	If $ergebnis <> "" Then
; 		MsgBox(0, $texts[$language][112], $ergebnis)
; 	Else
; 		MsgBox(0, $texts[$language][112], $texts[$language][113])
; 	EndIf
; EndFunc   ;==>Citrixprofile
 #CE

Func suche_nachbarn()
	$raum = $oUsr.physicalDeliveryOfficeName
	$raum = StringLeft($raum, StringLen($raum) - 3) & "*"

	; GUI aufbauen
	$Form4 = GUICreate($texts[$language][114], 550, 510)
	$liste = GUICtrlCreateListView($texts[$language][115], 5, 10, 540, 400)
	_GUICtrlListView_SetColumnWidth($liste, 0, 70)
	_GUICtrlListView_SetColumnWidth($liste, 1, 250)
	$btn_nachbarn_vergleich = GUICtrlCreateButton($texts[$language][116], 5, 415, 540, 20)
	$btn_nachbarn_wechseln = GUICtrlCreateButton($texts[$language][117], 5, 440, 540, 20)
	$btn_nachbarn_ok = GUICtrlCreateButton("Ok", 5, 465, 540, 35)

	; suche nach Nachbarn auf dieser Etage
	Dim $treffer_array[1]
	$z = ""
	$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(physicalDeliveryOfficeName=" & $raum & "));ADsPath;subtree"
	$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	Do
		$y = $objRecordSet.Fields(0).Value ; FQDN-Name des Users
		If Not StringInStr($y, "ou=Empfänger") Then ; übergehe alle, die es nur als Mailempfänger gibt (die sind nämlich ou=benutzer,ou=empfänger
			$o_temp_Usr = ObjGet($objRecordSet.Fields(0).Value) ; Retrieve the COM Object for the logged on user
			_ArrayAdd($treffer_array, $o_temp_Usr.physicalDeliveryOfficeName & "|" & $o_temp_Usr.sn & "," & $o_temp_Usr.givenName & "|" & $o_temp_Usr.samAccountName & "|" & $o_temp_Usr.telephoneNumber)
		EndIf
		$objRecordSet.MoveNext
	Until $objRecordSet.EOF
	_ArrayDelete($treffer_array, 0) ; das leere erste Feld löschen
	_ArraySort($treffer_array) ; Treffer sortieren

	$raum = $oUsr.physicalDeliveryOfficeName
	For $i = 0 To UBound($treffer_array) - 1 ; und alle Ergebnisse in Listview kopieren
		$x = GUICtrlCreateListViewItem($treffer_array[$i], $liste)
		$wert = StringSplit($treffer_array[$i], "|")
		If $wert[1] = $raum Then GUICtrlSetBkColor($x, 0xffff00) ; die Raumnummer des Users einfärben
	Next

	GUISetState(@SW_SHOW, $Form4)

	While 1
		$msg = GUIGetMsg()
		If ($msg = $btn_nachbarn_wechseln) Or ($msg = $btn_nachbarn_vergleich) Then
			$x = GUICtrlRead($liste)
			If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
				$y = GUICtrlRead($liste) ; die Control-ID des markierten Listview-Items holen
				$y = GUICtrlRead($y) ; das ListviewItem auslesen
				$tmp = StringSplit($y, "|") ; aufteilen in einzelne Felder
				If $msg = $btn_nachbarn_wechseln Then
					GUIDelete($Form4) ; Auswahl-GUI löschen
					establishUser($tmp[3])
				Else
					GUIDelete($Form4) ; Auswahl-GUI löschen
					compareGroup($tmp[3])
				EndIf
				ExitLoop
			EndIf
		EndIf

		If ($msg = $GUI_EVENT_CLOSE) Or ($msg = $btn_nachbarn_ok) Then
			GUISetState(@SW_SHOW, $Form1)
			GUIDelete($Form4) ; Auswahl-GUI löschen
			ExitLoop ; und raus aus der Schleife
		EndIf
	WEnd
EndFunc   ;==>suche_nachbarn

Func versionsinfo($pc_name)
	If Not Ping($pc_name) Then
		MsgBox(16, $texts[$language][118], $texts[$language][119])
		Return
	EndIf
	$pfad = "\\" & $pc_name & "\c$\"
	If Not FileExists($pfad & ".") Then
		MsgBox(16, $texts[$language][120], $texts[$language][121])
		Return
	EndIf

	$virenscanner = FileGetVersion($pfad & "\Programme\McAfee\VirusScan Enterprise\scan32.exe")
	$ergebnis = "Virenscanner: " & $virenscanner & @CRLF

	$empirum = FileGetVersion($pfad & "\EmpirumAgent\Packages\matrix42\PM2Client\12.0\pm2client.exe")
	If @error Then $empirum = FileGetVersion($pfad & "windows\system32\empirum\swdepot.exe")
	$ergebnis = $ergebnis & "Empirum: " & $empirum & @CRLF

	$office2003 = FileGetVersion($pfad & "\Programme\Microsoft Office\Office11\OUTLOOK.EXE")
	If Not @error Then $ergebnis = $ergebnis & "Office 2003: " & $office2003 & @CRLF

	$office2007 = FileGetVersion($pfad & "\Programme\Microsoft Office\Office12\OUTLOOK.EXE")
	If Not @error Then $ergebnis = $ergebnis & "Office 2007: " & $office2007 & @CRLF

	MsgBox(0, $texts[$language][122] & $pc_name, $ergebnis)
EndFunc   ;==>versionsinfo

Func softwareliste($pc_name)
	If Not Ping($pc_name) Then
		MsgBox(16, $texts[$language][118], $texts[$language][119])
		Return
	EndIf
	; auslesen der Registry Werte
	Dim $programme[1]

	$lauf = 1
	$pfad = "\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	While 1
		$var = RegEnumKey("\\" & $pc_name & $pfad, $lauf)
		If $var = "" Then ExitLoop ; Abbrechen der Schleife wenn Ende der Werte erreicht

		$temp = "\\" & $pc_name & $pfad & "\" & $var ; kompletter Pfad zur GUID
		$display = RegRead($temp, "DisplayName")
		If $display = "" Then $display = $var
		_ArrayAdd($programme, $display)
		$lauf += 1
	WEnd
	ReDim $programme[$lauf]
	_ArraySort($programme)
	_ArrayDisplay($programme, $texts[$language][123] & $pc_name)
EndFunc   ;==>softwareliste

Func remote_desktop($pc_name)
	Run("psservice \\" & $pc_name & " start RemoteRegistry")
	$reg_flag = RegRead("\\" & $pc_name & "\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server", "fDenyTSConnections")
	If $reg_flag Then
		RegWrite("\\" & $pc_name & "\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server", "fDenyTSConnections", "REG_DWORD", 0)
		Sleep(5000)
	EndIf
	Run("mstsc /f /v:" & $pc_name)
EndFunc   ;==>remote_desktop

Func userzaehlen()
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
EndFunc   ;==>userzaehlen

Func suche_dezernat()
	$raum = $oUsr.physicalDeliveryOfficeName
	$dezernat = $oUsr.department

	; GUI aufbauen
	$Form4 = GUICreate($texts[$language][114], 550, 510)
	$liste = GUICtrlCreateListView($texts[$language][115], 5, 10, 540, 400)
	_GUICtrlListView_SetColumnWidth($liste, 0, 70)
	_GUICtrlListView_SetColumnWidth($liste, 1, 250)
	$btn_nachbarn_vergleich = GUICtrlCreateButton($texts[$language][116], 5, 415, 540, 20)
	$btn_nachbarn_wechseln = GUICtrlCreateButton($texts[$language][117], 5, 440, 540, 20)
	$btn_nachbarn_ok = GUICtrlCreateButton("Ok", 5, 465, 540, 35)

	; suche nach Nachbarn auf dieser Etage
	Dim $treffer_array[1]
	$z = ""
	$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(department=" & $dezernat & "));ADsPath;subtree"
	$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	Do
		$y = $objRecordSet.Fields(0).Value ; FQDN-Name des Users
		If Not StringInStr($y, "ou=Empfänger") Then ; übergehe alle, die es nur als Mailempfänger gibt (die sind nämlich ou=benutzer,ou=empfänger
			$o_temp_Usr = ObjGet($objRecordSet.Fields(0).Value) ; Retrieve the COM Object for the logged on user
			_ArrayAdd($treffer_array, $o_temp_Usr.physicalDeliveryOfficeName & "|" & $o_temp_Usr.sn & "," & $o_temp_Usr.givenName & "|" & $o_temp_Usr.samAccountName & "|" & $o_temp_Usr.telephoneNumber)
		EndIf
		$objRecordSet.MoveNext
	Until $objRecordSet.EOF
	_ArrayDelete($treffer_array, 0) ; das leere erste Feld löschen
	_ArraySort($treffer_array) ; Treffer sortieren

	$raum = $oUsr.physicalDeliveryOfficeName
	For $i = 0 To UBound($treffer_array) - 1 ; und alle Ergebnisse in Listview kopieren
		$x = GUICtrlCreateListViewItem($treffer_array[$i], $liste)
		$wert = StringSplit($treffer_array[$i], "|")
		If $wert[1] = $raum Then GUICtrlSetBkColor($x, 0xffff00) ; die Raumnummer des Users einfärben
	Next

	GUISetState(@SW_SHOW, $Form4)

	While 1
		$msg = GUIGetMsg()
		If ($msg = $btn_nachbarn_wechseln) Or ($msg = $btn_nachbarn_vergleich) Then
			$x = GUICtrlRead($liste)
			If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
				$y = GUICtrlRead($liste) ; die Control-ID des markierten Listview-Items holen
				$y = GUICtrlRead($y) ; das ListviewItem auslesen
				$tmp = StringSplit($y, "|") ; aufteilen in einzelne Felder
				If $msg = $btn_nachbarn_wechseln Then
					GUIDelete($Form4) ; Auswahl-GUI löschen
					establishUser($tmp[3])
				Else
					GUIDelete($Form4) ; Auswahl-GUI löschen
					compareGroup($tmp[3])
				EndIf
				ExitLoop
			EndIf
		EndIf

		If ($msg = $GUI_EVENT_CLOSE) Or ($msg = $btn_nachbarn_ok) Then
			GUISetState(@SW_SHOW, $Form1)
			GUIDelete($Form4) ; Auswahl-GUI löschen
			ExitLoop ; und raus aus der Schleife
		EndIf
	WEnd
EndFunc   ;==>suche_dezernat

#CS Func suche_postkorb($suchbegriff)
; 	Return
; 	$connSBDAT.Open("Driver={SQL Server};Server=sqlserv.xxxxxxx.de;Database=X500;") ; mit der Datenbank verbinden
; 	If @error Then
; 		SplashTextOn("", $texts[$language][125], 250, 100)
; 		Sleep(2000)
; 		SplashOff()
; 		Return
; 	EndIf
; 	$suche = "SELECT UserID FROM T_DB2 WHERE Postkorb LIKE '" & $suchbegriff & "'"
; 	$rs.Open($suche, $connSBDAT)
; 	$tmp = $rs.Fields("UserID").Value
; 	$rs.Close
; 	$connSBDAT.Close
; 	Return ($tmp)
; EndFunc   ;==>suche_postkorb
#CE
