#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ActiveDirectory.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=Performs simple AD queries
#AutoIt3Wrapper_Res_Fileversion=1.0.0.7
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/sf
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/sf
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#AutoIt3Wrapper_run_obfuscator=y
;#Obfuscator_parameters=/striponly /om

If Not @Compiled Then
	Opt("TrayIconDebug", 1)
	Opt("TrayIconHide", 0)
EndIf

#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GuiStatusBar.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ComboConstants.au3>

#include <Array.au3>
#include <Date.au3>
#include <ColorConstants.au3>
#include <IE.au3>
#include <File.au3>
#include <Math.au3> ; importing the max command
#include <Misc.au3> ; importing _singleton
#include <INet.au3>
#include <String.au3>
#include <localization.au3> ; load translations

Opt("GUICloseOnESC", 0)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
Opt("WinTitleMatchMode", 2)
Opt("TrayIconDebug", 0)
Opt("TrayIconHide", 1)

#Region ; Define AD Constants
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
#EndRegion ; Define AD Constants

Global $objConnection = ObjCreate("ADODB.Connection") ; Create COM object to AD
If @error Then
	MsgBox(16, "Error", "Could not create ADODB Connection")
	Exit
EndIf
$objConnection.ConnectionString = "Provider=ADsDSOObject"
$objConnection.Open("Active Directory Provider") ; Open connection to AD

Global $ADSystemInfo = ObjCreate("ADSystemInfo")

#Region ; User Interface
Global $nMsg

$titel = "Userinfo " & FileGetVersion(@AutoItExe)
If @AutoItX64 Then
	$titel &= " [64 Bit]"
Else
	$titel &= " [32 Bit]"
EndIf

If @Compiled And _Singleton("userinfo", 1) = 0 Then ; Only start one instance of the application
	WinActivate($titel)
	Exit
EndIf

Global Const $guiHeight = 700 ; Main GUI Height
Global Const $guiWidth = 1000 ; Main GUI Width

Global Const $labelWidth = 90 ; default label width
Global Const $labelHeight = 17 ; default label height
Global Const $valueWidthLong = 300 ; default value width
Global Const $valueWidthShort = 30 ; default value short width
Global Const $valueWidthDate = 125 ; Width for date fields
Global Const $valueHeight = 17 ; default value height
Global Const $cellPadding = 5 ; default distance between ...

$col1LabelLeft = 10 ; Start of first label
$col1ValueLeft = $col1LabelLeft + $labelWidth + $cellPadding ; Start of first value
$col2LabelLeft = $col1ValueLeft + $valueWidthLong + $cellPadding ; Start of second label
$col2ValueLeft = $col2LabelLeft + $labelWidth + $cellPadding ; Start of second value

Global Const $buttonWidth = 140
Global Const $buttonBigHeight = 40
Global Const $buttonNormalHeight = 20
Global Const $buttonPadding = 5
Global Const $buttonLeft = $guiWidth - ($buttonWidth + (2 * $buttonPadding))

#Region - Create main window
$frmMain = GUICreate($titel, $guiWidth, $guiHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MAXIMIZEBOX)
GUISetIcon("shell32.dll", -171)
GUICtrlSetFont($frmMain, 6)

#Region ; Create all buttons

GUICtrlCreateGroup("Functions", $guiWidth - ($buttonWidth + (3 * $buttonPadding)), $buttonPadding, $buttonWidth + (2 * $buttonPadding), $guiHeight) ;Create Functions group
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)

; for some reason, the "flashing" of labels (when clicked) only works when lables are created after creating the buttons...
$top = 20
$btnRefreshQuery = GUICtrlCreateButton("Refresh", $buttonLeft, $top, $buttonWidth, $buttonBigHeight)
GUICtrlSetImage(-1, "shell32.dll", 255)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Refresh from AD")

$top += $buttonBigHeight
$btnNewQuery = GUICtrlCreateButton("New Search", $buttonLeft, $top, $buttonWidth, $buttonBigHeight)
GUICtrlSetImage(-1, "shell32.dll", 23)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Do another search")

$top += 60
$btnCountDeptUsers = GUICtrlCreateButton("Count Users of Department", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Investigating use...EC")

$top += $buttonNormalHeight
$btnDepartmentUsers = GUICtrlCreateButton("Show Users of Department", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "List all colleagues")

$top += $buttonNormalHeight
$btnCompareUsers = GUICtrlCreateButton("Compare Users", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Compare group membership with other userId")

$top += $buttonNormalHeight
$btnCopyGroups = GUICtrlCreateButton("Copy Group List", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Copy the list of group to clipboard as a single line...?")

$top += $buttonNormalHeight
$btnGroupMembers = GUICtrlCreateButton("Group members", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Show members of selected group")

$btnExit = GUICtrlCreateButton("Exit", $buttonLeft, $guiHeight - 3 * $buttonBigHeight, $buttonWidth, $buttonBigHeight)
GUICtrlSetImage(-1, "shell32.dll", 28)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Guess...try and find out!")

GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group
#EndRegion ; Create all buttons

#Region ; Create all info lables
$top = 10
GUICtrlCreateLabel("Name:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblName = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("UserID:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblUserId = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("OU:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblOU = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Telephone:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblTelephone = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Department:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblDepartment = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Domain:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblDomain = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Email:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblEmail = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Location:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblLocation = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Address:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblAddress = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Office:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblOffice = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Last update:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblLastUpdate = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)

GUICtrlCreateLabel("Valid till:", $col1ValueLeft + $valueWidthDate + $cellPadding, $top, $labelWidth / 2, $labelHeight)
$lblExpiration = GUICtrlCreateLabel("", $col1ValueLeft + $valueWidthDate + ($labelWidth / 2) + $cellPadding, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)

GUICtrlCreateLabel("PW Required:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblPwRequired = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Change:", $col2ValueLeft + $valueWidthDate + $cellPadding, $top, $labelWidth / 2, $labelHeight)
$lblPWChange = GUICtrlCreateLabel("", $col2ValueLeft + $valueWidthDate + ($labelWidth / 2) + $cellPadding, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Locked/Deact:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblLocked = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthShort, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("/", $cellPadding + $col1ValueLeft + $valueWidthShort, $top, 10, $labelHeight, $SS_CENTER)
$lblDeactivated = GUICtrlCreateLabel("-", $col1ValueLeft + $valueWidthShort + (4 * $cellPadding), $top, $valueWidthShort, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Created on:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblCreatedDate = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Failed count:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblBadLoginCount = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthShort, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Last logon:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblLastLogon = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Homedirectory:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblHomeDirectory = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Homedrive:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblHomeDrive = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthShort, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$top += 20
GUICtrlCreateLabel("Loginscript:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblLogonScript = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

GUICtrlCreateLabel("Profile Path:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblProfilePath = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthLong, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)
#EndRegion ; Create all info lables

#Region ; Create Listview
$top += 25
$listGroups = GUICtrlCreateListView("Groupname|samAccountName|Group type|Member of|Scope", $col1LabelLeft, $top, 2 * $labelWidth + 2 * $valueWidthLong + 3 * $cellPadding, 400)
_GUICtrlListView_SetColumnWidth($listGroups, 0, 240)
_GUICtrlListView_SetColumnWidth($listGroups, 1, 225)
#EndRegion ; Create Listview

#Region ; Create statusbar
$sbMain = _GUICtrlStatusBar_Create($frmMain)
GUICtrlSetResizing(-1, $GUI_DOCKSTATEBAR)
#EndRegion ; Create statusbar

#EndRegion - Create main window

#Region - Create Search AD frame
$frmSearchAD = GUICreate($titel, 370, 140, -1, -1)
GUISetIcon("shell32.dll", -171)
GUICtrlCreateLabel("UserID/Name/Tel.:", 5, 8, 100, 17)
$edit = GUICtrlCreateInput(@UserName, 110, 5, 250, 20)
GUICtrlCreateLabel("Wildcards:", 5, 35, 100, 17)
$rdbUserId = GUICtrlCreateRadio("UserID", 110, 35, 70, 20)
$rdbSurname = GUICtrlCreateRadio("Surname", 190, 35, 80, 20)
$rdbFirstName = GUICtrlCreateRadio("First name", 270, 35, 120, 20)
GUICtrlSetState($rdbUserId, $GUI_CHECKED)

$Enter_key = GUICtrlCreateDummy()
Dim $a_AccelKeys[1][2] = [["{ENTER}", $Enter_key]] ; Hotkey array for evaluating the Enter button on frmSearchAD
GUISetAccelerators($a_AccelKeys, $frmSearchAD)


GUICtrlCreateLabel("Domain:", 5, 65, 100, 17)
$cmbDomain = GUICtrlCreateCombo("", 110, 62, 250, 250, $CBS_DROPDOWNLIST)
$btnOk = GUICtrlCreateButton("&Ok", 110, 90, 120, 25)
$btnCancel = GUICtrlCreateButton("E&xit", 240, 90, 120, 25)
$sbSearchAD = _GUICtrlStatusBar_Create($frmSearchAD)
#EndRegion - Create Search AD frame

#EndRegion ; User Interface

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
Global $time ; Zeitstempel fürs Stoppen des Auslesens der Userdaten
Global $cancel ; Flag zum Beenden der Domänensuche
Global $loopflag ; dient zur Kontrolle der Schleife bei der Userauswahl
Global $MaxPasswordAge ; wie lang ein PW in der Domäne gültig ist

#Region ; Main program
listDomains()
setUser()
#EndRegion ; Main program

#Region ; Functions
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

Func _ADGetAccount($user)
	GUICtrlSetData($lblLastUpdate, Zeit($oUsr.PasswordLastChanged))
	If $oUsr.IsAccountLocked Then
		GUICtrlSetData($lblLocked, "Yes")
		GUICtrlSetColor($lblLocked, $color_red)
	Else
		GUICtrlSetData($lblLocked, "No")
		GUICtrlSetColor($lblLocked, $color_green) ;$color_green
	EndIf

	GUICtrlSetData($lblBadLoginCount, $oUsr.BadLoginCount)

	$intUAC = $oUsr.Get("userAccountControl")

	; Check if the user is allowed to change his password. Because no direct query is possible, this is done via the Security Descriptor
	$oSecDesc = $oUsr.Get("ntSecurityDescriptor")
	$oACL = $oSecDesc.DiscretionaryACL
	For $oACE In $oACL
		If ($oACE.ObjectType = $USER_CHANGE_PASSWORD) And (($oACE.Trustee = "Everyone") Or ($oACE.Trustee = "Iedereen")) Then
			If ($oACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT) Then
				GUICtrlSetData($lblPWChange, "Allowed")
				GUICtrlSetColor($lblPWChange, $color_green)
			Else
				GUICtrlSetData($lblPWChange, "Locked")
				GUICtrlSetColor($lblPWChange, $color_red)
			EndIf
		EndIf
	Next

	If BitAND(0x00020, $intUAC) Then ; PW not required flag is set
		GUICtrlSetData($lblPwRequired, "No")
	Else ; not set -> need PW
		GUICtrlSetData($lblPwRequired, "Yes")
	EndIf

	GUICtrlSetData($lblCreatedDate, Zeit($oUsr.whenCreated))
	GUICtrlSetData($lblLogonScript, $oUsr.scriptPath)

	If BitAND($intUAC, $ADS_UF_ACCOUNTDISABLE) Then ; Disabled flag is set
		GUICtrlSetData($lblDeactivated, "Yes")
		GUICtrlSetColor($lblDeactivated, $color_red)
	Else
		GUICtrlSetData($lblDeactivated, "No")
		GUICtrlSetColor($lblDeactivated, $color_green)
	EndIf


	GUICtrlSetColor($lblExpiration, 0x000000)
	$dummy = $oUsr.AccountExpirationDate
	$tmp = Zeit2($dummy) ; Converted time for DateDiff
	$dummy = Zeit($dummy) ; Converted time for output

	$tmp2 = _DateDiff("D", _NowCalcDate(), $tmp)
	If ($tmp2 < 1) And ($tmp2 > -148883) Then GUICtrlSetColor($lblExpiration, $color_red) ; Password is expired

	If ($dummy = "01.01.1601 02:00") Or ($dummy = "01.01.1970 00:00") Or ($dummy = "01.01.1601 01:00") Then
		$dummy = "forever"
		GUICtrlSetColor($lblExpiration, 0x000000)
	EndIf
	GUICtrlSetData($lblExpiration, $dummy)

EndFunc   ;==>_ADGetAccount

Func _ADGetGroupMembers($group)
	If $group = "" Then Return

	_GUICtrlStatusBar_SetText($sbMain, "Fetching members for " & $group & " ...this can take a while")
	SplashTextOn("", "Please wait while minions are rumbling through the AD", "-1", "-1", "-1", "-1", 33, "", "", "")

	$groupdn = _ADSamAccountNameToFQDN($group)
	If $groupdn = "" Then
		MsgBox(16, "Problem", "This is a distribution group and can not be interpreted")
		Return
	EndIf
	$objGroup = ObjGet("LDAP://" & $groupdn)

	Dim $arrayMembers[50000][2]
	$i = 0
	For $objmember In $objGroup.members
		$arrayMembers[$i][0] = StringReplace($objmember.displayname, " (LBV)", "") ; remove useless additional string from the Active Directory
		$arrayMembers[$i][1] = $objmember.samAccountName
		$i = $i + 1
	Next
	ReDim $arrayMembers[$i][2]

	_ArraySort($arrayMembers, 0, 0, 0, 1) ; sort on UserID
	$counter = 0
	$resultSet = ""
	For $strMember = 0 To UBound($arrayMembers) - 1
		$tmp = StringSplit($strMember, ",")
		$resultSet = $resultSet & "<td>" & $arrayMembers[$counter][1] & "<br>(" & $arrayMembers[$counter][0] & ")</td>"
		$counter = $counter + 1
		If Mod($counter, 5) = 0 Then $resultSet = $resultSet & "</tr><tr><th>&nbsp;</th>"
	Next
	$resultSet = StringReplace($resultSet, "\", "")

	$objGroup.GetInfo

	$intGroupType = $objGroup.GroupType

	If $intGroupType And $ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP Then
		$groupScope = "Domain local"
	ElseIf $intGroupType And $ADS_GROUP_TYPE_GLOBAL_GROUP Then
		$groupScope = "Global"
	ElseIf $intGroupType And $ADS_GROUP_TYPE_UNIVERSAL_GROUP Then
		$groupScope = "Universal"
	Else
		$groupScope = "Unknown"
	EndIf

	If $intGroupType And $ADS_GROUP_TYPE_SECURITY_ENABLED Then
		$groupType = "Security group"
	Else
		$groupType = "Distribution group"
	EndIf

	$created = $objGroup.whenCreated
	$created = StringLeft($created, 4) & "-" & StringMid($created, 5, 2) & "-" & StringMid($created, 7, 2) & " " & StringMid($created, 9, 2) & ":" & StringMid($created, 11, 2) & ":" & StringMid($created, 13, 2)

	$changed = $objGroup.whenChanged
	$changed = StringLeft($changed, 4) & "-" & StringMid($changed, 5, 2) & "-" & StringMid($changed, 7, 2) & " " & StringMid($changed, 9, 2) & ":" & StringMid($changed, 11, 2) & ":" & StringMid($changed, 13, 2)

	SplashOff()

	$data = FileOpen(@TempDir & "\groupMembers.html", 2)
	FileWriteLine($data, "<html><head>")
	FileWriteLine($data, "<style>")
	FileWriteLine($data, '  table.table{ font-family: "Trebuchet MS", sans-serif; font-size: 12px; font-weight: bold; line-height: 1.4em; font-style: normal; border-collapse:separate; }')
	FileWriteLine($data, "  .table thead th{ padding:8px 0px; color:#fff; text-align:center; background-color:#9DD929; border: 1px solid #93CE37; border-bottom:3px solid #9ED929; background:-webkit-gradient( linear, left bottom, left top, color-stop(0.02, rgb(123,192,67)), color-stop(0.51, rgb(139,198,66)), color-stop(0.87, rgb(158,217,41)) ); background: -moz-linear-gradient( center bottom, rgb(123,192,67) 2%, rgb(139,198,66) 51%, rgb(158,217,41) 87% ); -webkit-border-top-left-radius:5px; -webkit-border-top-right-radius:5px; -moz-border-radius:5px 5px 0px 0px; border-top-left-radius:5px; border-top-right-radius:5px; text-shadow:1px 1px 1px #568F23;}")
	FileWriteLine($data, "  .table tbody th{ padding:0px 5px; color:#fff; text-align:left;   background-color:#9DD929; border: 1px solid #93CE37; -moz-border-radius:5px 5px 5px 5px; -webkit-border-radius:5px; -webkit-border-radius:5px; border-radius:5px; }")
	FileWriteLine($data, "  .table tbody td{ padding:0px 5px; color:#666; text-align:left;   background-color:#DEF3CA; border: 2px solid #E7EFE0; -moz-border-radius:2px; -webkit-border-radius:2px; border-radius:2px; }")
	FileWriteLine($data, "  .table tfoot td{ padding:2px 0px; color:#666; text-align:center; font-size:10px; }")
	FileWriteLine($data, "</style>")
	FileWriteLine($data, "</head><title>Members of " & $group & "</title><body> <table class=table align=center><tbody>")
	FileWriteLine($data, "<thead>")
	FileWriteLine($data, "  <tr><th colspan=6>" & "Groupname" & ": " & $group & "</th></tr>")
	FileWriteLine($data, "</thead>")
	FileWriteLine($data, "<tbody>")
	FileWriteLine($data, "  <tr><th>" & "Created" & ":</th><td colspan=5>" & $created & "</td></tr>")
	FileWriteLine($data, "  <tr><th>" & "Last changed" & ":</th><td colspan=5>" & $changed & "</td></tr>")
	FileWriteLine($data, "  <tr><th>" & "Mail" & ":</th><td colspan=5>" & $objGroup.mail & "</td></tr>")
	FileWriteLine($data, "  <tr><th>" & "Group info" & ":</th><td colspan=5>" & $objGroup.info & "</td></tr>")
	FileWriteLine($data, "  <tr><th>" & "Description" & ":</th><td colspan=5>" & $objGroup.description & "</td></tr>")
	FileWriteLine($data, "  <tr><th>" & "Scope of the group" & ":</th><td colspan=5>" & $groupScope & "</td></tr>")
	FileWriteLine($data, "  <tr><th>" & "Group class" & ":</th><td colspan=5>" & $groupType & "</td></tr>")
	FileWriteLine($data, "  <tr><th>" & "Member count" & ":</th><td colspan=5>" & UBound($arrayMembers) & "</td></tr>")
	FileWriteLine($data, "  <tr><th>Groupmembers:</th>" & $resultSet & "</tr>")
	FileWriteLine($data, "</tbody>")
	FileWriteLine($data, "</table></body></html>")
	FileClose($data)
	ShellExecute(@TempDir & "\groupMembers.html")

	_GUICtrlStatusBar_SetText($sbMain, "Group membership created.")
	Return
EndFunc   ;==>_ADGetGroupMembers

Func _ADGetUserData($user)
	$oUsr.GetInfo() ; Refresh ADS-Cache
	$tmp = $oUsr.DisplayName
	If $tmp <> "" Then
		GUICtrlSetData($lblName, $oUsr.DisplayName)
	Else
		GUICtrlSetData($lblName, $oUsr.sn & ", " & $oUsr.givenName)
	EndIf

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

	GUICtrlSetData($lblOU, $tmp2)

	GUICtrlSetData($lblEmail, $oUsr.mail)
	GUICtrlSetData($lblTelephone, $oUsr.telephoneNumber)
	GUICtrlSetData($lblDepartment, StringReplace($oUsr.department, "&", "&&"))
	GUICtrlSetData($lblProfilePath, $oUsr.profilePath)

	$lw_m = $oUsr.homeDirectory
	GUICtrlSetData($lblHomeDrive, $oUsr.homeDrive)
	GUICtrlSetColor($lblHomeDirectory, -1)
	GUICtrlSetColor($lblHomeDrive, -1)
	If $lw_m <> "" Then
		GUICtrlSetData($lblHomeDirectory, $lw_m)
	EndIf

	GUICtrlSetData($lblLocation, StringStripWS($oUsr.postalCode & " " & $oUsr.l, 3))
	GUICtrlSetData($lblOffice, $oUsr.physicalDeliveryOfficeName)
	GUICtrlSetData($lblAddress, $oUsr.streetAddress)
EndFunc   ;==>_ADGetUserData

Func _ADGetUserGroups(ByRef $usergroups, $user = @UserName)
	$usergroups = $oUsr.GetEx("memberof")
	$count = UBound($usergroups)
	If $count = 0 Then
		GUICtrlSetData($btnCopyGroups, "Copy Group List") ; reset label of button
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
		$tmp = StringReplace($objGroup.name, "\/", "/") ; Masking undo for smart display
		$tmp = StringReplace($tmp, "CN=", "") & "|" & $objGroup.samAccountName & "|" & $tmp2 & "|" & $tmp3[1] & "|" & $tmp4
		GUICtrlCreateListViewItem($tmp, $listGroups)
	Next
	GUICtrlSetData($btnCopyGroups, "Copy Group List (" & $count & ")")
EndFunc   ;==>_ADGetUserGroups

Func _ADObjectExists($object)
	$flag = 0 ; für Auswertung der Radiobuttons
	$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sAMAccountName=" & $object & "));ADsPath;subtree"
	$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	Switch $objRecordSet.RecordCount
		Case 0 ; User wurde nicht UserID gefunden
			; nix tun, weitere Suchen ablaufen lassen
		Case 1 ; User wurde eindeutig identifiziert anhand des Namens
			GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
			Return 1
		Case Else ; mehrere User wurden über UserID gefunden
			If GUICtrlRead($rdbUserId) = $GUI_CHECKED Then $flag = 1 ; Ergebnis nur behalten wenn suche nach UserID aktiviert
	EndSwitch

	If $flag = 0 Then ; wenn bei der UserID Suche nichts gefunden wurde bzw. Suche nach Name/Vorname gewünscht
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sn=" & $object & "));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
		Switch $objRecordSet.RecordCount
			Case 0 ; User wurde nicht per Nachname gefunden
				; nix tun, weiter prüfen
			Case 1 ; User wurde eindeutig identifiziert anhand des Namens
				GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
				Return 1
			Case Else ; mehrere User wurden über UserID gefunden -> Userliste anzeigen lassen
				If GUICtrlRead($rdbSurname) = $GUI_CHECKED Then $flag = 1 ; Ergebnis nur behalten wenn suche nach Nachname aktiviert
		EndSwitch
	EndIf

	If ($flag = 0) And (GUICtrlRead($rdbFirstName) = $GUI_CHECKED) Then ; Wenn immer noch nichts gefunden, suche nach dem Vornamen
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(givenName=" & $object & "));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery)
	EndIf

	If $objRecordSet.RecordCount = 0 Then ; So far nothing has been found purely -> So we are looking for a phone number
		If StringIsDigit($object) Then $object = "*" & $object
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(telephoneNumber=" & $object & "));ADsPath;subtree" ; also suche nach Telefonnummer
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
	EndIf

	Switch $objRecordSet.RecordCount
		Case 0 ; User not found
			Return 0
		Case 1 ; User found by name
			GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
			Return 1
		Case Else ; Multiple Users found
			Dim $treffer_arry[1]
			$z = ""
			Do
				$y = $objRecordSet.Fields(0).Value ; FQDN-Name des Users
				If Not StringInStr($y, "ou=Recipients") Then ; skip all mail-only accounts (ou=user,ou=Recipients)
					$oUsr = ObjGet($objRecordSet.Fields(0).Value) ; Retrieve the COM Object for the logged on user
					_ArrayAdd($treffer_arry, $oUsr.sn & "," & $oUsr.givenName & "|" & $oUsr.samAccountName)
				EndIf
				$objRecordSet.MoveNext
			Until $objRecordSet.EOF
			If UBound($treffer_arry) = 2 Then ; Array hat nur 1 Element => die anderen Treffer waren reine Mail Recipients und wurden beim Übertragen ins Array übergangen
				$x = StringInStr($treffer_arry[1], "|")
				GUICtrlSetData($edit, StringMid($treffer_arry[1], $x + 1)) ; die Userkennung wird im Suchfenster eingetragen
				GUISetState(@SW_SHOW, $frmSearchAD) ; Fenster zur Benutzerwahl wieder einblenden
				Return -1
			Else ; frmUserSelect - Show this form to select the right user
				GUISetState(@SW_HIDE, $frmSearchAD)
				$frmUserSelect = GUICreate("Choose User", 350, 355)
				GUISetIcon("shell32.dll", -171)
				$Enter_key2 = GUICtrlCreateDummy()
				Dim $b_AccelKeys[1][2] = [["{ENTER}", $Enter_key2]] ; Hotkey-Array für das Auswerten der Enter-Taste in frmUserSelect
				GUISetAccelerators($b_AccelKeys, $frmUserSelect)
				$userList = GUICtrlCreateListView("Name or UserID to search", 5, 40, 340, 280)
				_GUICtrlListView_SetColumnWidth($userList, 0, 250)
				$btn_userwahl = GUICtrlCreateButton("Ok", 5, 325, 340, 25)
				GUICtrlCreateLabel(UBound($treffer_arry) - 1 & " users found. Please select one:", 5, 5, 290, 30)
				_ArrayDelete($treffer_arry, 0) ; das leere erste Feld löschen
				_ArraySort($treffer_arry) ; Treffer sortieren

				For $i = 0 To UBound($treffer_arry) - 1 ; und alle Ergebnisse in Listview kopieren
					GUICtrlCreateListViewItem($treffer_arry[$i], $userList)
				Next
				GUISetState(@SW_SHOW, $frmUserSelect)

				While 1
					$msg = GUIGetMsg()
					If $msg = $GUI_EVENT_CLOSE Then Exit
					If ($msg = $btn_userwahl) Or ($msg = $Enter_key2) Then
						$x = GUICtrlRead($userList)
						If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
							$y = GUICtrlRead($userList) ; die Control-ID des markierten Listview-Items holen
							$y = GUICtrlRead($y) ; das ListviewItem auslesen
							$tmp = StringSplit($y, "|") ; aufteilen und nur die UserID behalten
							$x = GUICtrlSetData($edit, $tmp[2]) ; die Userkennung wird im Suchfenster eingetragen
							ExitLoop
						EndIf
					EndIf
				WEnd
				GUIDelete($frmUserSelect) ; Destroy selection GUI
				GUISetState(@SW_SHOW, $frmSearchAD) ; Display window for user selection again
				Return -1 ; User was selected in form
			EndIf
	EndSwitch
EndFunc   ;==>_ADObjectExists

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

Func compareUsers($compareUser = "")
	Local $strQuery2, $objRecordSet2, $ldap_entry2
	Dim $groupCompare

	If $compareUser = "" Then
		$compareUser = InputBox($texts[$language][92], $texts[$language][93], "", "", 300, 120)
		If @error Then Return
	EndIf

	SplashTextOn("", "Please wait while minions are rumbling through the AD", "-1", "-1", "-1", "-1", 33, "", "", "")

	$strQuery2 = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(sAMAccountName=" & $compareUser & ");ADsPath;subtree"
	$objRecordSet2 = $objConnection.Execute($strQuery2) ; Retrieve the FQDN for the logged on user
	$ldap_entry2 = $objRecordSet2.fields(0).value
	$oUsr2 = ObjGet($ldap_entry2) ; Retrieve the COM Object for the logged on user
	$groupCompare = $oUsr2.GetEx("memberof")
	$oUsr2 = 0
	$count = UBound($groupCompare)
	If $count = 0 Then Return ; fange leere Gruppenlisten ab
	_ArrayInsert($groupCompare, 0, $count)
	For $i = 0 To $count
		$tmp = StringSplit($groupCompare[$i], ",")
		$tmp[1] = StringReplace($tmp[1], "CN=", "")
		$groupCompare[$i] = $tmp[1]
	Next
	_ArraySort($groupCompare, 0, 1)
	$missing1 = ""
	$missing2 = ""
	$commonGroups = ""

	For $i = 0 To _GUICtrlListView_GetItemCount($listGroups) - 1 ; durchlaufe alle Gruppen des geöffneten Users
		$flag = 0 ; per default hat der User die Gruppe, der Vergleichsuser nicht
		$name1 = _GUICtrlListView_GetItem($listGroups, $i, 0)
		$name1 = $name1[3]
		For $name2 In $groupCompare ; durchlaufe alle Gruppen des Vergleichs-Users
			If $name1 = $name2 Then ; wenn der Name der aktuellen Gruppe vom User auch beim Vergleichsuser vorhanden ist
				$flag = 1 ; setze das Flag
				$commonGroups = $commonGroups & $name1 & @CRLF ; und füge den Gruppennamen zur Liste der gemeinsamen Gruppen hinzu
				ExitLoop ; verlasse den Durchlauf der Gruppen des Vergleichs-Users
			EndIf
		Next
		; wenn Gruppen des Vergleichs-Users durchsucht wurden ohne Treffer, füge die Gruppe des Users zu Liste der Gruppen hinzu, die nur der User hat
		If $flag = 0 Then $missing1 = $missing1 & $name1 & @CRLF
	Next

	For $name2 In $groupCompare ; go through all the groups of comparison Users
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

	; soon in the two lists of groups that only 1 user has, the first 2 lines (the number of elements and a blank) away
	;	_ArrayDelete($missing1, 0)
	_ArrayDelete($missing2, 0)
	;	_ArrayDelete($missing1, 0)
	_ArrayDelete($missing2, 0)

	; Find the maximum number of entries for all 3 columns
	$max = 0
	$m1 = UBound($missing1)
	$m2 = UBound($missing2)
	$m3 = UBound($commonGroups)
	$max = _Max($max, $m1)
	$max = _Max($max, $m2)
	$max = _Max($max, $m3)

	SplashOff()

	$data = FileOpen(@TempDir & "\userCompare.html", 2)

	; Construct html file
	; write style
	FileWriteLine($data, "<html><head>")
	FileWriteLine($data, "<style>")
	FileWriteLine($data, '  table.table{ font-family: "Trebuchet MS", sans-serif; font-size: 12px; font-weight: bold; line-height: 1.4em; font-style: normal; border-collapse:separate; }')
	FileWriteLine($data, "  .table thead th{ padding:8px 0px; color:#fff; text-align:center; background-color:#9DD929; border: 1px solid #93CE37; border-bottom:3px; -moz-border-radius:5px 5px 0px 0px; -webkit-border-top-left-radius:5px; -webkit-border-top-right-radius:5px; border-top-left-radius:5px; border-top-right-radius:5px; solid #9ED929; background:-webkit-gradient( linear, left bottom, left top, color-stop(0.02, rgb(123,192,67)), color-stop(0.51, rgb(139,198,66)), color-stop(0.87, rgb(158,217,41)) ); background: -moz-linear-gradient( center bottom, rgb(123,192,67) 2%, rgb(139,198,66) 51%, rgb(158,217,41) 87% ); text-shadow:1px 1px 1px #568F23; }")
	FileWriteLine($data, "  .table tbody th{ padding:0px 5px; color:#fff; text-align:left;   background-color:#9DD929; border: 1px solid #93CE37; border-radius:5px; -moz-border-radius:5px 5px 5px 5px; -webkit-border-radius:5px; -webkit-border-radius:5px; }")
	FileWriteLine($data, "  .table tbody td{ padding:0px 5px; color:#666; text-align:left;   background-color:#DEF3CA; border: 2px solid #E7EFE0; border-radius:2px; -moz-border-radius:2px 0px 0px 0px; -webkit-border-radius:2px; }")
	FileWriteLine($data, "  .table tfoot td{ padding:2px 0px; color:#666; text-align:center; font-size:10px; }")
	FileWriteLine($data, "</style>")
	FileWriteLine($data, "</head>")
	FileWriteLine($data, "<body>")
	FileWriteLine($data, "  <table class=table align=center>")
	; write table header
	FileWriteLine($data, "    <thead>")
	FileWriteLine($data, "      <tr><th>Groups which only " & $user & " has</th> <th>Common Groups</th><th>Groups which only " & $compareUser & " has</th></tr>")
	FileWriteLine($data, "    </thead>")
	FileWriteLine($data, "    <tbody>")
	; write rows
	For $i = 1 To $max
		$s1 = ""
		$s2 = ""
		$s3 = ""
		If $i < $m1 Then $s1 = $missing1[$i]
		If $i < $m2 Then $s2 = $missing2[$i]
		If $i < $m3 Then $s3 = $commonGroups[$i]
		If $s1 = "" And $s2 = "" And $s3 = "" Then ContinueLoop
		FileWriteLine($data, "<tr><td>" & $s1 & '</td><th scope="row">' & $s3 & "</th><td>" & $s2 & "</td></tr>")
	Next
	FileWriteLine($data, "    </tbody>")
	; write table footer
	FileWriteLine($data, "<tfoot> <tr> <td>" & ($m1 / 2) - 1 & "</td> <td>" & ($m3 / 2) - 1 & "</td> <td>" & ($m2 / 2) - 1 & "</td> </tr> </tfoot> ")
	; write end tags
	FileWriteLine($data, "  </table>")
	FileWriteLine($data, "</body></html>")
	FileClose($data)
	ShellExecute(@TempDir & "\userCompare.html")

	_GUICtrlStatusBar_SetText($sbSearchAD, "Comparison of groups completed")
EndFunc   ;==>compareUsers

Func copyGroups() ; copy groups to clipboard
	For $i = 0 To UBound($groups) - 1
		$tmp = StringSplit($groups[$i], ",")
		$tmp[1] = StringReplace($tmp[1], "CN=", "")
		$groups[$i] = $tmp[1]
	Next
	_ArrayToClip($groups, @CRLF, 1)
EndFunc   ;==>copyGroups

Func copyToClipboard($clickValue)
	If $clickValue <> "" Then ; only if a text has been detected
		While StringLeft($clickValue, 1) = "|" ; strip leading hyphens
			$clickValue = StringMid($clickValue, 2)
		WEnd
		$clickValue = StringReplace($clickValue, "|", @CRLF)
		$clickValue = StringReplace($clickValue, @CRLF & @CRLF, "")
		ClipPut($clickValue) ; copy to clipboard
		_GUICtrlStatusBar_SetText($sbMain, $clickValue & " has been copied to the clipboard") ; update status bar
		$x = ControlGetFocus($titel) ; find out what was clicked
		If Not StringInStr($x, "SysListView32") Then ; If not something was clicked in the group list
			GUICtrlSetBkColor($nMsg, $color_red) ; color the entry red
			Sleep(200) ; wait a quarter of a second
			GUICtrlSetBkColor($nMsg, -1) ; Revert color of text
		EndIf
	EndIf
EndFunc   ;==>copyToClipboard

Func countUsers()
	$z = ""
	$total = 0
	For $i = 1 To 6
		$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(department=*0" & $i & "*));ADsPath;subtree"
		$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists
		$z = $z & "Department " & $i & ": " & $objRecordSet.RecordCount & @CRLF
		$total += $objRecordSet.RecordCount
	Next
	$z &= "===========================" & @CRLF & "Total: " & $total
	MsgBox(0, "Usercount per department ", $z)
EndFunc   ;==>countUsers

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
				Return 0
			Case 0 ; User not found
				_GUICtrlStatusBar_SetText($sbSearchAD, "The account " & $user & " does not exist in this domain.")
				$loopflag = 1
				Return 0
		EndSwitch
	EndIf
EndFunc   ;==>domainSearch

Func GUILoop() ; frmMain GUILoop
	While 1
		Global $nMsg
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $btnExit
				Exit
			Case $btnCopyGroups
				copyGroups()
			Case $btnGroupMembers
				$x = GUICtrlRead($listGroups)
				If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
					$y = GUICtrlRead($x) ; Read ListviewItem
					$tmp = StringSplit($y, "|") ; strip everything but groupname
					_ADGetGroupMembers($tmp[2])
				Else
					; Show error on screen
					If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
					$iMsgBoxAnswer = MsgBox(8192, "", "Please select a group from the list to retrieve the members", 10)
					Select
						Case $iMsgBoxAnswer = -1 ;Timeout
						Case Else ;OK
					EndSelect
				EndIf
			Case $btnCompareUsers
				compareUsers()
			Case $btnRefreshQuery
				$time = TimerInit() ; reset timer
				queryAD()
			Case $btnNewQuery
				setUser()
			Case $edit ; otherwise the text from the search edit field is copied to the clipboard when a hit on Wildcard search was carried out
				;
			Case $btnDepartmentUsers
				searchDepartmentUsers()
			Case $btnCountDeptUsers
				countUsers()
			Case -100 To 0
				;
			Case Else
				$clickValue = GUICtrlRead($nMsg)
				copyToClipboard($clickValue)
		EndSwitch
		;Sleep(50)
	WEnd
EndFunc   ;==>GUILoop

Func listDomains() ; discover all Domains in forest and select ours
	$oDS = ObjGet("WinNT:")
	$y = ""
	For $x In $oDS
		$y = $y & "|" & $x.Name
	Next
	GUICtrlSetData($cmbDomain, $y, "US")
EndFunc   ;==>listDomains

Func passwort_req($user)
	$oUsr.Put("userAccountControl", BitXOR($intUAC, 0x00020))
	$oUsr.SetInfo

	$oUsr.GetInfo
	$intUAC = $oUsr.Get("userAccountControl")
	If BitAND(0x00020, $intUAC) Then ; PW not required flag is set
		GUICtrlSetData($lblPwRequired, "No")
		GUICtrlSetColor($lblPwRequired, $color_red)
	Else ; not set -> Password required
		GUICtrlSetData($lblPwRequired, "Yes")
		GUICtrlSetColor($lblPwRequired, $color_green)
	EndIf
EndFunc   ;==>passwort_req

Func queryAD()
	$pc = ""
	$lastlogon = "-"
	GUICtrlSetData($lblHomeDrive, "")
	GUICtrlSetData($lblHomeDirectory, "")
	GUICtrlSetData($lblLogonScript, "")
	_GUICtrlListView_DeleteAllItems($listGroups)

	_ADGetUserData($user)
	_ADGetUserGroups($groups, $user)
	_ADGetAccount($user)

	$tmp = $oUsr.scriptpath

	If $lastlogon = "-" Then ; Domänen LastLogin auswerten
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

	$tmp = TimerDiff($time) / 1000
	$tmp2 = StringFormat("%.2f", $tmp)

	_GUICtrlStatusBar_SetText($sbSearchAD, "Userdata read (in " & $tmp2 & " seconds, with " & $tmp1 & " from ADS), Domaincontroller: " & $strHostServer)
EndFunc   ;==>queryAD

Func searchDepartmentUsers()
	$office = $oUsr.physicalDeliveryOfficeName
	$department = $oUsr.department

	; Create GUI
	Local $spacer = 5
	Local $frmWidth = 650
	Local $frmHeight = 550
	Local $btnWidth = $frmWidth - (2 * $spacer)
	Local $btnHeight = 30
	Local $lvWidth = $frmWidth - (2 * $spacer)
	Local $lvHeight = 400

	$frmColleagues = GUICreate("Colleagues", $frmWidth, $frmHeight)
	GUISetIcon("shell32.dll", -171)
	$userList = GUICtrlCreateListView("Username|UserID|Telephone|Office", $spacer, $spacer, $btnWidth, $lvHeight, $LVS_SORTASCENDING + $LVS_SINGLESEL + $LVS_SHOWSELALWAYS)
	_GUICtrlListView_SetColumnWidth($userList, 0, 140)
	_GUICtrlListView_SetColumnWidth($userList, 1, $LVSCW_AUTOSIZE)
	$btnColleagueCompare = GUICtrlCreateButton("Compare groups with this colleague", $spacer, ($frmHeight + 2*$spacer) - (3 * $btnHeight), $btnWidth, $btnHeight)
	$btnColleagueSwitch = GUICtrlCreateButton("Switch to this colleague", $spacer, ($frmHeight + 2*$spacer) - (2 * $btnHeight), $btnWidth, $btnHeight)
	$btnColleaguesOk = GUICtrlCreateButton("Ok", $spacer, ($frmHeight + 2* $spacer) - (1 * $btnHeight),
$btnWidth, $btnHeight)


	; looking for neighbors on this floor
	Dim $arrayFound[1]
	$z = ""
	$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(department=" & $department & "));ADsPath;subtree"
	$objRecordSet = $objConnection.Execute($strQuery) ; Retrieve the FQDN if it exists

	;Fill listView with users where ou<>Recipients
	Do
		$y = $objRecordSet.Fields(0).Value ; FQDN-Name Users
		If Not StringInStr($y, "ou=Recipients") Then
			$o_temp_Usr = ObjGet($objRecordSet.Fields(0).Value) ; Retrieve the COM Object for all users in search
			;_ArrayAdd($arrayFound, '"' & $o_temp_Usr.physicalDeliveryOfficeName & "|" & $o_temp_Usr.sn & "," & $o_temp_Usr.givenName & "|" & $o_temp_Usr.samAccountName & "|" & $o_temp_Usr.telephoneNumber & '"')
			$x = GUICtrlCreateListViewItem($o_temp_Usr.sn & "," & $o_temp_Usr.givenName & "|" & $o_temp_Usr.samAccountName & "|" & $o_temp_Usr.telephoneNumber & "|" & $o_temp_Usr.physicalDeliveryOfficeName, $userList)
		EndIf
		$objRecordSet.MoveNext
	Until $objRecordSet.EOF
	;_ArrayDelete($arrayFound, 0) ; delete the empty first field
	;_ArraySort($arrayFound) ; Sort hits

;~ 	For $i = 0 To UBound($arrayFound) - 1 ; copy all results in ListView
;~ 		$x = GUICtrlCreateListViewItem($arrayFound[$i], $userList)
;~ 		;$wert = StringSplit($arrayFound[$i], "|")
;~ 		;If $wert[1] = $office Then GUICtrlSetBkColor($x, 0xffff00) ; color the room number of the user
;~ 	Next

	GUISetState(@SW_SHOW, $frmColleagues)

	While 1
		$msg = GUIGetMsg()
		If ($msg = $btnColleagueSwitch) Or ($msg = $btnColleagueCompare) Then
			$x = GUICtrlRead($userList)
			If $x <> "" Then ; nur wenn auch ein User gewählt wurde...
				$y = GUICtrlRead($userList) ; die Control-ID des markierten Listview-Items holen
				$y = GUICtrlRead($y) ; das ListviewItem auslesen
				$tmp = StringSplit($y, "|") ; aufteilen in einzelne Felder
				If $msg = $btnColleagueSwitch Then
					GUIDelete($frmColleagues) ; Destroy GUI
					setUser($tmp[3])
				Else
					GUIDelete($frmColleagues) ; Destroy GUI
					compareUsers($tmp[3])
				EndIf
				ExitLoop
			EndIf
		EndIf

		If ($msg = $GUI_EVENT_CLOSE) Or ($msg = $btnColleaguesOk) Then
			GUISetState(@SW_SHOW, $frmMain)
			GUIDelete($frmColleagues) ; Destroy GUI
			ExitLoop
		EndIf
	WEnd
EndFunc   ;==>searchDepartmentUsers

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

Func Zeit($zeit) ; Umwandeln der AD-Zeit ins deutsche Format
	If $zeit = "" Then Return "---" ; User wurde grade frisch zurückgesetzt
	$tmp = StringMid($zeit, 7, 2) & "." & StringMid($zeit, 5, 2) & "." & StringLeft($zeit, 4) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2)
	Return $tmp
EndFunc   ;==>Zeit

Func Zeit2($zeit) ; Umwandeln der AD-Zeit ins englische Format für _DateAdd
	$tmp = StringLeft($zeit, 4) & "/" & StringMid($zeit, 5, 2) & "/" & StringMid($zeit, 7, 2) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2) & ":" & StringMid($zeit, 13, 2)
	Return $tmp
EndFunc   ;==>Zeit2

Func Zeit3($zeit) ; Converting English date format YYYY MM DD to German
	If $zeit = "0" Then Return "---" ; User has reset
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
#EndRegion ; Functions
