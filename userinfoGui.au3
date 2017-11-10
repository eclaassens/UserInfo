#Region -**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ActiveDirectory.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=Performs simple AD queries
#AutoIt3Wrapper_Res_Fileversion=1.0.0.9
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/sf
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/sf
#EndRegion -**** Directives created by AutoIt3Wrapper_GUI ****
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
#include <File.au3>
#include <Math.au3> ; importing the max command
#include <Misc.au3> ; importing _singleton
#include <INet.au3>
#include <String.au3>

Opt("GUICloseOnESC", 0)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
Opt("WinTitleMatchMode", 2)
Opt("TrayIconDebug", 0)
Opt("TrayIconHide", 1)

#Region - User Interface
Global $nMsg

$applicationTitle = "Userinfo " & FileGetVersion(@AutoItExe)
If @AutoItX64 Then
	$applicationTitle &= " [64 Bit]"
Else
	$applicationTitle &= " [32 Bit]"
EndIf

If @Compiled And _Singleton("Userinfo", 1) = 0 Then ; Only start one instance of the application
	WinActivate($applicationTitle)
	Exit
EndIf

#Region - Create main window
$frmMain = GUICreate($applicationTitle, 1000, 700, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MAXIMIZEBOX)
GUISetIcon("shell32.dll", -171)
GUICtrlSetFont($frmMain, 6)

#Region - Create all buttons

GUICtrlCreateGroup("Functions", 845, 5, 150, 700) ;Create Functions group
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)

$btnRefreshQuery = GUICtrlCreateButton("Refresh", 850, 20, 140, 40)
GUICtrlSetImage(-1, "shell32.dll", 255)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Refresh from AD")

$btnNewQuery = GUICtrlCreateButton("New Search", 850, 60, 140, 40)
GUICtrlSetImage(-1, "shell32.dll", 23)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Do another search")

$btnCountDeptUsers = GUICtrlCreateButton("Count Users of Department", 850, 120, 140, 20)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Count the number of users with the same department")

$btnDepartmentUsers = GUICtrlCreateButton("Show Users of Department", 850, 140, 140, 20)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "List all colleagues")

$btnCompareUsers = GUICtrlCreateButton("Compare Users", 850, 160, 140, 20)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Compare group membership with other userId")

$btnCopyGroups = GUICtrlCreateButton("Copy Group List", 850, 180, 140, 20)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Copy the list of groups to clipboard")

$btnGroupMembers = GUICtrlCreateButton("Group members", 850, 200, 140, 20)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Show members of selected group")

$btnExit = GUICtrlCreateButton("Exit", 850, 580, 140, 40)
GUICtrlSetImage(-1, "shell32.dll", 28)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Guess...try and find out!")

GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group
#EndRegion - Create all buttons

#Region - Create all info lables
$lblName = GUICtrlCreateLabel("Name:", 10, 10, 90, 17)
$lblNameValue = GUICtrlCreateLabel("fetching...", 105, 10, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblUserId = GUICtrlCreateLabel("UserID:", 410, 10, 90, 17)
$lblUserIdValue = GUICtrlCreateLabel("fetching...", 505, 10, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblOU = GUICtrlCreateLabel("OU:", 10, 30, 90, 17)
$lblOUValue = GUICtrlCreateLabel("fetching...", 105, 30, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblTelephone = GUICtrlCreateLabel("Telephone:", 410, 30, 90, 17)
$lblTelephoneValue = GUICtrlCreateLabel("fetching...", 505, 30, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblDepartment = GUICtrlCreateLabel("Department:", 10, 50, 90, 17)
$lblDepartmentValue = GUICtrlCreateLabel("fetching...", 105, 50, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblDomain = GUICtrlCreateLabel("Domain:", 410, 50, 90, 17)
$lblDomainValue = GUICtrlCreateLabel("fetching...", 505, 50, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblEmail = GUICtrlCreateLabel("Email:", 10, 70, 90, 17)
$lblEmailValue = GUICtrlCreateLabel("fetching...", 105, 70, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblLocation = GUICtrlCreateLabel("Location:", 410, 70, 90, 17)
$lblLocationValue = GUICtrlCreateLabel("fetching...", 505, 70, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblAddress = GUICtrlCreateLabel("Address:", 10, 90, 90, 17)
$lblAddressValue = GUICtrlCreateLabel("fetching...", 105, 90, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblOffice = GUICtrlCreateLabel("Office:", 410, 90, 90, 17)
$lblOfficeValue = GUICtrlCreateLabel("fetching...", 505, 90, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblLastUpdate = GUICtrlCreateLabel("Last update:", 10, 110, 90, 17)
$lblLastUpdateValue = GUICtrlCreateLabel("fetching...", 105, 110, 125, 17, $SS_SUNKEN)

$lblExpiration = GUICtrlCreateLabel("Valid till:", 105 + 125 + 5, 110, 90 / 2, 17)
$lblExpirationValue = GUICtrlCreateLabel("fetching...", 280, 110, 125, 17, $SS_SUNKEN)

$lblPwRequired = GUICtrlCreateLabel("PW Required:", 410, 110, 90, 17)
$lblPwRequiredValue = GUICtrlCreateLabel("fetching...", 505, 110, 125, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblPWChange = GUICtrlCreateLabel("Change:", 635, 110, 90 / 2, 17)
$lblPWChangeValue = GUICtrlCreateLabel("fetching...", 680, 110, 125, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblAccountLocked = GUICtrlCreateLabel("Locked/Deact:", 10, 130, 90, 17)
$lblAccountLockedValue = GUICtrlCreateLabel("-", 105, 130, 30, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblDeactivated = GUICtrlCreateLabel("/", 185, 130, 10, 17, $SS_CENTER)
$lblDeactivatedValue = GUICtrlCreateLabel("-", 155, 130, 30, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblCreatedDate = GUICtrlCreateLabel("Created on:", 410, 130, 90, 17)
$lblCreatedDateValue = GUICtrlCreateLabel("fetching...", 505, 130, 125, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblBadLoginCount = GUICtrlCreateLabel("Failed count:", 10, 150, 90, 17)
$lblBadLoginCountValue = GUICtrlCreateLabel("fetching...", 105, 150, 30, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblLastLogon = GUICtrlCreateLabel("Last logon:", 410, 150, 90, 17)
$lblLastLogonValue = GUICtrlCreateLabel("fetching...", 505, 150, 125, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblHomeDirectory = GUICtrlCreateLabel("Home directory:", 10, 170, 90, 17)
$lblHomeDirectoryValue = GUICtrlCreateLabel("fetching...", 105, 170, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblHomeDrive = GUICtrlCreateLabel("Home drive:", 410, 170, 90, 17)
$lblHomeDriveValue = GUICtrlCreateLabel("fetching...", 505, 170, 30, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblLogonScript = GUICtrlCreateLabel("Loginscript:", 10, 190, 90, 17)
$lblLogonScriptValue = GUICtrlCreateLabel("fetching...", 105, 190, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)

$lblProfilePath = GUICtrlCreateLabel("Profile Path:", 410, 190, 90, 17)
$lblProfilePathValue = GUICtrlCreateLabel("fetching...", 505, 190, 300, 17, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)
#EndRegion - Create all info lables

#Region - Create Listview
$listGroups = GUICtrlCreateListView("Groupname|samAccountName|Group type|Member of|Scope", 10, 215, 795, 400)
_GUICtrlListView_SetColumnWidth($listGroups, 0, 240)
_GUICtrlListView_SetColumnWidth($listGroups, 1, 225)
#EndRegion - Create Listview

#Region - Create statusbar
$sbMain = _GUICtrlStatusBar_Create($frmMain)
GUICtrlSetResizing(-1, $GUI_DOCKSTATEBAR)
#EndRegion - Create statusbar

#EndRegion - Create main window

#EndRegion - User Interface

#Region - Main program
GUILoop()
#EndRegion - Main program

#Region - Functions

Func GUILoop() ; frmMain GUILoop
	GUISetState(@SW_SHOW, $frmMain)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop

		EndSwitch
	WEnd
EndFunc   ;==>GUILoop
#EndRegion - Functions
