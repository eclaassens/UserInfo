If Not @Compiled Then
Opt("TrayIconDebug", 1)
Opt("TrayIconHide", 0)
EndIf
Global Const $CBS_DROPDOWNLIST = 0x3
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_CHECKED = 1
Global Const $GUI_DOCKAUTO = 0x0001
Global Const $GUI_DOCKRIGHT = 0x0004
Global Const $GUI_DOCKTOP = 0x0020
Global Const $GUI_DOCKBOTTOM = 0x0040
Global Const $GUI_DOCKWIDTH = 0x0100
Global Const $GUI_DOCKSTATEBAR = 0x0240
Global Const $LVIF_GROUPID = 0x00000100
Global Const $LVIF_IMAGE = 0x00000002
Global Const $LVIF_INDENT = 0x00000010
Global Const $LVIF_PARAM = 0x00000004
Global Const $LVIF_STATE = 0x00000008
Global Const $LVIS_CUT = 0x0004
Global Const $LVIS_DROPHILITED = 0x0008
Global Const $LVIS_FOCUSED = 0x0001
Global Const $LVIS_OVERLAYMASK = 0x0F00
Global Const $LVIS_SELECTED = 0x0002
Global Const $LVIS_STATEIMAGEMASK = 0xF000
Global Const $LVS_SHOWSELALWAYS = 0x0008
Global Const $LVS_SINGLESEL = 0x0004
Global Const $LVS_SORTASCENDING = 0x0010
Global Const $LVM_FIRST = 0x1000
Global Const $LVM_DELETEALLITEMS =($LVM_FIRST + 9)
Global Const $LVM_GETITEMA =($LVM_FIRST + 5)
Global Const $LVM_GETITEMW =($LVM_FIRST + 75)
Global Const $LVM_GETITEMCOUNT =($LVM_FIRST + 4)
Global Const $LVM_GETITEMTEXTA =($LVM_FIRST + 45)
Global Const $LVM_GETITEMTEXTW =($LVM_FIRST + 115)
Global Const $LVM_GETUNICODEFORMAT = 0x2000 + 6
Global Const $LVM_SETCOLUMNWIDTH =($LVM_FIRST + 30)
Global Const $LVSCW_AUTOSIZE = -1
Global Const $SS_CENTER = 0x1
Global Const $SS_SUNKEN = 0x1000
Global Const $WS_MAXIMIZEBOX = 0x00010000
Global Const $WS_SIZEBOX = 0x00040000
Global Const $WS_SYSMENU = 0x00080000
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $STR_ENTIRESPLIT = 1
Global Const $STR_NOCOUNT = 2
Global Enum $ARRAYFILL_FORCE_DEFAULT, $ARRAYFILL_FORCE_SINGLEITEM, $ARRAYFILL_FORCE_INT, $ARRAYFILL_FORCE_NUMBER, $ARRAYFILL_FORCE_PTR, $ARRAYFILL_FORCE_HWND, $ARRAYFILL_FORCE_STRING
Func _ArrayAdd(ByRef $aArray, $vValue, $iStart = 0, $sDelim_Item = "|", $sDelim_Row = @CRLF, $iForce = $ARRAYFILL_FORCE_DEFAULT)
If $iStart = Default Then $iStart = 0
If $sDelim_Item = Default Then $sDelim_Item = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iForce = Default Then $iForce = $ARRAYFILL_FORCE_DEFAULT
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS)
Local $hDataType = 0
Switch $iForce
Case $ARRAYFILL_FORCE_INT
$hDataType = Int
Case $ARRAYFILL_FORCE_NUMBER
$hDataType = Number
Case $ARRAYFILL_FORCE_PTR
$hDataType = Ptr
Case $ARRAYFILL_FORCE_HWND
$hDataType = Hwnd
Case $ARRAYFILL_FORCE_STRING
$hDataType = String
EndSwitch
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iForce = $ARRAYFILL_FORCE_SINGLEITEM Then
ReDim $aArray[$iDim_1 + 1]
$aArray[$iDim_1] = $vValue
Return $iDim_1
EndIf
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(5, 0, -1)
$hDataType = 0
Else
Local $aTmp = StringSplit($vValue, $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
If UBound($aTmp, $UBOUND_ROWS) = 1 Then
$aTmp[0] = $vValue
EndIf
$vValue = $aTmp
EndIf
Local $iAdd = UBound($vValue, $UBOUND_ROWS)
ReDim $aArray[$iDim_1 + $iAdd]
For $i = 0 To $iAdd - 1
If IsFunc($hDataType) Then
$aArray[$iDim_1 + $i] = $hDataType($vValue[$i])
Else
$aArray[$iDim_1 + $i] = $vValue[$i]
EndIf
Next
Return $iDim_1 + $iAdd - 1
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iStart < 0 Or $iStart > $iDim_2 - 1 Then Return SetError(4, 0, -1)
Local $iValDim_1, $iValDim_2 = 0, $iColCount
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 2 Then Return SetError(5, 0, -1)
$iValDim_1 = UBound($vValue, $UBOUND_ROWS)
$iValDim_2 = UBound($vValue, $UBOUND_COLUMNS)
$hDataType = 0
Else
Local $aSplit_1 = StringSplit($vValue, $sDelim_Row, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iValDim_1 = UBound($aSplit_1, $UBOUND_ROWS)
Local $aTmp[$iValDim_1][0], $aSplit_2
For $i = 0 To $iValDim_1 - 1
$aSplit_2 = StringSplit($aSplit_1[$i], $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iColCount = UBound($aSplit_2)
If $iColCount > $iValDim_2 Then
$iValDim_2 = $iColCount
ReDim $aTmp[$iValDim_1][$iValDim_2]
EndIf
For $j = 0 To $iColCount - 1
$aTmp[$i][$j] = $aSplit_2[$j]
Next
Next
$vValue = $aTmp
EndIf
If UBound($vValue, $UBOUND_COLUMNS) + $iStart > UBound($aArray, $UBOUND_COLUMNS) Then Return SetError(3, 0, -1)
ReDim $aArray[$iDim_1 + $iValDim_1][$iDim_2]
For $iWriteTo_Index = 0 To $iValDim_1 - 1
For $j = 0 To $iDim_2 - 1
If $j < $iStart Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = ""
ElseIf $j - $iStart > $iValDim_2 - 1 Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = ""
Else
If IsFunc($hDataType) Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = $hDataType($vValue[$iWriteTo_Index][$j - $iStart])
Else
$aArray[$iWriteTo_Index + $iDim_1][$j] = $vValue[$iWriteTo_Index][$j - $iStart]
EndIf
EndIf
Next
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS) - 1
EndFunc
Func _ArrayDelete(ByRef $aArray, $vRange)
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If IsArray($vRange) Then
If UBound($vRange, $UBOUND_DIMENSIONS) <> 1 Or UBound($vRange, $UBOUND_ROWS) < 2 Then Return SetError(4, 0, -1)
Else
Local $iNumber, $aSplit_1, $aSplit_2
$vRange = StringStripWS($vRange, 8)
$aSplit_1 = StringSplit($vRange, ";")
$vRange = ""
For $i = 1 To $aSplit_1[0]
If Not StringRegExp($aSplit_1[$i], "^\d+(-\d+)?$") Then Return SetError(3, 0, -1)
$aSplit_2 = StringSplit($aSplit_1[$i], "-")
Switch $aSplit_2[0]
Case 1
$vRange &= $aSplit_2[1] & ";"
Case 2
If Number($aSplit_2[2]) >= Number($aSplit_2[1]) Then
$iNumber = $aSplit_2[1] - 1
Do
$iNumber += 1
$vRange &= $iNumber & ";"
Until $iNumber = $aSplit_2[2]
EndIf
EndSwitch
Next
$vRange = StringSplit(StringTrimRight($vRange, 1), ";")
EndIf
If $vRange[1] < 0 Or $vRange[$vRange[0]] > $iDim_1 Then Return SetError(5, 0, -1)
Local $iCopyTo_Index = 0
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
For $i = 1 To $vRange[0]
$aArray[$vRange[$i]] = ChrW(0xFAB1)
Next
For $iReadFrom_Index = 0 To $iDim_1
If $aArray[$iReadFrom_Index] == ChrW(0xFAB1) Then
ContinueLoop
Else
If $iReadFrom_Index <> $iCopyTo_Index Then
$aArray[$iCopyTo_Index] = $aArray[$iReadFrom_Index]
EndIf
$iCopyTo_Index += 1
EndIf
Next
ReDim $aArray[$iDim_1 - $vRange[0] + 1]
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
For $i = 1 To $vRange[0]
$aArray[$vRange[$i]][0] = ChrW(0xFAB1)
Next
For $iReadFrom_Index = 0 To $iDim_1
If $aArray[$iReadFrom_Index][0] == ChrW(0xFAB1) Then
ContinueLoop
Else
If $iReadFrom_Index <> $iCopyTo_Index Then
For $j = 0 To $iDim_2
$aArray[$iCopyTo_Index][$j] = $aArray[$iReadFrom_Index][$j]
Next
EndIf
$iCopyTo_Index += 1
EndIf
Next
ReDim $aArray[$iDim_1 - $vRange[0] + 1][$iDim_2 + 1]
Case Else
Return SetError(2, 0, False)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS)
EndFunc
Func _ArrayInsert(ByRef $aArray, $vRange, $vValue = "", $iStart = 0, $sDelim_Item = "|", $sDelim_Row = @CRLF, $iForce = $ARRAYFILL_FORCE_DEFAULT)
If $vValue = Default Then $vValue = ""
If $iStart = Default Then $iStart = 0
If $sDelim_Item = Default Then $sDelim_Item = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iForce = Default Then $iForce = $ARRAYFILL_FORCE_DEFAULT
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
Local $hDataType = 0
Switch $iForce
Case $ARRAYFILL_FORCE_INT
$hDataType = Int
Case $ARRAYFILL_FORCE_NUMBER
$hDataType = Number
Case $ARRAYFILL_FORCE_PTR
$hDataType = Ptr
Case $ARRAYFILL_FORCE_HWND
$hDataType = Hwnd
Case $ARRAYFILL_FORCE_STRING
$hDataType = String
EndSwitch
Local $aSplit_1, $aSplit_2
If IsArray($vRange) Then
If UBound($vRange, $UBOUND_DIMENSIONS) <> 1 Or UBound($vRange, $UBOUND_ROWS) < 2 Then Return SetError(4, 0, -1)
Else
Local $iNumber
$vRange = StringStripWS($vRange, 8)
$aSplit_1 = StringSplit($vRange, ";")
$vRange = ""
For $i = 1 To $aSplit_1[0]
If Not StringRegExp($aSplit_1[$i], "^\d+(-\d+)?$") Then Return SetError(3, 0, -1)
$aSplit_2 = StringSplit($aSplit_1[$i], "-")
Switch $aSplit_2[0]
Case 1
$vRange &= $aSplit_2[1] & ";"
Case 2
If Number($aSplit_2[2]) >= Number($aSplit_2[1]) Then
$iNumber = $aSplit_2[1] - 1
Do
$iNumber += 1
$vRange &= $iNumber & ";"
Until $iNumber = $aSplit_2[2]
EndIf
EndSwitch
Next
$vRange = StringSplit(StringTrimRight($vRange, 1), ";")
EndIf
If $vRange[1] < 0 Or $vRange[$vRange[0]] > $iDim_1 Then Return SetError(5, 0, -1)
For $i = 2 To $vRange[0]
If $vRange[$i] < $vRange[$i - 1] Then Return SetError(3, 0, -1)
Next
Local $iCopyTo_Index = $iDim_1 + $vRange[0]
Local $iInsertPoint_Index = $vRange[0]
Local $iInsert_Index = $vRange[$iInsertPoint_Index]
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iForce = $ARRAYFILL_FORCE_SINGLEITEM Then
ReDim $aArray[$iDim_1 + $vRange[0] + 1]
For $iReadFromIndex = $iDim_1 To 0 Step -1
$aArray[$iCopyTo_Index] = $aArray[$iReadFromIndex]
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
$aArray[$iCopyTo_Index] = $vValue
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index < 1 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Return $iDim_1 + $vRange[0] + 1
EndIf
ReDim $aArray[$iDim_1 + $vRange[0] + 1]
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(5, 0, -1)
$hDataType = 0
Else
Local $aTmp = StringSplit($vValue, $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
If UBound($aTmp, $UBOUND_ROWS) = 1 Then
$aTmp[0] = $vValue
$hDataType = 0
EndIf
$vValue = $aTmp
EndIf
For $iReadFromIndex = $iDim_1 To 0 Step -1
$aArray[$iCopyTo_Index] = $aArray[$iReadFromIndex]
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
If $iInsertPoint_Index <= UBound($vValue, $UBOUND_ROWS) Then
If IsFunc($hDataType) Then
$aArray[$iCopyTo_Index] = $hDataType($vValue[$iInsertPoint_Index - 1])
Else
$aArray[$iCopyTo_Index] = $vValue[$iInsertPoint_Index - 1]
EndIf
Else
$aArray[$iCopyTo_Index] = ""
EndIf
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index = 0 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iStart < 0 Or $iStart > $iDim_2 - 1 Then Return SetError(6, 0, -1)
Local $iValDim_1, $iValDim_2
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 2 Then Return SetError(7, 0, -1)
$iValDim_1 = UBound($vValue, $UBOUND_ROWS)
$iValDim_2 = UBound($vValue, $UBOUND_COLUMNS)
$hDataType = 0
Else
$aSplit_1 = StringSplit($vValue, $sDelim_Row, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iValDim_1 = UBound($aSplit_1, $UBOUND_ROWS)
StringReplace($aSplit_1[0], $sDelim_Item, "")
$iValDim_2 = @extended + 1
Local $aTmp[$iValDim_1][$iValDim_2]
For $i = 0 To $iValDim_1 - 1
$aSplit_2 = StringSplit($aSplit_1[$i], $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
For $j = 0 To $iValDim_2 - 1
$aTmp[$i][$j] = $aSplit_2[$j]
Next
Next
$vValue = $aTmp
EndIf
If UBound($vValue, $UBOUND_COLUMNS) + $iStart > UBound($aArray, $UBOUND_COLUMNS) Then Return SetError(8, 0, -1)
ReDim $aArray[$iDim_1 + $vRange[0] + 1][$iDim_2]
For $iReadFromIndex = $iDim_1 To 0 Step -1
For $j = 0 To $iDim_2 - 1
$aArray[$iCopyTo_Index][$j] = $aArray[$iReadFromIndex][$j]
Next
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
For $j = 0 To $iDim_2 - 1
If $j < $iStart Then
$aArray[$iCopyTo_Index][$j] = ""
ElseIf $j - $iStart > $iValDim_2 - 1 Then
$aArray[$iCopyTo_Index][$j] = ""
Else
If $iInsertPoint_Index - 1 < $iValDim_1 Then
If IsFunc($hDataType) Then
$aArray[$iCopyTo_Index][$j] = $hDataType($vValue[$iInsertPoint_Index - 1][$j - $iStart])
Else
$aArray[$iCopyTo_Index][$j] = $vValue[$iInsertPoint_Index - 1][$j - $iStart]
EndIf
Else
$aArray[$iCopyTo_Index][$j] = ""
EndIf
EndIf
Next
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index = 0 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS)
EndFunc
Func _ArrayReverse(ByRef $aArray, $iStart = 0, $iEnd = 0)
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If Not IsArray($aArray) Then Return SetError(1, 0, 0)
If UBound($aArray, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(3, 0, 0)
If Not UBound($aArray) Then Return SetError(4, 0, 0)
Local $vTmp, $iUBound = UBound($aArray) - 1
If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(2, 0, 0)
For $i = $iStart To Int(($iStart + $iEnd - 1) / 2)
$vTmp = $aArray[$i]
$aArray[$i] = $aArray[$iEnd]
$aArray[$iEnd] = $vTmp
$iEnd -= 1
Next
Return 1
EndFunc
Func _ArraySort(ByRef $aArray, $iDescending = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0, $iPivot = 0)
If $iDescending = Default Then $iDescending = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
If $iPivot = Default Then $iPivot = 0
If Not IsArray($aArray) Then Return SetError(1, 0, 0)
Local $iUBound = UBound($aArray) - 1
If $iUBound = -1 Then Return SetError(5, 0, 0)
If $iEnd = Default Then $iEnd = 0
If $iEnd < 1 Or $iEnd > $iUBound Or $iEnd = Default Then $iEnd = $iUBound
If $iStart < 0 Or $iStart = Default Then $iStart = 0
If $iStart > $iEnd Then Return SetError(2, 0, 0)
If $iDescending = Default Then $iDescending = 0
If $iPivot = Default Then $iPivot = 0
If $iSubItem = Default Then $iSubItem = 0
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iPivot Then
__ArrayDualPivotSort($aArray, $iStart, $iEnd)
Else
__ArrayQuickSort1D($aArray, $iStart, $iEnd)
EndIf
If $iDescending Then _ArrayReverse($aArray, $iStart, $iEnd)
Case 2
If $iPivot Then Return SetError(6, 0, 0)
Local $iSubMax = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iSubItem > $iSubMax Then Return SetError(3, 0, 0)
If $iDescending Then
$iDescending = -1
Else
$iDescending = 1
EndIf
__ArrayQuickSort2D($aArray, $iDescending, $iStart, $iEnd, $iSubItem, $iSubMax)
Case Else
Return SetError(4, 0, 0)
EndSwitch
Return 1
EndFunc
Func __ArrayQuickSort1D(ByRef $aArray, Const ByRef $iStart, Const ByRef $iEnd)
If $iEnd <= $iStart Then Return
Local $vTmp
If($iEnd - $iStart) < 15 Then
Local $vCur
For $i = $iStart + 1 To $iEnd
$vTmp = $aArray[$i]
If IsNumber($vTmp) Then
For $j = $i - 1 To $iStart Step -1
$vCur = $aArray[$j]
If($vTmp >= $vCur And IsNumber($vCur)) Or(Not IsNumber($vCur) And StringCompare($vTmp, $vCur) >= 0) Then ExitLoop
$aArray[$j + 1] = $vCur
Next
Else
For $j = $i - 1 To $iStart Step -1
If(StringCompare($vTmp, $aArray[$j]) >= 0) Then ExitLoop
$aArray[$j + 1] = $aArray[$j]
Next
EndIf
$aArray[$j + 1] = $vTmp
Next
Return
EndIf
Local $L = $iStart, $R = $iEnd, $vPivot = $aArray[Int(($iStart + $iEnd) / 2)], $bNum = IsNumber($vPivot)
Do
If $bNum Then
While($aArray[$L] < $vPivot And IsNumber($aArray[$L])) Or(Not IsNumber($aArray[$L]) And StringCompare($aArray[$L], $vPivot) < 0)
$L += 1
WEnd
While($aArray[$R] > $vPivot And IsNumber($aArray[$R])) Or(Not IsNumber($aArray[$R]) And StringCompare($aArray[$R], $vPivot) > 0)
$R -= 1
WEnd
Else
While(StringCompare($aArray[$L], $vPivot) < 0)
$L += 1
WEnd
While(StringCompare($aArray[$R], $vPivot) > 0)
$R -= 1
WEnd
EndIf
If $L <= $R Then
$vTmp = $aArray[$L]
$aArray[$L] = $aArray[$R]
$aArray[$R] = $vTmp
$L += 1
$R -= 1
EndIf
Until $L > $R
__ArrayQuickSort1D($aArray, $iStart, $R)
__ArrayQuickSort1D($aArray, $L, $iEnd)
EndFunc
Func __ArrayQuickSort2D(ByRef $aArray, Const ByRef $iStep, Const ByRef $iStart, Const ByRef $iEnd, Const ByRef $iSubItem, Const ByRef $iSubMax)
If $iEnd <= $iStart Then Return
Local $vTmp, $L = $iStart, $R = $iEnd, $vPivot = $aArray[Int(($iStart + $iEnd) / 2)][$iSubItem], $bNum = IsNumber($vPivot)
Do
If $bNum Then
While($iStep *($aArray[$L][$iSubItem] - $vPivot) < 0 And IsNumber($aArray[$L][$iSubItem])) Or(Not IsNumber($aArray[$L][$iSubItem]) And $iStep * StringCompare($aArray[$L][$iSubItem], $vPivot) < 0)
$L += 1
WEnd
While($iStep *($aArray[$R][$iSubItem] - $vPivot) > 0 And IsNumber($aArray[$R][$iSubItem])) Or(Not IsNumber($aArray[$R][$iSubItem]) And $iStep * StringCompare($aArray[$R][$iSubItem], $vPivot) > 0)
$R -= 1
WEnd
Else
While($iStep * StringCompare($aArray[$L][$iSubItem], $vPivot) < 0)
$L += 1
WEnd
While($iStep * StringCompare($aArray[$R][$iSubItem], $vPivot) > 0)
$R -= 1
WEnd
EndIf
If $L <= $R Then
For $i = 0 To $iSubMax
$vTmp = $aArray[$L][$i]
$aArray[$L][$i] = $aArray[$R][$i]
$aArray[$R][$i] = $vTmp
Next
$L += 1
$R -= 1
EndIf
Until $L > $R
__ArrayQuickSort2D($aArray, $iStep, $iStart, $R, $iSubItem, $iSubMax)
__ArrayQuickSort2D($aArray, $iStep, $L, $iEnd, $iSubItem, $iSubMax)
EndFunc
Func __ArrayDualPivotSort(ByRef $aArray, $iPivot_Left, $iPivot_Right, $bLeftMost = True)
If $iPivot_Left > $iPivot_Right Then Return
Local $iLength = $iPivot_Right - $iPivot_Left + 1
Local $i, $j, $k, $iAi, $iAk, $iA1, $iA2, $iLast
If $iLength < 45 Then
If $bLeftMost Then
$i = $iPivot_Left
While $i < $iPivot_Right
$j = $i
$iAi = $aArray[$i + 1]
While $iAi < $aArray[$j]
$aArray[$j + 1] = $aArray[$j]
$j -= 1
If $j + 1 = $iPivot_Left Then ExitLoop
WEnd
$aArray[$j + 1] = $iAi
$i += 1
WEnd
Else
While 1
If $iPivot_Left >= $iPivot_Right Then Return 1
$iPivot_Left += 1
If $aArray[$iPivot_Left] < $aArray[$iPivot_Left - 1] Then ExitLoop
WEnd
While 1
$k = $iPivot_Left
$iPivot_Left += 1
If $iPivot_Left > $iPivot_Right Then ExitLoop
$iA1 = $aArray[$k]
$iA2 = $aArray[$iPivot_Left]
If $iA1 < $iA2 Then
$iA2 = $iA1
$iA1 = $aArray[$iPivot_Left]
EndIf
$k -= 1
While $iA1 < $aArray[$k]
$aArray[$k + 2] = $aArray[$k]
$k -= 1
WEnd
$aArray[$k + 2] = $iA1
While $iA2 < $aArray[$k]
$aArray[$k + 1] = $aArray[$k]
$k -= 1
WEnd
$aArray[$k + 1] = $iA2
$iPivot_Left += 1
WEnd
$iLast = $aArray[$iPivot_Right]
$iPivot_Right -= 1
While $iLast < $aArray[$iPivot_Right]
$aArray[$iPivot_Right + 1] = $aArray[$iPivot_Right]
$iPivot_Right -= 1
WEnd
$aArray[$iPivot_Right + 1] = $iLast
EndIf
Return 1
EndIf
Local $iSeventh = BitShift($iLength, 3) + BitShift($iLength, 6) + 1
Local $iE1, $iE2, $iE3, $iE4, $iE5, $t
$iE3 = Ceiling(($iPivot_Left + $iPivot_Right) / 2)
$iE2 = $iE3 - $iSeventh
$iE1 = $iE2 - $iSeventh
$iE4 = $iE3 + $iSeventh
$iE5 = $iE4 + $iSeventh
If $aArray[$iE2] < $aArray[$iE1] Then
$t = $aArray[$iE2]
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
If $aArray[$iE3] < $aArray[$iE2] Then
$t = $aArray[$iE3]
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
If $aArray[$iE4] < $aArray[$iE3] Then
$t = $aArray[$iE4]
$aArray[$iE4] = $aArray[$iE3]
$aArray[$iE3] = $t
If $t < $aArray[$iE2] Then
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
EndIf
If $aArray[$iE5] < $aArray[$iE4] Then
$t = $aArray[$iE5]
$aArray[$iE5] = $aArray[$iE4]
$aArray[$iE4] = $t
If $t < $aArray[$iE3] Then
$aArray[$iE4] = $aArray[$iE3]
$aArray[$iE3] = $t
If $t < $aArray[$iE2] Then
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
EndIf
EndIf
Local $iLess = $iPivot_Left
Local $iGreater = $iPivot_Right
If(($aArray[$iE1] <> $aArray[$iE2]) And($aArray[$iE2] <> $aArray[$iE3]) And($aArray[$iE3] <> $aArray[$iE4]) And($aArray[$iE4] <> $aArray[$iE5])) Then
Local $iPivot_1 = $aArray[$iE2]
Local $iPivot_2 = $aArray[$iE4]
$aArray[$iE2] = $aArray[$iPivot_Left]
$aArray[$iE4] = $aArray[$iPivot_Right]
Do
$iLess += 1
Until $aArray[$iLess] >= $iPivot_1
Do
$iGreater -= 1
Until $aArray[$iGreater] <= $iPivot_2
$k = $iLess
While $k <= $iGreater
$iAk = $aArray[$k]
If $iAk < $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
ElseIf $iAk > $iPivot_2 Then
While $aArray[$iGreater] > $iPivot_2
$iGreater -= 1
If $iGreater + 1 = $k Then ExitLoop 2
WEnd
If $aArray[$iGreater] < $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $aArray[$iGreater]
$iLess += 1
Else
$aArray[$k] = $aArray[$iGreater]
EndIf
$aArray[$iGreater] = $iAk
$iGreater -= 1
EndIf
$k += 1
WEnd
$aArray[$iPivot_Left] = $aArray[$iLess - 1]
$aArray[$iLess - 1] = $iPivot_1
$aArray[$iPivot_Right] = $aArray[$iGreater + 1]
$aArray[$iGreater + 1] = $iPivot_2
__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 2, True)
__ArrayDualPivotSort($aArray, $iGreater + 2, $iPivot_Right, False)
If($iLess < $iE1) And($iE5 < $iGreater) Then
While $aArray[$iLess] = $iPivot_1
$iLess += 1
WEnd
While $aArray[$iGreater] = $iPivot_2
$iGreater -= 1
WEnd
$k = $iLess
While $k <= $iGreater
$iAk = $aArray[$k]
If $iAk = $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
ElseIf $iAk = $iPivot_2 Then
While $aArray[$iGreater] = $iPivot_2
$iGreater -= 1
If $iGreater + 1 = $k Then ExitLoop 2
WEnd
If $aArray[$iGreater] = $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iPivot_1
$iLess += 1
Else
$aArray[$k] = $aArray[$iGreater]
EndIf
$aArray[$iGreater] = $iAk
$iGreater -= 1
EndIf
$k += 1
WEnd
EndIf
__ArrayDualPivotSort($aArray, $iLess, $iGreater, False)
Else
Local $iPivot = $aArray[$iE3]
$k = $iLess
While $k <= $iGreater
If $aArray[$k] = $iPivot Then
$k += 1
ContinueLoop
EndIf
$iAk = $aArray[$k]
If $iAk < $iPivot Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
Else
While $aArray[$iGreater] > $iPivot
$iGreater -= 1
WEnd
If $aArray[$iGreater] < $iPivot Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $aArray[$iGreater]
$iLess += 1
Else
$aArray[$k] = $iPivot
EndIf
$aArray[$iGreater] = $iAk
$iGreater -= 1
EndIf
$k += 1
WEnd
__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 1, True)
__ArrayDualPivotSort($aArray, $iGreater + 1, $iPivot_Right, False)
EndIf
EndFunc
Func _ArrayToClip(Const ByRef $aArray, $sDelim_Col = "|", $iStart_Row = -1, $iEnd_Row = -1, $sDelim_Row = @CRLF, $iStart_Col = -1, $iEnd_Col = -1)
Local $sResult = _ArrayToString($aArray, $sDelim_Col, $iStart_Row, $iEnd_Row, $sDelim_Row, $iStart_Col, $iEnd_Col)
If @error Then Return SetError(@error, 0, 0)
If ClipPut($sResult) Then Return 1
Return SetError(-1, 0, 0)
EndFunc
Func _ArrayToString(Const ByRef $aArray, $sDelim_Col = "|", $iStart_Row = -1, $iEnd_Row = -1, $sDelim_Row = @CRLF, $iStart_Col = -1, $iEnd_Col = -1)
If $sDelim_Col = Default Then $sDelim_Col = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iStart_Row = Default Then $iStart_Row = -1
If $iEnd_Row = Default Then $iEnd_Row = -1
If $iStart_Col = Default Then $iStart_Col = -1
If $iEnd_Col = Default Then $iEnd_Col = -1
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If $iStart_Row = -1 Then $iStart_Row = 0
If $iEnd_Row = -1 Then $iEnd_Row = $iDim_1
If $iStart_Row < -1 Or $iEnd_Row < -1 Then Return SetError(3, 0, -1)
If $iStart_Row > $iDim_1 Or $iEnd_Row > $iDim_1 Then Return SetError(3, 0, "")
If $iStart_Row > $iEnd_Row Then Return SetError(4, 0, -1)
Local $sRet = ""
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
For $i = $iStart_Row To $iEnd_Row
$sRet &= $aArray[$i] & $sDelim_Col
Next
Return StringTrimRight($sRet, StringLen($sDelim_Col))
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iStart_Col = -1 Then $iStart_Col = 0
If $iEnd_Col = -1 Then $iEnd_Col = $iDim_2
If $iStart_Col < -1 Or $iEnd_Col < -1 Then Return SetError(5, 0, -1)
If $iStart_Col > $iDim_2 Or $iEnd_Col > $iDim_2 Then Return SetError(5, 0, -1)
If $iStart_Col > $iEnd_Col Then Return SetError(6, 0, -1)
For $i = $iStart_Row To $iEnd_Row
For $j = $iStart_Col To $iEnd_Col
$sRet &= $aArray[$i][$j] & $sDelim_Col
Next
$sRet = StringTrimRight($sRet, StringLen($sDelim_Col)) & $sDelim_Row
Next
Return StringTrimRight($sRet, StringLen($sDelim_Row))
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return 1
EndFunc
Global Const $MEM_COMMIT = 0x00001000
Global Const $MEM_RESERVE = 0x00002000
Global Const $PAGE_READWRITE = 0x00000004
Global Const $MEM_RELEASE = 0x00008000
Global Const $PROCESS_VM_OPERATION = 0x00000008
Global Const $PROCESS_VM_READ = 0x00000010
Global Const $PROCESS_VM_WRITE = 0x00000020
Global Const $SE_PRIVILEGE_ENABLED = 0x00000002
Global Enum $SECURITYANONYMOUS = 0, $SECURITYIDENTIFICATION, $SECURITYIMPERSONATION, $SECURITYDELEGATION
Global Const $TOKEN_QUERY = 0x00000008
Global Const $TOKEN_ADJUST_PRIVILEGES = 0x00000020
Func _WinAPI_GetLastError(Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
Local $aResult = DllCall("kernel32.dll", "dword", "GetLastError")
Return SetError($_iCurrentError, $_iCurrentExtended, $aResult[0])
EndFunc
Func _Security__AdjustTokenPrivileges($hToken, $bDisableAll, $tNewState, $iBufferLen, $tPrevState = 0, $pRequired = 0)
Local $aCall = DllCall("advapi32.dll", "bool", "AdjustTokenPrivileges", "handle", $hToken, "bool", $bDisableAll, "struct*", $tNewState, "dword", $iBufferLen, "struct*", $tPrevState, "struct*", $pRequired)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__ImpersonateSelf($iLevel = $SECURITYIMPERSONATION)
Local $aCall = DllCall("advapi32.dll", "bool", "ImpersonateSelf", "int", $iLevel)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__LookupPrivilegeValue($sSystem, $sName)
Local $aCall = DllCall("advapi32.dll", "bool", "LookupPrivilegeValueW", "wstr", $sSystem, "wstr", $sName, "int64*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[3]
EndFunc
Func _Security__OpenThreadToken($iAccess, $hThread = 0, $bOpenAsSelf = False)
If $hThread = 0 Then
Local $aResult = DllCall("kernel32.dll", "handle", "GetCurrentThread")
If @error Then Return SetError(@error + 10, @extended, 0)
$hThread = $aResult[0]
EndIf
Local $aCall = DllCall("advapi32.dll", "bool", "OpenThreadToken", "handle", $hThread, "dword", $iAccess, "bool", $bOpenAsSelf, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[4]
EndFunc
Func _Security__OpenThreadTokenEx($iAccess, $hThread = 0, $bOpenAsSelf = False)
Local $hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then
Local Const $ERROR_NO_TOKEN = 1008
If _WinAPI_GetLastError() <> $ERROR_NO_TOKEN Then Return SetError(20, _WinAPI_GetLastError(), 0)
If Not _Security__ImpersonateSelf() Then Return SetError(@error + 10, _WinAPI_GetLastError(), 0)
$hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then Return SetError(@error, _WinAPI_GetLastError(), 0)
EndIf
Return $hToken
EndFunc
Func _Security__SetPrivilege($hToken, $sPrivilege, $bEnable)
Local $iLUID = _Security__LookupPrivilegeValue("", $sPrivilege)
If $iLUID = 0 Then Return SetError(@error + 10, @extended, False)
Local Const $tagTOKEN_PRIVILEGES = "dword Count;align 4;int64 LUID;dword Attributes"
Local $tCurrState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iCurrState = DllStructGetSize($tCurrState)
Local $tPrevState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iPrevState = DllStructGetSize($tPrevState)
Local $tRequired = DllStructCreate("int Data")
DllStructSetData($tCurrState, "Count", 1)
DllStructSetData($tCurrState, "LUID", $iLUID)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tCurrState, $iCurrState, $tPrevState, $tRequired) Then Return SetError(2, @error, False)
DllStructSetData($tPrevState, "Count", 1)
DllStructSetData($tPrevState, "LUID", $iLUID)
Local $iAttributes = DllStructGetData($tPrevState, "Attributes")
If $bEnable Then
$iAttributes = BitOR($iAttributes, $SE_PRIVILEGE_ENABLED)
Else
$iAttributes = BitAND($iAttributes, BitNOT($SE_PRIVILEGE_ENABLED))
EndIf
DllStructSetData($tPrevState, "Attributes", $iAttributes)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tPrevState, $iPrevState, $tCurrState, $tRequired) Then Return SetError(3, @error, False)
Return True
EndFunc
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" &((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $tagSECURITY_ATTRIBUTES = "dword Length;ptr Descriptor;bool InheritHandle"
Global Const $tagMEMMAP = "handle hProc;ulong_ptr Size;ptr Mem"
Func _MemFree(ByRef $tMemMap)
Local $pMemory = DllStructGetData($tMemMap, "Mem")
Local $hProcess = DllStructGetData($tMemMap, "hProc")
Local $bResult = _MemVirtualFreeEx($hProcess, $pMemory, 0, $MEM_RELEASE)
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
If @error Then Return SetError(@error, @extended, False)
Return $bResult
EndFunc
Func _MemInit($hWnd, $iSize, ByRef $tMemMap)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error + 10, @extended, 0)
Local $iProcessID = $aResult[2]
If $iProcessID = 0 Then Return SetError(1, 0, 0)
Local $iAccess = BitOR($PROCESS_VM_OPERATION, $PROCESS_VM_READ, $PROCESS_VM_WRITE)
Local $hProcess = __Mem_OpenProcess($iAccess, False, $iProcessID, True)
Local $iAlloc = BitOR($MEM_RESERVE, $MEM_COMMIT)
Local $pMemory = _MemVirtualAllocEx($hProcess, 0, $iSize, $iAlloc, $PAGE_READWRITE)
If $pMemory = 0 Then Return SetError(2, 0, 0)
$tMemMap = DllStructCreate($tagMEMMAP)
DllStructSetData($tMemMap, "hProc", $hProcess)
DllStructSetData($tMemMap, "Size", $iSize)
DllStructSetData($tMemMap, "Mem", $pMemory)
Return $pMemory
EndFunc
Func _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
Local $aResult = DllCall("kernel32.dll", "bool", "ReadProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pSrce, "struct*", $pDest, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemWrite(ByRef $tMemMap, $pSrce, $pDest = 0, $iSize = 0, $sSrce = "struct*")
If $pDest = 0 Then $pDest = DllStructGetData($tMemMap, "Mem")
If $iSize = 0 Then $iSize = DllStructGetData($tMemMap, "Size")
Local $aResult = DllCall("kernel32.dll", "bool", "WriteProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pDest, $sSrce, $pSrce, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
Local $aResult = DllCall("kernel32.dll", "ptr", "VirtualAllocEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
Local $aResult = DllCall("kernel32.dll", "bool", "VirtualFreeEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func __Mem_OpenProcess($iAccess, $bInherit, $iProcessID, $bDebugPriv = False)
Local $aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iProcessID)
If @error Then Return SetError(@error + 10, @extended, 0)
If $aResult[0] Then Return $aResult[0]
If Not $bDebugPriv Then Return 0
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 20, @extended, 0)
_Security__SetPrivilege($hToken, "SeDebugPrivilege", True)
Local $iError = @error
Local $iLastError = @extended
Local $iRet = 0
If Not @error Then
$aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iProcessID)
$iError = @error
$iLastError = @extended
If $aResult[0] Then $iRet = $aResult[0]
_Security__SetPrivilege($hToken, "SeDebugPrivilege", False)
If @error Then
$iError = @error + 30
$iLastError = @extended
EndIf
Else
$iError = @error + 40
EndIf
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
Return SetError($iError, $iLastError, $iRet)
EndFunc
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aResult = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aResult[$iReturn]
Return $aResult
EndFunc
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global $__g_aInProcess_WinAPI[64][2] = [[0, 0]]
Func _WinAPI_CreateWindowEx($iExStyle, $sClass, $sName, $iStyle, $iX, $iY, $iWidth, $iHeight, $hParent, $hMenu = 0, $hInstance = 0, $pParam = 0)
If $hInstance = 0 Then $hInstance = _WinAPI_GetModuleHandle("")
Local $aResult = DllCall("user32.dll", "hwnd", "CreateWindowExW", "dword", $iExStyle, "wstr", $sClass, "wstr", $sName, "dword", $iStyle, "int", $iX, "int", $iY, "int", $iWidth, "int", $iHeight, "hwnd", $hParent, "handle", $hMenu, "handle", $hInstance, "struct*", $pParam)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetDlgCtrlID($hWnd)
Local $aResult = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetModuleHandle($sModuleName)
Local $sModuleNameType = "wstr"
If $sModuleName = "" Then
$sModuleName = 0
$sModuleNameType = "ptr"
EndIf
Local $aResult = DllCall("kernel32.dll", "handle", "GetModuleHandleW", $sModuleNameType, $sModuleName)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindowThreadProcessId($hWnd, ByRef $iPID)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error, @extended, 0)
$iPID = $aResult[2]
Return $aResult[0]
EndFunc
Func _WinAPI_InProcess($hWnd, ByRef $hLastWnd)
If $hWnd = $hLastWnd Then Return True
For $iI = $__g_aInProcess_WinAPI[0][0] To 1 Step -1
If $hWnd = $__g_aInProcess_WinAPI[$iI][0] Then
If $__g_aInProcess_WinAPI[$iI][1] Then
$hLastWnd = $hWnd
Return True
Else
Return False
EndIf
EndIf
Next
Local $iPID
_WinAPI_GetWindowThreadProcessId($hWnd, $iPID)
Local $iCount = $__g_aInProcess_WinAPI[0][0] + 1
If $iCount >= 64 Then $iCount = 1
$__g_aInProcess_WinAPI[0][0] = $iCount
$__g_aInProcess_WinAPI[$iCount][0] = $hWnd
$__g_aInProcess_WinAPI[$iCount][1] =($iPID = @AutoItPID)
Return $__g_aInProcess_WinAPI[$iCount][1]
EndFunc
Global Const $_UDF_GlobalIDs_OFFSET = 2
Global Const $_UDF_GlobalID_MAX_WIN = 16
Global Const $_UDF_STARTID = 10000
Global Const $_UDF_GlobalID_MAX_IDS = 55535
Global Const $__UDFGUICONSTANT_WS_VISIBLE = 0x10000000
Global Const $__UDFGUICONSTANT_WS_CHILD = 0x40000000
Global $__g_aUDF_GlobalIDs_Used[$_UDF_GlobalID_MAX_WIN][$_UDF_GlobalID_MAX_IDS + $_UDF_GlobalIDs_OFFSET + 1]
Func __UDF_GetNextGlobalID($hWnd)
Local $nCtrlID, $iUsedIndex = -1, $bAllUsed = True
If Not WinExists($hWnd) Then Return SetError(-1, -1, 0)
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] <> 0 Then
If Not WinExists($__g_aUDF_GlobalIDs_Used[$iIndex][0]) Then
For $x = 0 To UBound($__g_aUDF_GlobalIDs_Used, $UBOUND_COLUMNS) - 1
$__g_aUDF_GlobalIDs_Used[$iIndex][$x] = 0
Next
$__g_aUDF_GlobalIDs_Used[$iIndex][1] = $_UDF_STARTID
$bAllUsed = False
EndIf
EndIf
Next
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] = $hWnd Then
$iUsedIndex = $iIndex
ExitLoop
EndIf
Next
If $iUsedIndex = -1 Then
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] = 0 Then
$__g_aUDF_GlobalIDs_Used[$iIndex][0] = $hWnd
$__g_aUDF_GlobalIDs_Used[$iIndex][1] = $_UDF_STARTID
$bAllUsed = False
$iUsedIndex = $iIndex
ExitLoop
EndIf
Next
EndIf
If $iUsedIndex = -1 And $bAllUsed Then Return SetError(16, 0, 0)
If $__g_aUDF_GlobalIDs_Used[$iUsedIndex][1] = $_UDF_STARTID + $_UDF_GlobalID_MAX_IDS Then
For $iIDIndex = $_UDF_GlobalIDs_OFFSET To UBound($__g_aUDF_GlobalIDs_Used, $UBOUND_COLUMNS) - 1
If $__g_aUDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = 0 Then
$nCtrlID =($iIDIndex - $_UDF_GlobalIDs_OFFSET) + 10000
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = $nCtrlID
Return $nCtrlID
EndIf
Next
Return SetError(-1, $_UDF_GlobalID_MAX_IDS, 0)
EndIf
$nCtrlID = $__g_aUDF_GlobalIDs_Used[$iUsedIndex][1]
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][1] += 1
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][($nCtrlID - 10000) + $_UDF_GlobalIDs_OFFSET] = $nCtrlID
Return $nCtrlID
EndFunc
Global $__g_hLVLastWnd
Func _GUICtrlListView_DeleteAllItems($hWnd)
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
Local $vCID = 0
If IsHWnd($hWnd) Then
$vCID = _WinAPI_GetDlgCtrlID($hWnd)
Else
$vCID = $hWnd
$hWnd = GUICtrlGetHandle($hWnd)
EndIf
If $vCID < $_UDF_STARTID Then
Local $iParam = 0
For $iIndex = _GUICtrlListView_GetItemCount($hWnd) - 1 To 0 Step -1
$iParam = _GUICtrlListView_GetItemParam($hWnd, $iIndex)
If GUICtrlGetState($iParam) > 0 And GUICtrlGetHandle($iParam) = 0 Then
GUICtrlDelete($iParam)
EndIf
Next
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
EndIf
Return _SendMessage($hWnd, $LVM_DELETEALLITEMS) <> 0
EndFunc
Func _GUICtrlListView_GetItem($hWnd, $iIndex, $iSubItem = 0)
Local $aItem[8]
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "Mask", BitOR($LVIF_GROUPID, $LVIF_IMAGE, $LVIF_INDENT, $LVIF_PARAM, $LVIF_STATE))
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "StateMask", -1)
_GUICtrlListView_GetItemEx($hWnd, $tItem)
Local $iState = DllStructGetData($tItem, "State")
If BitAND($iState, $LVIS_CUT) <> 0 Then $aItem[0] = BitOR($aItem[0], 1)
If BitAND($iState, $LVIS_DROPHILITED) <> 0 Then $aItem[0] = BitOR($aItem[0], 2)
If BitAND($iState, $LVIS_FOCUSED) <> 0 Then $aItem[0] = BitOR($aItem[0], 4)
If BitAND($iState, $LVIS_SELECTED) <> 0 Then $aItem[0] = BitOR($aItem[0], 8)
$aItem[1] = __GUICtrlListView_OverlayImageMaskToIndex($iState)
$aItem[2] = __GUICtrlListView_StateImageMaskToIndex($iState)
$aItem[3] = _GUICtrlListView_GetItemText($hWnd, $iIndex, $iSubItem)
$aItem[4] = DllStructGetData($tItem, "Image")
$aItem[5] = DllStructGetData($tItem, "Param")
$aItem[6] = DllStructGetData($tItem, "Indent")
$aItem[7] = DllStructGetData($tItem, "GroupID")
Return $aItem
EndFunc
Func _GUICtrlListView_GetItemCount($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETITEMCOUNT)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETITEMCOUNT, 0, 0)
EndIf
EndFunc
Func _GUICtrlListView_GetItemEx($hWnd, ByRef $tItem)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
$iRet = _SendMessage($hWnd, $LVM_GETITEMW, 0, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem, $tMemMap)
_MemWrite($tMemMap, $tItem)
If $bUnicode Then
_SendMessage($hWnd, $LVM_GETITEMW, 0, $pMemory, 0, "wparam", "ptr")
Else
_SendMessage($hWnd, $LVM_GETITEMA, 0, $pMemory, 0, "wparam", "ptr")
EndIf
_MemRead($tMemMap, $pMemory, $tItem, $iItem)
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_GETITEMW, 0, $pItem)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_GETITEMA, 0, $pItem)
EndIf
EndIf
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_GetItemParam($hWnd, $iIndex)
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "Mask", $LVIF_PARAM)
DllStructSetData($tItem, "Item", $iIndex)
_GUICtrlListView_GetItemEx($hWnd, $tItem)
Return DllStructGetData($tItem, "Param")
EndFunc
Func _GUICtrlListView_GetItemText($hWnd, $iIndex, $iSubItem = 0)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[4096]")
Else
$tBuffer = DllStructCreate("char Text[4096]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "TextMax", 4096)
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tItem, "Text", $pBuffer)
_SendMessage($hWnd, $LVM_GETITEMTEXTW, $iIndex, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem + 4096, $tMemMap)
Local $pText = $pMemory + $iItem
DllStructSetData($tItem, "Text", $pText)
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
If $bUnicode Then
_SendMessage($hWnd, $LVM_GETITEMTEXTW, $iIndex, $pMemory, 0, "wparam", "ptr")
Else
_SendMessage($hWnd, $LVM_GETITEMTEXTA, $iIndex, $pMemory, 0, "wparam", "ptr")
EndIf
_MemRead($tMemMap, $pText, $tBuffer, 4096)
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
DllStructSetData($tItem, "Text", $pBuffer)
If $bUnicode Then
GUICtrlSendMsg($hWnd, $LVM_GETITEMTEXTW, $iIndex, $pItem)
Else
GUICtrlSendMsg($hWnd, $LVM_GETITEMTEXTA, $iIndex, $pItem)
EndIf
EndIf
Return DllStructGetData($tBuffer, "Text")
EndFunc
Func _GUICtrlListView_GetUnicodeFormat($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETUNICODEFORMAT) <> 0
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETUNICODEFORMAT, 0, 0) <> 0
EndIf
EndFunc
Func __GUICtrlListView_OverlayImageMaskToIndex($iMask)
Return BitShift(BitAND($LVIS_OVERLAYMASK, $iMask), 8)
EndFunc
Func _GUICtrlListView_SetColumnWidth($hWnd, $iCol, $iWidth)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_SETCOLUMNWIDTH, $iCol, $iWidth)
Else
Return GUICtrlSendMsg($hWnd, $LVM_SETCOLUMNWIDTH, $iCol, $iWidth)
EndIf
EndFunc
Func __GUICtrlListView_StateImageMaskToIndex($iMask)
Return BitShift(BitAND($iMask, $LVIS_STATEIMAGEMASK), 12)
EndFunc
Global Const $__STATUSBARCONSTANT_WM_USER = 0X400
Global Const $SB_GETUNICODEFORMAT = 0x2000 + 6
Global Const $SB_ISSIMPLE =($__STATUSBARCONSTANT_WM_USER + 14)
Global Const $SB_SETPARTS =($__STATUSBARCONSTANT_WM_USER + 4)
Global Const $SB_SETTEXTA =($__STATUSBARCONSTANT_WM_USER + 1)
Global Const $SB_SETTEXTW =($__STATUSBARCONSTANT_WM_USER + 11)
Global Const $SB_SETTEXT = $SB_SETTEXTA
Global Const $SB_SIMPLEID = 0xff
Global $__g_hSBLastWnd
Global Const $__STATUSBARCONSTANT_ClassName = "msctls_statusbar32"
Global Const $__STATUSBARCONSTANT_WM_SIZE = 0x05
Func _GUICtrlStatusBar_Create($hWnd, $vPartEdge = -1, $vPartText = "", $iStyles = -1, $iExStyles = 0x00000000)
If Not IsHWnd($hWnd) Then Return SetError(1, 0, 0)
Local $iStyle = BitOR($__UDFGUICONSTANT_WS_CHILD, $__UDFGUICONSTANT_WS_VISIBLE)
If $iStyles = -1 Then $iStyles = 0x00000000
If $iExStyles = -1 Then $iExStyles = 0x00000000
Local $aPartWidth[1], $aPartText[1]
If @NumParams > 1 Then
If IsArray($vPartEdge) Then
$aPartWidth = $vPartEdge
Else
$aPartWidth[0] = $vPartEdge
EndIf
If @NumParams = 2 Then
ReDim $aPartText[UBound($aPartWidth)]
Else
If IsArray($vPartText) Then
$aPartText = $vPartText
Else
$aPartText[0] = $vPartText
EndIf
If UBound($aPartWidth) <> UBound($aPartText) Then
Local $iLast
If UBound($aPartWidth) > UBound($aPartText) Then
$iLast = UBound($aPartText)
ReDim $aPartText[UBound($aPartWidth)]
Else
$iLast = UBound($aPartWidth)
ReDim $aPartWidth[UBound($aPartText)]
For $x = $iLast To UBound($aPartWidth) - 1
$aPartWidth[$x] = $aPartWidth[$x - 1] + 75
Next
$aPartWidth[UBound($aPartText) - 1] = -1
EndIf
EndIf
EndIf
If Not IsHWnd($hWnd) Then $hWnd = HWnd($hWnd)
If @NumParams > 3 Then $iStyle = BitOR($iStyle, $iStyles)
EndIf
Local $nCtrlID = __UDF_GetNextGlobalID($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Local $hWndSBar = _WinAPI_CreateWindowEx($iExStyles, $__STATUSBARCONSTANT_ClassName, "", $iStyle, 0, 0, 0, 0, $hWnd, $nCtrlID)
If @error Then Return SetError(@error, @extended, 0)
If @NumParams > 1 Then
_GUICtrlStatusBar_SetParts($hWndSBar, UBound($aPartWidth), $aPartWidth)
For $x = 0 To UBound($aPartText) - 1
_GUICtrlStatusBar_SetText($hWndSBar, $aPartText[$x], $x)
Next
EndIf
Return $hWndSBar
EndFunc
Func _GUICtrlStatusBar_GetUnicodeFormat($hWnd)
Return _SendMessage($hWnd, $SB_GETUNICODEFORMAT) <> 0
EndFunc
Func _GUICtrlStatusBar_IsSimple($hWnd)
Return _SendMessage($hWnd, $SB_ISSIMPLE) <> 0
EndFunc
Func _GUICtrlStatusBar_Resize($hWnd)
_SendMessage($hWnd, $__STATUSBARCONSTANT_WM_SIZE)
EndFunc
Func _GUICtrlStatusBar_SetParts($hWnd, $aParts = -1, $aPartWidth = 25)
Local $tParts, $iParts = 1
If IsArray($aParts) <> 0 Then
$aParts[UBound($aParts) - 1] = -1
$iParts = UBound($aParts)
$tParts = DllStructCreate("int[" & $iParts & "]")
For $x = 0 To $iParts - 2
DllStructSetData($tParts, 1, $aParts[$x], $x + 1)
Next
DllStructSetData($tParts, 1, -1, $iParts)
ElseIf IsArray($aPartWidth) <> 0 Then
$iParts = UBound($aPartWidth)
$tParts = DllStructCreate("int[" & $iParts & "]")
For $x = 0 To $iParts - 2
DllStructSetData($tParts, 1, $aPartWidth[$x], $x + 1)
Next
DllStructSetData($tParts, 1, -1, $iParts)
ElseIf $aParts > 1 Then
$iParts = $aParts
$tParts = DllStructCreate("int[" & $iParts & "]")
For $x = 1 To $iParts - 1
DllStructSetData($tParts, 1, $aPartWidth * $x, $x)
Next
DllStructSetData($tParts, 1, -1, $iParts)
Else
$tParts = DllStructCreate("int")
DllStructSetData($tParts, $iParts, -1)
EndIf
If _WinAPI_InProcess($hWnd, $__g_hSBLastWnd) Then
_SendMessage($hWnd, $SB_SETPARTS, $iParts, $tParts, 0, "wparam", "struct*")
Else
Local $iSize = DllStructGetSize($tParts)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iSize, $tMemMap)
_MemWrite($tMemMap, $tParts)
_SendMessage($hWnd, $SB_SETPARTS, $iParts, $pMemory, 0, "wparam", "ptr")
_MemFree($tMemMap)
EndIf
_GUICtrlStatusBar_Resize($hWnd)
Return True
EndFunc
Func _GUICtrlStatusBar_SetText($hWnd, $sText = "", $iPart = 0, $iUFlag = 0)
Local $bUnicode = _GUICtrlStatusBar_GetUnicodeFormat($hWnd)
Local $iBuffer = StringLen($sText) + 1
Local $tText
If $bUnicode Then
$tText = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tText = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
DllStructSetData($tText, "Text", $sText)
If _GUICtrlStatusBar_IsSimple($hWnd) Then $iPart = $SB_SIMPLEID
Local $iRet
If _WinAPI_InProcess($hWnd, $__g_hSBLastWnd) Then
$iRet = _SendMessage($hWnd, $SB_SETTEXTW, BitOR($iPart, $iUFlag), $tText, 0, "wparam", "struct*")
Else
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iBuffer, $tMemMap)
_MemWrite($tMemMap, $tText)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $SB_SETTEXTW, BitOR($iPart, $iUFlag), $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $SB_SETTEXT, BitOR($iPart, $iUFlag), $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Return $iRet <> 0
EndFunc
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Func _DateDiff($sType, $sStartDate, $sEndDate)
$sType = StringLeft($sType, 1)
If StringInStr("d,m,y,w,h,n,s", $sType) = 0 Or $sType = "" Then
Return SetError(1, 0, 0)
EndIf
If Not _DateIsValid($sStartDate) Then
Return SetError(2, 0, 0)
EndIf
If Not _DateIsValid($sEndDate) Then
Return SetError(3, 0, 0)
EndIf
Local $asStartDatePart[4], $asStartTimePart[4], $asEndDatePart[4], $asEndTimePart[4]
_DateTimeSplit($sStartDate, $asStartDatePart, $asStartTimePart)
_DateTimeSplit($sEndDate, $asEndDatePart, $asEndTimePart)
Local $aDaysDiff = _DateToDayValue($asEndDatePart[1], $asEndDatePart[2], $asEndDatePart[3]) - _DateToDayValue($asStartDatePart[1], $asStartDatePart[2], $asStartDatePart[3])
Local $iTimeDiff, $iYearDiff, $iStartTimeInSecs, $iEndTimeInSecs
If $asStartTimePart[0] > 1 And $asEndTimePart[0] > 1 Then
$iStartTimeInSecs = $asStartTimePart[1] * 3600 + $asStartTimePart[2] * 60 + $asStartTimePart[3]
$iEndTimeInSecs = $asEndTimePart[1] * 3600 + $asEndTimePart[2] * 60 + $asEndTimePart[3]
$iTimeDiff = $iEndTimeInSecs - $iStartTimeInSecs
If $iTimeDiff < 0 Then
$aDaysDiff = $aDaysDiff - 1
$iTimeDiff = $iTimeDiff + 24 * 60 * 60
EndIf
Else
$iTimeDiff = 0
EndIf
Select
Case $sType = "d"
Return $aDaysDiff
Case $sType = "m"
$iYearDiff = $asEndDatePart[1] - $asStartDatePart[1]
Local $iMonthDiff = $asEndDatePart[2] - $asStartDatePart[2] + $iYearDiff * 12
If $asEndDatePart[3] < $asStartDatePart[3] Then $iMonthDiff = $iMonthDiff - 1
$iStartTimeInSecs = $asStartTimePart[1] * 3600 + $asStartTimePart[2] * 60 + $asStartTimePart[3]
$iEndTimeInSecs = $asEndTimePart[1] * 3600 + $asEndTimePart[2] * 60 + $asEndTimePart[3]
$iTimeDiff = $iEndTimeInSecs - $iStartTimeInSecs
If $asEndDatePart[3] = $asStartDatePart[3] And $iTimeDiff < 0 Then $iMonthDiff = $iMonthDiff - 1
Return $iMonthDiff
Case $sType = "y"
$iYearDiff = $asEndDatePart[1] - $asStartDatePart[1]
If $asEndDatePart[2] < $asStartDatePart[2] Then $iYearDiff = $iYearDiff - 1
If $asEndDatePart[2] = $asStartDatePart[2] And $asEndDatePart[3] < $asStartDatePart[3] Then $iYearDiff = $iYearDiff - 1
$iStartTimeInSecs = $asStartTimePart[1] * 3600 + $asStartTimePart[2] * 60 + $asStartTimePart[3]
$iEndTimeInSecs = $asEndTimePart[1] * 3600 + $asEndTimePart[2] * 60 + $asEndTimePart[3]
$iTimeDiff = $iEndTimeInSecs - $iStartTimeInSecs
If $asEndDatePart[2] = $asStartDatePart[2] And $asEndDatePart[3] = $asStartDatePart[3] And $iTimeDiff < 0 Then $iYearDiff = $iYearDiff - 1
Return $iYearDiff
Case $sType = "w"
Return Int($aDaysDiff / 7)
Case $sType = "h"
Return $aDaysDiff * 24 + Int($iTimeDiff / 3600)
Case $sType = "n"
Return $aDaysDiff * 24 * 60 + Int($iTimeDiff / 60)
Case $sType = "s"
Return $aDaysDiff * 24 * 60 * 60 + $iTimeDiff
EndSelect
EndFunc
Func _DateIsLeapYear($iYear)
If StringIsInt($iYear) Then
Select
Case Mod($iYear, 4) = 0 And Mod($iYear, 100) <> 0
Return 1
Case Mod($iYear, 400) = 0
Return 1
Case Else
Return 0
EndSelect
EndIf
Return SetError(1, 0, 0)
EndFunc
Func _DateIsValid($sDate)
Local $asDatePart[4], $asTimePart[4]
_DateTimeSplit($sDate, $asDatePart, $asTimePart)
If Not StringIsInt($asDatePart[1]) Then Return 0
If Not StringIsInt($asDatePart[2]) Then Return 0
If Not StringIsInt($asDatePart[3]) Then Return 0
$asDatePart[1] = Int($asDatePart[1])
$asDatePart[2] = Int($asDatePart[2])
$asDatePart[3] = Int($asDatePart[3])
Local $iNumDays = _DaysInMonth($asDatePart[1])
If $asDatePart[1] < 1000 Or $asDatePart[1] > 2999 Then Return 0
If $asDatePart[2] < 1 Or $asDatePart[2] > 12 Then Return 0
If $asDatePart[3] < 1 Or $asDatePart[3] > $iNumDays[$asDatePart[2]] Then Return 0
If $asTimePart[0] < 1 Then Return 1
If $asTimePart[0] < 2 Then Return 0
If $asTimePart[0] = 2 Then $asTimePart[3] = "00"
If Not StringIsInt($asTimePart[1]) Then Return 0
If Not StringIsInt($asTimePart[2]) Then Return 0
If Not StringIsInt($asTimePart[3]) Then Return 0
$asTimePart[1] = Int($asTimePart[1])
$asTimePart[2] = Int($asTimePart[2])
$asTimePart[3] = Int($asTimePart[3])
If $asTimePart[1] < 0 Or $asTimePart[1] > 23 Then Return 0
If $asTimePart[2] < 0 Or $asTimePart[2] > 59 Then Return 0
If $asTimePart[3] < 0 Or $asTimePart[3] > 59 Then Return 0
Return 1
EndFunc
Func _DateTimeSplit($sDate, ByRef $aDatePart, ByRef $iTimePart)
Local $sDateTime = StringSplit($sDate, " T")
If $sDateTime[0] > 0 Then $aDatePart = StringSplit($sDateTime[1], "/-.")
If $sDateTime[0] > 1 Then
$iTimePart = StringSplit($sDateTime[2], ":")
If UBound($iTimePart) < 4 Then ReDim $iTimePart[4]
Else
Dim $iTimePart[4]
EndIf
If UBound($aDatePart) < 4 Then ReDim $aDatePart[4]
For $x = 1 To 3
If StringIsInt($aDatePart[$x]) Then
$aDatePart[$x] = Int($aDatePart[$x])
Else
$aDatePart[$x] = -1
EndIf
If StringIsInt($iTimePart[$x]) Then
$iTimePart[$x] = Int($iTimePart[$x])
Else
$iTimePart[$x] = 0
EndIf
Next
Return 1
EndFunc
Func _DateToDayValue($iYear, $iMonth, $iDay)
If Not _DateIsValid(StringFormat("%04d/%02d/%02d", $iYear, $iMonth, $iDay)) Then
Return SetError(1, 0, "")
EndIf
If $iMonth < 3 Then
$iMonth = $iMonth + 12
$iYear = $iYear - 1
EndIf
Local $i_FactorA = Int($iYear / 100)
Local $i_FactorB = Int($i_FactorA / 4)
Local $i_FactorC = 2 - $i_FactorA + $i_FactorB
Local $i_FactorE = Int(1461 *($iYear + 4716) / 4)
Local $i_FactorF = Int(153 *($iMonth + 1) / 5)
Local $iJulianDate = $i_FactorC + $iDay + $i_FactorE + $i_FactorF - 1524.5
Return $iJulianDate
EndFunc
Func _NowCalcDate()
Return @YEAR & "/" & @MON & "/" & @MDAY
EndFunc
Func _DaysInMonth($iYear)
Local $aDays = [12, 31,(_DateIsLeapYear($iYear) ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
Return $aDays
EndFunc
Global Const $COLOR_GREEN = 0x008000
Global Const $COLOR_RED = 0xFF0000
Func _Max($iNum1, $iNum2)
If Not IsNumber($iNum1) Then Return SetError(1, 0, 0)
If Not IsNumber($iNum2) Then Return SetError(2, 0, 0)
Return($iNum1 > $iNum2) ? $iNum1 : $iNum2
EndFunc
Func _Singleton($sOccurrenceName, $iFlag = 0)
Local Const $ERROR_ALREADY_EXISTS = 183
Local Const $SECURITY_DESCRIPTOR_REVISION = 1
Local $tSecurityAttributes = 0
If BitAND($iFlag, 2) Then
Local $tSecurityDescriptor = DllStructCreate("byte;byte;word;ptr[4]")
Local $aRet = DllCall("advapi32.dll", "bool", "InitializeSecurityDescriptor", "struct*", $tSecurityDescriptor, "dword", $SECURITY_DESCRIPTOR_REVISION)
If @error Then Return SetError(@error, @extended, 0)
If $aRet[0] Then
$aRet = DllCall("advapi32.dll", "bool", "SetSecurityDescriptorDacl", "struct*", $tSecurityDescriptor, "bool", 1, "ptr", 0, "bool", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aRet[0] Then
$tSecurityAttributes = DllStructCreate($tagSECURITY_ATTRIBUTES)
DllStructSetData($tSecurityAttributes, 1, DllStructGetSize($tSecurityAttributes))
DllStructSetData($tSecurityAttributes, 2, DllStructGetPtr($tSecurityDescriptor))
DllStructSetData($tSecurityAttributes, 3, 0)
EndIf
EndIf
EndIf
Local $aHandle = DllCall("kernel32.dll", "handle", "CreateMutexW", "struct*", $tSecurityAttributes, "bool", 1, "wstr", $sOccurrenceName)
If @error Then Return SetError(@error, @extended, 0)
Local $aLastError = DllCall("kernel32.dll", "dword", "GetLastError")
If @error Then Return SetError(@error, @extended, 0)
If $aLastError[0] = $ERROR_ALREADY_EXISTS Then
If BitAND($iFlag, 1) Then
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aHandle[0])
If @error Then Return SetError(@error, @extended, 0)
Return SetError($aLastError[0], $aLastError[0], 0)
Else
Exit -1
EndIf
EndIf
Return $aHandle[0]
EndFunc
$language = 1
dim $texts[2][200]
$texts[0][0] = "Konnte ADODDB-Connection nicht erstellen"
$texts[0][1] = "Funktionen"
$texts[0][2] = "Neu Laden"
$texts[0][3] = "Nachbarn"
$texts[0][4] = "User wechseln"
$texts[0][5] = " alles kopieren"
$texts[0][6] = " alles drucken"
$texts[0][7] = "NC mit VNC öffnen"
$texts[0][8] = "Citrix-Profile säubern"
$texts[0][9] = "Zähle User pro Abteilung"
$texts[0][10] = "Zeige alle aus Dezernat"
$texts[0][11] = "Gruppenvergleich"
$texts[0][12] = "Gruppenliste kopieren"
$texts[0][13] = "Gruppen-Info"
$texts[0][14] = " c:\ des PCs"
$texts[0][15] = "Ereignisanzeige"
$texts[0][16] = "Remote Desktop"
$texts[0][17] = "Versionscheck"
$texts[0][18] = "PC-Softwareinfo"
$texts[0][19] = " Beenden"
$texts[0][20] = "Name des Users:"
$texts[0][21] = "UserID:"
$texts[0][22] = "Anrede:"
$texts[0][23] = "F-Nummer:"
$texts[0][24] = "Veröffentlicht auf"
$texts[0][25] = "Tel.Nr.:"
$texts[0][26] = "Abteilung:"
$texts[0][27] = "Domain:"
$texts[0][28] = "email:"
$texts[0][29] = "Ort:"
$texts[0][30] = "Adresse:"
$texts[0][31] = "Raum:"
$texts[0][32] = "Namen der PCs|IP-Adresse|Mac-Adresse|?|Namen der NCs|IP-Adresse|Mac-Adresse|?"
$texts[0][33] = "Letzte Änderung:"
$texts[0][34] = "Haltbar bis:"
$texts[0][35] = "PW nötig:"
$texts[0][36] = "PW ändern:"
$texts[0][37] = "Gesperrt/Deaktiv:"
$texts[0][38] = "Entsperren"
$texts[0][39] = "Erzeugt am:"
$texts[0][40] = "Fehlversuche:"
$texts[0][41] = "letzter Login:"
$texts[0][42] = "Home-Verzeichnis:"
$texts[0][43] = "Home-Lw:"
$texts[0][44] = "Loginscript:"
$texts[0][45] = "Profilpfad:"
$texts[0][46] = "Geboren am:"
$texts[0][47] = "TS-Profil:"
$texts[0][48] = "Postkorb:"
$texts[0][49] = "AmtsBez.:"
$texts[0][50] = "PWChange:"
$texts[0][51] = "Geschlecht:"
$texts[0][52] = "SAP PersNr:"
$texts[0][53] = "Name der Gruppe|samAccountName|Art der Gruppe|Mitglied von|Gültigkeit"
$texts[0][54] = "Nachname"
$texts[0][55] = "Vorname"
$texts[0][56] = "Abbrechen"
$texts[0][57] = "Habe die neue Version installiert."
$texts[0][58] = "Warte auf Eingabe"
$texts[0][59] = "Lese die Userdaten aus..."
$texts[0][60] = " wurde ins Clipboard kopiert"
$texts[0][61] = "Die Domäne "
$texts[0][62] = " kann nicht angesprochen werden."
$texts[0][63] = "Der Account "
$texts[0][64] = " existiert nicht in der Domäne."
$texts[0][65] = "noch nie"
$texts[0][66] = "Userdaten wurden ausgelesen (in "
$texts[0][67] = " Sekunden, davon "
$texts[0][68] = " fürs ADS), Domänencontroller: "
$texts[0][69] = "Benutzer wählen"
$texts[0][70] = "Name des Anwenders/der Anwenderin|UserID"
$texts[0][71] = "Es wurden "
$texts[0][72] = " User mit diesem Suchbegriff gefunden. Bitte den genauen User auswählen:"
$texts[0][73] = "32Bit Programm unter 64 Bit OS -> nicht auslesbar!"
$texts[0][74] = "Ja"
$texts[0][75] = "Nein"
$texts[0][76] = "erlaubt"
$texts[0][77] = "gesperrt"
$texts[0][78] = "für immer"
$texts[0][79] = "Diese Gruppe ist eine Distributionsgruppe und wird von Userinfo (noch) nicht interpretiert"
$texts[0][80] = "Gruppenname"
$texts[0][81] = "angelegt"
$texts[0][82] = "letzte Änderung"
$texts[0][83] = "Mail"
$texts[0][84] = "Gruppeninfo"
$texts[0][85] = "Beschreibung"
$texts[0][86] = "Bereich der Gruppe"
$texts[0][87] = "Art der Gruppe"
$texts[0][88] = "Anzahl Mitglieder"
$texts[0][89] = "Mitglieder nach Namen"
$texts[0][90] = "Mitglieder nach UserID"
$texts[0][91] = "Gruppenmitglieder wurden ins Clipboard kopiert"
$texts[0][92] = "Mit wem vergleichen"
$texts[0][93] = "Mit welcher UserID soll verglichen werden?"
$texts[0][94] = "Gruppen die nur "
$texts[0][95] = " hat"
$texts[0][96] = "Gemeinsame Gruppen"
$texts[0][97] = "Vergleich der Benutzergruppen wurde ausgeführt"
$texts[0][98] = "Account gesperrt"
$texts[0][99] = "Account deaktiviert"
$texts[0][100] = "Domänencontroller"
$texts[0][101] = "Namen der PCs"
$texts[0][102] = "Namen der NCs"
$texts[0][103] = "Gruppen"
$texts[0][104] = "Keine Verbindung zur PC-Datenbank möglich - entweder Server weg oder Account gesperrt!"
$texts[0][105] = "Keine Verbindung zur NC-Datenbank möglich"
$texts[0][106] = "Mehrere Geräte in Datenbank - Bitte wählen"
$texts[0][107] = "Name des Rechners|IP-Adresse"
$texts[0][108] = "Kein RealVNC gefunden"
$texts[0][109] = "Der unter C:\Programme\RealVNC\VNC4\vncviewer.exe erwartete RealVNC Viewer ist nicht vorhanden"
$texts[0][110] = "-> verschoben"
$texts[0][111] = "-> nicht verschoben"
$texts[0][112] = "Löschergebnis"
$texts[0][113] = "Es wurden keine Verzeichnisse gefunden"
$texts[0][114] = "Nachbarn anzeigen"
$texts[0][115] = "Raum|Name des Anwenders/der Anwenderin|UserID|Telefon"
$texts[0][116] = "Gruppenvergleich mit diesem Nachbarn"
$texts[0][117] = "Wechsel der Info-Anzeige zu diesem Nachbarn"
$texts[0][118] = "Offline"
$texts[0][119] = "Der PC ist offline"
$texts[0][120] = "Zugriffsproblem"
$texts[0][121] = "Kein Zugriff möglich auf das Laufwerk C"
$texts[0][122] = "Softwarestand des PCs "
$texts[0][123] = "Software auf dem PC "
$texts[0][124] = "Anzahl der Mitarbeiter pro Abteilung"
$texts[0][125] = "Keine Verbindung zur X500 möglich"
$texts[1][ 0] = "Could not create ADODB Connection"
$texts[1][ 1] = "Functions"
$texts[1][ 2] = "Refresh"
$texts[1][ 3] = "Neighbours"
$texts[1][ 4] = "New Search"
$texts[1][ 5] = "Copy everything"
$texts[1][ 6] = "Print everything"
$texts[1][ 7] = "View NC with VNC"
$texts[1][ 8] = "Clean Citrix Profiles"
$texts[1][ 9] = "Count Users of Department"
$texts[1][10] = "Show Users of Department"
$texts[1][11] = "Compare Groups"
$texts[1][12] = "Copy Group List"
$texts[1][13] = "Group-Info"
$texts[1][14] = "c:\ from PC"
$texts[1][15] = "Eventlog"
$texts[1][16] = "Remote Desktop"
$texts[1][17] = "Check Versions"
$texts[1][18] = "PC-Softwareinfo"
$texts[1][19] = " Exit"
$texts[1][20] = "Name of the User:"
$texts[1][21] = "UserID:"
$texts[1][22] = "Title:"
$texts[1][23] = "F-Number:"
$texts[1][24] = "Released on:"
$texts[1][25] = "Phone:"
$texts[1][26] = "Department:"
$texts[1][27] = "Domain:"
$texts[1][28] = "email:"
$texts[1][29] = "Location:"
$texts[1][30] = "Address:"
$texts[1][31] = "Room:"
$texts[1][32] = "Names of PCs|IP-Adress|Mac-Adress|?|Names of NCs|IP-Adress|Mac-Adress|?"
$texts[1][33] = "Last changed:"
$texts[1][34] = "Valid till:"
$texts[1][35] = "PW req.:"
$texts[1][36] = "PW change:"
$texts[1][37] = "Locked/Deact:"
$texts[1][38] = "Unlock"
$texts[1][39] = "Created on:"
$texts[1][40] = "Errorcount:"
$texts[1][41] = "Last logon:"
$texts[1][42] = "Homedirectory:"
$texts[1][43] = "Homedrive:"
$texts[1][44] = "Loginscript:"
$texts[1][45] = "Profile Path:"
$texts[1][46] = "Born on:"
$texts[1][47] = "TS-Profile:"
$texts[1][48] = "Jobstorage:"
$texts[1][49] = "Rank:"
$texts[1][50] = "PWChange:"
$texts[1][51] = "Sex:"
$texts[1][52] = "SAP No.:"
$texts[1][53] = "Name of Group|samAccountName|Group type|Member of|Scope"
$texts[1][54] = "Surname"
$texts[1][55] = "Christian name"
$texts[1][56] = "Cancel"
$texts[1][57] = "New version is installed."
$texts[1][58] = "Waiting for input"
$texts[1][59] = "Reading user data..."
$texts[1][60] = " has been copied to the clipboard"
$texts[1][61] = "The domain "
$texts[1][62] = " could not be reached."
$texts[1][63] = "The account "
$texts[1][64] = " does not exist in the domain."
$texts[1][65] = "never ever"
$texts[1][66] = "Userdata read (in "
$texts[1][67] = " seconds, with "
$texts[1][68] = " from ADS), Domaincontroller: "
$texts[1][69] = "Choose User"
$texts[1][70] = "Name or UserID to search"
$texts[1][71] = "We've found "
$texts[1][72] = " Users with this criteria. Please select the desired one:"
$texts[1][73] = "32Bit Program under 64 Bit OS -> can't read value!"
$texts[1][74] = "Yes"
$texts[1][75] = "No"
$texts[1][76] = "allowed"
$texts[1][77] = "locked"
$texts[1][78] = "forever"
$texts[1][79] = "This is a distribution group and can not (yet) be interpreted"
$texts[1][80] = "Groupname"
$texts[1][81] = "created"
$texts[1][82] = "last changed"
$texts[1][83] = "mail"
$texts[1][84] = "Group info"
$texts[1][85] = "Description"
$texts[1][86] = "Scope of the group"
$texts[1][87] = "Group class"
$texts[1][88] = "Member count"
$texts[1][89] = "Groupmembers by Name"
$texts[1][90] = "Groupmembers by UserID"
$texts[1][91] = "Groupmembers have been copied to the clipboard"
$texts[1][92] = "Compare with"
$texts[1][93] = "Which UserID shall the user be compared with?"
$texts[1][94] = "Groups which only "
$texts[1][95] = " has"
$texts[1][96] = "Common Groups"
$texts[1][97] = "Comparison of groups completed"
$texts[1][98] = "Account locked"
$texts[1][99] = "Account deactivated"
$texts[1][100] = "Domain controller"
$texts[1][100] = "Names of PCs"
$texts[1][101] = "Names of NCs"
$texts[1][103] = "Gruppen"
$texts[1][104] = "No connection to the PC Database"
$texts[1][105] = "No connection to the NC Database"
$texts[1][106] = "Multiple devices in database - please choose"
$texts[1][107] = "Name of PC|IP-Adress"
$texts[1][108] = "No RealVNC found"
$texts[1][109] = "The expected file at C:\Programme\RealVNC\VNC4\vncviewer.exe"
$texts[1][110] = "-> moved"
$texts[1][111] = "-> not moved"
$texts[1][112] = "Deletion result"
$texts[1][113] = "No directories have been found"
$texts[1][114] = "Show neighbours"
$texts[1][115] = "Room|Name of the User|UserID|Telephone"
$texts[1][116] = "Compare groups with this neighbour"
$texts[1][117] = "Switch to this neighbour"
$texts[1][118] = "Offline"
$texts[1][119] = "The PC is offline"
$texts[1][120] = "Access problem"
$texts[1][121] = "Drive C:\ cannot be accessed"
$texts[1][122] = "Softwarestatus of this PC "
$texts[1][123] = "Software on this PC "
$texts[1][124] = "Usercount per department "
$texts[1][125] = "Could not connect to X500 Database"
Opt("GUICloseOnESC", 0)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
Opt("WinTitleMatchMode", 2)
Opt("TrayIconDebug", 0)
Opt("TrayIconHide", 1)
Global Const $ADS_GROUP_TYPE_GLOBAL_GROUP = 0x2
Global Const $ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 0x4
Global Const $ADS_GROUP_TYPE_UNIVERSAL_GROUP = 0x8
Global Const $ADS_GROUP_TYPE_SECURITY_ENABLED = 0x80000000
Global Const $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = 0x5
Global Const $ADS_UF_ACCOUNTDISABLE = 2
Global Const $USER_CHANGE_PASSWORD = "{AB721A53-1E2f-11D0-9819-00AA0040529b}"
Const $Member_SchemaIDGuid = "{BF9679C0-0DE6-11D0-A285-00AA003049E2}"
Global $objConnection = ObjCreate("ADODB.Connection")
If @error Then
MsgBox(16, "Error", "Could not create ADODB Connection")
Exit
EndIf
$objConnection.ConnectionString = "Provider=ADsDSOObject"
$objConnection.Open("Active Directory Provider")
Global $ADSystemInfo = ObjCreate("ADSystemInfo")
Global $nMsg
$titel = "Userinfo " & FileGetVersion(@AutoItExe)
If @AutoItX64 Then
$titel &= " [64 Bit]"
Else
$titel &= " [32 Bit]"
EndIf
If @Compiled And _Singleton("Userinfo", 1) = 0 Then
WinActivate($titel)
Exit
EndIf
Global Const $guiHeight = 700
Global Const $guiWidth = 1000
Global Const $labelWidth = 90
Global Const $labelHeight = 17
Global Const $valueWidthLong = 300
Global Const $valueWidthShort = 30
Global Const $valueWidthDate = 125
Global Const $valueHeight = 17
Global Const $cellPadding = 5
$col1LabelLeft = 10
$col1ValueLeft = $col1LabelLeft + $labelWidth + $cellPadding
$col2LabelLeft = $col1ValueLeft + $valueWidthLong + $cellPadding
$col2ValueLeft = $col2LabelLeft + $labelWidth + $cellPadding
Global Const $buttonWidth = 140
Global Const $buttonBigHeight = 40
Global Const $buttonNormalHeight = 20
Global Const $buttonPadding = 5
Global Const $buttonLeft = $guiWidth -($buttonWidth +(2 * $buttonPadding))
$frmMain = GUICreate($titel, $guiWidth, $guiHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MAXIMIZEBOX)
GUISetIcon("shell32.dll", -171)
GUICtrlSetFont($frmMain, 6)
GUICtrlCreateGroup("Functions", $guiWidth -($buttonWidth +(3 * $buttonPadding)), $buttonPadding, $buttonWidth +(2 * $buttonPadding), $guiHeight)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)
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
GUICtrlSetTip(-1, "Count the number of users with the same department")
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
GUICtrlSetTip(-1, "Copy the list of groups to clipboard")
$top += $buttonNormalHeight
$btnGroupMembers = GUICtrlCreateButton("Group members", $buttonLeft, $top, $buttonWidth, $buttonNormalHeight)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Show members of selected group")
$btnExit = GUICtrlCreateButton("Exit", $buttonLeft, $guiHeight - 3 * $buttonBigHeight, $buttonWidth, $buttonBigHeight)
GUICtrlSetImage(-1, "shell32.dll", 28)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH)
GUICtrlSetTip(-1, "Guess...try and find out!")
GUICtrlCreateGroup("", -99, -99, 1, 1)
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
$lblExpiration = GUICtrlCreateLabel("", $col1ValueLeft + $valueWidthDate +($labelWidth / 2) + $cellPadding, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)
GUICtrlCreateLabel("PW Required:", $col2LabelLeft, $top, $labelWidth, $labelHeight)
$lblPwRequired = GUICtrlCreateLabel("", $col2ValueLeft, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)
GUICtrlCreateLabel("Change:", $col2ValueLeft + $valueWidthDate + $cellPadding, $top, $labelWidth / 2, $labelHeight)
$lblPWChange = GUICtrlCreateLabel("", $col2ValueLeft + $valueWidthDate +($labelWidth / 2) + $cellPadding, $top, $valueWidthDate, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)
$top += 20
GUICtrlCreateLabel("Locked/Deact:", $col1LabelLeft, $top, $labelWidth, $labelHeight)
$lblAccountLocked = GUICtrlCreateLabel("", $col1ValueLeft, $top, $valueWidthShort, $valueHeight, $SS_SUNKEN)
GUICtrlSetFont(-1, -1, 800)
GUICtrlCreateLabel("/", $cellPadding + $col1ValueLeft + $valueWidthShort, $top, 10, $labelHeight, $SS_CENTER)
$lblDeactivated = GUICtrlCreateLabel("-", $col1ValueLeft + $valueWidthShort +(4 * $cellPadding), $top, $valueWidthShort, $valueHeight, $SS_SUNKEN)
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
$top += 25
$listGroups = GUICtrlCreateListView("Groupname|samAccountName|Group type|Member of|Scope", $col1LabelLeft, $top, 2 * $labelWidth + 2 * $valueWidthLong + 3 * $cellPadding, 400)
_GUICtrlListView_SetColumnWidth($listGroups, 0, 240)
_GUICtrlListView_SetColumnWidth($listGroups, 1, 225)
$sbMain = _GUICtrlStatusBar_Create($frmMain)
GUICtrlSetResizing(-1, $GUI_DOCKSTATEBAR)
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
Dim $a_AccelKeys[1][2] = [["{ENTER}", $Enter_key]]
GUISetAccelerators($a_AccelKeys, $frmSearchAD)
GUICtrlCreateLabel("Domain:", 5, 65, 100, 17)
$cmbDomain = GUICtrlCreateCombo("", 110, 62, 250, 250, $CBS_DROPDOWNLIST)
$btnOk = GUICtrlCreateButton("&Ok", 110, 90, 120, 25)
$btnCancel = GUICtrlCreateButton("E&xit", 240, 90, 120, 25)
$sbSearchAD = _GUICtrlStatusBar_Create($frmSearchAD)
Global $oMyError = ""
$oMyError = ObjEvent("AutoIt.Error", "_ADDoError")
Dim $groups
Global $lw_m = ""
Global $pc = ""
Global $user = ""
Global $objRootDSE
Global $strDNSDomain
Global $strHostServer
Global $strConfiguration
Global $intUAC
Global $oUsr
Global $objRecordSet
Global $time
Global $loopflag
Global $MaxPasswordAge
listDomains()
setUser()
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
MsgBox(262144, "", "We intercepted a COM Error !" & @CRLF & "Number is: " & @TAB & $HexNumber & @CRLF & "Windescription is: " & @TAB & $oMyError.windescription & @CRLF & "err.description is: " & @TAB & $oMyError.description & @CRLF & "err.source is: " & @TAB & $oMyError.source & @CRLF & "err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & "err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF & "err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & "err.retcode is: " & @TAB & $oMyError.retcode & @CRLF & "Script Line number is: " & @TAB & $oMyError.scriptline)
Select
Case $oMyError.windescription = "Access is denied"
$objConnection.Close("Active Directory Provider")
$objConnection.Open("Active Directory Provider")
SetError(2)
Case 1
SetError(1)
EndSelect
EndFunc
Func _ADGetAccount($user)
GUICtrlSetData($lblLastUpdate, Zeit($oUsr.PasswordLastChanged))
If $oUsr.IsAccountLocked Then
GUICtrlSetData($lblAccountLocked, "Yes")
GUICtrlSetColor(-1, $color_red)
Else
GUICtrlSetData($lblAccountLocked, "No")
GUICtrlSetColor(-1, $color_green)
EndIf
GUICtrlSetData($lblBadLoginCount, $oUsr.BadLoginCount)
$oSecDesc = $oUsr.Get("ntSecurityDescriptor")
$oACL = $oSecDesc.DiscretionaryACL
For $oACE In $oACL
If($oACE.ObjectType = $USER_CHANGE_PASSWORD) And(($oACE.Trustee = "Everyone") Or($oACE.Trustee = "Iedereen")) Then
If($oACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT) Then
GUICtrlSetData($lblPWChange, "Allowed")
GUICtrlSetColor($lblPWChange, $color_green)
Else
GUICtrlSetData($lblPWChange, "Locked")
GUICtrlSetColor($lblPWChange, $color_red)
EndIf
EndIf
Next
$intUAC = $oUsr.Get("userAccountControl")
If BitAND(0x00020, $intUAC) Then
GUICtrlSetData($lblPwRequired, "No")
Else
GUICtrlSetData($lblPwRequired, "Yes")
EndIf
GUICtrlSetData($lblCreatedDate, Zeit($oUsr.whenCreated))
GUICtrlSetData($lblLogonScript, $oUsr.scriptPath)
If BitAND($intUAC, $ADS_UF_ACCOUNTDISABLE) Then
GUICtrlSetData($lblDeactivated, "Yes")
GUICtrlSetColor($lblDeactivated, $color_red)
Else
GUICtrlSetData($lblDeactivated, "No")
GUICtrlSetColor($lblDeactivated, $color_green)
EndIf
GUICtrlSetColor($lblExpiration, 0x000000)
$dummy = $oUsr.AccountExpirationDate
$tmp = Zeit2($dummy)
$dummy = Zeit($dummy)
$tmp2 = _DateDiff("D", _NowCalcDate(), $tmp)
If($tmp2 < 1) And($tmp2 > -148883) Then GUICtrlSetColor($lblExpiration, $color_red)
If($dummy = "01.01.1601 02:00") Or($dummy = "01.01.1970 00:00") Or($dummy = "01.01.1601 01:00") Then
$dummy = " forever"
GUICtrlSetColor($lblExpiration, 0x000000)
EndIf
GUICtrlSetData($lblExpiration, $dummy)
EndFunc
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
$arrayMembers[$i][0] = StringReplace($objmember.displayname, " (LBV)", "")
$arrayMembers[$i][1] = $objmember.samAccountName
$i = $i + 1
Next
ReDim $arrayMembers[$i][2]
_ArraySort($arrayMembers, 0, 0, 0, 1)
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
EndFunc
Func _ADGetUserData($user)
$oUsr.GetInfo()
$tmp = $oUsr.DisplayName
If $tmp <> "" Then
GUICtrlSetData($lblName, $oUsr.DisplayName)
Else
GUICtrlSetData($lblName, $oUsr.sn & ", " & $oUsr.givenName)
EndIf
$tmp = $oUsr.distinguishedName
$tmp = StringReplace($tmp, "\,", "ÃµÃµ")
$tmp = StringSplit($tmp, ",")
$tmp2 = ""
For $i = 1 To $tmp[0]
$tmp[$i] = StringReplace($tmp[$i], "ÃµÃµ", ",")
$tmp[$i] = StringMid($tmp[$i], 4)
Next
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
EndFunc
Func _ADGetUserGroups(ByRef $usergroups, $user = @UserName)
$usergroups = $oUsr.GetEx("memberof")
$count = UBound($usergroups)
If $count = 0 Then
GUICtrlSetData($btnCopyGroups, "Copy Group List")
Return
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
$tmp = StringReplace($objGroup.name, "\/", "/")
$tmp = StringReplace($tmp, "CN=", "") & "|" & $objGroup.samAccountName & "|" & $tmp2 & "|" & $tmp3[1] & "|" & $tmp4
GUICtrlCreateListViewItem($tmp, $listGroups)
Next
GUICtrlSetData($btnCopyGroups, "Copy Group List (" & $count & ")")
EndFunc
Func _ADObjectExists($object)
$flag = 0
$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sAMAccountName=" & $object & "));ADsPath;subtree"
$objRecordSet = $objConnection.Execute($strQuery)
Switch $objRecordSet.RecordCount
Case 0
Case 1
GUISetState(@SW_SHOW, $frmSearchAD)
Return 1
Case Else
If GUICtrlRead($rdbUserId) = $GUI_CHECKED Then $flag = 1
EndSwitch
If $flag = 0 Then
$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(sn=" & $object & "));ADsPath;subtree"
$objRecordSet = $objConnection.Execute($strQuery)
Switch $objRecordSet.RecordCount
Case 0
Case 1
GUISetState(@SW_SHOW, $frmSearchAD)
Return 1
Case Else
If GUICtrlRead($rdbSurname) = $GUI_CHECKED Then $flag = 1
EndSwitch
EndIf
If($flag = 0) And(GUICtrlRead($rdbFirstName) = $GUI_CHECKED) Then
$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(givenName=" & $object & "));ADsPath;subtree"
$objRecordSet = $objConnection.Execute($strQuery)
EndIf
If $objRecordSet.RecordCount = 0 Then
If StringIsDigit($object) Then $object = "*" & $object
$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(telephoneNumber=" & $object & "));ADsPath;subtree"
$objRecordSet = $objConnection.Execute($strQuery)
EndIf
Switch $objRecordSet.RecordCount
Case 0
Return 0
Case 1
GUISetState(@SW_SHOW, $frmSearchAD)
Return 1
Case Else
Dim $treffer_arry[1]
$z = ""
Do
$y = $objRecordSet.Fields(0).Value
If Not StringInStr($y, "ou=Recipients") Then
$oUsr = ObjGet($objRecordSet.Fields(0).Value)
_ArrayAdd($treffer_arry, $oUsr.sn & "," & $oUsr.givenName & "|" & $oUsr.samAccountName)
EndIf
$objRecordSet.MoveNext
Until $objRecordSet.EOF
If UBound($treffer_arry) = 2 Then
$x = StringInStr($treffer_arry[1], "|")
GUICtrlSetData($edit, StringMid($treffer_arry[1], $x + 1))
GUISetState(@SW_SHOW, $frmSearchAD)
Return -1
Else
GUISetState(@SW_HIDE, $frmSearchAD)
$frmUserSelect = GUICreate("Choose User", 350, 355)
GUISetIcon("shell32.dll", -171)
$Enter_key2 = GUICtrlCreateDummy()
Dim $b_AccelKeys[1][2] = [["{ENTER}", $Enter_key2]]
GUISetAccelerators($b_AccelKeys, $frmUserSelect)
$userList = GUICtrlCreateListView("Name or UserID to search", 5, 40, 340, 280)
_GUICtrlListView_SetColumnWidth($userList, 0, 250)
$btn_userwahl = GUICtrlCreateButton("Ok", 5, 325, 340, 25)
GUICtrlCreateLabel(UBound($treffer_arry) - 1 & " users found. Please select one:", 5, 5, 290, 30)
_ArrayDelete($treffer_arry, 0)
_ArraySort($treffer_arry)
For $i = 0 To UBound($treffer_arry) - 1
GUICtrlCreateListViewItem($treffer_arry[$i], $userList)
Next
GUISetState(@SW_SHOW, $frmUserSelect)
While 1
$msg = GUIGetMsg()
If $msg = $GUI_EVENT_CLOSE Then Exit
If($msg = $btn_userwahl) Or($msg = $Enter_key2) Then
$x = GUICtrlRead($userList)
If $x <> "" Then
$y = GUICtrlRead($userList)
$y = GUICtrlRead($y)
$tmp = StringSplit($y, "|")
$x = GUICtrlSetData($edit, $tmp[2])
ExitLoop
EndIf
EndIf
WEnd
GUIDelete($frmUserSelect)
GUISetState(@SW_SHOW, $frmSearchAD)
Return -1
EndIf
EndSwitch
EndFunc
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
EndFunc
Func compareUsers($compareUser = "")
Local $strQuery2, $objRecordSet2, $ldap_entry2
Dim $groupCompare
If $compareUser = "" Then
$compareUser = InputBox($texts[$language][92], $texts[$language][93], "", "", 300, 120)
If @error Then Return
EndIf
SplashTextOn("", "Please wait while minions are rumbling through the AD", "-1", "-1", "-1", "-1", 33, "", "", "")
$strQuery2 = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(sAMAccountName=" & $compareUser & ");ADsPath;subtree"
$objRecordSet2 = $objConnection.Execute($strQuery2)
$ldap_entry2 = $objRecordSet2.fields(0).value
$oUsr2 = ObjGet($ldap_entry2)
$groupCompare = $oUsr2.GetEx("memberof")
$oUsr2 = 0
$count = UBound($groupCompare)
If $count = 0 Then Return
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
For $i = 0 To _GUICtrlListView_GetItemCount($listGroups) - 1
$flag = 0
$name1 = _GUICtrlListView_GetItem($listGroups, $i, 0)
$name1 = $name1[3]
For $name2 In $groupCompare
If $name1 = $name2 Then
$flag = 1
$commonGroups = $commonGroups & $name1 & @CRLF
ExitLoop
EndIf
Next
If $flag = 0 Then $missing1 = $missing1 & $name1 & @CRLF
Next
For $name2 In $groupCompare
$flag = 0
For $i = 0 To _GUICtrlListView_GetItemCount($listGroups) - 1
$name1 = _GUICtrlListView_GetItem($listGroups, $i, 0)
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
_ArrayDelete($missing2, 0)
_ArrayDelete($missing2, 0)
$max = 0
$m1 = UBound($missing1)
$m2 = UBound($missing2)
$m3 = UBound($commonGroups)
$max = _Max($max, $m1)
$max = _Max($max, $m2)
$max = _Max($max, $m3)
SplashOff()
$data = FileOpen(@TempDir & "\userCompare.html", 2)
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
FileWriteLine($data, "    <thead>")
FileWriteLine($data, "      <tr><th>Groups which only " & $user & " has</th> <th>Common Groups</th><th>Groups which only " & $compareUser & " has</th></tr>")
FileWriteLine($data, "    </thead>")
FileWriteLine($data, "    <tbody>")
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
FileWriteLine($data, "<tfoot> <tr> <td>" &($m1 / 2) - 1 & "</td> <td>" &($m3 / 2) - 1 & "</td> <td>" &($m2 / 2) - 1 & "</td> </tr> </tfoot> ")
FileWriteLine($data, "  </table>")
FileWriteLine($data, "</body></html>")
FileClose($data)
ShellExecute(@TempDir & "\userCompare.html")
_GUICtrlStatusBar_SetText($sbSearchAD, "Comparison of groups completed")
EndFunc
Func copyGroups()
For $i = 0 To UBound($groups) - 1
$tmp = StringSplit($groups[$i], ",")
$tmp[1] = StringReplace($tmp[1], "CN=", "")
$groups[$i] = $tmp[1]
Next
_ArrayToClip($groups, @CRLF, 1)
EndFunc
Func copyToClipboard($clickValue)
If $clickValue <> "" Then
While StringLeft($clickValue, 1) = "|"
$clickValue = StringMid($clickValue, 2)
WEnd
$clickValue = StringReplace($clickValue, "|", @CRLF)
$clickValue = StringReplace($clickValue, @CRLF & @CRLF, "")
ClipPut($clickValue)
_GUICtrlStatusBar_SetText($sbMain, "[ " & $clickValue & "] has been copied to the clipboard")
$x = ControlGetFocus($titel)
If Not StringInStr($x, "SysListView32") Then
GUICtrlSetBkColor($nMsg, $color_red)
Sleep(200)
GUICtrlSetBkColor($nMsg, -1)
EndIf
EndIf
EndFunc
Func countDepartmentUsers($department)
$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(department=" & $department & "));ADsPath;subtree"
$objRecordSet = $objConnection.Execute($strQuery)
Return $objRecordSet.RecordCount
EndFunc
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
$strDNSDomain = $objRootDSE.Get("defaultNamingContext")
$strHostServer = $objRootDSE.Get("dnsHostName")
$strConfiguration = $objRootDSE.Get("ConfigurationNamingContext")
Switch _ADObjectExists($user)
Case 1
$loopflag = 0
$x = ObjGet("WinNT://" & $domain)
$MaxPasswordAge = $x.MaxPasswordAge
Return 1
Case -1
ControlClick($titel, "", $btnOk)
Return 0
Case 0
_GUICtrlStatusBar_SetText($sbSearchAD, "The account " & $user & " does not exist in this domain.")
$loopflag = 1
Return 0
EndSwitch
EndIf
EndFunc
Func GUILoop()
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
If $x <> "" Then
$y = GUICtrlRead($x)
$tmp = StringSplit($y, "|")
_ADGetGroupMembers($tmp[2])
Else
If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
$iMsgBoxAnswer = MsgBox(8192, "", "Please select a group from the list to retrieve the members", 10)
Select
Case $iMsgBoxAnswer = -1
Case Else
EndSelect
EndIf
Case $btnCompareUsers
compareUsers()
Case $btnRefreshQuery
$time = TimerInit()
queryAD()
Case $btnNewQuery
setUser()
Case $edit
Case $btnDepartmentUsers
searchDepartmentUsers()
Case $btnCountDeptUsers
MsgBox(4160, "", "There are " & countDepartmentUsers(GUICtrlRead($lblDepartment)) & " people working at " & GUICtrlRead($lblDepartment))
Case -100 To 0
Case Else
$clickValue = GUICtrlRead($nMsg)
copyToClipboard($clickValue)
EndSwitch
WEnd
EndFunc
Func listDomains()
$oDS = ObjGet("WinNT:")
$y = ""
For $x In $oDS
$y = $y & "|" & $x.Name
Next
GUICtrlSetData($cmbDomain, $y, "US")
EndFunc
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
Local $tmp = $oUsr.scriptpath
If $lastlogon = "-" Then
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
EndFunc
Func searchDepartmentUsers()
$office = $oUsr.physicalDeliveryOfficeName
$department = $oUsr.department
Local $spacer = 5
Local $frmWidth = 650
Local $frmHeight = 550
Local $btnWidth = $frmWidth -(2 * $spacer)
Local $btnHeight = 30
Local $lvWidth = $frmWidth -(2 * $spacer)
Local $lvHeight =($frmHeight - $spacer) -(3 *($btnHeight + $spacer))
$frmColleagues = GUICreate("Colleagues", $frmWidth, $frmHeight)
GUISetIcon("shell32.dll", -171)
$userList = GUICtrlCreateListView("Username|UserID|Telephone|Office|Department", $spacer, $spacer, $btnWidth, $lvHeight, $LVS_SORTASCENDING + $LVS_SINGLESEL + $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth($userList, 0, 150)
_GUICtrlListView_SetColumnWidth($userList, 1, $LVSCW_AUTOSIZE)
$btnColleagueCompare = GUICtrlCreateButton("Compare groups with this colleague", $spacer,($frmHeight - $spacer) -(3 * $btnHeight), $btnWidth, $btnHeight)
$btnColleagueSwitch = GUICtrlCreateButton("Switch to this colleague", $spacer,($frmHeight - $spacer) -(2 * $btnHeight), $btnWidth, $btnHeight)
$btnColleagueOk = GUICtrlCreateButton("Ok", $spacer,($frmHeight - $spacer) -(1 * $btnHeight), $btnWidth, $btnHeight)
Dim $arrayFound[1]
$strQuery = "<LDAP://" & $strHostServer & "/" & $strDNSDomain & ">;(&(objectCategory=person)(objectclass=user)(department=" & $department & "));ADsPath;subtree"
$objRecordSet = $objConnection.Execute($strQuery)
Do
$y = $objRecordSet.Fields(0).Value
If Not StringInStr($y, "ou=Recipients") Then
$o_temp_Usr = ObjGet($objRecordSet.Fields(0).Value)
$x = GUICtrlCreateListViewItem($o_temp_Usr.sn & "," & $o_temp_Usr.givenName & "|" & $o_temp_Usr.samAccountName & "|" & $o_temp_Usr.telephoneNumber & "|" & $o_temp_Usr.physicalDeliveryOfficeName & "|" & $department, $userList)
EndIf
$objRecordSet.MoveNext
Until $objRecordSet.EOF
GUISetState(@SW_SHOW, $frmColleagues)
While 1
$msg = GUIGetMsg()
If($msg = $btnColleagueSwitch) Or($msg = $btnColleagueCompare) Then
$x = GUICtrlRead($userList)
If $x <> "" Then
$y = GUICtrlRead($x)
$tmp = StringSplit($y, "|")
If $msg = $btnColleagueSwitch Then
GUIDelete($frmColleagues)
setUser($tmp[2])
Else
GUIDelete($frmColleagues)
compareUsers($tmp[2])
EndIf
ExitLoop
Else
MsgBox(64, "", "Please select a user")
EndIf
EndIf
If($msg = $GUI_EVENT_CLOSE) Or($msg = $btnColleagueOk) Then
GUISetState(@SW_SHOW, $frmMain)
GUIDelete($frmColleagues)
ExitLoop
EndIf
WEnd
EndFunc
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
$ldap_entry = $objRecordSet.fields(0).value
$oUsr = ObjGet($ldap_entry)
$user = $oUsr.samAccountName
GUISetState(@SW_SHOW, $frmMain)
GUISetCursor(15, 1, $frmMain)
_GUICtrlStatusBar_SetText($sbSearchAD, "Reading user data...")
queryAD()
GUISetCursor(2)
GUILoop()
EndFunc
Func Zeit($zeit)
If $zeit = "" Then Return "---"
$tmp = StringMid($zeit, 7, 2) & "." & StringMid($zeit, 5, 2) & "." & StringLeft($zeit, 4) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2)
Return $tmp
EndFunc
Func Zeit2($zeit)
$tmp = StringLeft($zeit, 4) & "/" & StringMid($zeit, 5, 2) & "/" & StringMid($zeit, 7, 2) & " " & StringMid($zeit, 9, 2) & ":" & StringMid($zeit, 11, 2) & ":" & StringMid($zeit, 13, 2)
Return $tmp
EndFunc
