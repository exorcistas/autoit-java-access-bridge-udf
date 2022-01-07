#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: JavaAccessBridge_Extended
    Description...........: Access Bridge extended functions for Core/CoreEx
	
	Dependencies..........: <JavaAccessBridge_Core.au3>; 
	                        <JavaAccessBridge_CoreEx.au3>; 
							<JavaAccessBridge_Globals.au3>

    Author................: exorcistas@github.com
    Modified..............: 2020-03-12
    Version...............: v0.1.2b
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <JavaAccessBridge_CoreEx.au3>

#Region EXTENDED_FUNCTIONS_LIST
#cs	===================================================================================================================================
	_JAB_Ext_RunJavaApplication($_sFullExePath)
	_JAB_Ext_AttachToWindow($_sWinTitle, $_iTimeout = 30000)
	_JAB_Ext_InvokeAction($_oVMID, $_oAC, $_sActionName)
	_JAB_Ext_GetChildAC($_oVMID, $_oParentAC, $_index = 0)
	_JAB_Ext_GetParentAC($_oVMID, $_oAC)
	_JAB_Ext_GetSiblingAC($_oVMID, $_oAC, $_RelativeIndex = 1)
	_JAB_Ext_GetACProperty($_oVMID, $_oAC, $_sPropertyName)
	_JAB_Ext_FindACByProperties($_oVMID, $_oParentAC, $_sProperties)
	_JAB_Ext_PasteValueFromClipboard($_oVMID, $_oAC, $sValue)
	_JAB_Ext_CopyValueToClipboard($_oVMID, $_oAC)
#ce	===================================================================================================================================
#EndRegion CORE-EXTENDED_FUNCTIONS_LIST

#Region EXTENDED_FUNCTIONS

	Func _JAB_Ext_RunJavaApplication($_sFullExePath)
		If NOT FileExists($_sFullExePath) Then Return SetError(1, 0, 0)

		Local $_pID = ShellExecute($_sFullExePath)
		Return $_pID
	EndFunc

	Func _JAB_Ext_AttachToWindow($_sWinTitle, $_iTimeout = 30000)
		Local $_hTimer = TimerInit()
		Do
			$_hwnd = __JAB_GetWindowHandle($_sWinTitle)
				If @error Then Return SetError(1, @error, 0)
			Sleep(100)
		Until (($_hwnd <> 0) OR (TimerDiff($_hTimer) > $_iTimeout))

		If NOT __JAB_IsJavaWindow($_hwnd) Then Return SetError(2, @error, 0)

		Return $_hwnd
	EndFunc

	Func _JAB_Ext_InvokeAction($_oVMID, $_oAC, $_sActionName)
		Local $_aActions = _JAB_GetAccessibleActions($_oVMID, $_oAC)

			If $_JAB_DEBUG = $_JAB_DEBUG_EXT Then ConsoleWrite(@CRLF & ">> [_JAB_Ext_InvokeAction]: " & $_sActionName & @CRLF)
			
		Local $_index = _ArraySearch($_aActions, $_sActionName)
				If @error Then Return SetError(@error, 0, 0)
		
		_JAB_DoAccessibleActions($_oVMID, $_oAC, $_aActions[$_index][1])
	EndFunc

	Func _JAB_Ext_GetChildAC($_oVMID, $_oParentAC, $_index = 0)
		Local $_aChildAC = _JAB_GetAccessibleChildFromContext($_oVMID, $_oParentAC, $_index)
			If @error Then Return SetError(@error, 0, 0)
		Local $_oChildAC = $_aChildAC[0][1]

		Return $_oChildAC
	EndFunc

	Func _JAB_Ext_GetParentAC($_oVMID, $_oAC)
		Local $_aParentAC = _JAB_GetAccessibleParentFromContext($_oVMID, $_oAC)
			If @error Then Return SetError(@error, 0, 0)
		Local $_oParentAC = $_aParentAC[0][1]

		Return $_oParentAC
	EndFunc

	Func _JAB_Ext_GetSiblingAC($_oVMID, $_oAC, $_RelativeIndex = 1)
		Local $_iCurrentIndex = _JAB_Ext_GetACProperty($_oVMID, $_oAC, "indexInParent")
			If @error Then Return SetError(1, @error, 0)

		Local $_oParentAC = _JAB_Ext_GetParentAC($_oVMID, $_oAC)
			If @error Then Return SetError(2, @error, 0)

		Local $_iChildrenCount = _JAB_Ext_GetACProperty($_oVMID, $_oAC, "childrenCount")
			If @error Then Return SetError(3, @error, 0)

			If $_JAB_DEBUG = $_JAB_DEBUG_EXT Then 
				ConsoleWrite(@CRLF & ">> [_JAB_Ext_GetSiblingAC]: " & @CRLF & _ 
									">> Parent childrenCount: " & $_iChildrenCount & _
									">> Child indexInParent: " & $_iCurrentIndex & @CRLF)
			EndIf

		;-- if range matches
		If (($_iCurrentIndex + $_RelativeIndex >= 0) AND ($_iCurrentIndex + $_RelativeIndex <= $_iChildrenCount-1)) Then
			Local $_oSiblingAC = _JAB_Ext_GetChildAC($_oVMID, $_oParentAC, $_iCurrentIndex + $_RelativeIndex)
				If @error Then Return SetError(4, @error, 0)
			Return $_oSiblingAC
		EndIf

		Return SetError(1, 0, 0)
	EndFunc

	Func _JAB_Ext_GetACProperty($_oVMID, $_oAC, $_sPropertyName)
		Local $_aACInfo = _JAB_GetAccessibleContextInfo($_oVMID, $_oAC)
			If @error Then Return SetError(1, @error, 0)

		Local $_index = _ArraySearch($_aACInfo, $_sPropertyName)
			If @error Then Return SetError(2, @error, 0)
		Local $_sPropertyValue = $_aACInfo[$_index][1]

			If $_JAB_DEBUG = $_JAB_DEBUG_EXT Then 
				ConsoleWrite(@CRLF & ">> [_JAB_Ext_GetACProperty]: " & @CRLF & _ 
									">> " & $_sPropertyName & ": " & $_sPropertyValue & @CRLF)
			EndIf

		Return $_sPropertyValue
	EndFunc

	Func _JAB_Ext_FindACByProperties($_oVMID, $_oParentAC, $_sProperties)
		Local $_iChildrenCount = _JAB_Ext_GetACProperty($_oVMID, $_oParentAC, "childrenCount")
			If @error Then Return SetError(1, @error, 0)
			If $_iChildrenCount = 0 Then Return 0

		;-- create property filter:
		Local $_aProperties = StringSplit($_sProperties, ";")
		Local $_aFilter[0][2]
		For $i = 1 To $_aProperties[0]
			_ArrayAdd($_aFilter, $_aProperties[$i], 0, ":", "|")
		Next

		Local $_oFoundAC = 0
		Local $_iPropMatch = 0
		Local $_oChildAC = $_oFoundAC

		;-- iterate through AC children:
		For $i = 0 To $_iChildrenCount-1

			$_oChildAC = _JAB_Ext_GetChildAC($_oVMID, $_oParentAC, $i)
				If @error Then Return SetError(2, @error, 0)
			;-- iterate property filter:
			For $z = 0 To Ubound($_aFilter)-1
				If _JAB_Ext_GetACProperty($_oVMID, $_oChildAC, $_aFilter[$z][0]) = $_aFilter[$z][1] Then $_iPropMatch = $_iPropMatch + 1
					If @error Then Return SetError(3, @error, 0)
			Next
			If ($_iPropMatch = Ubound($_aFilter)) Then Return $_oChildAC

			$_iPropMatch = 0
			;-- if not matched, look in children:
			$_oChildAC = _JAB_Ext_FindACByProperties($_oVMID, $_oChildAC, $_sProperties)
				If @error Then Return SetError(4, @error, 0)
			If $_oChildAC Then Return $_oChildAC
		Next

		Return 0
	EndFunc

	Func _JAB_Ext_PasteValueFromClipboard($_oVMID, $_oAC, $sValue)
		ClipPut($sValue)
		Local $_aActionsToDo[3]
		$_aActionsToDo[0] = "select-all"
		$_aActionsToDo[1] = "set-writable"
		$_aActionsToDo[2] = "paste-from-clipboard"
		_JAB_DoAccessibleActions($_oVMID, $_oAC, $_aActionsToDo)
			If @error Then Return SetError(@error, 0, 0)
	EndFunc

	Func _JAB_Ext_CopyValueToClipboard($_oVMID, $_oAC)
		ClipPut("")

		Local $_aActionsToDo[2]
		$_aActionsToDo[0] = "select-all"
		$_aActionsToDo[1] = "copy-to-clipboard"
		_JAB_DoAccessibleActions($_oVMID, $_oAC, $_aActionsToDo)
			If @error Then Return SetError(@error, 0, 0)
		
		$_sValue = ClipGet()
		Return $_sValue
	EndFunc

#EndRegion EXTENDED_FUNCTIONS