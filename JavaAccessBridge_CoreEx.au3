#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: JavaAccessBridge_CoreEx
    Description...........: Access Bridge Core DLL expanded (mediator) functions
	
    Dependencies..........: <JavaAccessBridge_Core.au3>; <JavaAccessBridge_Globals.au3>
    Documentation.........: https://docs.oracle.com/javase/9/access/jaapi.htm
                            https://docs.oracle.com/javase/8/docs/api/javax/accessibility/package-summary.html
                            https://docs.oracle.com/javase/accessbridge/

    Author................: exorcistas@github.com
    Modified..............: 2020-03-11
    Version...............: v0.1.2b
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <WinAPI.au3>
#include <Memory.au3>
#include <JavaAccessBridge_Core.au3>

#Region CORE-EX_FUNCTIONS_LIST
#cs	===================================================================================================================================
%% INITIALIZATION/SHUTDOWN %%
    _JAB_Startup($_bDebug = $_JAB_DEBUG)
    _JAB_CloseDown()

%% GATEWAY_EX %%
    _JAB_GetAccessibleContextFromHWND($_hwnd)

%% GENERAL_EX %%
    _JAB_GetVersionInfo($_vmID)

%% ACCESSIBLE_CONTEXT-EX %%
    _JAB_GetAccessibleContextInfo($_oVMID, $_oAccessibleContext)
    _JAB_GetAccessibleChildFromContext($_oVMID, $_oAccessibleContext, $_index = 0)
    _JAB_GetAccessibleParentFromContext($_oVMID, $_oAccessibleContext)

%% ACCESSIBLE_TEXT-EX %%
    _JAB_GetAccessibleTextInfo($_oVMID, $_oAccessibleContext, $xPos = 0, $yPos = 0)

%% ADDITIONAL_TEXT-EX %%
    _JAB_SelectTextRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex)

%% ACCESSIBLE_TABLE-EX %%
    _JAB_GetAccessibleTableInfo($_oVMID, $_oParentAccessibleContext)
    _JAB_GetAccessibleTableCellInfo($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn)

%% ACCESSIBLE_ACTIONS-EX %%
    _JAB_GetAccessibleActions($_oVMID, $_oAccessibleContext)
    _JAB_DoAccessibleActions($_oVMID, $_oAccessibleContext, $_vActionsToDo)

%% UTILITY-EX %%
    _JAB_GetTopLevelObject($_oVMID, $_oAccessibleContext)
    _JAB_GetObjectDepth($_oVMID, $_oAccessibleContext)
    _JAB_GetActiveDescendent($_oVMID, $_oAccessibleContext)
    _JAB_GetVisibleChildren($_oVMID, $_oAccessibleContext, $_iStartIndex = 0)
#ce	===================================================================================================================================
#EndRegion CORE-EX_FUNCTIONS_LIST

#Region CORE-EX_FUNCTIONS

    #Region INITIALIZATION/SHUTDOWN-EX

        Func _JAB_Startup($_bDebug = $_JAB_DEBUG)

            __JAB_SetEnvironment($_bDebug)
            __JAB_OpenBridgeDLL()
            __JAB_EnableSwitch()
            __JAB_Windows_run()

        EndFunc

        Func _JAB_CloseDown()
            ;~ __JAB_EnableSwitch(False)
           __JAB_CloseBridgeDLL()
        EndFunc

    #EndRegion INITIALIZATION/SHUTDOWN-EX

    #Region GATEWAY-EX

        Func _JAB_GetAccessibleContextFromHWND($_hwnd)
            Local $_oVMID, $_oAccessibleContext
            Local $_aResult = __JAB_GetAccessibleContextFromHWND($_hwnd, $_oVMID, $_oAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)
            
                Local $_aAccContext[0][2]
                    _ArrayAdd($_aAccContext, "Success|" & $_aResult[0])
                    _ArrayAdd($_aAccContext, "HWND|" & $_aResult[1])
                    _ArrayAdd($_aAccContext, "vmID|" & $_aResult[2])
                    _ArrayAdd($_aAccContext, "AccessibleContext|" & $_aResult[3])

                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleContextFromHWND]: " & $_aAccContext[0][1] & @CRLF & _
                                        ">> Window handle:    " & $_aAccContext[1][1] & @CRLF & _
                                        ">> Virtual Machine ID (vmID):    " & $_aAccContext[2][1] & @CRLF & _
                                        ">> Accessible Context (AC):    " & $_aAccContext[3][1] & @CRLF)
                EndIf

            Return $_aAccContext
        EndFunc

    #EndRegion GATEWAY_EX

    #Region GENERAL-EX
        
        Func _JAB_GetVersionInfo($_vmID)
            Local $_tAccessBridgeVersionInfo = DllStructCreate($tagAccessBridgeVersionInfo)
            Local $_ptr_AccessBridgeVersionInfo = DllStructGetPtr($_tAccessBridgeVersionInfo)

            Local $_aResult = __JAB_GetVersionInfo($_vmID, $_ptr_AccessBridgeVersionInfo)
                If @error Then Return SetError(@error, 0, 0)

            Local $_aVersionInfo[0][2]
                _ArrayAdd($_aVersionInfo, "$tAccessBridgeVersionInfo|" & $_aResult[2])
                _ArrayAdd($_aVersionInfo, "VMVersion|" & DllStructGetData($_tAccessBridgeVersionInfo, "VMVersion"))
                _ArrayAdd($_aVersionInfo, "bridgeJavaClassVersion|" & DllStructGetData($_tAccessBridgeVersionInfo, "bridgeJavaClassVersion"))
                _ArrayAdd($_aVersionInfo, "bridgeJavaDLLVersion|" & DllStructGetData($_tAccessBridgeVersionInfo, "bridgeJavaDLLVersion"))
                _ArrayAdd($_aVersionInfo, "bridgeWinDLLVersion|" & DllStructGetData($_tAccessBridgeVersionInfo, "bridgeWinDLLVersion"))
                
                If $_JAB_DEBUG Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_CEx_GetVersionInfo]: " & $_aVersionInfo[0][1] & @CRLF & _
                                        ">> VMVersion: " & $_aVersionInfo[1][1] & @CRLF & _
                                        ">> bridgeJavaClassVersion: " & $_aVersionInfo[2][1] & @CRLF & _
                                        ">> bridgeJavaDLLVersion: " & $_aVersionInfo[3][1] & @CRLF & _
                                        ">> bridgeWinDLLVersion: " & $_aVersionInfo[4][1] & @CRLF)
                EndIf
            
            Return $_aVersionInfo
        EndFunc

    #EndRegion GENERAL_EX

    #Region ACCESSIBLE_CONTEXT-EX

        Func _JAB_GetAccessibleContextInfo($_oVMID, $_oAccessibleContext)
            Local $_tAccessibleContextInfo = DllStructCreate($tagAccessibleContextInfo)
            Local $_ptr_AccessibleContextInfo = DllStructGetPtr($_tAccessibleContextInfo)

            Local $_aResult = __JAB_GetAccessibleContextInfo($_oVMID, $_oAccessibleContext, $_ptr_AccessibleContextInfo)
                If @error Then Return SetError(@error, 0, 0)
            
            Local $_aAccContextInfo[0][2]
                _ArrayAdd($_aAccContextInfo, "name|" & DllStructGetData($_tAccessibleContextInfo, "name"))
                _ArrayAdd($_aAccContextInfo, "description|" & DllStructGetData($_tAccessibleContextInfo, "description"))
                _ArrayAdd($_aAccContextInfo, "role|" & DllStructGetData($_tAccessibleContextInfo, "role"))
                _ArrayAdd($_aAccContextInfo, "role_en_US|" & DllStructGetData($_tAccessibleContextInfo, "role_en_US"))
                _ArrayAdd($_aAccContextInfo, "states|" & DllStructGetData($_tAccessibleContextInfo, "states"))
                _ArrayAdd($_aAccContextInfo, "states_en_US|" & DllStructGetData($_tAccessibleContextInfo, "states_en_US"))
                _ArrayAdd($_aAccContextInfo, "indexInParent|" & DllStructGetData($_tAccessibleContextInfo, "indexInParent"))
                _ArrayAdd($_aAccContextInfo, "childrenCount|" & DllStructGetData($_tAccessibleContextInfo, "childrenCount"))
                _ArrayAdd($_aAccContextInfo, "xCoord|" & DllStructGetData($_tAccessibleContextInfo, "x"))
                _ArrayAdd($_aAccContextInfo, "yCoord|" & DllStructGetData($_tAccessibleContextInfo, "y"))
                _ArrayAdd($_aAccContextInfo, "width|" & DllStructGetData($_tAccessibleContextInfo, "width"))
                _ArrayAdd($_aAccContextInfo, "height|" & DllStructGetData($_tAccessibleContextInfo, "height"))
                _ArrayAdd($_aAccContextInfo, "accessibleComponent|" & DllStructGetData($_tAccessibleContextInfo, "accessibleComponent"))
                _ArrayAdd($_aAccContextInfo, "accessibleAction|" & DllStructGetData($_tAccessibleContextInfo, "accessibleAction"))
                _ArrayAdd($_aAccContextInfo, "accessibleSelection|" & DllStructGetData($_tAccessibleContextInfo, "accessibleSelection"))
                _ArrayAdd($_aAccContextInfo, "accessibleText|" & DllStructGetData($_tAccessibleContextInfo, "accessibleText"))
                _ArrayAdd($_aAccContextInfo, "accessibleInterfaces|" & DllStructGetData($_tAccessibleContextInfo, "accessibleInterfaces"))
            
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleContextInfo]: " & $_aResult[0] & @CRLF & _
                                        ">> name: " & $_aAccContextInfo[0][1] & @CRLF & _
                                        ">> description: " & $_aAccContextInfo[1][1] & @CRLF & _
                                        ">> role: " & $_aAccContextInfo[2][1] & @CRLF & _
                                        ">> role_en_US: " & $_aAccContextInfo[3][1] & @CRLF & _
                                        ">> states: " & $_aAccContextInfo[4][1] & @CRLF & _
                                        ">> states_en_US: " & $_aAccContextInfo[5][1] & @CRLF & _
                                        ">> indexInParent: " & $_aAccContextInfo[6][1] & @CRLF & _
                                        ">> childrenCount: " & $_aAccContextInfo[7][1] & @CRLF & _
                                        ">> xCoord: " & $_aAccContextInfo[8][1] & @CRLF & _
                                        ">> yCoord: " & $_aAccContextInfo[9][1] & @CRLF & _
                                        ">> width: " & $_aAccContextInfo[10][1] & @CRLF & _
                                        ">> height: " & $_aAccContextInfo[11][1] & @CRLF & _
                                        ">> accessibleComponent: " & $_aAccContextInfo[12][1] & @CRLF & _
                                        ">> accessibleAction: " & $_aAccContextInfo[13][1] & @CRLF & _
                                        ">> accessibleSelection: " & $_aAccContextInfo[14][1] & @CRLF & _
                                        ">> accessibleText: " & $_aAccContextInfo[15][1] & @CRLF & _
                                        ">> accessibleInterfaces: " & $_aAccContextInfo[16][1] & @CRLF)
                EndIf

            Return $_aAccContextInfo
        EndFunc

        Func _JAB_GetAccessibleChildFromContext($_oVMID, $_oAccessibleContext, $_index = 0)
            Local $_aResult = __JAB_GetAccessibleChildFromContext($_oVMID, $_oAccessibleContext, $_index)
                If @error Then Return SetError(@error, 0, 0)
            
            Local $_aAccChild[0][2]
                    _ArrayAdd($_aAccChild, "AccessibleContextChild|" & $_aResult[0])
                    _ArrayAdd($_aAccChild, "vmID|" & $_aResult[1])
                    _ArrayAdd($_aAccChild, "AccessibleContextParent|" & $_aResult[2])
                    _ArrayAdd($_aAccChild, "index|" & $_aResult[3])

                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleChildFromContext]: " & @CRLF & _
                                        ">> AccessibleContextChild:    " & $_aAccChild[0][1] & @CRLF & _
                                        ">> Virtual Machine ID (vmID):    " & $_aAccChild[1][1] & @CRLF & _
                                        ">> AccessibleContextParent:    " & $_aAccChild[2][1] & @CRLF & _
                                        ">> index:    " & $_aAccChild[3][1] & @CRLF)
                EndIf

            Return $_aAccChild
        EndFunc

        Func _JAB_GetAccessibleParentFromContext($_oVMID, $_oAccessibleContext)
            Local $_aResult = __JAB_GetAccessibleParentFromContext($_oVMID, $_oAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)

            Local $_aAccParent[0][2]
                    _ArrayAdd($_aAccParent, "AccessibleContextParent|" & $_aResult[0])
                    _ArrayAdd($_aAccParent, "vmID|" & $_aResult[1])
                    _ArrayAdd($_aAccParent, "AccessibleContextChild|" & $_aResult[2])

                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleParentFromContext]: " & @CRLF & _
                                        ">> AccessibleContextParent:    " & $_aAccParent[0][1] & @CRLF & _
                                        ">> Virtual Machine ID (vmID):    " & $_aAccParent[1][1] & @CRLF & _
                                        ">> AccessibleContextChild:    " & $_aAccParent[2][1] & @CRLF)
                EndIf

            Return $_aAccParent
        EndFunc

        Func _JAB_GetAccessibleContextAt($_oVMID, $_oParentAccessibleContext, $xMouse = -1, $yMouse = -1)
            If (($xMouse < 0) OR ($yMouse < 0)) Then
                $xMouse = MouseGetPos(0)
                $yMouse = MouseGetPos(1)
            EndIf
            Local $tStruct = DllStructCreate($tagPOINT)
            DllStructSetData($tStruct, "x", $xMouse)
            DllStructSetData($tStruct, "y", $yMouse)
            Local $_oNewAccessibleContext

            Local $_aResult = __JAB_GetAccessibleContextAt($_oVMID, $_oParentAccessibleContext, $xMouse, $yMouse, $_oNewAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)

            Local $_aNewAccessibleContext[0][2]
                _ArrayAdd($_aNewAccessibleContext, "Success|" & $_aResult[0])
                _ArrayAdd($_aNewAccessibleContext, "vmID|" & $_aResult[1])
                _ArrayAdd($_aNewAccessibleContext, "AccessibleContextParent|" & $_aResult[2])
                _ArrayAdd($_aNewAccessibleContext, "xCoord|" & $_aResult[3])
                _ArrayAdd($_aNewAccessibleContext, "yCoord|" & $_aResult[4])
                _ArrayAdd($_aNewAccessibleContext, "NewAccessibleContext|" & $_aResult[5])
    
                If $_JAB_DEBUG Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleContextAt]: " & $_aNewAccessibleContext[0][1] & @CRLF & _
                                        ">> vmID: " & $_aNewAccessibleContext[1][1] & @CRLF & _
                                        ">> AccessibleContextParent: " & $_aNewAccessibleContext[2][1] & @CRLF & _
                                        ">> xCoord: " & $_aNewAccessibleContext[3][1] & @CRLF & _
                                        ">> yCoord: " & $_aNewAccessibleContext[4][1] & @CRLF & _
                                        ">> NewAccessibleContext: " & $_aNewAccessibleContext[5][1] & @CRLF)
                EndIf
            
            Return $_aNewAccessibleContext
        EndFunc

        Func _JAB_GetAccessibleContextWithFocus($_hwnd)
            Local $_oVMID, $_oAccessibleContext
            Local $_aResult = __JAB_GetAccessibleContextWithFocus($_hwnd, $_oVMID, $_oAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)
            
            Local $_aAccContext[0][2]
            _ArrayAdd($_aAccContext, "Success|" & $_aResult[0])
            _ArrayAdd($_aAccContext, "HWND|" & $_aResult[1])
            _ArrayAdd($_aAccContext, "vmID|" & $_aResult[2])
            _ArrayAdd($_aAccContext, "AccessibleContext|" & $_aResult[3])

                If $_JAB_DEBUG Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleContextWithFocus]: " & $_aAccContext[0][1] & @CRLF & _
                                        ">> Window handle:    " & $_aAccContext[1][1] & @CRLF & _
                                        ">> Virtual Machine ID (vmID):    " & $_aAccContext[2][1] & @CRLF & _
                                        ">> Accessible Context (AC):    " & $_aAccContext[3][1] & @CRLF)
                EndIf

            Return $_aAccContext
        EndFunc

    #EndRegion ACCESSIBLE_CONTEXT_EX

    #Region ACCESSIBLE_TABLE-EX

        Func _JAB_GetAccessibleTableInfo($_oVMID, $_oParentAccessibleContext)
            Local $_tAccessibleTableInfo = DllStructCreate($tagAccessibleTableInfo)
            Local $_ptr_AccessibleTableInfo = DllStructGetPtr($_tAccessibleTableInfo)

            Local $_aResult = __JAB_GetAccessibleTableInfo($_oVMID, $_oParentAccessibleContext, $_ptr_AccessibleTableInfo)
                If @error Then Return SetError(@error, 0, 0)
            
            Local $_aAccTableInfo[0][2]
                _ArrayAdd($_aAccTableInfo, "caption|" & DllStructGetData($_tAccessibleTableInfo, "caption"))
                _ArrayAdd($_aAccTableInfo, "summary|" & DllStructGetData($_tAccessibleTableInfo, "summary"))
                _ArrayAdd($_aAccTableInfo, "rowCount|" & DllStructGetData($_tAccessibleTableInfo, "rowCount"))
                _ArrayAdd($_aAccTableInfo, "columnCount|" & DllStructGetData($_tAccessibleTableInfo, "columnCount"))
                _ArrayAdd($_aAccTableInfo, "accessibleContext|" & DllStructGetData($_tAccessibleTableInfo, "accessibleContext"))
                _ArrayAdd($_aAccTableInfo, "accessibleTable|" & DllStructGetData($_tAccessibleTableInfo, "accessibleTable"))
            
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleTableInfo]: " & $_aResult[0] & @CRLF & _
                                        ">> caption: " & $_aAccTableInfo[0][1] & @CRLF & _
                                        ">> summary: " & $_aAccTableInfo[0][1] & @CRLF & _
                                        ">> rowCount: " & $_aAccTableInfo[1][1] & @CRLF & _
                                        ">> columnCount: " & $_aAccTableInfo[2][1] & @CRLF & _
                                        ">> accessibleContext: " & $_aAccTableInfo[3][1] & @CRLF & _
                                        ">> accessibleTable: " & $_aAccTableInfo[4][1] & @CRLF)
                EndIf

            Return $_aAccTableInfo
        EndFunc

        Func _JAB_GetAccessibleTableCellInfo($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn)
            Local $_tAccessibleTableCellInfo = DllStructCreate($tagAccessibleTableCellInfo)
            Local $_ptr_AccessibleTableCellInfo = DllStructGetPtr($_tAccessibleTableCellInfo)

            Local $_aResult = __JAB_GetAccessibleTableCellInfo($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn, $_ptr_AccessibleTableCellInfo)
                If @error Then Return SetError(@error, 0, 0)
            
            Local $_aAccTableCellInfo[0][2]
                _ArrayAdd($_aAccTableCellInfo, "accessibleContext|" & DllStructGetData($_tAccessibleTableCellInfo, "accessibleContext"))
                _ArrayAdd($_aAccTableCellInfo, "index|" & DllStructGetData($_tAccessibleTableCellInfo, "index"))
                _ArrayAdd($_aAccTableCellInfo, "row|" & DllStructGetData($_tAccessibleTableCellInfo, "row"))
                _ArrayAdd($_aAccTableCellInfo, "column|" & DllStructGetData($_tAccessibleTableCellInfo, "column"))
                _ArrayAdd($_aAccTableCellInfo, "rowExtent|" & DllStructGetData($_tAccessibleTableCellInfo, "rowExtent"))
                _ArrayAdd($_aAccTableCellInfo, "columnExtent|" & DllStructGetData($_tAccessibleTableCellInfo, "columnExtent"))
                _ArrayAdd($_aAccTableCellInfo, "isSelected|" & DllStructGetData($_tAccessibleTableCellInfo, "isSelected"))
            
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleTableInfo]: " & $_aResult[0] & @CRLF & _
                                        ">> accessibleContext: " & $_aAccTableCellInfo[0][1] & @CRLF & _
                                        ">> index: " & $_aAccTableCellInfo[1][1] & @CRLF & _
                                        ">> row: " & $_aAccTableCellInfo[2][1] & @CRLF & _
                                        ">> column: " & $_aAccTableCellInfo[3][1] & @CRLF & _
                                        ">> rowExtent: " & $_aAccTableCellInfo[4][1] & @CRLF & _
                                        ">> columnExtent: " & $_aAccTableCellInfo[5][1] & @CRLF & _
                                        ">> isSelected: " & $_aAccTableCellInfo[6][1] & @CRLF)
                EndIf

            Return $_aAccTableCellInfo
        EndFunc

    #EndRegion ACCESSIBLE_TABLE_EX

    #Region ACCESSIBLE_TEXT-EX

        Func _JAB_GetAccessibleTextInfo($_oVMID, $_oAccessibleContext, $xPos = 0, $yPos = 0)
            Local $_tAccessibleTextInfo = DllStructCreate($tagAccessibleTextInfo)
            Local $_ptr_AccessibleTextInfo = DllStructGetPtr($_tAccessibleTextInfo)

            Local $_aResult = __JAB_GetAccessibleTextInfo($_oVMID, $_oAccessibleContext, $_ptr_AccessibleTextInfo, $xPos, $yPos)
                If @error Then Return SetError(@error, 0, 0)

            Local $_aAccTextInfo[0][2]
                _ArrayAdd($_aAccTextInfo, "charCount|" & DllStructGetData($_tAccessibleTextInfo, "charCount"))
                _ArrayAdd($_aAccTextInfo, "caretIndex|" & DllStructGetData($_tAccessibleTextInfo, "caretIndex"))
                _ArrayAdd($_aAccTextInfo, "indexAtPoint|" & DllStructGetData($_tAccessibleTextInfo, "indexAtPoint"))
            
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleTextInfo]: " & $_aResult[0] & @CRLF & _
                                        ">> charCount: " & $_aAccTextInfo[0][1] & @CRLF & _
                                        ">> caretIndex: " & $_aAccTextInfo[0][1] & @CRLF & _
                                        ">> indexAtPoint: " & $_aAccTextInfo[1][1] & @CRLF)
                EndIf

            Return $_aAccTextInfo
        EndFunc

    #EndRegion ACCESSIBLE_TEXT-EX

    #Region ADDITIONAL_TEXT-EX

        Func _JAB_SelectTextRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex)
            $_aResult = __JAB_SelectTextRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex)
                If @error Then Return SetError(@error, 0, 0)

            If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [_JAB_SelectTextRange]: " & $_aResult[0] & @CRLF)
            
            Return $_aResult[0]
        EndFunc

    #EndRegion ADDITIONAL_TEXT-EX

    #Region ACCESSIBLE_ACTIONS-EX

        Func _JAB_GetAccessibleActions($_oVMID, $_oAccessibleContext)
            Local $_tAccessibleActions = DllStructCreate($tagAccessibleActions)
            Local $_ptr_AccessibleActions = DllStructGetPtr($_tAccessibleActions)

            Local $_aResult = __JAB_GetAccessibleActions($_oVMID, $_oAccessibleContext, $_ptr_AccessibleActions)
                If @error Then Return SetError(@error, 0, 0)

            Local $_aAccActions[0][2]
                _ArrayAdd($_aAccActions, "actionsCount|" & DllStructGetData($_tAccessibleActions, "actionsCount"))
                If $_aAccActions[0][1] > 0 Then
                    For $i = 2 To $_aAccActions[0][1] + 1
                        _ArrayAdd($_aAccActions, $i & "|" & DllStructGetData($_tAccessibleActions, $i))
                    Next
                EndIf
                
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetAccessibleActions]: " & $_aResult[0] & @CRLF & _
                                        ">> actionsCount: " & $_aAccActions[0][1] & @CRLF)
                    For $i = 1 To $_aAccActions[0][1]
                        ConsoleWrite(">> action [" & $i & "]: " & $_aAccActions[$i][1] & @CRLF)
                    Next
                EndIf

            Return $_aAccActions
        EndFunc

        Func _JAB_DoAccessibleActions($_oVMID, $_oAccessibleContext, $_vActionsToDo)

            Local $_tAccessibleActionsToDo = DllStructCreate($tagAccessibleActionsToDo)
            Local $_iActionCount = 1

            ;-- if actions in array:
            If VarGetType($_vActionsToDo) = "array" Then
                Local $_iActionCount = Ubound($_vActionsToDo)
                DllStructSetData($_tAccessibleActionsToDo, "actionsCount", $_iActionCount)
                For $i = 0 To $_iActionCount-1
                    DllStructSetData($_tAccessibleActionsToDo, $i+2, $_vActionsToDo[$i])
                Next
            Else
                ;-- if action is string:
                DllStructSetData($_tAccessibleActionsToDo, "actionsCount", 1)
                DllStructSetData($_tAccessibleActionsToDo, 2, $_vActionsToDo)
            EndIf

            Local $_aResult = __JAB_DoAccessibleActions($_oVMID, $_oAccessibleContext, $_tAccessibleActionsToDo)
                If @error Then Return SetError(@error, 0, 0)

                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_DoAccessibleActions]: " & $_aResult[0] & @CRLF & _
                                        ">> actionsCount: " & $_iActionCount & @CRLF & _
                                        ">> failed action index: " & $_aResult[4] & @CRLF)
                EndIf

            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_ACTIONS-EX

    #Region UTILITY-EX

        Func _JAB_GetTopLevelObject($_oVMID, $_oAccessibleContext)
            $_aResult = __JAB_GetTopLevelObject($_oVMID, $_oAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)

            If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [_JAB_GetTopLevelObject]: " & $_aResult[2] & @CRLF)
            
            Return $_aResult[0]
        EndFunc

        Func _JAB_GetObjectDepth($_oVMID, $_oAccessibleContext)
            $_aResult = __JAB_GetObjectDepth($_oVMID, $_oAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)
            
            If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetObjectDepth]: " & $_aResult[2] & @CRLF & _
                                        ">> depth level: " & $_aResult[0] & @CRLF)
            EndIf

            Return $_aResult[0]
        EndFunc

        Func _JAB_GetActiveDescendent($_oVMID, $_oAccessibleContext)
            $_aResult = __JAB_GetActiveDescendent($_oVMID, $_oAccessibleContext)
                 If @error Then Return SetError(@error, 0, 0)

            If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [_JAB_GetActiveDescendent]: " & $_aResult[2] & @CRLF)
            
            Return $_aResult[0]
        EndFunc

        Func _JAB_GetVisibleChildren($_oVMID, $_oAccessibleContext, $_iStartIndex = 0)
            Local $_tVisibleChildenInfo = DllStructCreate($tagVisibleChildenInfo)
            Local $_ptr_VisibleChildenInfo = DllStructGetPtr($_tVisibleChildenInfo)

            $_aResult = __JAB_GetVisibleChildren($_oVMID, $_oAccessibleContext, $_iStartIndex, $_ptr_VisibleChildenInfo)
                If @error Then Return SetError(@error, 0, 0)

            Local $_aVisibleChildren[0][2]
                _ArrayAdd($_aVisibleChildren, "returnedChildrenCount|" & DllStructGetData($_tVisibleChildenInfo, "returnedChildrenCount"))
                _ArrayAdd($_aVisibleChildren, "children|" & DllStructGetData($_tVisibleChildenInfo, "children"))
            
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [_JAB_GetVisibleChildren]: " & $_aResult[0] & @CRLF & _
                                        ">> returnedChildrenCount: " & $_aVisibleChildren[0][1] & @CRLF & _
                                        ">> children: " & $_aVisibleChildren[1][1] & @CRLF)
                EndIf

            Return $_aVisibleChildren
        EndFunc

    #EndRegion UTILITY-EX
        
#EndRegion CORE-EX_FUNCTIONS