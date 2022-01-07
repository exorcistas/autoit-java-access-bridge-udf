#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: JavaAccessBridge_Core
    Description...........: The Java Access Bridge API enables you to develop assistive technology apps for the Windows OS that work with Java applications. 
                            It contains native methods that enable you to view and manipulate information about GUI elements in a Java application, 
                            which is forwarded to your assistive technology application through Java Access Bridge.
	
    Dependencies..........: JRE 8/9 1.8.0 (windowsaccessbridge.dll)
    Documentation.........: https://docs.oracle.com/javase/9/access/jaapi.htm
                            https://docs.oracle.com/javase/8/docs/api/javax/accessibility/package-summary.html
                            https://docs.oracle.com/javase/accessbridge/

    Author................: exorcistas@github.com
    Modified..............: 2020-03-11
    Version...............: v0.7.5.1rc
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <Array.au3>
#include <JavaAccessBridge_Globals.au3>

#Region CORE_FUNCTIONS_LIST
#cs	===================================================================================================================================
%% INITIALIZATION/SHUTDOWN %%
    __JAB_SetEnvironment($_bDebug = $_JAB_DEBUG)
    __JAB_OpenBridgeDLL()
    __JAB_CloseBridgeDLL()
    __JAB_EnableSwitch($_bEnable = True)
    __JAB_Windows_run()

%% GATEWAY %%
    __JAB_GetWindowHandle($_sWinTitle)
    __JAB_IsJavaWindow($_hwnd)
    __JAB_GetAccessibleContextFromHWND($_hwnd, ByRef $_oVMID, ByRef $_oAccessibleContext)

%% GENERAL %%
    __JAB_ReleaseJavaObject($_oVMID, $_oJavaObject)
    __JAB_GetVersionInfo($_oVMID, ByRef $_ptr_AccessBridgeVersionInfo)

%% ACCESSIBLE_CONTEXT %%
    __JAB_GetAccessibleContextInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextInfo)
    __JAB_GetAccessibleChildFromContext($_oVMID, $_oParentAccessibleContext, $_index)
    __JAB_GetAccessibleParentFromContext($_oVMID, $_oChildAccessibleContext)
    __JAB_GetAccessibleContextAt($_oVMID, $_oParentAccessibleContext, $xCoord, $yCoord, ByRef $_ptr_AccessibleContext)
    __JAB_GetAccessibleContextWithFocus($_hwnd, ByRef $_oVMID, ByRef $_ptr_AccessibleContext)
    __JAB_GetHWNDFromAccessibleContext($_oVMID, $_oAccessibleContext)

%% ACCESSIBLE_TEXT %%
    __JAB_GetAccessibleTextInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextInfo, $xPos, $yPos)
    __JAB_GetAccessibleTextItems($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextItemsInfo, $_index)
    __JAB_GetAccessibleTextSelectionInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextSelectionInfo)
    __JAB_GetAccessibleTextAttributes($_oVMID, $_oAccessibleContext, $_index, ByRef $_ptr_AccessibleTextAttributesInfo)
    __JAB_GetAccessibleTextRect($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextRectInfo, $_index)
    __JAB_GetAccessibleTextRange($_oVMID, $_oAccessibleContext, $_iStartPos, $_iEndPos, $_sText, $_iLength)
    __JAB_GetAccessibleTextLineBounds($_oVMID, $_oAccessibleContext, $_index, ByRef $_ptr_startIndex, ByRef $_ptr_endIndex)

%% ADDITIONAL_TEXT_FUNCTIONS %%
    __JAB_SelectTextRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex)
    __JAB_GetTextAttributesInRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex, ByRef $_ptr_AccessibleTextAttributesInfo, ByRef $_ptr_length)
    __JAB_SetCaretPosition($_oVMID, $_oAccessibleContext, $_iPos)
    __JAB_GetCaretLocation($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextRectInfo, $_iPos)
    __JAB_SetTextContents($_oVMID, $_oAccessibleContext, $_sText)

%% ACCESSIBLE_TABLE %%
    __JAB_GetAccessibleTableInfo($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
    __JAB_GetAccessibleTableCellInfo($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn, ByRef $_ptr_AccessibleTableCellInfo)
    __JAB_GetAccessibleTableRowHeader($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
    __JAB_GetAccessibleTableColumnHeader($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
    __JAB_GetAccessibleTableRowDescription($_oVMID, $_oParentAccessibleContext, $_iRow)
    __JAB_GetAccessibleTableColumnDescription($_oVMID, $_oParentAccessibleContext, $_iColumn)
    __JAB_GetAccessibleTableRowSelectionCount($_oVMID, $_oAccessibleTable)
    __JAB_IsAccessibleTableRowSelected($_oVMID, $_oAccessibleTable, $_iRow)
    __JAB_GetAccessibleTableRowSelections($_oVMID, $_oAccessibleTable, $_iCount, ByRef $_ptr_Selections)
    __JAB_GetAccessibleTableColumnSelectionCount($_oVMID, $_oAccessibleTable)
    __JAB_IsAccessibleTableColumnSelected($_oVMID, $_oAccessibleTable, $_iColumn)
    __JAB_GetAccessibleTableColumnSelections($_oVMID, $_oAccessibleTable, $_iCount, ByRef $_ptr_Selections)
    __JAB_GetAccessibleTableRow($_oVMID, $_oAccessibleTable, $_index)
    __JAB_GetAccessibleTableColumn($_oVMID, $_oAccessibleTable, $_index)
    __JAB_GetAccessibleTableIndex($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn)

%% ACCESSIBLE_RELATION %%
    __JAB_GetAccessibleRelationSet($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleRelationSetInfo)

%% ACCESSIBLE_HYPERTEXT %%
    __JAB_GetAccessibleHypertext($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleHypertextInfo)
    __JAB_ActivateAccessibleHyperlink($_oVMID, $_oAccessibleContext, $_oAccessibleHyperlink)
    __JAB_GetAccessibleHyperlinkCount($_oVMID, $_oAccessibleHypertext)
    __JAB_GetAccessibleHypertextExt($_oVMID, $_oAccessibleContext, $_iStartIndex, ByRef $_ptr_AccessibleHypertextInfo)
    __JAB_GetAccessibleHypertextLinkIndex($_oVMID, $_oAccessibleHypertext, $_index)
    __JAB_GetAccessibleHyperlink($_oVMID, $_oAccessibleHypertext, $_index, ByRef $_ptr_AccessibleHypertextInfo)

%% ACCESSIBLE_KEY_BINDINGS %%
    __JAB_GetAccessibleKeyBindings($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleKeyBindings)

%% ACCESSIBLE_ICONS %%
    __JAB_GetAccessibleIcons($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleIcons)

%% ACCESSIBLE_ACTIONS %%
    __JAB_GetAccessibleActions($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextActions)
    __JAB_DoAccessibleActions($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextActionsToDo)

%% UTILITY %%
    __JAB_IsSameObject($_oVMID, $_oAccessibleContext1, $_oAccessibleContext2)
    __JAB_GetParentWithRole($_oVMID, $_oAccessibleContext, $_sRole)
    __JAB_GetParentWithRoleElseRoot($_oVMID, $_oAccessibleContext, $_sRole)
    __JAB_GetTopLevelObject($_oVMID, $_oAccessibleContext)
    __JAB_GetObjectDepth($_oVMID, $_oAccessibleContext)
    __JAB_GetActiveDescendent($_oVMID, $_oAccessibleContext)
    __JAB_RequestFocus($_oVMID, $_oAccessibleContext)
    __JAB_GetVisibleChildrenCount($_oVMID, $_oAccessibleContext)
    __JAB_GetVisibleChildren($_oVMID, $_oAccessibleContext, $_iStartIndex, ByRef $_ptr_VisibleChildrenInfo)
    __JAB_GetEventsWaiting()

%% ACCESSIBLE_VALUE %%
    __JAB_GetCurrentAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
    __JAB_GetMaximumAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
    __JAB_GetMinimumAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)

%% ACCESSIBLE_SELECTION %%
    __JAB_AddAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
    __JAB_ClearAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection)
    __JAB_GetAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
    __JAB_GetAccessibleSelectionCountFromContext($_oVMID, $_oAccessibleSelection)
    __JAB_IsAccessibleChildSelectedFromContext($_oVMID, $_oAccessibleSelection, $_index)
    __JAB_RemoveAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
    __JAB_SelectAllAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection)
#ce	===================================================================================================================================
#EndRegion CORE_FUNCTIONS_LIST

#Region CORE_FUNCTIONS 

    #Region INITIALIZATION/SHUTDOWN
        ; ## Functions to start/stop Java Access Bridge API ##

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_SetEnvironment($_bDebug = $_JAB_DEBUG)
            Description ...: Presets global variables required to run JAB.
            Syntax.........: -

            Global vars....: $_JAB_DEBUG; $_JAB_VERSION; $_JAB_HOME_DIR; $_JAB_WindowsAccessBridge_DLL
            Parameters.....: -
            Return values..: -

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-09
        #ce ===============================================================================================================================
        Func __JAB_SetEnvironment($_bDebug = $_JAB_DEBUG)
            $_JAB_DEBUG = $_bDebug

            Local $_sRegistryHome = "HKLM\SOFTWARE\"
            Local $_sRegJavaDir = "JavaSoft\Java Runtime Environment"
            Local $_sRegPath32bit = $_sRegistryHome & "Wow6432Node\" & $_sRegJavaDir
            Local $_sRegPath = $_sRegistryHome & $_sRegJavaDir

            Switch @OSArch
                Case "X86"
                    $_JAB_VERSION = RegRead($_sRegPath, "CurrentVersion")
                    $_JAB_HOME_DIR = RegRead($_sRegPath & "\" & $_JAB_VERSION, "JavaHome")
                    $_JAB_WindowsAccessBridge_DLL = "WindowsAccessBridge-32.dll"

                Case "X64"
                    If (NOT @AutoItX64) Then
                        $_JAB_VERSION = RegRead($_sRegPath32bit, "CurrentVersion")
                        $_JAB_HOME_DIR = RegRead($_sRegPath32bit & "\" & $_JAB_VERSION, "JavaHome")
                        $_JAB_WindowsAccessBridge_DLL = "WindowsAccessBridge-32.dll"
                    Else
                        $_JAB_VERSION = RegRead($_sRegPath, "CurrentVersion")
                        $_JAB_HOME_DIR = RegRead($_sRegPath & "\" & $_JAB_VERSION, "JavaHome")
                        $_JAB_WindowsAccessBridge_DLL = "WindowsAccessBridge-64.dll"
                    EndIf
            EndSwitch
            If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                ConsoleWrite(@CRLF & ">> [__JAB_SetEnvironment] OSArch:" & @TAB & @OSArch & @TAB & "| AutoItX64: " & @AutoItX64 & @CRLF & _
                                     ">> Java Version:    " & $_JAB_VERSION & @CRLF & _
                                     ">> Java Home:    " & $_JAB_HOME_DIR & @CRLF & _
                                     ">> WindowsAccessBridge:    " & $_JAB_WindowsAccessBridge_DLL & @CRLF)
            EndIf
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_OpenBridgeDLL()
            Description ...: Opens Access Bridge DLL for JAB API
            Syntax.........: -

            Global vars....: $_JAB_hAccessBridgeDLL; $_JAB_HOME_DIR; $_JAB_WindowsAccessBridge_DLL
            Parameters.....: -
            Return values..: Opened DLL handle ($_JAB_hAccessBridgeDLL)

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-05
        #ce ===============================================================================================================================
        Func __JAB_OpenBridgeDLL()
            $_JAB_hAccessBridgeDLL = DllOpen($_JAB_HOME_DIR & "\bin\" & $_JAB_WindowsAccessBridge_DLL)
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [__JAB_OpenBridgeDLL]: " & $_JAB_hAccessBridgeDLL & @CRLF)

            If $_JAB_hAccessBridgeDLL = -1 Then SetError(1)
            Return SetError(@error, 0, $_JAB_hAccessBridgeDLL)
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_CloseBridgeDLL()
            Description ...: Closes Access Bridge DLL
            Syntax.........: -

            Global vars....: $_JAB_hAccessBridgeDLL
            Parameters.....: -
            Return values..: -

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-09
        #ce ===============================================================================================================================
        Func __JAB_CloseBridgeDLL()
            DllClose($_JAB_hAccessBridgeDLL)
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [__JAB_CloseBridgeDLL]" & @CRLF)
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_EnableSwitch($_bEnable = True)
            Description ...: Enabled/disables JABSwitch. Required to be enabled before launching JAVA application
            Syntax.........: -

            Global vars....: $_JAB_HOME_DIR
            Parameters.....: $_bEnable (TRUE/FALSE) - to enable or disable JABSwitch
            Return values..: -

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-05
        #ce ===============================================================================================================================
        Func __JAB_EnableSwitch($_bEnable = True)
            Local $_sEnable = ($_bEnable) ? "enable" : "disable"
            Local $_sDebug = ($_JAB_DEBUG) ? @SW_SHOW : @SW_HIDE

            ;// TO DO: Handle messages received from initialize
            Runwait($_JAB_HOME_DIR & "\bin\jabswitch /" & $_sEnable, "" , @SW_HIDE)
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [__JAB_EnableSwitch]:" & @TAB & $_sEnable & @CRLF)
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_Windows_run()
            Description ...: Initializes Java Bridge.
            Syntax.........: -

            Global vars....: $_JAB_hAccessBridgeDLL
            Parameters.....: -
            Return values..: Success: 1; Failure: 0 + @error

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-06
        #ce ===============================================================================================================================
        Func __JAB_Windows_run()
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "Windows_run")    
                If @error Then Return SetError(@error, 0, 0)

            If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [__JAB_Windows_run]: " & $_aResult[0] & @CRLF)
            
            Return $_aResult[0]
        EndFunc
    #EndRegion INITIALIZATION/SHUTDOWN

    #Region GENERAL
        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_ReleaseJavaObject($_oVMID, $_oJavaObject)

            Description ...: Release the memory used by the Java_Object object, where object is an object returned to you by Java Access Bridge. 
                             Java Access Bridge automatically maintains a reference to all Java objects that it returns to you in the JVM so they are not garbage collected. 
                             To prevent memory leaks, you must call ReleaseJavaObject on all Java objects returned to you by Java Access Bridge once you are finished with them. 
                             See JavaFerret.c for an illustration of how to do this.

            Syntax.........: void ReleaseJavaObject(long vmID, Java_Object object);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oJavaObject - AccessibleContext/Java object to release;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================
        Func __JAB_ReleaseJavaObject($_oVMID, $_oJavaObject)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "NONE", "releaseJavaObject", "long", $_oVMID, $c_JOBJECT64, $_oJavaObject)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetVersionInfo($_oVMID, ByRef $_ptr_AccessBridgeVersionInfo)

            Description ...: Gets the version information of the instance of Java Access Bridge instance your application is using. 
                             You can use this information to determine the available functionality of your version of Java Access Bridge.
                             In order to determine the version of the JVM, you need to pass in a valid vmID; 
                             otherwise all that is returned is the version of the WindowsAccessBridge.DLL file to which your application is connected.

            Syntax.........: BOOL GetVersionInfo(long vmID, AccessBridgeVersionInfo *info);

            Global vars....: $_JAB_hAccessBridgeDLL
            Parameters.....: $_oVMID - JVM ID variable;
                             $_ptr_AccessBridgeVersionInfo - pointer to AccessBridgeVersionInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================
        Func __JAB_GetVersionInfo($_oVMID, ByRef $_ptr_AccessBridgeVersionInfo)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getVersionInfo", "long", $_oVMID, "struct*", $_ptr_AccessBridgeVersionInfo)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc
    #EndRegion GENERAL

    #Region GATEWAY
        ;## Typically called functions before calling any other Java Access Bridge API function ##

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetWindowHandle($_sWinTitle)
            Description ...: Returns Window Handle from Window Title
            Syntax.........: See function: {WinGetHandle}

            Global vars....: -
            Parameters.....: $_sWinTitle - windows title string to search
            Return values..: $_hwnd - window handle

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-09
        #ce ===============================================================================================================================
        Func __JAB_GetWindowHandle($_sWinTitle)
            Local $_bExists = WinExists($_sWinTitle)
                If NOT $_bExists Then Return SetError(@error, 0, 0)
            
            Local $_hwnd = WinGetHandle($_sWinTitle)
            WinActivate($_hwnd)
            $_hwnd = WinWaitActive($_hwnd, "", 30)
                If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then 
                    ConsoleWrite(@CRLF & ">> [__JAB_GetWindowHandle]:" & @CRLF & _
                                         ">> Window title:    " & $_sWinTitle & @CRLF & _
                                         ">> Window handle:    " & $_hwnd & @CRLF)
                EndIf

            Return $_hwnd
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_IsJavaWindow($_hwnd)
            Description ...: Checks to see if the given window implements the Java Accessibility API.

            Syntax.........: BOOL IsJavaWindow(HWND window);

            Global vars....: $_JAB_hAccessBridgeDLL
            Parameters.....: $_hwnd - window handle
            Return values..: Success: 1; Failure: 0 + @error

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-09
        #ce ===============================================================================================================================
        Func __JAB_IsJavaWindow($_hwnd)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "isJavaWindow", "hwnd", $_hwnd)
                If @error Then Return SetError(@error, 0, 0)
                    
            If $_JAB_DEBUG = $_JAB_DEBUG_CORE Then ConsoleWrite(@CRLF & ">> [__JAB_IsJavaWindow]: " & $_aResult[0] & @CRLF)
                        
            Return $_aResult[0]
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleContextFromHWND($_hwnd, ByRef $_oVMID, ByRef $_oAccessibleContext)

            Description ...: Gets the AccessibleContext and vmID values for the given window. 
                             Many Java Access Bridge functions require the AccessibleContext and vmID values.

            Syntax.........: BOOL GetAccessibleContextFromHWND(HWND target, long *vmID, AccessibleContext *ac);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_hwnd - window handle;
                             $_oVMID - variable to hold JVM ID cookie;
                             $_oAccessibleContext - variable to hold master Accessible Context cookie;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-09
        #ce ===============================================================================================================================
        Func __JAB_GetAccessibleContextFromHWND($_hwnd, ByRef $_oVMID, ByRef $_oAccessibleContext)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleContextFromHWND", "hwnd", $_hwnd, "long*", $_oVMID, $cP_JOBJECT64, $_oAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc
    #EndRegion GATEWAY

    #Region ACCESSIBLE_CONTEXT
        #cs
            These functions provide the core of the Java Accessibility API that is exposed by Java Access Bridge.

            The functions GetAccessibleContextAt and GetAccessibleContextWithFocus retrieve an AccessibleContext object, 
            which is a magic cookie (really a Java Object reference) to an Accessible object and a JVM cookie. 
            You use these two cookies to reference objects through Java Access Bridge. 
            Most Java Access Bridge API functions require that you pass in these two parameters.

            The function GetAccessibleContextInfo returns detailed information about an AccessibleContext object belonging to the JVM. 
            In order to improve performance, the various distinct methods in the Java Accessibility API are collected together 
            into a few routines in the Java Access Bridge API and returned in struct values. 
            These struct values are defined in the file AccessBridgePackages.h and are described in the section "API Callbacks".

            The functions GetAccessibleChildFromContext and GetAccessibleParentFromContext enable you to walk the GUI component hierarchy, 
            retrieving the nth child, or the parent, of a particular GUI object.
        #ce    

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleContextInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextInfo)
            Description ...: Retrieves an AccessibleContextInfo object of the AccessibleContext object ac.

            Syntax.........: BOOL GetAccessibleContextInfo(long vmID, AccessibleContext ac, AccessibleContextInfo *info);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleContextInfo - pointer to AccessibleContextInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-09
        #ce ===============================================================================================================================
        Func __JAB_GetAccessibleContextInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextInfo)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleContextInfo", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleContextInfo)
                If @error Then Return SetError(@error, 0, 0)

            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleChildFromContext($_oVMID, $_oParentAccessibleContext, $_index)
            Description ...: Returns an AccessibleContext object that represents the nth child of the object ac, 
                             where n is specified by the value index.

            Syntax.........: AccessibleContext GetAccessibleChildFromContext(long vmID, AccessibleContext ac, jint index);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Parent Accessible Context variable;
                             $_index - numeric index of child (start from 0);

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================   
        Func __JAB_GetAccessibleChildFromContext($_oVMID, $_oParentAccessibleContext, $_index)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getAccessibleChildFromContext", "long", $_oVMID, $c_JOBJECT64, $_oParentAccessibleContext, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleParentFromContext($_oVMID, $_oChildAccessibleContext)
            Description ...: Returns an AccessibleContext object that represents the parent of object ac.

            Syntax.........: AccessibleContext GetAccessibleParentFromContext(long vmID, AccessibleContext ac);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Child Accessible Context variable;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================  
        Func __JAB_GetAccessibleParentFromContext($_oVMID, $_oChildAccessibleContext)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getAccessibleParentFromContext", "long", $_oVMID, $c_JOBJECT64, $_oChildAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleContextAt($_oVMID, $_oParentAccessibleContext, $xCoord, $yCoord, ByRef $_ptr_AccessibleContext)
            Description ...: Retrieves an AccessibleContext object of the window or object that is under the mouse pointer.

            Syntax.........: BOOL GetAccessibleContextAt(long vmID, AccessibleContext acParent, jint x, jint y, AccessibleContext *ac)

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Parent/Master Accessible Context variable;
                             $xCoord, $yCoord - mouse coordinates;
                             $_ptr_AccessibleContext - pointer to AccessibleContext struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================
        Func __JAB_GetAccessibleContextAt($_oVMID, $_oParentAccessibleContext, $xCoord, $yCoord, ByRef $_ptr_AccessibleContext)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleContextAt", "long", $_oVMID, $c_JOBJECT64, $_oParentAccessibleContext, "int", $xCoord, "int", $yCoord, $cP_JOBJECT64, $_ptr_AccessibleContext)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: _JAB_GetAccessibleContextWithFocus($_hwnd, ByRef $_oVMID, ByRef $_AccessibleContext)
            Description ...: Retrieves an AccessibleContext object of the window or object that has the focus.

            Syntax.........: BOOL GetAccessibleContextWithFocus(HWND window, long *vmID, AccessibleContext *ac);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_hwnd - window handle;
                             $_oVMID - variable to hold JVM ID cookie;
                             $_oAccessibleContext - variable to hold master Accessible Context cookie;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================
        Func __JAB_GetAccessibleContextWithFocus($_hwnd, ByRef $_oVMID, ByRef $_ptr_AccessibleContext)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleContextWithFocus", "hwnd", $_hwnd, "long*", $_oVMID, $cP_JOBJECT64, $_ptr_AccessibleContext)
                If @error Then Return SetError(@error, 0, 0)

            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetHWNDFromAccessibleContext($_oVMID, $_oAccessibleContext)
            Description ...: Returns the HWND from the AccessibleContextof a top-level window.

            Syntax.........: HWND getHWNDFromAccessibleContext(long vmID, AccessibleContext ac);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================  
        Func __JAB_GetHWNDFromAccessibleContext($_oVMID, $_oAccessibleContext)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "hwnd", "getHWNDFromAccessibleContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_CONTEXT

    #Region ACCESSIBLE_TEXT
        #cs
            These functions get AccessibleText information provided by the Java Accessibility API, broken down into seven chunks for efficiency. 
            An AccessibleContext has AccessibleText information contained within it 
            if the flag accessibleText in the AccessibleContextInfo data structure is set to TRUE. 
            The struct values used in these functions are defined in the file AccessBridgePackages.h and are described in the section "API Callbacks".
        #ce

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTextInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextInfo, $xPos, $yPos)
            Description ...: Retrieves text parameters listed in AccessibleTextInfo struct

            Syntax.........: BOOL GetAccessibleTextInfo(long vmID, AccessibleText at, AccessibleTextInfo *textInfo, jint x, jint y);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTextInfo - pointer to AccessibleTextInfo struct;
                             $xPos, $yPos - coordinates of text;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================  
        Func __JAB_GetAccessibleTextInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextInfo, $xPos, $yPos)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTextInfo", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleTextInfo, "int", $xPos, "int", $yPos)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTextItems($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextItemsInfo, $_index)
            Description ...: 

            Syntax.........: BOOL GetAccessibleTextItems(long vmID, AccessibleText at, AccessibleTextItemsInfo *textItems, jint index);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTextItemsInfo - pointer to AccessibleTextItemsInfo struct;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================  
        Func __JAB_GetAccessibleTextItems($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextItemsInfo, $_index)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTextItems", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleTextItemsInfo, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTextSelectionInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextSelectionInfo)
            Description ...: 

            Syntax.........: BOOL GetAccessibleTextSelectionInfo(long vmID, AccessibleText at, AccessibleTextSelectionInfo *textSelection);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTextSelectionInfo - pointer to _ptr_AccessibleTextSelectionInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce ===============================================================================================================================  
        Func __JAB_GetAccessibleTextSelectionInfo($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextSelectionInfo)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTextSelectionInfo", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleTextSelectionInfo)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTextAttributes($_oVMID, $_oAccessibleContext, $_index, ByRef $_ptr_AccessibleTextAttributesInfo)
            Description ...: 

            Syntax.........: char *GetAccessibleTextAttributes(long vmID, AccessibleText at, jint index, AccessibleTextAttributesInfo *attributes);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_index;
                             $_ptr_AccessibleTextAttributesInfo - pointer to AccessibleTextAttributesInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTextAttributes($_oVMID, $_oAccessibleContext, $_index, ByRef $_ptr_AccessibleTextAttributesInfo)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTextAttributes", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_index, "struct*", $_ptr_AccessibleTextAttributesInfo)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTextRect($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextRectInfo, $_index)
            Description ...: 

            Syntax.........: BOOL GetAccessibleTextRect(long vmID, AccessibleText at, AccessibleTextRectInfo *rectInfo, jint index);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTextRectInfo - pointer to AccessibleTextRectInfo struct;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTextRect($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextRectInfo, $_index)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTextRect", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleTextRectInfo, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTextRange($_oVMID, $_oAccessibleContext, $_iStartPos, $_iEndPos, $_sText, $_iLength)
            Description ...: 

            Syntax.........: BOOL GetAccessibleTextRange(long vmID, AccessibleText at, jint start, jint end, wchar_t *text, short len);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_iStartPos, $_iEndPos - start/end points;
                             $_sText - text to select;
                             $_iLength - length to select;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTextRange($_oVMID, $_oAccessibleContext, $_iStartPos, $_iEndPos, $_sText, $_iLength)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTextRange", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_iStartPos, "int", $_iEndPos, "wstr", $_sText, "short", $_iLength)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTextLineBounds($_oVMID, $_oAccessibleContext, $_index, ByRef $_ptr_startIndex, ByRef $_ptr_endIndex)
            Description ...: 

            Syntax.........: BOOL GetAccessibleTextLineBounds(long vmID, AccessibleText at, jint index, jint *startIndex, jint *endIndex);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_index;
                             $_ptr_startIndex, $_ptr_endIndex;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTextLineBounds($_oVMID, $_oAccessibleContext, $_index, ByRef $_ptr_startIndex, ByRef $_ptr_endIndex)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTextLineBounds", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_index, "int*", $_ptr_startIndex, "int*", $_ptr_endIndex)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_TEXT

    #Region ADDITIONAL_TEXT
        
        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_SelectTextRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex)
            Description ...: Selects text between two indices. 
                             Selection includes the text at the start index and the text at the end index. 
                             Returns whether successful.

            Syntax.........: BOOL selectTextRange(const long vmID, const AccessibleContext accessibleContext, const int startIndex, const int endIndex);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_iStartIndex, $_iEndIndex;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_SelectTextRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "selectTextRange", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_iStartIndex, "int", $_iEndIndex)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetTextAttributesInRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex, ByRef $_ptr_AccessibleTextAttributesInfo, ByRef $_ptr_length)
            Description ...: Get text attributes between two indices. 
                             The attribute list includes the text at the start index and the text at the end index. 
                             Returns whether successful.

            Syntax.........: BOOL getTextAttributesInRange(const long vmID, const AccessibleContext accessibleContext, const int startIndex, const int endIndex, AccessibleTextAttributesInfo *attributes, short *len);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_iStartIndex, $_iEndIndex;
                             $_ptr_AccessibleTextAttributesInfo - pointer to AccessibleTextAttributesInfo struct;
                             $_ptr_length;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetTextAttributesInRange($_oVMID, $_oAccessibleContext, $_iStartIndex, $_iEndIndex, ByRef $_ptr_AccessibleTextAttributesInfo, ByRef $_ptr_length)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getTextAttributesInRange", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_iStartIndex, "int", $_iEndIndex, "struct*", $_ptr_AccessibleTextAttributesInfo, "short*", $_ptr_length)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_SetCaretPosition($_oVMID, $_oAccessibleContext, $_iPos)
            Description ...: Set the caret to a text position. Returns whether successful.

            Syntax.........: BOOL setCaretPosition(const long vmID, const AccessibleContext accessibleContext, const int position);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_iPos - position index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_SetCaretPosition($_oVMID, $_oAccessibleContext, $_iPos)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "setCaretPosition", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_iPos)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetCaretLocation($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextRectInfo, $_iPos)
            Description ...: Gets the text caret location.

            Syntax.........: BOOL getCaretLocation(long vmID, AccessibleContext ac, AccessibleTextRectInfo *rectInfo, jint index);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTextRectInfo - pointer to AccessibleTextRectInfo struct;
                             $_iPos - position index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetCaretLocation($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleTextRectInfo, $_iPos)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getCaretLocation", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleTextRectInfo, "int", $_iPos)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_SetTextContents($_oVMID, $_oAccessibleContext, $_sText)
            Description ...: Sets editable text contents. 
                             The AccessibleContext must implement AccessibleEditableText and be editable. 
                             The maximum text length that can be set is MAX_STRING_SIZE - 1. 
                             Returns whether successful.

            Syntax.........: BOOL setTextContents (const long vmID, const AccessibleContext accessibleContext, const wchar_t *text);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_sText - text to set;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_SetTextContents($_oVMID, $_oAccessibleContext, $_sText)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "setTextContents", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "wstr", $_sText)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

    #EndRegion ADDITIONAL_TEXT

    #Region ACCESSIBLE_TABLE

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableInfo($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
            Description ...: Returns information about the table, for example, caption, summary, row and column count, and the AccessibleTable.

            Syntax.........: BOOL getAccessibleTableInfo(long vmID, AccessibleContext acParent, AccessibleTableInfo *tableInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTableInfo - pointer to AccessibleTableInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableInfo($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTableInfo", "long", $_oVMID, $c_JOBJECT64, $_oParentAccessibleContext, "struct*", $_ptr_AccessibleTableInfo)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableCellInfo($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn, ByRef $_ptr_AccessibleTableCellInfo)
            Description ...: Returns information about the specified table cell. The row and column specifiers are zero-based.

            Syntax.........: BOOL getAccessibleTableCellInfo(long vmID, AccessibleTable accessibleTable, jint row, jint column, AccessibleTableCellInfo *tableCellInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_iRow, $_iColumn;
                             $_ptr_AccessibleTableCellInfo - pointer to AccessibleTableCellInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableCellInfo($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn, ByRef $_ptr_AccessibleTableCellInfo)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTableCellInfo", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_iRow, "int", $_iColumn, "struct*", $_ptr_AccessibleTableCellInfo)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableRowHeader($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
            Description ...: Returns the table row headers of the specified table as a table.

            Syntax.........: BOOL getAccessibleTableRowHeader(long vmID, AccessibleContext acParent, AccessibleTableInfo *tableInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTableInfo - pointer to AccessibleTableInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableRowHeader($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTableRowHeader", "long", $_oVMID, $c_JOBJECT64, $_oParentAccessibleContext, "struct*", $_ptr_AccessibleTableInfo)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableColumnHeader($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
            Description ...: Returns the table column headers of the specified table as a table.

            Syntax.........: BOOL getAccessibleTableColumnHeader(long vmID, AccessibleContext acParent, AccessibleTableInfo *tableInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleTableInfo - pointer to AccessibleTableInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableColumnHeader($_oVMID, $_oParentAccessibleContext, ByRef $_ptr_AccessibleTableInfo)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTableColumnHeader", "long", $_oVMID, $c_JOBJECT64, $_oParentAccessibleContext, "struct*", $_ptr_AccessibleTableInfo)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableRowDescription($_oVMID, $_oParentAccessibleContext, $_iRow)
            Description ...: Returns the description of the specified row in the specified table. The row specifier is zero-based.

            Syntax.........: AccessibleContext getAccessibleTableRowDescription(long vmID, AccessibleContext acParent, jint row);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Accessible Context variable;
                             $_iRow;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableRowDescription($_oVMID, $_oParentAccessibleContext, $_iRow)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getAccessibleTableRowDescription", "long", $_oVMID, $c_JOBJECT64, $_oParentAccessibleContext, "int", $_iRow)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc
        
        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableColumnDescription($_oVMID, $_oParentAccessibleContext, $_iColumn)
            Description ...: Returns the description of the specified column in the specified table. The column specifier is zero-based.

            Syntax.........: AccessibleContext getAccessibleTableColumnDescription(long vmID, AccessibleContext acParent, jint column);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oParentAccessibleContext - Accessible Context variable;
                             $_iColumn;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableColumnDescription($_oVMID, $_oParentAccessibleContext, $_iColumn)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getAccessibleTableColumnDescription", "long", $_oVMID, $c_JOBJECT64, $_oParentAccessibleContext, "int", $_iColumn)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableRowSelectionCount($_oVMID, $_oAccessibleTable)
            Description ...: Returns how many rows in the table are selected.

            Syntax.........: jint getAccessibleTableRowSelectionCount(long vmID, AccessibleTable table);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableRowSelectionCount($_oVMID, $_oAccessibleTable)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleTableRowSelectionCount", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_IsAccessibleTableRowSelected($_oVMID, $_oAccessibleTable, $_iRow)

            Description ...: Returns true if the specified zero based row is selected.

            Syntax.........: BOOL isAccessibleTableRowSelected(long vmID, AccessibleTable table, jint row);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_iRow;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_IsAccessibleTableRowSelected($_oVMID, $_oAccessibleTable, $_iRow)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "isAccessibleTableRowSelected", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_iRow)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableRowSelections($_oVMID, $_oAccessibleTable, $_iCount, ByRef $_ptr_Selections)
            Description ...: Returns an array of zero based indices of the selected rows.

            Syntax.........: BOOL getAccessibleTableRowSelections(long vmID, AccessibleTable table, jint count, jint *selections);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_iCount;
                             $_ptr_Selections - pointer to selections;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableRowSelections($_oVMID, $_oAccessibleTable, $_iCount, ByRef $_ptr_Selections)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTableRowSelections", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_iCount, "int*", $_ptr_Selections)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableColumnSelectionCount($_oVMID, $_oAccessibleTable)
            Description ...: Returns how many columns in the table are selected.

            Syntax.........: jint getAccessibleTableColumnSelectionCount(long vmID, AccessibleTable table);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableColumnSelectionCount($_oVMID, $_oAccessibleTable)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleTableColumnSelectionCount", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_IsAccessibleTableColumnSelected($_oVMID, $_oAccessibleTable, $_iColumn)
            Description ...: Returns true if the specified zero based column is selected.

            Syntax.........: BOOL isAccessibleTableColumnSelected(long vmID, AccessibleTable table, jint column);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_iColumn;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-04
        #ce =============================================================================================================================== 
        Func __JAB_IsAccessibleTableColumnSelected($_oVMID, $_oAccessibleTable, $_iColumn)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "isAccessibleTableColumnSelected", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_iColumn)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableColumnSelections($_oVMID, $_oAccessibleTable, $_iCount, ByRef $_ptr_Selections)
            Description ...: Returns an array of zero based indices of the selected columns.

            Syntax.........: BOOL getAccessibleTableColumnSelections(long vmID, AccessibleTable table, jint count, jint *selections);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_iCount;
                             $_ptr_Selections - pointer to selections;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableColumnSelections($_oVMID, $_oAccessibleTable, $_iCount, ByRef $_ptr_Selections)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleTableColumnSelections", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_iCount, "int*", $_ptr_Selections)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableRow($_oVMID, $_oAccessibleTable, $_index)
            Description ...: Returns the row number of the cell at the specified cell index. The values are zero based.

            Syntax.........: jint getAccessibleTableRow(long vmID, AccessibleTable table, jint index);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableRow($_oVMID, $_oAccessibleTable, $_index)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleTableRow", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableColumn($_oVMID, $_oAccessibleTable, $_index)
            Description ...: Returns the column number of the cell at the specified cell index. The values are zero based.

            Syntax.........: jint getAccessibleTableColumn(long vmID, AccessibleTable table, jint index);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableColumn($_oVMID, $_oAccessibleTable, $_index)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleTableColumn", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleTableIndex($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn)
            Description ...: Returns the index in the table of the specified row and column offset. The values are zero based.

            Syntax.........: jint getAccessibleTableIndex(long vmID, AccessibleTable table, jint row, jint column);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleTable - Accessible Context - Table variable;
                             $_iRow, $_iColumn;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleTableIndex($_oVMID, $_oAccessibleTable, $_iRow, $_iColumn)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleTableIndex", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleTable, "int", $_iRow, "int", $_iColumn)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_TABLE

    #Region ACCESSIBLE_RELATION
    
        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleRelationSet($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleRelationSetInfo)
            Description ...: Returns information about an object's related objects.

            Syntax.........: BOOL getAccessibleRelationSet(long vmID, AccessibleContext accessibleContext, AccessibleRelationSetInfo *relationSetInfo)

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context;
                             $_ptr_AccessibleRelationSetInfo - pointer to AccessibleRelationSetInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleRelationSet($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleRelationSetInfo)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleRelationSet", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleRelationSetInfo)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_RELATION

    #Region ACCESSIBLE_HYPERTEXT
    
        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleHypertext($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleHypertextInfo)
            Description ...: Returns hypertext information associated with a component.

            Syntax.........: BOOL getAccessibleHypertext(long vmID, AccessibleContext accessibleContext, AccessibleHypertextInfo *hypertextInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context;
                             $_ptr_AccessibleHypertextInfo - pointer to AccessibleHypertextInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleHypertext($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleHypertextInfo)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleHypertext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleHypertextInfo)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_ActivateAccessibleHyperlink($_oVMID, $_oAccessibleContext, $_oAccessibleHyperlink)
            Description ...: Requests that a hyperlink be activated.

            Syntax.........: BOOL activateAccessibleHyperlink(long vmID, AccessibleContext accessibleContext, AccessibleHyperlink accessibleHyperlink);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context;
                             $_oAccessibleHyperlink - Accessible Context - AccessibleHyperlink;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_ActivateAccessibleHyperlink($_oVMID, $_oAccessibleContext, $_oAccessibleHyperlink)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "activateAccessibleHyperlink", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, $c_JOBJECT64, $_oAccessibleHyperlink)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleHyperlinkCount($_oVMID, $_oAccessibleHypertext)
            Description ...: Returns the number of hyperlinks in a component. Maps to AccessibleHypertext.getLinkCount. Returns -1 on error.

            Syntax.........: jint getAccessibleHyperlinkCount(const long vmID, const AccessibleHypertext hypertext);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleHypertext - Accessible Context - AccessibleHypertext;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleHyperlinkCount($_oVMID, $_oAccessibleHypertext)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleHyperlinkCount", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleHypertext)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleHypertextExt($_oVMID, $_oAccessibleContext, $_iStartIndex, ByRef $_ptr_AccessibleHypertextInfo)
            Description ...: Iterates through the hyperlinks in a component. 
                             Returns hypertext information for a component starting at hyperlink index nStartIndex. 
                             No more than MAX_HYPERLINKS AccessibleHypertextInfo objects will be returned for each call to this method. 
                             Returns FALSE on error.

            Syntax.........: BOOL getAccessibleHypertextExt(const long vmID, const AccessibleContext accessibleContext, const jint nStartIndex, AccessibleHypertextInfo *hypertextInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context;
                             $_iStartIndex;
                             $_ptr_AccessibleHypertextInfo - pointer to AccessibleHypertextInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleHypertextExt($_oVMID, $_oAccessibleContext, $_iStartIndex, ByRef $_ptr_AccessibleHypertextInfo)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleHypertextExt", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_iStartIndex, "struct*", $_ptr_AccessibleHypertextInfo)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleHypertextLinkIndex($_oVMID, $_oAccessibleHypertext, $_index)
            Description ...: Returns the index into an array of hyperlinks that is associated with a character index in document. 
                             Maps to AccessibleHypertext.getLinkIndex. Returns -1 on error.

            Syntax.........: jint getAccessibleHypertextLinkIndex(const long vmID, const AccessibleHypertext hypertext, const jint nIndex);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleHypertext - Accessible Context - AccessibleHypertext variable;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleHypertextLinkIndex($_oVMID, $_oAccessibleHypertext, $_index)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleHypertextLinkIndex", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleHypertext, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleHyperlink($_oVMID, $_oAccessibleHypertext, $_index, ByRef $_ptr_AccessibleHypertextInfo)
            Description ...: Returns the nth hyperlink in a document. Maps to AccessibleHypertext.getLink. Returns FALSE on error.

            Syntax.........: BOOL getAccessibleHyperlink(const long vmID, const AccessibleHypertext hypertext, const jint nIndex, AccessibleHypertextInfo *hyperlinkInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleHypertext - Accessible Context - AccessibleHypertext variable;
                             $_index;
                             $_ptr_AccessibleHypertextInfo - pointer to AccessibleHypertextInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleHyperlink($_oVMID, $_oAccessibleHypertext, $_index, ByRef $_ptr_AccessibleHypertextInfo)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleHyperlink", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleHypertext, "int", $_index, "struct*", $_ptr_AccessibleHypertextInfo)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc
        
    #EndRegion ACCESSIBLE_HYPERTEXT

    #Region ACCESSIBLE_KEY_BINDINGS

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleKeyBindings($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleKeyBindings)
            Description ...: Returns a list of key bindings associated with a component.

            Syntax.........: BOOL getAccessibleKeyBindings(long vmID, AccessibleContext accessibleContext, AccessibleKeyBindings *keyBindings);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context;
                             $_ptr_AccessibleKeyBindings - pointer to AccessibleKeyBindings struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleKeyBindings($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleKeyBindings)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleKeyBindings", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleKeyBindings)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_KEY_BINDINGS

    #Region ACCESSIBLE_ICONS
    
        #cs #FUNCTION# ====================================================================================================================
            Name...........: _JAB_GetAccessibleIcons($_vmID, $_AccessibleContext, ByRef $icons)
            Description ...: Returns a list of icons associate with a component.

            Syntax.........: BOOL getAccessibleIcons(long vmID, AccessibleContext accessibleContext, AccessibleIcons *icons);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context;
                             $_ptr_AccessibleIcons - pointer to AccessibleIcons struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleIcons($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleIcons)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleIcons", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleIcons)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_ICONS

    #Region ACCESSIBLE_ACTIONS

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleActions($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextActions)

            Description ...: Returns a list of actions that a component can perform.

            Syntax.........: BOOL getAccessibleActions(long vmID, AccessibleContext accessibleContext, AccessibleActions *actions);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleContextActions - pointer to AccessibleContextActions struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleActions($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextActions)
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getAccessibleActions", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleContextActions)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_DoAccessibleActions($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextActionsToDo)

            Description ...: Request that a list of AccessibleActions be performed by a component. 
                             Returns TRUE if all actions are performed. 
                             Returns FALSE when the first requested action fails 
                             in which case "failure" contains the index of the action that failed.

            Syntax.........: BOOL doAccessibleActions(long vmID, AccessibleContext accessibleContext, AccessibleActionsToDo *actionsToDo, jint *failure);     

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext - Accessible Context variable;
                             $_ptr_AccessibleContextActionsToDo - pointer to AccessibleContextActionsToDo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-10
        #ce =============================================================================================================================== 
        Func __JAB_DoAccessibleActions($_oVMID, $_oAccessibleContext, ByRef $_ptr_AccessibleContextActionsToDo)
            Local $_iFailureIndex   ;-- index of failed action
            Local $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "doAccessibleActions", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "struct*", $_ptr_AccessibleContextActionsToDo,  "int*", $_iFailureIndex)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc
    #EndRegion ACCESSIBLE_ACTIONS

    #Region UTILITY
    
        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_IsSameObject($_oVMID, $_oAccessibleContext1, $_oAccessibleContext2)
            Description ...: Returns whether two object references refer to the same object.

            Syntax.........: BOOL IsSameObject(long vmID, JOBJECT64 obj1, JOBJECT64 obj2);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext1, $_oAccessibleContext2 - Accessible Context variables;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_IsSameObject($_oVMID, $_oAccessibleContext1, $_oAccessibleContext2)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "isSameObject", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext1, $c_JOBJECT64, $_oAccessibleContext2)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetParentWithRole($_oVMID, $_oAccessibleContext, $_sRole)
            Description ...: Returns the AccessibleContext with the specified role that is the ancestor of a given object. 
                             The role is one of the role strings defined in Java Access Bridge API Data Stuctures. 
                             If there is no ancestor object that has the specified role, returns (AccessibleContext)0.

            Syntax.........: AccessibleContext getParentWithRole (const long vmID, const AccessibleContext accessibleContext, const wchar_t *role);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;
                             $_sRole - role description;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetParentWithRole($_oVMID, $_oAccessibleContext, $_sRole)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getParentWithRole", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "wstr", $_sRole)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetParentWithRoleElseRoot($_oVMID, $_oAccessibleContext, $_sRole)
            Description ...: Returns the AccessibleContext with the specified role that is the ancestor of a given object. 
                             The role is one of the role strings defined in Java Access Bridge API Data Stuctures. 
                             If an object with the specified role does not exist, returns the top level object for the Java window. 
                             Returns (AccessibleContext)0 on error.

            Syntax.........: AccessibleContext getParentWithRoleElseRoot(const long vmID, const AccessibleContext accessibleContext, const wchar_t *role);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;
                             $_sRole - role description;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetParentWithRoleElseRoot($_oVMID, $_oAccessibleContext, $_sRole)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getParentWithRoleElseRoot", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "wstr", $_sRole)
                If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetTopLevelObject($_oVMID, $_oAccessibleContext)
            Description ...: Returns the AccessibleContext for the top level object in a Java window. 
                             This is same AccessibleContext that is obtained from GetAccessibleContextFromHWND for that window. 
                             Returns (AccessibleContext)0 on error.

            Syntax.........: AccessibleContext getTopLevelObject (const long vmID, const AccessibleContext accessibleContext);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetTopLevelObject($_oVMID, $_oAccessibleContext)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getTopLevelObject", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext)
                 If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetObjectDepth($_oVMID, $_oAccessibleContext)
            Description ...: Returns how deep in the object hierarchy a given object is. 
                             The top most object in the object hierarchy has an object depth of 0. Returns -1 on error.

            Syntax.........: int getObjectDepth (const long vmID, const AccessibleContext accessibleContext);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetObjectDepth($_oVMID, $_oAccessibleContext)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getObjectDepth", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext)
                 If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetActiveDescendent($_oVMID, $_oAccessibleContext)
            Description ...: Returns the AccessibleContext of the current ActiveDescendent of an object. 
                             This method assumes the ActiveDescendent is the component that is currently selected in a container object.

            Syntax.........: AccessibleContext getActiveDescendent (const long vmID, const AccessibleContext accessibleContext);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;


            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetActiveDescendent($_oVMID, $_oAccessibleContext)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getActiveDescendent", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext)
                 If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_RequestFocus($_oVMID, $_oAccessibleContext)
            Description ...: Request focus for a component. Returns whether successful.

            Syntax.........: BOOL requestFocus(const long vmID, const AccessibleContext accessibleContext);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_RequestFocus($_oVMID, $_oAccessibleContext)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "requestFocus", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext)
                 If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetVisibleChildrenCount($_oVMID, $_oAccessibleContext)
            Description ...: Returns the number of visible children of a component. Returns -1 on error.

            Syntax.........: int getVisibleChildrenCount(const long vmID, const AccessibleContext accessibleContext);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetVisibleChildrenCount($_oVMID, $_oAccessibleContext)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getVisibleChildrenCount", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext)
                 If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetVisibleChildren($_oVMID, $_oAccessibleContext, $_iStartIndex, ByRef $_ptr_VisibleChildrenInfo)
            Description ...: Gets the visible children of an AccessibleContext. Returns whether successful.

            Syntax.........: BOOL getVisibleChildren(const long vmID, const AccessibleContext accessibleContext, const int startIndex, VisibleChildrenInfo *visibleChildrenInfo);

            Global vars....: $_JAB_hAccessBridgeDLL; $cP_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleContext;
                             $_iStartIndex;
                             $_ptr_VisibleChildrenInfo - pointer to VisibleChildrenInfo struct;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetVisibleChildren($_oVMID, $_oAccessibleContext, $_iStartIndex, ByRef $_ptr_VisibleChildrenInfo)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getVisibleChildren", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleContext, "int", $_iStartIndex, "struct*", $_ptr_VisibleChildrenInfo)
            If @error Then Return SetError(@error, 0, 0)
                 If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetEventsWaiting()
            Description ...: Gets the number of events waiting to fire.

            Syntax.........: int getEventsWaiting();

            Global vars....: $_JAB_hAccessBridgeDLL
            Parameters.....: -

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-04
        #ce =============================================================================================================================== 
        Func __JAB_GetEventsWaiting()
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getEventsWaiting")
                 If @error Then Return SetError(@error, 0, 0)
            
            Return $_aResult
        EndFunc

    #EndRegion UTILITY

    #Region ACCESSIBLE_VALUE
        #cs
            These functions get AccessibleValue information provided by the Java Accessibility API. 
            An AccessibleContext object has AccessibleValue information contained within it 
            if the flag accessibleValue in the AccessibleContextInfo data structure is set to TRUE. 
            The values returned are strings (char *value) because there is no way to tell in advance if the value is an integer, 
            a floating point value, or some other object that subclasses the Java language construct java.lang.Number.
        #ce

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetCurrentAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
            Description ...: 

            Syntax.........: BOOL GetCurrentAccessibleValueFromContext(long vmID, AccessibleValue av, wchar_t *value, short len);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleValue - Accessible Context - AccessibleValue;
                             $_sValue - variable to store value;
                             $_iLength - string length;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetCurrentAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getCurrentAccessibleValueFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleValue, "wstr", $_sValue, "short", $_iLength)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetMaximumAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
            Description ...: 

            Syntax.........: BOOL GetMaximumAccessibleValueFromContext(long vmID, AccessibleValue av, wchar_ *value, short len);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleValue - Accessible Context - AccessibleValue;
                             $_sValue - variable to store value;
                             $_iLength - string length;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetMaximumAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getMaximumAccessibleValueFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleValue, "wstr", $_sValue, "short", $_iLength)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetMinimumAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
            Description ...: 

            Syntax.........: BOOL GetMinimumAccessibleValueFromContext(long vmID, AccessibleValue av, wchar_ *value, short len);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleValue - Accessible Context - AccessibleValue;
                             $_sValue - variable to store value;
                             $_iLength - string length;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetMinimumAccessibleValueFromContext($_oVMID, $_oAccessibleValue, $_sValue, $_iLength)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "getMinimumAccessibleValueFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleValue, "wstr", $_sValue, "short", $_iLength)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_VALUE

    #Region ACCESSIBLE_SELECTION
        #cs
            These functions get and manipulate AccessibleSelection information provided by the Java Accessibility API. 
            An AccessibleContext has AccessibleSelection information contained within it 
            if the flag accessibleSelection in the AccessibleContextInfo data structure is set to TRUE. 
            The AccessibleSelection support is the first place where the user interface can be manipulated, as opposed to being queries, 
            through adding and removing items from a selection. 
            Some of the functions use an index that is in child coordinates, while other use selection coordinates. 
            For example, add to remove from a selection by passing child indicies (for example, add the fourth child to the selection). 
            On the other hand, enumerating the selected children is done in selection coordinates 
            (for example, get the AccessibleContext of the first object selected).
        #ce

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_AddAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
            Description ...: 

            Syntax.........: void AddAccessibleSelectionFromContext(long vmID, AccessibleSelection as, int i);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleSelection - Accessible Context - AccessibleSelection;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_AddAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "NONE", "addAccessibleSelectionFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleSelection, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_ClearAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection)
            Description ...: 

            Syntax.........: void ClearAccessibleSelectionFromContext(long vmID, AccessibleSelection as);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleSelection - Accessible Context - AccessibleSelection;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_ClearAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "NONE", "clearAccessibleSelectionFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleSelection)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
            Description ...: 

            Syntax.........: jobject GetAccessibleSelectionFromContext(long vmID, AccessibleSelection as, int i);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleSelection - Accessible Context - AccessibleSelection;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "ptr:cdecl", "getAccessibleSelectionFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleSelection, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_GetAccessibleSelectionCountFromContext($_oVMID, $_oAccessibleSelection)
            Description ...: 

            Syntax.........: int GetAccessibleSelectionCountFromContext(long vmID, AccessibleSelection as);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleSelection - Accessible Context - AccessibleSelection;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_GetAccessibleSelectionCountFromContext($_oVMID, $_oAccessibleSelection)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "int:cdecl", "getAccessibleSelectionCountFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleSelection)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: __JAB_IsAccessibleChildSelectedFromContext($_oVMID, $_oAccessibleSelection, $_index)
            Description ...: 

            Syntax.........: BOOL IsAccessibleChildSelectedFromContext(long vmID, AccessibleSelection as, int i);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleSelection - Accessible Context - AccessibleSelection;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_IsAccessibleChildSelectedFromContext($_oVMID, $_oAccessibleSelection, $_index)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "bool:cdecl", "isAccessibleChildSelectedFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleSelection, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: _JAB_RemoveAccessibleSelectionFromContext($_vmID, $as, $i)
            Description ...: 

            Syntax.........: void RemoveAccessibleSelectionFromContext(long vmID, AccessibleSelection as, int i);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleSelection - Accessible Context - AccessibleSelection;
                             $_index;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_RemoveAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection, $_index)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "NONE", "removeAccessibleSelectionFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleSelection, "int", $_index)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

        #cs #FUNCTION# ====================================================================================================================
            Name...........: _JAB_SelectAllAccessibleSelectionFromContext($_vmID, $as)
            Description ...: 

            Syntax.........: void SelectAllAccessibleSelectionFromContext(long vmID, AccessibleSelection as);

            Global vars....: $_JAB_hAccessBridgeDLL; $c_JOBJECT64
            Parameters.....: $_oVMID - JVM ID variable;
                             $_oAccessibleSelection - Accessible Context - AccessibleSelection;

            Author ........: exorcistas@github.com
            Modified.......: 2020-03-11
        #ce =============================================================================================================================== 
        Func __JAB_SelectAllAccessibleSelectionFromContext($_oVMID, $_oAccessibleSelection)
            $_aResult = DllCall($_JAB_hAccessBridgeDLL, "NONE", "selectAllAccessibleSelectionFromContext", "long", $_oVMID, $c_JOBJECT64, $_oAccessibleSelection)
                If @error Then Return SetError(@error, 0, 0)
                
            Return $_aResult
        EndFunc

    #EndRegion ACCESSIBLE_SELECTION

#EndRegion CORE_FUNCTIONS