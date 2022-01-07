#include-once

Global $_JAB_HOME_DIR, $_JAB_WindowsAccessBridge_DLL, $_JAB_hAccessBridgeDLL
Global Enum _
      $_JAB_DEBUG_NONE = 0, _
		$_JAB_DEBUG_CORE, _
      $_JAB_DEBUG_EXT

Global $_JAB_DEBUG = $_JAB_DEBUG_CORE


#cs # Java Access Bridge API Data Stuctures # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   The Java Access Bridge API data structures are contained in the file AccessBridgePackages.h.

   Important Data Structures

   https://docs.oracle.com/javase/9/access/jaapi.htm#GUID-A4B488F9-622F-4A5C-932D-7C61BA2FFC86

   There are data structures in this file that you do not need (and can ignore); they are used as part of the inter-process communication mechanism of the two Java Access Bridge DLLs. The data structures of importance are as follows:

   #define MAX_STRING_SIZE     1024
   #define SHORT_STRING_SIZE    256

   typedef struct AccessibleContextInfoTag {
   wchar_ name[MAX_STRING_SIZE];        // the AccessibleName of the object
   wchar_ description[MAX_STRING_SIZE]; // the AccessibleDescription of the object
   wchar_ role[SHORT_STRING_SIZE];      // localized AccesibleRole string
   wchar_ states[SHORT_STRING_SIZE];    // localized AccesibleStateSet string
                                          //   (comma separated)
   jint indexInParent                   // index of object in parent
   jint childrenCount                   // # of children, if any
   jint x;                              // screen x-axis co-ordinate in pixels
   jint y;                              // screen y-axis co-ordinate in pixels
   jint width;                          // pixel width of object
   jint height;                         // pixel height of object
   BOOL accessibleComponent;            // flags for various additional
   BOOL accessibleAction;               // Java Accessibility interfaces
   BOOL accessibleSelection;            // FALSE if this object doesn't
   BOOL accessibleText;                 // implement the additional interface
   BOOL accessibleInterfaces;           // new bitfield containing additional
                                          //   interface flags
   } AccessibleContextInfo;
   
   typedef struct AccessibleTextInfoTag {
   jint charCount;       // # of characters in this text object
   jint caretIndex;      // index of caret
   jint indexAtPoint;    // index at the passsed in point
   } AccessibleTextInfo;

   typedef struct AccessibleTextItemsInfoTag {
   wchar_t letter;
   wchar_t word[SHORT_STRING_SIZE];
   wchar_t sentence[MAX_STRING_SIZE];
   } AccessibleTextItemsInfo;
   
   typedef struct AccessibleTextSelectionInfoTag {
   jint selectionStartIndex;
   jint selectionEndIndex;
   wchar_t selectedText[MAX_STRING_SIZE];
   } AccessibleTextSelectionInfo;
   
   typedef struct AccessibleTextRectInfoTag  {
   jint x;          // bounding recttangle of char at index, x-axis co-ordinate
   jint y;          // y-axis co-ordinate
   jint width;      // bounding rectangle width
   jint height;     // bounding rectangle height
   } AccessibleTextRectInfo;
   
   typedef struct AccessibleTextAttributesInfoTag {
   BOOL bold;
   BOOL italic;
   BOOL underline;
   BOOL strikethrough;
   BOOL superscript;
   BOOL subscript;
   wchar_t backgroundColor[SHORT_STRING_SIZE];
   wchar_t foregroundColor[SHORT_STRING_SIZE];
   wchar_t fontFamily[SHORT_STRING_SIZE];
   jint fontSize;
   jint alignment;
   jint bidiLevel;
   jfloat firstLineIndent;
   jfloat leftIndent;
   jfloat rightIndent;
   jfloat lineSpacing;
   jfloat spaceAbove;
   jfloat spaceBelow;
   wchar_t fullAttributesString[MAX_STRING_SIZE];
   } AccessibleTextAttributesInfo;

   typedef struct AccessibleTableInfoTag  {
   JOBJECT64 caption;  // AccesibleContext
   JOBJECT64 summary;  // AccessibleContext
   jint rowCount;
   jint columnCount;
   JOBJECT64 accessibleContext;
   JOBJECT64 accessibleTable;
   } AccessibleTableInfo;

   typedef struct AccessibleTableCellInfoTag  {
   JOBJECT64  accessibleContext;
   jint       index;
   jint       row;
   jint       column;
   jint       rowExtent;
   jint       columnExtent;
   jboolean   isSelected;
   } AccessibleTableCellInfo;

   typedef struct AccessibleRelationSetInfoTag {
   jint relationCount;
   AccessibleRelationInfo relations[MAX_RELATIONS];
   } AccessibleRelationSetInfo;

   typedef struct AccessibleRelationInfoTag {
   wchar_t key[SHORT_STRING_SIZE];
   jint targetCount;
   JOBJECT64 targets[MAX_RELATION_TARGETS];  // AccessibleContexts
   } AccessibleRelationInfo;


   typedef struct AccessibleHypertextInfoTag {
   jint linkCount;                                 // number of hyperlinks
   AccessibleHyperlinkInfo links[MAX_HYPERLINKS];  // the hyperlinks
   JOBJECT64 accessibleHypertext;                  // AccessibleHypertext object
   } AccessibleHypertextInfo;

   typedef struct AccessibleHyperlinkInfoTag {
   wchar_t text[SHORT_STRING_SIZE]; // the hyperlink text
   jint startIndex;                 // index in the hypertext document where the link begins
   jint endIndex;                   // index in the hypertext document where the link ends
   JOBJECT64 accessibleHyperlink;   // AccessibleHyperlink object
   } AccessibleHyperlinkInfo;

   typedef struct AccessibleKeyBindingsTag {
   int keyBindingsCount; // number of key bindings
   AccessibleKeyBindingInfo keyBindingInfo[MAX_KEY_BINDINGS];
   } AccessibleKeyBindings;

   typedef struct AccessibleKeyBindingInfoTag {
   jchar character; // the key character
   jint modifiers;  // the key modifiers
   } AccessibleKeyBindingInfo;

   typedef struct  AccessibleIconsTag {
   jint iconsCount;                            // number of icons
   AccessibleIconInfo iconInfo[MAX_ICON_INFO]; // the icons
   } AccessibleIcons;

   typedef struct AccessibleIconInfoTag {
   wchar_t description[SHORT_STRING_SIZE]; // icon description
   jint height;                            // icon height
   jint width;                             // icon width
   } AccessibleIconInfo;

   typedef struct AccessibleActionsTag {
   jint actionsCount;                                // number of actions
   AccessibleActionInfo actionInfo[MAX_ACTION_INFO]; // the action information
   } AccessibleActions;

   typedef struct AccessibleActionInfoTag {
   wchar_t name[SHORT_STRING_SIZE]; // action name
   } AccessibleActionInfo;

   typedef struct AccessibleActionsToDoTag {
   jint actionsCount;                               // number of actions to do
   AccessibleActionInfo actions[MAX_ACTIONS_TO_DO]; // the accessible actions to do
   } AccessibleActionsToDo;

   typedef struct VisibleChildrenInfoTag {
   int returnedChildrenCount;                        // number of children returned
   AccessibleContext children[MAX_VISIBLE_CHILDREN]; // the visible children
   } VisibleChildrenInfo;
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Global Const $c_JOBJECT64 = (@AutoItX64) ? ("UINT64") : ("UINT") ;~ JOBJECT64
Global Const $cP_JOBJECT64 = (@AutoItX64) ? ("UINT64*") : ("UINT*") ;~ POINTER(JOBJECT64)

#Region ACCESSIBLE_INFORMATION
   Global Const $MAX_BUFFER_SIZE = 10240
   Global Const $MAX_STRING_SIZE = 1024
   Global Const $SHORT_STRING_SIZE = 256

   Global Const $tagAccessBridgeVersionInfo = _
      "WCHAR VMversion[256];" & _                 ; output of "java -version"
      "WCHAR bridgeJavaClassVersion[256];" & _    ; version of the AccessBridge.class
      "WCHAR bridgeJavaDLLVersion[256];" & _      ; version of JavaAccessBridge.dll
      "WCHAR bridgeWinDLLVersion[256]"            ; version of WindowsAccessBridge.dll

   Global Const $tagAccessibleContextInfo = _
      "WCHAR name[1024];" & _             ; the AccessibleName of the object
      "WCHAR description[1024];" & _      ; the AccessibleDescription of the object
      "WCHAR role[256];" & _              ; localized AccesibleRole string
      "WCHAR role_en_US[256];" & _        ; AccesibleRole string in the en_US locale
      "WCHAR states[256];" & _            ; localized AccesibleStateSet string (comma separated)
      "WCHAR states_en_US[256];" & _      ; AccesibleStateSet string in the en_US locale (comma separated)
      "INT indexInParent;" & _            ; index of object in parent
      "INT childrenCount;" & _            ; # of children, if any
      "INT x;" & _                        ; screen x-axis co-ordinate in pixels
      "INT y;" & _                        ; screen y-axis co-ordinate in pixels
      "INT width;" & _                    ; pixel width of object
      "INT height;" & _                   ; pixel height of object
      "BOOL accessibleComponent;" & _     ; flags for various additional Java Accessibility interfaces
      "BOOL accessibleAction;" & _
      "BOOL accessibleSelection;" & _     ; FALSE if this object does not implement the additional interface in question
      "BOOL accessibleText;" & _
      "BOOL accessibleInterfaces"         ; new bitfield containing additional interface flags

   Global Const $tagAccessibleTextInfo = _
      "INT charCount;" & _                ; # of characters in this text object
      "INT caretIndex;" & _               ; index of caret
      "INT indexAtPoint"                  ; index at the passsed in point

   Global Const $tagAccessibleTextItemsInfo = _
      "WCHAR letter;" & _
      "WCHAR word[256];" & _
      "WCHAR sentence[1024]"

   Global Const $tagAccessibleTextSelectionInfo = _
      "INT selectionStartIndex;" & _
      "INT selectionEndIndex;" & _
      "WCHAR selectedText[1024]"

   Global Const $tagAccessibleTextRectInfo = _
      "INT x;" & _                        ; bounding recttangle of char at index, x-axis co-ordinate
      "INT y;" & _                        ; y-axis co-ordinate
      "INT width;" & _                    ; bounding rectangle width
      "INT height"                        ; bounding rectangle height

   Global Const $tagAccessibleTextAttributesInfo = _
      "BOOL bold;" & _
      "BOOL italic;" & _
      "BOOL underline;" & _
      "BOOL strikethrough;" & _
      "BOOL superscript;" & _
      "BOOL subscript;" & _
      "WCHAR backgroundColor[256];" & _
      "WCHAR foregroundColor[256];" & _
      "WCHAR fontFamily[256];" & _
      "INT fontSize;" & _
      "INT alignment;" & _
      "INT bidiLevel;" & _
      "FLOAT firstLineIndent;" & _
      "FLOAT LeftIndent;" & _
      "FLOAT rightIndent;" & _
      "FLOAT lineSpacing;" & _
      "FLOAT spaceAbove;" & _
      "FLOAT spaceBelow;" & _
      "WCHAR fullAttributesString[1024]"

#EndRegion ACCESSIBLE_INFORMATION

#Region ACCESSIBLE_TABLE
   Global Const $MAX_TABLE_SELECTIONS = 64

   Global Const $tagAccessibleTableInfo = _
      "UINT64 caption;" & _               ; AccessibleContext
      "UINT64 summary;" & _               ; AccessibleContext
      "INT rowCount;" & _
      "INT columnCount;" & _
      "UINT64 accessibleContext;" & _
      "UINT64 accessibleTable"

   Global Const $tagAccessibleTableCellInfo = _
      "UINT64 accessibleContext;" & _
      "INT index;" & _
      "INT row;" & _
      "INT column;" & _
      "INT rowExtent;" & _
      "INT columnExtent;" & _
      "BOOL isSelected"

#EndRegion ACCESSIBLE_TABLE

#Region ACCESSIBLE_RELATION_SET
   Global Const $MAX_RELATION_TARGETS = 25
   Global Const $MAX_RELATIONS = 5

   Global Const $tagAccessibleRelationInfo = _
      "WCHAR key[256];" & _
      "INT targetCount;" & _
      "UINT64 targets[25]"                ; AccessibleContexts

   Local $tag_Relations = ""
      For $i = 1 To 5
         If $tag_Relations = "" Then
            $tag_Relations = $tagAccessibleRelationInfo
         Else
            $tag_Relations = $tag_Relations & ";" & $tagAccessibleRelationInfo
         EndIf
      Next
   Global Const $tagAccessibleRelationSetInfo = _
      "INT relationCount;" & _
      $tag_Relations                      ; $tAccessibleRelationInfo relations[5]

#EndRegion ACCESSIBLE_RELATION_SET

#Region ACCESSIBLE_HYPERTEXT
   Global Const $MAX_HYPERLINKS = 64   ;~ maximum number of hyperlinks returned

   Global Const $tagAccessibleHyperlinkInfo = _
      "WCHAR text[256];" & _              ; the hyperlink text
      "INT startIndex;" & _               ; index in the hypertext document where the link begins
      "INT endIndex;" & _                 ; index in the hypertext document where the link ends
      "UINT64 accessibleHyperlink"        ; AccessibleHyperlink object

   Local $tag_Links = ""
      For $i = 1 To 64
         If $tag_Links = "" Then
            $tag_Links = $tagAccessibleHyperlinkInfo
         Else
            $tag_Links = $tag_Links & ";" & $tagAccessibleHyperlinkInfo
         EndIf
      Next
   Global Const $tagAccessibleHypertextInfo = _
      "INT linkCount;" & _                ; number of hyperlinks
      $tag_Links & ";" & _                ; the hyperlinks ,$tAccessibleHyperlinkInfo links[64]
      "UINT64 accessibleHypertext"        ; AccessibleHypertext object

#EndRegion ACCESSIBLE_HYPERTEXT

#Region ACCESSIBLE_KEY_BINDINGS
   Global Const $MAX_KEY_BINDINGS = 8

   Global Const $tagAccessibleKeyBindingInfo = _
      "WCHAR character;" & _              ; the key character
      "INT modifiers"                     ; the key modifiers

   Local $tag_KeyBindingInfo = ""
      For $i = 1 To 8
         If $tag_KeyBindingInfo  = "" Then
            $tag_KeyBindingInfo  = $tagAccessibleKeyBindingInfo
         Else
            $tag_KeyBindingInfo  = $tag_KeyBindingInfo  & ";" & $tagAccessibleKeyBindingInfo
         EndIf
      Next
   Global Const $tagAccessibleKeyBindings = _
      "INT keyBindingsCount;" & _         ; number of key bindings
      $tag_KeyBindingInfo                 ; $tAccessibleKeyBindingInfo keyBindingInfo[8]

#EndRegion ACCESSIBLE_KEY_BINDINGS

#Region ACCESSIBLE_ICON
   Global Const $MAX_ICON_INFO = 8

   Global Const $tagAccessibleIconInfo = _
                           "WCHAR description[256];" & _       ; icon description
                           "INT height;" & _                   ; icon height
                           "INT width"                         ; icon width

   Local $tag_IconInfo = ""
      For $i = 1 To 8
         If $tag_IconInfo  = "" Then
            $tag_IconInfo  = $tagAccessibleIconInfo
         Else
            $tag_IconInfo  = $tag_IconInfo  & ";" & $tagAccessibleIconInfo
         EndIf
      Next
   Global Const $tagAccessibleIcon = _
                           "INT iconsCount;" & _               ; number of icons
                           $tag_IconInfo                       ; the icons ,$tAccessibleIconInfo iconInfo[8]

#EndRegion ACCESSIBLE_ICON

#Region ACCESSIBLE_ACTION
   Global Const $MAX_ACTION_INFO = 256
   Global Const $MAX_ACTIONS_TO_DO = 32

   Global Const $tagAccessibleActionInfo = "WCHAR name[256]"

   Local $tag_ActionInfo = ""
      For $i = 1 To 256
         If $tag_ActionInfo = "" Then
            $tag_ActionInfo = $tagAccessibleActionInfo
         Else
            $tag_ActionInfo = $tag_ActionInfo & ";" & $tagAccessibleActionInfo
         EndIf
      Next
   Global Const $tagAccessibleActions = _
      "INT actionsCount;" & _
      $tag_ActionInfo                     ; $tAccessibleActionInfo actionInfo[256]

   Local $tag_Actions = ""
      For $i = 1 To 32
         If $tag_Actions = "" Then
            $tag_Actions = $tagAccessibleActionInfo
         Else
            $tag_Actions = $tag_Actions & ";" & $tagAccessibleActionInfo
         EndIf
      Next
   Global Const $tagAccessibleActionsToDo = _
      "INT actionsCount;" & _
      $tag_Actions                            ; $tAccessibleActions actions[32]

#EndRegion ACCESSIBLE_ACTION

#Region VISIBLE_CHILDREN
   Global Const $MAX_VISIBLE_CHILDREN = 256

   Global Const $tagVisibleChildenInfo = _
      "INT returnedChildrenCount;" & _    ; number of children returned
      "UINT64 children[256]"              ; the visible children

#EndRegion VISIBLE_CHILDREN