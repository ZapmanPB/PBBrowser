; *******************************************************************************
;
;  File for declarations of constants and global variables for PBBrowser
;
; *******************************************************************************
IncludePath  "Zapman libraries"
XIncludeFile "ZapmanCommon.pb"
XIncludeFile "Alert.pb"
XIncludeFile "TOM_Functions.pb"
XIncludeFile "Pipe.pb"
XIncludeFile "FastStringFunctions.pb"
XIncludeFile "ExpressionEvaluator.pb"
XIncludeFile "ChooseLanguage.pb"
;
Global MyAppDataFolder$ = GetSystemFolder(#CSIDL_COMMON_APPDATA) + "\" + #NomProg + "\" ; Address of the data folder.
Global PBBrowserPrefile$ = MyAppDataFolder$ + #NomProg + ".prefs" ; Address of the preferences file.
Global BOM ; Will contain the BOM found when reading the files.
Global PBBListOfFiles$ = "" ; Will contain the list of files linked to the main file.
Global PBBListOfBinaries$ = "" ; Will contain the list of Binary files linked to the main file.
Global PBBFirstLaunch ; Indicates the first launch of the application.
Global PureBasicProgAddr$, PBUnderCursor$, PBBFicPrincipalPB$, FicActualPB$, PBBTempFile$ ; Will contain the arguments sent by the PureBasic editor.
;
; ****************************************************************************** 
;
Global PBFunctionList$       ; Will contain the list of native functions.
Global PBFunctionListLCase$  ; Same list in LCase() version.
;
; ****************************************************************************** 
;
;           Variables intended to be modified by the user
;
Global PBBDefaultCursor = LoadCursor_(0, #IDC_ARROW) ; Standard arrow cursor to be displayed over RichEdit gadgets.
Global PBBGreyColor = RGB(120, 120, 120)                ; Color used in the "Details" panel to display everything following the element name.
Global PBBpcolorNotUsed = RGB(130, 140, 140)            ; Light gray for elements not used by the code.
Global PBBDarkRedColor = RGB(180, 0, 0)                 ; For coloring the element name in the 'Details' and 'Found in...' panels.
Global PBBSetValueColor = RGB(150, 90, 0)               ; For coloring 'Set', 'Read', 'Ret', 'Param' and 'Poke' mentions in 'Found in...'.
Global PBBTitleColor = RGB(150, 150, 150)            ; Color applied to "Source" and "Expression" in the main panel, as well as "Included Files" in the "Files" panel.
Global PBBTitleFont = LoadFont(100, "Segoe UI", 9, #PB_Font_Bold | #PB_Font_HighQuality) ; Same for the font.
Global PBBProgressLegendFont = LoadFont(101, "Segoe UI", 7, #PB_Font_HighQuality)           ; Font used for the text above the progress bar.
Global ZapmanGadgetsFont = LoadFont(#PB_Any, "Segoe UI", 9)                                  ; Font used for all gadgets
;
Global PBB_Bullet$ = "○ "                            ; Bullet character used in 'Details' and 'Found in...'.
Global PBB_LineBreak$ = Chr(13) + PBB_Bullet$
;
; The code is displayed in 'Consolas', size 10.
; The 'Consolas' font is a variant of 'Courier', but it has
; the advantage of displaying a more extended character set
; that fully exploits the possibilities of UTF-8.
Global PBBCodeFontStyle$ = "Name(Consolas),Bold(0),Italic(0),Size(10)"
; 
Global PBBDefaultFontStyle$ ; Default font for the content of the panels (defined at the beginning of PBBrowser.pb).
;
; ****************************************************************************** 
;
;                Parameters for the arrow buttons and the code
;
; ****************************************************************************** 
;
#PBBRTFMarker$ = "PBBMark_:" ; Will be used to identify images inserted in RichEdit gadgets.
#PBBLeftArrowMarker$ = "Previous" ; Used only in the code as the name for the << navigation button.
#PBBRightArrowMarker$ = "Next"    ; Used only in the code as the name for the >> navigation button.
#PBBLeftArrow$ = "◄";"←"          ; Substitution character when the OS version doesn't support embedding images in text.
#PBBRightArrow$ = "►" ;"→"        ; Substitution character when the OS version doesn't support embedding images in text.
#PBBArrowsFontStyle$ = "Bold,Size(12)" ; Font for the substitution characters.
;
Global CharButtons
If OSVersion() <= #PB_OS_Windows_Server_2008_R2
  ; The program is running on Windows 7 or an earlier version,
  ; unable to display an image in RichEdit. As an alternative
  ; solution, we will display the characters #PBBLeftArrow$ and #PBBRightArrow$
  CharButtons = #True
Else
  CharButtons = #False
EndIf
;
#OutOfElementsName = "OutOfElements"
;
; *********************************************************
;    List of PureBasic Keywords that can preceed a variable name
;
#OperatorAsWord$ = ",If,ElseIf,While,Until,Select,Case,Default,With,To,ProcedureReturn,And,Or,Xor,Not,Protected,Global,Static,Shared,Define,"
;
; *********************************************************
;            Definition of element types
;
Enumeration PBBTypes
  #PBBProcedure       ; Element of type 'Procedure'.
  #PBBStructure       ; Same for structures
  #PBBMacro           ; Same for macros
  #PBBEnumeration     ; Same for enumerations
  #PBBInterface       ; Same for interfaces
  #PBBLabel           ; Same for labels
  #PBBConstante       ; Same for constants
  #PBBVariable        ; Same for variables
  #EndEnumPBBElementTypes ; Signals the end of the elements found in the files
  #PBBNativeFunction  ; Same for native PureBasic functions
  #PBBBasicKeyword    ; Same for Basic keywords (While, Wend, If, Then, etc.)
  #EndEnumPBBTypes    ; Signals the end of the list
  #NotAnElement
EndEnumeration
;
; Assign a name for each type:
TypesList$ = "Procedure,Structure,Macro,Enumeration,Interface,Label,Constante,Variable,,PBNativeFunction,PBKeyword"
;
; Register these names in the PBBTypeNames$() array
Global Dim PBBTypeNames$(#EndEnumPBBTypes - 1)
For ct = 0 To (#EndEnumPBBTypes - 1)
  PBBTypeNames$(ct) = StringField(TypesList$, ct + 1, ",")
Next
;
; The element #PBBVariable will be the only one that will be split into several
; categories. These categories are listed below:
#PBBVariableSpecies$ = "Global,Shared"
;
; ************************************************************
;                   Element arrays
;
; The following arrays will be used to complete various stages of exploration,
; in order to obtain the lists of elements (procedures, structures, enumerations, etc.) that
; appear in the examined files.
;
; ListOfAllElementsNbr() will contain the number of elements
; found for a given element type.
Global Dim ListOfAllElementsNbr(#EndEnumPBBElementTypes - 1)
;
; ListOfUsedElementsNbr() will contain the number of elements
; used for a given element type.
Global Dim ListOfUsedElementsNbr(#EndEnumPBBElementTypes - 1)
;
; ListOfAllElementsColor() will contain the colors to be used to color
; each element in the code.
Global Dim ListOfAllElementsColor(#EndEnumPBBTypes - 1)
;
; The ListCompletionAll and ListCompletionUsed arrays will indicate if the
; arrays above have already been completed by the file exploration.
Global Dim ListCompletionAll(#EndEnumPBBElementTypes - 1)
Global Dim ListCompletionUsed(#EndEnumPBBElementTypes - 1)
;
; For the first exploration phase, the following array will contain the stage at which
; the completion procedure was interrupted during background work.
Global Dim ListCompletionStage(#EndEnumPBBElementTypes - 1)
;
; For the first exploration phase, the following array will contain the file name
; being examined during background work.
Global Dim ListCompletionReference$(#EndEnumPBBElementTypes - 1)
;
;
; **************************************************************
;    Values for the completion state of the element list
;
Enumeration ListCompletionState ; Current state of completion for a given element type.
  #ListCompletion_Undone
  #ListCompletion_Pending
  #ListCompletion_StageCompleted
  #ListCompletion_Done
  #ListCompletion_Printed
EndEnumeration
;
Enumeration CompletionState ; General completion state
  #Completion_Completed
  #Completion_Uncomplete
  #Completion_Error
EndEnumeration
;
; *************************************************************
;    Structure of the 'ElementsList()' which will contain
;        the list of various elements of the examined code.
;
Structure Elements
  Name.s                ; Name of the element.
  NameLCase.s           ; Name of the element in lowercase, and without the type, for structured or typed variables.
  Type.l                ; Element type (procedure, structure, constant, variable, etc.).
  Used.b                ; Indicates if the element is marked as 'used' by the program.
  FileName.s            ; .pb file where the element is defined.
  Declaration.s         ; Declaration line of the element
  DeclarationDetails.s  ; Used only for constants (contains the name of the enumeration where
                        ;   the constant is declared) and procedures (contains the line where the procedure is declared).
  VariableSpecies.s     ; Used only for variables. Specifies the category (Global, Shared, List, Array, Var, Map).
  StartingLine.l        ; Line number where the element is declared in the code
  EndingLine.l          ; Line number at the end of the element ('EndProcedure', for example).
  StartingPos.l         ; Exact position of the start of the element in the code.
  EndingPos.l           ; Position of the end of the element.
  Parents.s             ; List of indices of 'Parents' (elements that contain the current element).
  Children.s            ; List of indices of 'Children' (elements contained by the current element).
  OutOfElementLinePos.s ; Positions (files and line numbers) where the element is referenced in the main code.
  Value.s               ; Used only for constants (constant value).
  Comment.s             ; Contains the comments for elements that have no code.
EndStructure
;
Global NewList ElementsList.Elements()
;
; *********************************************************
; Definition of the different panels of the 'PannelGadgets'
;
Enumeration PBBPanels ; Numbers assigned to panels
  ; MainPBBPanel
  #FilePBBPanel
  #ListPBBPanel
  #DetailPBBPanel
  #FoundInFilesPBBPanel
  ;
  ; ListPBBPanel
  #ProcPBBPanel
  #StructurePBBPanel
  #MacroPBBPanel
  #EnumPBBPanel
  #InterfacePBBPanel
  #LabelPBBPanel
  #ConstantePBBPanel
  #VariablePBBPanel
  #EndEnumPBBPanels
EndEnumeration
;
; Assign a name for each panel:
Global Dim PBBPanelNames$(#EndEnumPBBPanels - 1)
TypesList$ = "Files,Lists,Details,FoundIn,Procedures,Structures,Macros,Enumerations,Interfaces,Labels,Constantes,Variables"
; Store these names in the PBBPanelNames$() array
For ct = 0 To (#EndEnumPBBPanels - 1)
  PBBPanelNames$(ct) = StringField(TypesList$, ct + 1, ",")
Next
;
; The following arrays will provide details about each panel:
Global Dim NoPBBPanel(#EndEnumPBBPanels - 1)
Global Dim NoGadgetPBBPanel(#EndEnumPBBPanels - 1)
Global Dim NoREGadgetOfPBBPanel(#EndEnumPBBPanels - 1)
;
; The following will create a bridge between ListPBBPanel
; and element types, specifying which types of elements are displayed
; by each panel.
Global Dim TypeElementOfPBBPanel(#EndEnumPBBPanels - 1)
TypeElementOfPBBPanel(#FilePBBPanel) = #NotAnElement
TypeElementOfPBBPanel(#ListPBBPanel) = #NotAnElement
TypeElementOfPBBPanel(#DetailPBBPanel) = #NotAnElement
TypeElementOfPBBPanel(#FoundInFilesPBBPanel) = #NotAnElement
TypeElementOfPBBPanel(#ProcPBBPanel) = #PBBProcedure
TypeElementOfPBBPanel(#StructurePBBPanel) = #PBBStructure
TypeElementOfPBBPanel(#MacroPBBPanel) = #PBBMacro
TypeElementOfPBBPanel(#EnumPBBPanel) = #PBBEnumeration
TypeElementOfPBBPanel(#InterfacePBBPanel) = #PBBInterface
TypeElementOfPBBPanel(#LabelPBBPanel) = #PBBLabel
TypeElementOfPBBPanel(#ConstantePBBPanel) = #PBBConstante
TypeElementOfPBBPanel(#VariablePBBPanel) = #PBBVariable
;
; *********************************************************
;
; The following enumerations will be used to manage
; background tasks.
;
Enumeration PBB_PriorityMode ; Managing display for background tasks.
  #WorkInBackGround
  #ShowCompletionWindow
  #FinishCompletionNow
EndEnumeration
;
Enumeration PBB_BackgroundTasksState ; Execution state of background tasks.
  #BackgroundTasksUncompleted
  #BackgroundTasksHavePriority
  #BackgroundTasksCompleted
  #BackgroundTasksMustRestart
EndEnumeration
;
; ****************************************************************************** 
;
;                        For keyboard shortcuts:
;
Enumeration PBB_Menus
  #PBBMenu_Return
  #PBBMenu_Escape
  #PBBMenu_Find
  #PBBMenu_THAT
  #PBBMenu_Up
  #PBBMenu_Down
  #PBBMenu_PageUp
  #PBBMenu_PageDown
  #PBBMenu_SetShortcut
  #PBBMenu_NextPanel
  #PBBMenu_PreviousPanel
  #PBBMenu_NextPage
  #PBBMenu_PreviousPage
EndEnumeration
;
Enumeration REMenuItems
  #REM_ZoomIn = 100
  #REM_ZoomOut
  #REM_ZoomReset
  #REM_FindInPannel
  #REM_CopyAll
  #REM_SaveAsText
  #REM_SaveAsRTF
EndEnumeration
;
; ******************************************************************************
;
;    For storing the gadget numbers of the main window:
;
Structure PBBGadgets
  PBBWindow.i
  AdrTitle.i
  Adr_gadget.i
  BChangeAdresse.i
  ProcNameTitle.i
  SearchedExpression_gadget.i
  BSearchExpression.i
  GLine.i
  TProgressBar.i
  EProgressBar.i
  BStats.i
  BRefresh.i
  MainPanelGadget.i
  ListsPGadget.i
  BREMenu.i
  REMenu.i
  BAbout.i
  BStick.i
  BLanguage.i
  BQuit.i
  IWhiteOver.i
  Disabled.i
  Menu.i
  ToolTip.i
EndStructure
;
Global GPBBGadgets.PBBGadgets;
;
; ******************************************************************************
;
;                  For managing item searches:
;
Structure LastSearchDetails
  ElementName$
  ElementType.i
  TypeName$
EndStructure
;
Enumeration TypeOfSearch
  #NoSearch = -1
  #DoManualSearch = -2
  #DoProgrammedSearch = -3
  #FindAnyType = -4
  #FindFollowedByParentheses = -5
  #FindNotFollowedByParentheses = -6

EndEnumeration
;
; ******************************************************************************
;
; For managing navigation within the 'Details' and 'Found in...' panels:
;
Structure SearchDetails
  Search_Index.i
  Search_Details.LastSearchDetails
  Search_DetailScrollPos.Point
  Search_FoundInScrollPos.Point
EndStructure
;
Global NewList ListOfSearchs.SearchDetails()
;
; ******************************************************************************
;
;               For managing file cache memory:
;
Enumeration ValuesForCheckFileUpdating
  #FileIsUpdated
  #FileWasUpToDate
  #FileMustBeUpdated
  #FileDoesntExist
  #FileHasBeenDeleted
EndEnumeration
;
; ******************************************************************************
;
; The following is a backup copy of the PureBasic preferences file,
; for values essential to PBBrowser. This allows it to run
; on a machine where PureBasic is not installed.

Global PBPrefSecure$
PBPrefSecure$ = "ASMKeywordColor = 7490450" + Chr(13) + "BackgroundColor = 14680063" + Chr(13) + "BasicKeywordColor = 6710784" + Chr(13)
PBPrefSecure$ + "CommentColor = 11184640" + Chr(13) + "ConstantColor = 7490450" + Chr(13) + "LabelColor = 0" + Chr(13) + "NormalTextColor = 0" + Chr(13)
PBPrefSecure$ + "NumberColor = 0" + Chr(13) + "OperatorColor = 0" + Chr(13) + "PointerColor = 0" + Chr(13) + "PureKeywordColor = 6710784" + Chr(13)
PBPrefSecure$ + "SeparatorColor = 0" + Chr(13) + "StringColor = 14643200" + Chr(13) + "StructureColor = 0" + Chr(13) + "LineNumberColor = 8421504" + Chr(13)
PBPrefSecure$ + "LineNumberBackColor = 14155775" + Chr(13) + "MarkerColor = 11184640" + Chr(13) + "CurrentLineColor = 12058623" + Chr(13)
PBPrefSecure$ + "SelectionColor = 14120960" + Chr(13) + "SelectionFrontColor = 16777215" + Chr(13) + "CursorColor = 0" + Chr(13)
PBPrefSecure$ + "Debugger_LineColor = 16771304" + Chr(13) + "Debugger_LineSymbolColor = 16771304" + Chr(13) + "Debugger_ErrorColor = 255" + Chr(13)
PBPrefSecure$ + "Debugger_ErrorSymbolColor = 255" + Chr(13) + "Debugger_BreakPointColor = 16776960" + Chr(13)
PBPrefSecure$ + "Debugger_BreakpoinSymbolColor = 16776960" + Chr(13) + "DisabledBackColor = 16316664" + Chr(13) + "GoodBraceColor = 6710784" + Chr(13)
PBPrefSecure$ + "BadBraceColor = 255" + Chr(13) + "ProcedureBackColor = 14680063" + Chr(13) + "CustomKeywordColor = 6684672" + Chr(13)
PBPrefSecure$ + "Debugger_WarningColor = 53503" + Chr(13) + "Debugger_WarningSymbolColor = 53503" + Chr(13) + "IndentColor = 11184640" + Chr(13)
PBPrefSecure$ + "ModuleColor = 0" + Chr(13) + "SelectionRepeatColor = 5746176" + Chr(13) + "PlainBackground = 14680063" + Chr(13)
PBPrefSecure$ + "EditorFontName = Consolas" + Chr(13) + "EditorFontSize = 10"
;
; ******************************************************************************
;
; The following files will be embedded into the executable, and then extracted
; to the preferences folder on the first launch of the executable.
;
IncludePath  "..\Images\"
;
DataSection
  BNext: 
  IncludeBinary "Next.jpg"
  BNextEnd: 
  
  BPrevious: 
  IncludeBinary "Previous.jpg"
  BPreviousEnd: 
  
  BNoNext: 
  IncludeBinary "NoNext.jpg"
  BNoNextEnd: 
  
  BNoPrevious: 
  IncludeBinary "NoPrevious.jpg"
  BNoPreviousEnd: 
    
  BRefresh: 
  IncludeBinary "Refresh.jpg"
  BRefreshEnd: 
  ;
  IncludePath  "..\Catalogs\"
  ;
  BCatalogFR: 
  IncludeBinary "Francais\PBBrowser.catalog"
  BCatalogFREnd: 
  
  BIntroFR: 
  IncludeBinary "Francais\IntroPBBrowser.rtf"
  BIntroFREnd: 
  
  BCatalogEN: 
  IncludeBinary "English\PBBrowser.catalog"
  BCatalogENEnd: 
  
  BIntroEN: 
  IncludeBinary "English\IntroPBBrowser.rtf"
  BIntroENEnd: 
  
  BCatalogIT: 
  IncludeBinary "Italiano\PBBrowser.catalog"
  BCatalogITEnd: 
  
  BIntroIT: 
  IncludeBinary "Italiano\IntroPBBrowser.rtf"
  BIntroITEnd:
  
  BCatalogRU: 
  IncludeBinary "Russian\PBBrowser.catalog"
  BCatalogRUEnd: 
  
  BIntroRU: 
  IncludeBinary "Russian\IntroPBBrowser.rtf"
  BIntroRUEnd: 
  ;
  IncludePath  "..\"
  ;
  PBFunction: 
  IncludeBinary "PBFunctionList.Data"
  PBFunctionEnd: 
    
  APIFunction: 
  IncludeBinary "APIFunctionListing.txt"
  APIFunctionEnd: 
EndDataSection
;
; ******************************************************************************
;
; The list of Basic keywords is not long. Rather than keeping it in a file,
; it is stored here as data.
;
DataSection
	PBBasicKeyWords: 
  Data.s "And", "Array", "As", "Align", "Break", "CallDebugger", "Case", "CompilerCase", "CompilerDefault", "CompilerElse", "CompilerElseIf", "CompilerEndIf", "CompilerEndSelect"
  Data.s "CompilerError", "CompilerIf", "CompilerSelect", "Continue", "Data", "DataSection", "EndDataSection", "Debug", "DebugLevel", "Declare", "DeclareC"
  Data.s "DeclareCDLL", "DeclareDLL", "Default", "Define", "Dim", "DisableASM", "DisableDebugger", "DisableExplicit", "DeclareModule", "Else", "ElseIf", "EnableASM"
  Data.s "EnableDebugger", "EnableExplicit", "End", "EndEnumeration", "EndIf", "EndImport", "EndInterface", "EndMacro", "EndProcedure", "EndDeclareModule", "EndModule"
  Data.s "EndSelect", "EndStructure", "EndStructureUnion", "EndWith", "Enumeration", "Extends", "FakeReturn", "For", "Next", "ForEach"
  Data.s "ForEver", "Global", "Gosub", "Goto", "If", "Import", "ImportC", "IncludeBinary", "IncludeFile", "IncludePath", "Interface", "List", "Macro", "Map", "MacroExpandedCount"
  Data.s "Module", "NewList", "Not", "Or", "Procedure", "ProcedureC", "ProcedureCDLL", "ProcedureDLL", "ProcedureReturn", "Protected", "Prototype"
  Data.s "PrototypeC", "Read", "ReDim", "Repeat", "Until", "Restore", "Return", "Runtime", "Select", "Shared", "Static", "Step", "Structure", "StructureUnion"
  Data.s "Swap", "To", "Wend", "While", "With", "XIncludeFile", "XOr", "UseModule", "UnuseModule", "UndefineMacro"
  Data.s "EndPBBasicKeyWords"
EndDataSection
; IDE Options = PureBasic 6.10 LTS (Windows - x86)
; CursorPosition = 80
; FirstLine = 60
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb