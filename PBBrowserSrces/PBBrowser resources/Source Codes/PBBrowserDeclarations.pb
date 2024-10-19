; *******************************************************************************
;
;  Fichier des déclarations de constantes et de variables globales de PBBrowser
;
; *******************************************************************************
XIncludeFile "Zapman libraries\ZapmanCommon.pb"
XIncludeFile "Zapman libraries\RichEdit_Library.pb"
XIncludeFile "Zapman libraries\TOM_Functions.pb"
XIncludeFile "Zapman libraries\Pipe.pb"
XIncludeFile "Zapman libraries\FastStringFunctions.pb"
XIncludeFile "Zapman libraries\ExpressionEvaluator.pb"
XIncludeFile "Zapman libraries\ChooseLanguage.pb"
;
Global MyAppDataFolder$ = GetSystemFolder(#CSIDL_COMMON_APPDATA) + "\" + #NomProg + "\" ; Adresse du dossier des données.
Global PBBrowserPrefile$ = MyAppDataFolder$ + #NomProg + ".prefs" ; Adresse du fichier des préférences.
Global BOM ; Contiendra le BOM trouvé à la lecture des fichiers.
Global PBBListOfFiles$ = "" ; Contiendra la liste des fichiers liés au fichier principal.
Global PBBFirstLaunch ; Signale le premier lancement de l'application.
Global OutOfElementContent.String\s = ""  ; Contiendra une compilation du code hors procédure, pour l'ensemble des fichiers.
Global PureBasicProgAdr$, PBUnderCursor$, FicPrincipalPB$, FicActualPB$, TempFile$ ; Contiendra les arguments envoyés par l'éditeur de PureBasic.
                                                                               ;
#OutOfElementsName = "OutOfElements"
;
; *********************************************************
;            Définition des types d'éléments
;
Enumeration PBBTypes
  #PBBProcedure       ; Élement de type 'Procedure'.
  #PBBStructure       ; idem pour les structures
  #PBBMacro           ; idem pour les macros
  #PBBEnumeration     ; idem pour les énumérations
  #PBBInterface       ; idem pour les interfaces
  #PBBLabel           ; idem pour les labels
  #PBBConstante       ; idem pour les constantes
  #PBBVariable        ; idem pour les variables
  #EndEnumPBBElementTypes ; Signale la fin des éléments trouvés dans les fichiers
  #PBBNativeFunction  ; idem pour les fonctions natives de PureBasic
  #PBBBasicKeyword    ; idem pour les mots-clé du Basic (While, Wend, If, Then, etc.)
  #EndEnumPBBTypes    ; Signale la fin de la liste
EndEnumeration
; On attribue un nom pour chaque type :
Global Dim PBBTypeNames$(#EndEnumPBBTypes - 1)
TypesList$ = "Procedure,Structure,Macro,Enumeration,Interface,Label,Constante,Variable,,PBNativeFunction,PBKeyword"
; On enregistre ces noms dans le tableau PBBTypeNames$()
For ct = 0 To (#EndEnumPBBTypes - 1)
  PBBTypeNames$(ct) = StringField(TypesList$, ct + 1, ",")
Next
;
; L'élément #PBBVariable sera le seul qui sera éclaté en plusieurs
; catégories. Ces catégories figurent ci-dessous :
#PBBVariableSpecies$ = "Global,Shared"
;
; ************************************************************
;                   Tableaux d'éléments
;
; Les tableaux qui vont suivre permettront de compléter les diverses étapes d'exploration,
; afin d'obtenir les listes des éléments (procédures, structures, énumérations, etc.) qui
; figurent dans les fichiers examinés.
; Ces tableaux sont nombreux, mais comme leur dimension est très petite, ils n'occuperont
; qu'une place raisonnablement limitée en mémoire.
;
; ListOfAllElements$() contiendra la liste complète des éléments
; qui ont été trouvés dans la liste des fichiers.
Global Dim ListOfAllElements$(#EndEnumPBBElementTypes - 1)
;
; ListOfAllElementsNbr() contiendra le nombre d'éléments
; trouvés pour un type d'élément donné.
Global Dim ListOfAllElementsNbr(#EndEnumPBBElementTypes - 1)
;
; ListOfUsedElements$ contiendra la liste des éléments (procédures, structures, etc.)
; réellements utilisés par le fichier principal.
; Il peut sembler idiot de travailler sur une liste séparée de ListOfAllElements$(),
; d'autant que ListOfAllElements$() est également mise à jour afin de signaler les
; élements utilisés ou non (voir le champs #EL_Parent, figurant dans la structure
; de cette liste). Cependant, le fait de gérer une liste séparée permet de gagner
; beaucoup de temps pendant la phase d'exploration (jusqu'à 50% du temps est économisé),
; car cette liste permet d'éviter une recherche afin de déterminer si un élément
; est utilisé ou non. De plus, elle permet de savoir quels sont les derniers
; éléments a avoir été notés comme utilisés, puisqu'ils sont toujours placés
; en fin de liste. Cette information permet d'optimiser considérablement la
; procédure 'SetListOfUsedElements()' qui se contente d'examiner la fin de cette
; liste en boucle jusqu'à être sûr que tous les éléments utilisés ont été identifiés.
Global Dim ListOfUsedElements$(#EndEnumPBBElementTypes - 1)
;
; Les tableaux ListCompletionAll et ListCompletionUsed indiqueront si les
; tableaux qui précèdent ont déjà été complétés par l'exploration des fichiers.
Global Dim ListCompletionAll(#EndEnumPBBElementTypes - 1)
Global Dim ListCompletionUsed(#EndEnumPBBElementTypes - 1)
;
; Pour ListOfAllElements$, le tableau suivant contiendra l'étape à laquelle
; la procédure de complétion s'est interrompue lors du travail en tâche de fond.
Global Dim ListCompletionStage(#EndEnumPBBElementTypes - 1)
;
; Pour ListOfAllElements$, le tableau suivant comportera le nom de fichier
; en cours d'examen lors du travail en tâche de fond.
; Pour ListOfUsedElements$, il contiendra le dernier élément de référence
; (celui dont on examine les codes, pour déterminer si l'élément principal
; est utilisé ou non).
Global Dim ListCompletionReference$(#EndEnumPBBElementTypes - 1)
;
; Le tableau suivant permettra de mesurer la completion de
; ListOfUsedElements$() :
Global Dim ListCompletionStepsCompleted(#EndEnumPBBElementTypes - 1)
;
; **************************************************************
;         Valeurs pour l'état de complétion des listes
;
; L'énumération qui va suivre sera utilisée pour renseigner les tableaux
; 'ListCompletionAll()' et 'ListCompletionUsed()'
Enumeration ListCompletionState ; État courant de la complétion pour une liste donnée.
  #ListCompletion_Undone
  #ListCompletion_ExamAll
  #ListCompletion_RebootFromTabRef
  #ListCompletion_Pending
  #ListCompletion_DoublePending
  #ListCompletion_StageCompleted
  #ListCompletion_Done
  #ListCompletion_Printed
  #ListCompletion_DoNot
EndEnumeration
;
Enumeration CompletionState ; État général de la complétion
  #Completion_Completed
  #Completion_Uncomplete
  #Completion_Error
EndEnumeration
;
; *************************************************************
;           Organisation des données dans les listes
;        'ListOfAllElements$()' et 'ListOfUsedElements$
; (L'organisation est strictement la même dans les deux listes).
;
; Les listes sont organisées comme une suite de valeurs séparées par
; le caractère tabulation (Chr(9)).
; Chaque ligne se termine par un retour chariot (Chr(13)).
; La somme des lignes forme une chaîne de caractère qui représente
; la totalité des éléments trouvés pour un type d'élément particulier.
; Par exemple, ListOfAllElements$(#PBBProcedure) contiendra la liste
; complète des procédures trouvées dans les fichiers.
; Chaque ligne de ListOfAllElements$() est organisée comme suit :
; 'Nom de l'élément (de la structure, de la procédure, etc.)' + Chr(9)
;       + 'Adresse du fichier contenant l'élément' + Chr(9)
;       + 'Déclaration complète de l'élement (avec arguments, s'il y en a)'  + Chr(9)
;       + 'N° de ligne du début dans le fichier'  + Chr(9)
;       + 'N° de ligne de la fin dans le fichier' + Chr(9)
;       etc.
;
Enumeration StructureOfElementsLists
  ; Les procédures de PBBrowser sont conçues de sorte que
  ; l'ordre et le nombre des composants de cette énumération
  ; puisse être modifiés. À deux exceptions près : 
  ; - #EL_ElementNameLCase doit être en première position avec la valeur '1',
  ; - #EL_EndOfLine doit figurer en dernière position.
  #EL_ElementNameLCase = 1
  #EL_ElementName
  #EL_FileName     ; Nom du fichier ou l'élément figure (où il est défini).
  #EL_CompleteElementDeclaration ; Déclaration de l'élément avec ses paramètres,
                                 ; s'il y en a.
  #EL_StartingLine ; Numéro de ligne ou l'élément est déclaré dans le code
  #EL_EndingLine   ; Numéro de ligne de la fin de l'élément ('EndProcedure', par exemple).
  #EL_StartingPos  ; Position exacte du début de l'élément dans le code.
  #EL_EndingPos    ; Position de la fin de l'élément.
  #EL_Parent       ; Ce champs contiendra une références à l'élément qui a permis de décider
                   ; que l'élément courant est utilisé. Il sera renseigné comme suit :
                   ;     ElementName/ElementType
  #EL_DeclarationDetails     ; Utilisé seulement pour les constantes et les procédures.
  #EL_Value           ; Utilisé seulement pour les constantes.
  #EL_VariableSpecies ; Utilisé seulement pour les variables.
  #EL_Comment     ; Contiendra les commentaires pour les éléments qui n'ont pas de code
                  ; (Constante et Label, par exemple).
  #EL_EndOfLine
EndEnumeration
;
; *********************************************************
; Définition des différents panneaux des 'PannelGadgets'
;
Enumeration PBBPanels ; Numéros attribués au panneaux
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
; On attribue un nom pour chaque panneau :
Global Dim PBBPanelNames$(#EndEnumPBBPanels - 1)
TypesList$ = "Files,Lists,Details,FoundIn,Procedures,Structures,Macros,Enumerations,Interfaces,Labels,Constantes,Variables"
; On enregistre ces noms dans le tableau PBBPanelNames$()
For ct = 0 To (#EndEnumPBBPanels - 1)
  PBBPanelNames$(ct) = StringField(TypesList$, ct + 1, ",")
Next
; Les tableaux suivants permettront de fournir des détails sur chaque panneau
Global Dim NoPBBPanel(#EndEnumPBBPanels - 1)
Global Dim NoGadgetPBBPanel(#EndEnumPBBPanels - 1)
Global Dim NoREGadgetOfPBBPanel(#EndEnumPBBPanels - 1)
; Ce qui suit va permettre d'établir un pont entre les pages ListPBBPanel
; et les types d'éléments, en indiquant quels types d'éléments sont affichés
; par chaque panneau.
Global Dim TypeElementOfPBBPanel(#EndEnumPBBPanels - 1)
TypeElementOfPBBPanel(#ProcPBBPanel) = #PBBProcedure
TypeElementOfPBBPanel(#StructurePBBPanel) = #PBBStructure
TypeElementOfPBBPanel(#MacroPBBPanel) = #PBBMacro
TypeElementOfPBBPanel(#EnumPBBPanel) = #PBBEnumeration
TypeElementOfPBBPanel(#InterfacePBBPanel) = #PBBInterface
TypeElementOfPBBPanel(#LabelPBBPanel) = #PBBLabel
TypeElementOfPBBPanel(#ConstantePBBPanel) = #PBBConstante
TypeElementOfPBBPanel(#VariablePBBPanel) = #PBBVariable
;
; Les énumérations suivantes seront utilisées pour gérer
; les tâches qui s'exécutent en tâches de fond.
;
Enumeration PBB_PriorityMode ; Gestion de l'affichage pour les tâches de fonds.
  #WorkInBackGround
  #ShowCompletionWindow
  #FinishCompletionNow
EndEnumeration
;
Enumeration PBB_BackgroundTasksState ; État d'exécution des tâches de fonds.
  #BackgroundTasksUncompleted
  #BackgroundTasksHavePriority
  #BackgroundTasksCompleted
  #BackgroundTasksMustRestart
EndEnumeration
;
; Pour les raccourcis clavier :
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
; Pour la mémorisation des numéros des gadgets de la fenêtre principale :
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
Structure LastSearchDetails
  ElementName$
  ElementType.i
  TypeName$
EndStructure
;
Global GPBBGadgets.PBBGadgets;
;
#PBBRTFMarker$ = "PBBMark_:" ; Sera utilisé pour identifier les images insérées dans les gadgets RichEdit.
#PBBLeftArrowMarker$ = "Previous"
#PBBRightArrowMarker$ = "Next"
#PBBLeftArrow$ = "←"
#PBBRightArrow$ = "→"

;
Enumeration ValuesForExactValueSearch
  #NoSearchOrManualSearch
  #DoProgrammedSearch
  #DoProgSearchAndPrintResult
EndEnumeration
;
Enumeration ValuesForCheckFileUpdating
  #FileIsUpdated
  #FileWasUpToDate
  #FileMustBeUpdated
  #FileDoesntExist
  #FileHasBeenDeleted
EndEnumeration
;
Global PBFunctionList$       ; Contiendra la liste des fonctions natives.
Global PBFunctionListLCase$  ; Idem en version LCase()
;
Global PBBDefaultCursor = LoadCursor_(0, #IDC_ARROW) ; Le curseur flèche standard qui sera affiché au-dessus des gadgets RichEdit.
Global GreyColor = RGB(120, 120, 120)
Global pcolorNotUsed = RGB(130, 140, 140) ; Gris pâle pour les procédures qui figurent dans ListOfAllElements$() mais pas dans ListOfUsedElementsForPrint$()
Global DarkRedColor = RGB(180, 0, 0)
Global SetValueColor = RGB(150, 90, 0)
Global PBBTitleColor = RGB(150, 150, 150)
Global PBBTitleFont = LoadFont(100, "Segoe UI", 9, #PB_Font_Bold | #PB_Font_HighQuality)
Global ProgressLegendFont = LoadFont(101, "Segoe UI", 7, #PB_Font_HighQuality)
;
; Ce sui suit est une copie de sécurité du fichier des préférences de PureBasic,
; pour les valeurs indispensables à PBBrowser. Cela permet de le faire fonctionner
; sur une machine où PureBasic n'est pas installé.
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
; Les fichiers suivants seront intégrés dans l'exécutable, puis extraits dans
; le dossier des préférences au premier lancement de l'exécutable.
;
DataSection
  BNext: 
  IncludeBinary "..\Images\Next.jpg"
  BNextEnd: 
  
  BPrevious: 
  IncludeBinary "..\Images\Previous.jpg"
  BPreviousEnd: 
  
  BNoNext: 
  IncludeBinary "..\Images\NoNext.jpg"
  BNoNextEnd: 
  
  BNoPrevious: 
  IncludeBinary "..\Images\NoPrevious.jpg"
  BNoPreviousEnd: 
    
  BRefresh: 
  IncludeBinary "..\Images\Refresh.jpg"
  BRefreshEnd: 
  
  BCatalogFR: 
  IncludeBinary "..\Catalogs\Français\PBBrowser.catalog"
  BCatalogFREnd: 
  
  BIntroFR: 
  IncludeBinary "..\Catalogs\Français\IntroPBBrowser.rtf"
  BIntroFREnd: 
  
  BCatalogEN: 
  IncludeBinary "..\Catalogs\English\PBBrowser.catalog"
  BCatalogENEnd: 
  
  BIntroEN: 
  IncludeBinary "..\Catalogs\English\IntroPBBrowser.rtf"
  BIntroENEnd: 
  
  BCatalogIT: 
  IncludeBinary "..\Catalogs\Italiano\PBBrowser.catalog"
  BCatalogITEnd: 
  
  BIntroIT: 
  IncludeBinary "..\Catalogs\Italiano\IntroPBBrowser.rtf"
  BIntroITEnd: 
  
  PBFunction: 
  IncludeBinary "..\PBFunctionList.Data"
  PBFunctionEnd: 
    
  APIFunction: 
  IncludeBinary "..\APIFunctionListing.txt"
  APIFunctionEnd: 
EndDataSection
;
;
; La liste des mots-clés du Basic n'est pas longue. Plutôt que de la conserver
; dans un fichier, elle est stockée ici sous forme de data.
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
; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; CursorPosition = 325
; FirstLine = 313
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb