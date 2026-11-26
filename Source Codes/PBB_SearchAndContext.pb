; **********************************************************************************
;
;                         PBBrowser_SearchAndContext.pb
;
;        Contextual AST Analysis, Layout Formatting & Homonym Disambiguation
;                                 Nov. 20/2026
;
; **********************************************************************************
;
; ROLE & ARCHITECTURE:
; Contextual resolution and search orchestration engine of PBBrowser.
; Evaluates SearchDetails compatibility, inspects code lines via a local AST to
; classify occurrences and format UI layout markers (BuildLayout), deduces scope
; declarations upstream (GetInfosFromContext), iterates over project matches, and
; manages homonym collection and scope disambiguation (GetInfosFromData).
;
; FILE ORGANIZATION:
;   • SearchDetails Model & Compatibility Rules - Initializes, links, compares, and upgrades SearchDetails.
;   • Local AST Tokenization & BuildLayout      - Tokenizes code lines to deduce roles and UI layout strings.
;   • Contextual Scope Deduction                - Scans local and global ranges to resolve declarations.
;   • Project-Wide Sequence Match Iterator      - Locates valid code/expression matches across sequences.
;   • Homonym Collection & Scope Helpers        - Populates SearchMenuChoices() and synthetic 'All...' scopes.
;   • Interactive Scope Selection & Lookup      - Synchronizes Search/LookUp UI and disambiguates homonyms.
;
; **********************************************************************************
;
;-                   SearchDetails model & compatibility rules
;
; **********************************************************************************
;
Procedure PrepareSearchDetailsElement(*SearchVar.SearchDetails, Expression$)
  ;
  ; Role: Resets '*SearchVar' and initializes its raw and purified name fields from 'Expression$'.
  ; Flags composite expressions via 'NotAnElement' and 'SearchPurifiedOnly'.
  ;
  Protected NotAnElement
  ;
  ClearStructure(*SearchVar, SearchDetails)
  *SearchVar\ElementName$         = Expression$
  *SearchVar\ElementPurifiedName$ = PurifyForSearch(Expression$, @NotAnElement)
  If NotAnElement
    *SearchVar\SearchPurifiedOnly = #True
    *SearchVar\NotAnElement       = #True
  EndIf
  *SearchVar\HasParentheses       = -1
EndProcedure
;
Procedure.s ComputeExpressionDetails(*LastSearch.SearchDetails)
  ;
  ; Role: Builds and returns a formatted description of the element's type, species,
  ; enclosing parent container, and module (e.g., "Procedure in Module TextMod") from '*LastSearch'.
  ;
  Protected Type$, Articulation$
  ;
  If *LastSearch\SearchPurifiedOnly
    If *LastSearch\NotAnElement = #False
      Type$ = GetTextFromCatalogPB("All")
    EndIf
  Else
    If *LastSearch\ElementSpecies$
      Type$ = *LastSearch\ElementSpecies$
      AddStringElement(Type$, *LastSearch\ElementSubSpecies$, " ")
    ElseIf *LastSearch\TypeOfElement > #PBBNoType And *LastSearch\TypeOfElement < #PBBNotAnElement
      Type$ = GetTextFromCatalogPB(PBBTypeNames$(*LastSearch\TypeOfElement))
    EndIf
    ;
    If *LastSearch\DontCheckPosition And *LastSearch\ParentTypeAndName$ <> ""
      ; For #PBBField, qualify the container kind ("Field of Structure" / "Field of Interface") before "in all...":
      If *LastSearch\TypeOfElement = #PBBField
        AddStringElement(Type$, Mid(*LastSearch\ParentTypeAndName$, 4, Len(*LastSearch\ParentTypeAndName$) - 4), #Space$ + GetTextFromCatalogPB("of") + #Space$)
      EndIf
      AddStringElement(Type$, GetTextFromCatalogPB(*LastSearch\ParentTypeAndName$), #Space$ + GetTextFromCatalogPB("in") + #Space$)
    Else
      If *LastSearch\TypeOfElement = #PBBField
        Articulation$ = GetTextFromCatalogPB("of")
      Else
        Articulation$ = GetTextFromCatalogPB("in")
      EndIf
      ;
      If *LastSearch\ParentTypeAndName$ <> "" And
         (*LastSearch\TypeOfElement = #PBBField Or *LastSearch\TypeOfElement = #PBBVariable Or
          *LastSearch\TypeOfElement = #PBBConstant Or *LastSearch\TypeOfElement = #PBBImported)
        AddStringElement(Type$, *LastSearch\ParentTypeAndName$, #Space$ + Articulation$ + #Space$)
      EndIf
      ;
      ; Append enclosing module name when the element belongs to a Module or DeclareModule
      If *LastSearch\ModuleName$ <> "" And *LastSearch\TypeOfElement <> #PBBModule
        AddStringElement(Type$, PBBTypeNames$(#PBBModule) + #Space$ + *LastSearch\ModuleName$, #Space$ + GetTextFromCatalogPB("in") + #Space$)
      EndIf
    EndIf
  EndIf
  ;
  If *LastSearch\InDeadZone
    AddStringElement(Type$, GetTextFromCatalogPB("inactiveCompilerZone"), " ")
  EndIf
  ProcedureReturn Type$
EndProcedure
;
Procedure FillSearchDetailsFromElement(*SrcElement.Elements, *SearchDetails.SearchDetails, *CED.CED_Structure = 0, ParamToken$ = "")
  ;
  ; Role: Populates *SearchDetails from an AST element record, a secondary CED declaration,
  ; or a formal parameter token (when ParamToken$ is provided on a Procedure/Macro *SrcElement).
  ;
  Protected ContainerID, HasPar
  Protected DotType.String
  ;
  ClearStructure(*SearchDetails, SearchDetails)
  ;
  If ParamToken$ <> ""
    ; Unpack a formal parameter belonging to *SrcElement (#PBBProcedure or #PBBMacro), preserving original type casing
    *SearchDetails\ElementPurifiedName$ = PurifyElementName(ParamToken$, @HasPar, @DotType)
    If *SrcElement\Type = #PBBProcedure And DotType\s = "" And Right(*SearchDetails\ElementPurifiedName$, 1) <> "$"
      DotType\s = "i"
      If HasPar
        ParamToken$ = Left(ParamToken$, Len(ParamToken$) - 2) + ".i()"
      Else
        ParamToken$ + ".i"
      EndIf
    EndIf
    *SearchDetails\ElementName$         = ParamToken$
    *SearchDetails\VarDotType$          = DotType\s
    *SearchDetails\TypeOfElement        = #PBBVariable
    *SearchDetails\ElementSpecies$      = "Param"
    *SearchDetails\ElementSubSpecies$   = "Var"
    *SearchDetails\HasParentheses       = HasPar
    *SearchDetails\ElementFile$         = *SrcElement\FileName
    *SearchDetails\ElementPosition      = *SrcElement\StartingPos
    *SearchDetails\ParentTypeAndName$   = PBBTypeNames$(*SrcElement\Type) + #Space$ + *SrcElement\Name
    *SearchDetails\StartingContextLimit = *SrcElement\StartingPos
    *SearchDetails\EndingContextLimit   = *SrcElement\EndingPos
  Else
    ; Standard AST element or secondary CED record (populate FixedID strictly in List Mode)
    If ExplorationState = #ExplorationState_Completed And ListSize(ElementList())
      *SearchDetails\FixedID           = *SrcElement\FixedID
    EndIf
    *SearchDetails\ElementName$         = *SrcElement\Name
    *SearchDetails\ElementPurifiedName$ = *SrcElement\PurifiedName
    *SearchDetails\TypeOfElement        = *SrcElement\Type
    *SearchDetails\ElementSpecies$      = *SrcElement\VariableSpecie
    *SearchDetails\ElementSubSpecies$   = *SrcElement\VariableSubSpecie
    *SearchDetails\VarDotType$          = StringField(*SrcElement\Name, 2, ".")
    *SearchDetails\VarDotType$          = StringField(*SearchDetails\VarDotType$, 1, "(")
    *SearchDetails\HasParentheses       = *SrcElement\HasParentheses
    If *CED
      *SearchDetails\ElementFile$       = *CED\Filename
      *SearchDetails\ElementPosition    = *CED\StartingPos
      ContainerID                       = *CED\DirectContainer
    Else
      *SearchDetails\ElementFile$       = *SrcElement\FileName
      *SearchDetails\ElementPosition    = *SrcElement\StartingPos
      ContainerID                       = *SrcElement\LocalVarContainer
    EndIf
    PushListPosition(ElementList())
    If ContainerID > 0
      If SelectPBBElement(ContainerID)
        ; Use the untranslated PureBasic keyword from PBBTypeNames$() so Colorize_REContent()
        ; and ClickOverREGadget() consistently recognize the container keyword across all languages.
        *SearchDetails\ParentTypeAndName$ = PBBTypeNames$(ElementList()\Type) + #Space$ + ElementList()\Name
        ; Context boundary limits apply strictly to local procedural variables (#PBBVariable).
        ; Structure and interface fields (#PBBField) retain global-like zero boundaries.
        If *SrcElement\Type = #PBBVariable
          *SearchDetails\StartingContextLimit = ElementList()\StartingPos
          *SearchDetails\EndingContextLimit   = ElementList()\EndingPos
        EndIf
      EndIf
    ElseIf *SrcElement\Type = #PBBVariable And *SrcElement\VariableSpecie = "Shared"
      *SearchDetails\ParentTypeAndName$ = GetTextFromCatalogPB(#OutOfElementsName)
    EndIf
    PopListPosition(ElementList())
  EndIf
  ;
  *SearchDetails\ModuleName$ = GetEnclosingModuleName(*SrcElement\EnclosingModuleID)
EndProcedure
;
Procedure ResolveSearchDetailsElement(*Search.SearchDetails)
  ;
  ; Role: Resolves a SearchDetails specification to its concrete FixedID in ElementList()
  ; using FindElementInList() (or O(1) SelectPBBElement() when *Search\FixedID is already known).
  ; Translates *Search\ModuleName$and *Search\ParentTypeAndName$ into their corresponding container IDs,
  ; caches the resolved ID in *Search\FixedID when DontCheckPosition = #False (or on a unique 'Shared' variable),
  ; and leaves ElementList() positioned on the matched element.
  ; Output: Returns the resolved FixedID (> 0), or 0 if no single element matches.
  ;
  Protected ModuleContainer = -1, MacroProcID = 0, TargetContainerID = 0
  Protected ParentPrefix$, ParentName_LCase$, CandidateIDs$, IDStr$
  Protected ctType, IDIdx, MatchedID
  ;
  If *Search = 0 Or *Search\ElementPurifiedName$ = "" Or *Search\SearchPurifiedOnly Or
     (*Search\DontCheckPosition And
      (*Search\TypeOfElement <> #PBBVariable Or *Search\ElementSpecies$ <> "Shared" Or *Search\ParentTypeAndName$ = "AllModules"))
    ProcedureReturn 0
  EndIf
  ;
  ; 1. Fast-path O(1) resolution when FixedID is already known
  If *Search\FixedID > 0 And SelectPBBElement(*Search\FixedID) And ElementList()\StartingPos > 0
    ProcedureReturn *Search\FixedID
  EndIf
  ;
  ; 2. Resolve enclosing module namespace ID when ModuleName$ is specified
  If *Search\ModuleName$ <> ""
    ModuleContainer = GetModuleID(*Search\ModuleName$)
  ElseIf *Search\DontCheckPosition
    ; Multi-procedure 'Shared' selection without an explicit module: ensure at most one 'Shared' element matches across modules
    If FindMapElement(ElementMapIndex(), *Search\ElementPurifiedName$) = 0
      ProcedureReturn 0
    EndIf
    CandidateIDs$ = ElementMapIndex()
    Repeat
      IDIdx + 1
      IDStr$= StringField(CandidateIDs$, IDIdx, ",")
      If IDStr$ = "" : Break : EndIf
      If SelectPBBElement(Val(IDStr$)) And ElementList()\StartingPos > 0 And
         ElementList()\Type = #PBBVariable And ElementList()\VariableSpecie = "Shared" And
         (*Search\HasParentheses = -1 Or ElementList()\HasParentheses = *Search\HasParentheses)
        If MatchedID > 0
          ProcedureReturn 0
        EndIf
        MatchedID = ElementList()\FixedID
      EndIf
    ForEver
    If MatchedID > 0 And SelectPBBElement(MatchedID)
      *Search\FixedID = MatchedID
    EndIf
    ProcedureReturn MatchedID
  EndIf
  ;
  ; 3. Resolve parent container ID (Procedure, Macro, Structure, Interface, Enumeration, or OutOfElements)
  If *Search\ParentTypeAndName$ <> "" And *Search\DontCheckPosition = #False
    If *Search\ParentTypeAndName$ = GetTextFromCatalogPB(#OutOfElementsName)
      MacroProcID = -1
    Else
      ParentPrefix$= StringField(*Search\ParentTypeAndName$, 1, " ")
      ParentName_LCase$= PurifyElementName(Mid(*Search\ParentTypeAndName$, Len(ParentPrefix$) + 2))
      If ParentName_LCase$ <> ""
        For ctType = #PBBProcedure To #PBBImport
          If LCase(PBBTypeNames$(ctType)) = LCase(ParentPrefix$)
            If ctType = #PBBProcedure Or ctType = #PBBMacro
              MacroProcID = FindElementInList(ParentName_LCase$, ctType, 0, "", 0, -1, "", 0, ModuleContainer)
            Else
              TargetContainerID = FindElementInList(ParentName_LCase$, ctType, 0, "", 0, -1, "", 0, ModuleContainer)
            EndIf
            Break
          EndIf
        Next
      EndIf
    EndIf
  EndIf
  ;
  ; 4. Query FindElementInList() and cache the resulting FixedID
  MatchedID = FindElementInList(*Search\ElementPurifiedName$, *Search\TypeOfElement, 0, *Search\ElementFile$, MacroProcID, *Search\HasParentheses,
                                LCase(*Search\VarDotType$), TargetContainerID, ModuleContainer,
                                *Search\ElementSpecies$, *Search\ElementSubSpecies$)
  If MatchedID > 0
    *Search\FixedID = MatchedID
  EndIf
  ;
  ProcedureReturn MatchedID
EndProcedure
;
; ----------------------------------------------------------------------------------
;            ---- Compatibility evaluation & trait inheritance ----
; ----------------------------------------------------------------------------------
;
Procedure AreSearchDetailsCompatible(*Candidate.SearchDetails, *Filter.SearchDetails)
  ;
  ; Role: Evaluates semantic, morphological, and geographical compatibility between a candidate
  ; expression and an active filter. Rejects collisions across different structures early.
  ;
  ; 0. Fast-path O(1) FixedID divergence check (valid only when both possess unique resolved IDs)
  If *Filter\FixedID > 0 And *Candidate\FixedID > 0 And *Filter\DontCheckPosition = #False And *Candidate\DontCheckPosition = #False
    If *Filter\FixedID <> *Candidate\FixedID
      ProcedureReturn #False
    EndIf
  EndIf
  ;
  ; 1. Strict fundamental identity and typing checks
  If *Filter\ElementPurifiedName$ <> "" And *Candidate\ElementPurifiedName$ <> "" And Not CompareExpressionsWithJoker(*Filter\ElementPurifiedName$, *Candidate\ElementPurifiedName$)
    ProcedureReturn #False
  EndIf
  ;
  ; Prevent raw text searches ("All") from merging with semantic AST searches in history
  If *Filter\SearchPurifiedOnly <> *Candidate\SearchPurifiedOnly
    ProcedureReturn #False
  EndIf
  ;
  If *Filter\SearchPurifiedOnly
    ProcedureReturn #True
  EndIf
  ;
  If *Filter\TypeOfElement <> 0 And *Candidate\TypeOfElement <> 0 And *Filter\TypeOfElement <> *Candidate\TypeOfElement
    ProcedureReturn #False
  EndIf
  ;
  ; Enforce module namespace isolation once both records are anchored to a code position, module, or global PB symbol
  If (*Filter\DontCheckPosition = #False Or (*Filter\ElementSpecies$ = "Shared" And *Filter\ModuleName$ <> "")) And
     (*Candidate\DontCheckPosition = #False Or (*Candidate\ElementSpecies$ = "Shared" And *Candidate\ModuleName$ <> "")) And
     (*Filter\ElementPosition > 0 Or *Filter\ModuleName$ <> "" Or
      *Filter\TypeOfElement = #PBBNativeFunction Or *Filter\TypeOfElement = #PBBBasicKeyword) And
     (*Candidate\ElementPosition > 0 Or *Candidate\ModuleName$ <> "" Or
      *Candidate\TypeOfElement = #PBBNativeFunction Or *Candidate\TypeOfElement = #PBBBasicKeyword)
    If LCase(*Filter\ModuleName$) <> LCase(*Candidate\ModuleName$) And
       Not ((*Filter\TypeOfElement = #PBBField Or *Candidate\TypeOfElement = #PBBField) And
            ((*Filter\ParentTypeAndName$ = "" And *Filter\ModuleName$ = "") Or
             (*Candidate\ParentTypeAndName$ = "" And *Candidate\ModuleName$ = "")))
      ProcedureReturn #False
    EndIf
  EndIf
  ;
  ; Parent container check:
  ; - When both records share the same DontCheckPosition state: enforces exact Structure/Interface match
  ;   for #PBBField, exact "All<Container>s" match when both have DontCheckPosition = #True, and exact
  ;   Enumeration match when both #PBBConstant records reside inside an Enumeration block.
  ; - When *Filter\DontCheckPosition = #True and *Candidate\DontCheckPosition = #False:
  ;   ensures the candidate's parent container type matches the target "All<ContainerType>s" category.
  If *Filter\ParentTypeAndName$ <> "" And *Candidate\ParentTypeAndName$ <> ""
    If *Filter\DontCheckPosition = *Candidate\DontCheckPosition
      If *Filter\TypeOfElement = #PBBField Or *Filter\DontCheckPosition
        If LCase(*Filter\ParentTypeAndName$) <> LCase(*Candidate\ParentTypeAndName$)
          ProcedureReturn #False
        EndIf
      ElseIf *Filter\TypeOfElement = #PBBConstant
        Protected EnumPrefix$ = LCase(PBBTypeNames$(#PBBEnumeration)) + #Space$
        If Left(LCase(*Filter\ParentTypeAndName$), Len(EnumPrefix$)) = EnumPrefix$ And
           Left(LCase(*Candidate\ParentTypeAndName$), Len(EnumPrefix$)) = EnumPrefix$ And
           LCase(*Filter\ParentTypeAndName$) <> LCase(*Candidate\ParentTypeAndName$)
          ProcedureReturn #False
        EndIf
      EndIf
    ElseIf *Filter\DontCheckPosition And *Filter\ParentTypeAndName$ <> "AllModules"
      Protected ExpectedContainer$ = LCase(Mid(*Filter\ParentTypeAndName$, 4, Len(*Filter\ParentTypeAndName$) - 4)) + #Space$
      If (*Filter\TypeOfElement = #PBBField Or *Filter\TypeOfElement = #PBBVariable) And
         *Candidate\ParentTypeAndName$ <> GetTextFromCatalogPB(#OutOfElementsName) And
         Left(LCase(*Candidate\ParentTypeAndName$), Len(ExpectedContainer$)) <> ExpectedContainer$
        ProcedureReturn #False
      EndIf
    EndIf
  EndIf
  ;
  ; Macro Structural Immunity: Macros operate as raw text replacements before compilation.
  ; Morphological mismatches (parentheses or explicit typing) are therefore bypassed.
  Protected IsMacroInvolved = Bool(*Filter\TypeOfElement = #PBBMacro Or *Candidate\TypeOfElement = #PBBMacro)
  ;
  If Not IsMacroInvolved
    If *Filter\HasParentheses <> -1 And *Candidate\HasParentheses <> -1 And *Filter\HasParentheses <> *Candidate\HasParentheses
      ProcedureReturn #False
    EndIf
    ;
    If *Filter\VarDotType$ <> "" And *Candidate\VarDotType$ <> "" And LCase(*Filter\VarDotType$) <> LCase(*Candidate\VarDotType$)
      ProcedureReturn #False
    EndIf
  EndIf
  ;
  ; 2. SubSpecies check with "Var" wildcard tolerance
  Protected FSub$ = *Filter\ElementSubSpecies$
  Protected CSub$ = *Candidate\ElementSubSpecies$
  Protected FSp$ = *Filter\ElementSpecies$
  Protected CSp$ = *Candidate\ElementSpecies$
  Protected IsFGlobal = Bool(FSp$ = "Global" Or FSp$ = "Threaded")
  Protected IsCGlobal = Bool(CSp$ = "Global" Or CSp$ = "Threaded")
  Protected IsFLocalShield = Bool(FSp$ = "Protected" Or FSp$ = "Static" Or FSp$ = "Param")
  Protected IsCLocalShield = Bool(CSp$ = "Protected" Or CSp$ = "Static" Or CSp$ = "Param")
  Protected FLim = *Filter\StartingContextLimit
  Protected CLim = *Candidate\StartingContextLimit
  ;
  If FSub$ <> "" And FSub$ <> "Var" And CSub$ <> "" And CSub$ <> "Var"
    If FSub$ <> CSub$
      ProcedureReturn #False
    EndIf
  EndIf
  ;
  ; 3. Phantom context and multi-procedure scope resolution
  If *Filter\ElementPosition = 0 Or *Candidate\ElementPosition = 0 Or *Filter\DontCheckPosition Or *Candidate\DontCheckPosition
    If *Filter\TypeOfElement = #PBBVariable And *Candidate\TypeOfElement = #PBBVariable
      ; Strictly reject collisions between local procedural variables and global declarations
      If (IsFLocalShield And IsCGlobal) Or (IsCLocalShield And IsFGlobal)
        ProcedureReturn #False
      EndIf
      ; When position checks are bypassed, accept identical species across distinct procedure scopes
      If *Filter\DontCheckPosition Or *Candidate\DontCheckPosition
        If FSp$ = "Shared" And CLim = 0
          ProcedureReturn Bool(CSp$ = "Shared" Or CSp$ = "Define" Or CSp$ = "Implicit" Or CSp$ = "")
        ElseIf CSp$ = "Shared" And FLim = 0
          ProcedureReturn Bool(FSp$ = "Shared" Or FSp$ = "Define" Or FSp$ = "Implicit" Or FSp$ = "")
        ElseIf FSp$ = "Define" And CLim > 0
          ProcedureReturn Bool(CSp$ = "Define" Or CSp$ = "Implicit" Or CSp$ = "")
        ElseIf CSp$ = "Define" And FLim > 0
          ProcedureReturn Bool(FSp$ = "Define" Or FSp$ = "Implicit" Or FSp$ = "")
        EndIf
        ProcedureReturn Bool(FSp$ = CSp$ Or FSp$ = "" Or CSp$ = "")
      EndIf
    EndIf
    ; Phantom contexts for non-variables bypass subsequent spatial position assertions
    If Not (*Filter\TypeOfElement = #PBBVariable Or *Candidate\TypeOfElement = #PBBVariable)
      ProcedureReturn #True
    EndIf
    If *Filter\ElementPosition = 0 Or *Candidate\ElementPosition = 0
      ProcedureReturn #True
    EndIf
  EndIf
  ;
  ; 4. Non-Variable Scope Bypass
  If Not (*Filter\TypeOfElement = #PBBVariable Or *Candidate\TypeOfElement = #PBBVariable)
    If *Filter\DontCheckPosition = #False
      If *Filter\TypeOfElement = #PBBMacro And *Filter\ElementPosition > 0 And *Candidate\ElementPosition > 0
        ProcedureReturn IsPositionBefore(*Filter\ElementFile$, *Filter\ElementPosition, *Candidate\ElementFile$, *Candidate\ElementPosition)
      ElseIf *Candidate\TypeOfElement = #PBBMacro And *Candidate\ElementPosition > 0 And *Filter\ElementPosition > 0
        ProcedureReturn IsPositionBefore(*Candidate\ElementFile$, *Candidate\ElementPosition, *Filter\ElementFile$, *Filter\ElementPosition)
      EndIf
    EndIf
    ProcedureReturn #True
  EndIf
  ;
  ; 5. Scope Evaluation Initializer
  Protected SameScope = Bool(FLim = CLim And *Filter\ElementFile$ = *Candidate\ElementFile$)
  ;
  If SameScope
    If CSp$ = "Implicit" And FSp$ = "Protected"
      If *Filter\ElementPosition > *Candidate\ElementPosition
        ProcedureReturn #False
      EndIf
    EndIf
    If CSp$ = "Protected" And FSp$ = "Implicit"
      If *Candidate\ElementPosition > *Filter\ElementPosition
        ProcedureReturn #False
      EndIf
    EndIf
    ProcedureReturn #True
  EndIf
  ;
  ; 6. The Local Procedural Shield ('Protected', 'Static', 'Param')
  If IsFLocalShield Or IsCLocalShield
    ProcedureReturn #False
  EndIf
  ;
  ; 7. Global / Threaded Chronology Rule
  If IsFGlobal Or IsCGlobal
    If IsFGlobal And IsCGlobal
      ProcedureReturn Bool(FSp$ = CSp$)
    EndIf
    If IsFGlobal
      ProcedureReturn IsPositionBefore(*Filter\ElementFile$, *Filter\ElementPosition, *Candidate\ElementFile$, *Candidate\ElementPosition)
    Else
      ProcedureReturn IsPositionBefore(*Candidate\ElementFile$, *Candidate\ElementPosition, *Filter\ElementFile$, *Filter\ElementPosition)
    EndIf
  EndIf
  ;
  ; 8. Distinct procedures with DontCheckPosition = #False are strictly isolated
  If FLim > 0 And CLim > 0
    ProcedureReturn #False
  EndIf
  ;
  ; 9. Main scope (0) vs Local scope (> 0) bridge for 'Shared' type resolution in GetInfosFromContext.
  ;    Disabled when the user explicitly selected the 'OutOfElements' entry in the popup menu.
  If (FLim = 0 And CLim > 0) Or (FLim > 0 And CLim = 0)
    If *Filter\ParentTypeAndName$ = GetTextFromCatalogPB(#OutOfElementsName)
      ProcedureReturn #False
    EndIf
    If (FLim > 0 And FSp$ = "Shared" And (CSp$ = "Define" Or CSp$ = "Implicit")) Or
       (CLim > 0 And CSp$ = "Shared" And (FSp$ = "Define" Or FSp$ = "Implicit"))
      ProcedureReturn #True
    EndIf
    ProcedureReturn #False
  EndIf
  ;
  ; 10. Both in Main scope across different files (FLim = 0 And CLim = 0)
  If FLim = 0 And CLim = 0
    ProcedureReturn #True
  EndIf
  ;
  ; 11. Undecidable fallback across different scopes
  If FSp$ = "Implicit" Or FSp$ = "" Or CSp$ = "Implicit" Or CSp$ = ""
    ProcedureReturn #True
  EndIf
  ;
  ProcedureReturn #False
  ;
EndProcedure
;
Procedure InheritStrongSearchTraits(*Source.SearchDetails, *Dest.SearchDetails)
  ;
  ; Role: Safely upgrades weak or missing traits in '*Dest' from '*Source' once compatibility is verified.
  ;
  If Not AreSearchDetailsCompatible(*Dest, *Source)
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowBuildLayout
        OPDReport("       NOT COMPATIBLE")
      EndIf
    CompilerEndIf
    ProcedureReturn #False
  EndIf
  ;
  ; 1. Inherit fundamental identity
  If *Source\ElementPurifiedName$ <> ""
    *Dest\ElementPurifiedName$ = *Source\ElementPurifiedName$
  EndIf
  If Len(*Source\ElementName$) > Len(*Dest\ElementName$) Or FindString(*Dest\ElementName$, #PBBJoker$)
    *Dest\ElementName$ = *Source\ElementName$
    *Dest\ElementPurifiedName$ = *Source\ElementPurifiedName$
  EndIf
  ;
  ; 2. Inherit fundamental type (e.g., #PBBProcedure, #PBBVariable) and List-Mode FixedID
  If *Dest\TypeOfElement = 0 And *Source\TypeOfElement <> 0
    *Dest\TypeOfElement = *Source\TypeOfElement
  EndIf
  If ExplorationState = #ExplorationState_Completed And ListSize(ElementList()) And
     *Dest\DontCheckPosition = #False And *Source\DontCheckPosition = #False And
     *Source\FixedID And *Dest\FixedID = 0
    *Dest\FixedID = *Source\FixedID
  EndIf
  ;
  ; 3. Inherit explicit typing (e.g., .s, .MyStructure)
  If *Dest\VarDotType$ = "" And *Source\VarDotType$ <> "" And Right(StringField(*Dest\ElementName$, 1, "("), 1) <> "$"
    *Dest\VarDotType$ = *Source\VarDotType$
  EndIf
  ;
  ; 4. Inherit physical parentheses signature
  If *Dest\ElementPurifiedName$ = "" And *Dest\ElementPosition = 0
    *Dest\HasParentheses = -1
  EndIf
  ;
  If *Dest\HasParentheses = -1 And *Source\HasParentheses <> -1
    *Dest\HasParentheses = *Source\HasParentheses
  EndIf
  ;
  ; 5. SubSpecies Hierarchy ('Var' or empty is upgraded by anything else)
  If (*Dest\ElementSubSpecies$ = "" Or *Dest\ElementSubSpecies$ = "Var") And *Source\ElementSubSpecies$ <> ""
    *Dest\ElementSubSpecies$ = *Source\ElementSubSpecies$
  EndIf
  ;
  ; 6. Module Namespace & Structural Parent Inheritance
  Protected CopyCoordinates = #False
  ;
  If *Dest\ParentTypeAndName$ <> "AllModules" And *Source\ModuleName$ <> ""
    If *Dest\ModuleName$ = "" Or (*Source\FixedID > 0 And LCase(*Dest\ModuleName$) = LCase(*Source\ModuleName$))
      *Dest\ModuleName$ = *Source\ModuleName$
    EndIf
  EndIf
  ;
  If *Dest\ParentTypeAndName$ = "" And *Source\ParentTypeAndName$ <> ""
    ; Do not overwrite main-scope variable context (StartingContextLimit = 0 at a known position) with a procedure parent
    If *Dest\TypeOfElement <> #PBBVariable Or *Dest\ElementPosition = 0 Or *Dest\StartingContextLimit > 0
      *Dest\ParentTypeAndName$ = *Source\ParentTypeAndName$
    EndIf
    If *Dest\TypeOfElement = #PBBField
      CopyCoordinates = #True
    EndIf
  EndIf
  ;
  ; 7. Species Hierarchy Upgrade & Coordinate Transfer
  Protected DestSp$ = *Dest\ElementSpecies$
  Protected SrcSp$  = *Source\ElementSpecies$
  ;
  ; If destination has no physical position assigned yet
  If *Dest\ElementPosition = 0 And *Dest\ElementFile$ = ""
    CopyCoordinates = #True
  EndIf
  ;
  If *Dest\TypeOfElement = #PBBVariable
    ; Standard variable scope hierarchy upgrade
    If SrcSp$ <> "" And SrcSp$ <> DestSp$
      If DestSp$ = "" Or DestSp$ = "Implicit"
        *Dest\ElementSpecies$ = *Source\ElementSpecies$
        CopyCoordinates = #True
      ElseIf (DestSp$ = "Shared" Or DestSp$ = "Define") And (SrcSp$ = "Global" Or SrcSp$ = "Threaded")
        *Dest\ElementSpecies$ = *Source\ElementSpecies$
        CopyCoordinates = #True
      ElseIf DestSp$ = "Define" And SrcSp$ = "Shared"
        *Dest\ElementSpecies$ = "Shared"
      EndIf
    EndIf
  ElseIf *Dest\TypeOfElement = #PBBField
    ; Pure species assignment for fields without ambiguity tagging
    If *Dest\ElementSpecies$ = "" And *Source\ElementSpecies$ <> ""
      *Dest\ElementSpecies$ = *Source\ElementSpecies$
      CopyCoordinates = #True
    EndIf
  EndIf
  ;
  ; Transfer physical coordinates only when *Source holds a concrete single-scope location
  If CopyCoordinates And *Source\ElementPosition > 0 And *Source\DontCheckPosition = #False
    *Dest\ElementPosition      = *Source\ElementPosition
    *Dest\ElementFile$         = *Source\ElementFile$
    *Dest\StartingContextLimit = *Source\StartingContextLimit
    *Dest\EndingContextLimit   = *Source\EndingContextLimit
    If ExplorationState = #ExplorationState_Completed And ListSize(ElementList()) And *Source\FixedID > 0
      *Dest\FixedID            = *Source\FixedID
    EndIf
  EndIf
  If *Dest\DontCheckPosition Or ExplorationState <> #ExplorationState_Completed Or ListSize(ElementList()) = 0
    *Dest\FixedID = 0
  EndIf
  ;
  ; 8. The Absolute Safety Net
  ; Global/Threaded variables and non-variable elements (such as Constants) MUST NEVER
  ; retain procedural context limits or a Procedure/Macro parent from a local usage site.
  If (*Dest\TypeOfElement = #PBBVariable And (*Dest\ElementSpecies$ = "Global" Or *Dest\ElementSpecies$ = "Threaded")) Or
     (*Dest\TypeOfElement <> 0 And *Dest\TypeOfElement <> #PBBVariable)
    *Dest\StartingContextLimit = 0
    *Dest\EndingContextLimit   = 0
    If *Dest\DontCheckPosition = #False
      If *Dest\TypeOfElement = #PBBConstant
        If Left(*Dest\ParentTypeAndName$, Len(PBBTypeNames$(#PBBEnumeration)) + 1) <> PBBTypeNames$(#PBBEnumeration) + #Space$
          *Dest\ParentTypeAndName$ = ""
        EndIf
      ElseIf *Dest\TypeOfElement <> #PBBField And *Dest\TypeOfElement <> #PBBImported
        *Dest\ParentTypeAndName$ = ""
      EndIf
    EndIf
  EndIf
  ;
  ; Rebuild full name if additional type or parenthesis traits were acquired
  If *Dest\VarDotType$ And FindString(*Dest\ElementName$, ".") = 0 And Right(StringField(*Dest\ElementName$, 1, "("), 1) <> "$"
    *Dest\ElementName$ = StringField(*Dest\ElementName$, 1, "(") + "." + *Dest\VarDotType$
  EndIf
  If *Dest\HasParentheses = 1 And FindString(*Dest\ElementName$, "(") = 0
    *Dest\ElementName$ + "()"
  EndIf
  ;
  ProcedureReturn #True
  ;
EndProcedure
;
Procedure IsSearchDetailsCompleted(*Candidate.SearchDetails, AcceptImplicit)
  ;
  ; Role: Determines whether '*Candidate' holds complete semantic metadata (including #PBBNotAnElement rejections).
  ;
  ; Reject incomplete candidates or unresolved wildcard patterns still containing #PBBJoker$:
  If *Candidate\ElementName$ = "" Or *Candidate\ElementPurifiedName$ = "" Or *Candidate\TypeOfElement = 0 Or
     FindString(*Candidate\ElementPurifiedName$, #PBBJoker$)
    ProcedureReturn #False
  EndIf
  ;
  If *Candidate\SearchPurifiedOnly Or *Candidate\TypeOfElement = #PBBMacro Or *Candidate\TypeOfElement = #PBBNotAnElement
    ProcedureReturn #True
  EndIf
  ;
  If *Candidate\HasParentheses = -1
    ProcedureReturn #False
  EndIf
  ;
  ; A structure or interface field is complete once its parent container is resolved:
  If *Candidate\TypeOfElement = #PBBField
    ProcedureReturn Bool(*Candidate\ParentTypeAndName$ <> "")
  ElseIf *Candidate\TypeOfElement <> #PBBVariable
    ProcedureReturn #True
  EndIf
  ;
  ; Variables require explicit declaration species and subspecies validation:
  If *Candidate\ElementSpecies$ = "" Or *Candidate\ElementSubSpecies$ = ""
    ProcedureReturn #False
  ElseIf *Candidate\ElementSpecies$ = "Implicit"
    ProcedureReturn Bool(AcceptImplicit <> 0)
  EndIf
  ;
  ProcedureReturn #True
EndProcedure
;
;
; **********************************************************************************
;
;-                    Local AST tokenization & BuildLayout
;
; **********************************************************************************
;
Enumeration ScopeShadowingModes
  #PBB_Shadow_Reset
  #PBB_Shadow_Register
  #PBB_Shadow_Query
EndEnumeration
;
Procedure.s ManageProcMacroScopeShadowing(Mode, ScopeStart, VarName$, HasParentheses = #False, ActiveSpecies$ = "")
  ;
  ; ROLE:
  ; Tracks and queries the active variable species within procedural scopes.
  ; Leverages chronological compiler sequence traversal: an identifier declared
  ; as 'Protected', 'Static', 'Param', or 'Shared' propagates its species to all
  ; subsequent usages within the same procedure scope.
  ;
  Static CachedScopeStart
  Static CachedVarName$
  Static CachedSpecies$
  ;
  ; Append "()" when the token has parentheses to distinguish scalar variables from lists, arrays, maps, or procedures
  If HasParentheses
    VarName$ + "()"
  EndIf
  ;
  Select Mode
    Case #PBB_Shadow_Reset
      CachedScopeStart = 0
      CachedVarName$   = ""
      CachedSpecies$   = ""
      ProcedureReturn ""
      ;
    Case #PBB_Shadow_Register
      ; Normalize identifier name to ensure case-insensitive scope tracking
      CachedScopeStart = ScopeStart
      CachedVarName$   = LCase(VarName$)
      CachedSpecies$   = ActiveSpecies$
      ProcedureReturn CachedSpecies$
      ;
    Case #PBB_Shadow_Query
      ; If outside a procedure, or entering a new procedure where no local declaration
      ; was encountered yet, the identifier natively defaults to outer scope
      If ScopeStart = 0 Or ScopeStart <> CachedScopeStart Or LCase(VarName$) <> CachedVarName$
        ProcedureReturn ""
      EndIf
      ProcedureReturn CachedSpecies$
      ;
  EndSelect
  ;
  ProcedureReturn ""
EndProcedure
;
Structure PBB_MacroEventStruct
  MacroName$    ; Lowercase identifier of the declared or undefined macro.
  FileName$     ; Source file hosting the Macro or UndefineMacro statement.
  PosInCode.i   ; 1-based character offset of the statement inside FileName$.
  VirtualPos.i  ; Absolute 1-based coordinate in the Virtual Whole Project File.
  HasParen.b    ; #True if the macro declaration includes an opening parenthesis '('.
  IsUndefine.b  ; #True for 'UndefineMacro', #False for 'Macro'.
EndStructure
;
Procedure CheckIfMacroOverridesNative(MacroName$, FileName$, PosInCode, *OutMacroFile.String = 0, HasParentheses = -1)
  ;
  ; Role: Determines whether 'MacroName$' is governed by an active user Macro at 'PosInCode' in 'FileName$'.
  ;       Resolves in O(1) via ElementMapIndex()/MacroVirtualPos in List Mode (unless PBB_HasUndefineMacro = #True),
  ;       or via an incremental virtual-position timeline cache ('MacroEvents()') in standalone Text Mode.
  ;
  Shared PBB_CompilationSequence(), PBB_FileListing()
  Static NewList MacroEvents.PBB_MacroEventStruct()
  Static ScannedUpToVirtualPos.i, ScannedUpToFile$, ScannedUpToPos.i, *ScannedUpToSeqElement
  ;
  Protected FIPS.FIPSPosStruct
  Protected ActiveMacroPos, ActiveMacroHasParen, BestVirtualPos, IDIndex, HasPrecedingMacro
  Protected TargetVirtualPos, EventVirtualPos, IsUndefine
  Protected ActiveMacroFile$, TargetMacro_LCase$, CandidateIDs$, IDStr$, ExtractedName$
  Protected *Char.Character, *SavedSeq, *SavedFile
  ;
  If *OutMacroFile : *OutMacroFile\s = "" : EndIf
  If MacroName$ = "" Or FileName$ = "" Or PosInCode <= 0
    ClearList(MacroEvents())
    ScannedUpToVirtualPos = 0 : ScannedUpToFile$ = "" : ScannedUpToPos = 0 : *ScannedUpToSeqElement = 0
    ProcedureReturn 0
  EndIf
  TargetMacro_LCase$ = LCase(MacroName$)
  ;
  ; 1. O(1) List Mode resolution: returns directly when PBB_HasUndefineMacro = #False
  If ExplorationState = #ExplorationState_Completed And ListSize(ElementList())
    If FindMapElement(ElementMapIndex(), TargetMacro_LCase$) = 0 : ProcedureReturn 0 : EndIf
    CandidateIDs$ = ElementMapIndex()
    If PBB_HasUndefineMacro = #False
      TargetVirtualPos = ComputePositionInVirtualWholeProjectFile(FileName$, PosInCode)
    EndIf
    Repeat
      IDIndex + 1
      IDStr$ = StringField(CandidateIDs$, IDIndex, ",")
      If IDStr$ = "" : Break : EndIf
      If SelectPBBElement(Val(IDStr$)) And ElementList()\Type = #PBBMacro And ElementList()\StartingPos > 0
        If HasParentheses = 0 And ElementList()\HasParentheses = 1 : Continue : EndIf
        If PBB_HasUndefineMacro = #False
          If ElementList()\MacroVirtualPos <= TargetVirtualPos And ElementList()\MacroVirtualPos >= BestVirtualPos
            BestVirtualPos   = ElementList()\MacroVirtualPos
            ActiveMacroPos   = ElementList()\StartingPos
            ActiveMacroFile$ = ElementList()\FileName
          EndIf
        ElseIf IsPositionBefore(ElementList()\FileName, ElementList()\StartingPos, FileName$, PosInCode)
          HasPrecedingMacro = #True
          Break
        EndIf
      EndIf
    ForEver
    If PBB_HasUndefineMacro = #False
      If ActiveMacroPos > 0 And *OutMacroFile : *OutMacroFile\s = ActiveMacroFile$ : EndIf
      ProcedureReturn ActiveMacroPos
    ElseIf HasPrecedingMacro = #False
      ProcedureReturn 0
    EndIf
  EndIf
  ;
  ; 2. Text Mode (or project containing UndefineMacro): incremental virtual-position timeline scan
  TargetVirtualPos = ComputePositionInVirtualWholeProjectFile(FileName$, PosInCode)
  If TargetVirtualPos > ScannedUpToVirtualPos
    *SavedSeq      = @PBB_CompilationSequence()
    *SavedFile     = @PBB_FileListing()
    FIPS\ScanMode  = #FIPSMode_Global
    FIPS\EndLimit  = PosInCode
    If ScannedUpToVirtualPos > 0 And *ScannedUpToSeqElement <> 0
      FIPS\IsLocalScan       = #False
      FIPS\FIPSContinue      = #True
      FIPS\FileName$         = ScannedUpToFile$
      FIPS\PosInCode         = ScannedUpToPos - 1
      FIPS\CurrentSeqElement = *ScannedUpToSeqElement
      FIPS\TargetFileName$   = FileName$
    Else
      FIPS\FileName$ = FileName$
    EndIf
    While FindInWholeProjectSequences("macro,undefinemacro", @FIPS, #CheckAll)
      IsUndefine     = Bool(FastMid(FIPS\SearchString_LCase\s, FIPS\PosInCode, 13) = "undefinemacro")
      *Char          = PosToChar(@FIPS\SearchString_LCase\s, FIPS\PosInCode + 5 + 8 * IsUndefine)
      ExtractedName$ = ExtractOrSkipNextToken(@FIPS\SearchString_LCase\s, @*Char)
      If ExtractedName$ <> ""
        EventVirtualPos = ComputePositionInVirtualWholeProjectFile(FIPS\FileName$, FIPS\PosInCode)
        If ListSize(MacroEvents()) = 0 Or MacroEvents()\VirtualPos < EventVirtualPos
          LastElement(MacroEvents())
          AddElement(MacroEvents())
          MacroEvents()\MacroName$ = ExtractedName$
          MacroEvents()\FileName$  = FIPS\FileName$
          MacroEvents()\PosInCode  = FIPS\PosInCode
          MacroEvents()\VirtualPos = EventVirtualPos
          MacroEvents()\HasParen   = Bool(*Char\c = '(')
          MacroEvents()\IsUndefine = IsUndefine
        EndIf
      EndIf
      FIPS\FIPSContinue = #True
    Wend
    ScannedUpToVirtualPos  = TargetVirtualPos
    ScannedUpToFile$       = FileName$
    ScannedUpToPos         = PosInCode
    *ScannedUpToSeqElement = 0
    ForEach PBB_CompilationSequence()
      If PBB_CompilationSequence()\FileName$ = FileName$ And
         PosInCode >= PBB_CompilationSequence()\LocalStartPos And PosInCode <= PBB_CompilationSequence()\LocalEndPos
        *ScannedUpToSeqElement = @PBB_CompilationSequence()
        Break
      EndIf
    Next
    If *SavedSeq  : ChangeCurrentElement(PBB_CompilationSequence(), *SavedSeq) : EndIf
    If *SavedFile : ChangeCurrentElement(PBB_FileListing(), *SavedFile)        : EndIf
  EndIf
  ;
  ; 3. Replay cached Macro / UndefineMacro events up to TargetVirtualPos
  ForEach MacroEvents()
    If MacroEvents()\VirtualPos > TargetVirtualPos : Break : EndIf
    If MacroEvents()\MacroName$ = TargetMacro_LCase$
      If MacroEvents()\IsUndefine
        If MacroEvents()\FileName$ = FileName$ And PosInCode >= MacroEvents()\PosInCode And
           PosInCode <= MacroEvents()\PosInCode + 18 + Len(TargetMacro_LCase$)
          Break
        EndIf
        ActiveMacroPos = 0 : ActiveMacroFile$ = "" : ActiveMacroHasParen = #False
      Else
        ActiveMacroPos      = MacroEvents()\PosInCode
        ActiveMacroFile$    = MacroEvents()\FileName$
        ActiveMacroHasParen = MacroEvents()\HasParen
      EndIf
    EndIf
  Next
  If ActiveMacroPos > 0 And HasParentheses = 0 And ActiveMacroHasParen
    ActiveMacroPos = 0 : ActiveMacroFile$ = ""
  EndIf
  If ActiveMacroPos > 0 And *OutMacroFile : *OutMacroFile\s = ActiveMacroFile$ : EndIf
  ProcedureReturn ActiveMacroPos
EndProcedure
;
Define NewList ContextFoundExpression.SearchDetails()
;
EnumerationBinary
  #PBB_Action_Layout      ; 1 -> Build and populate *LayoutDefinition
  #PBB_Action_Collect     ; 2 -> Update the ContextFoundExpression() list
  #PBB_Action_CollectAll  ; 4 -> Force context extraction even if data is incomplete or out of bounds
  #PBB_Action_Reset       ; 8 -> Clear caches and lists before processing
EndEnumeration
;
Procedure BuildLayout(*Filter.SearchDetails, FileName$, *SearchString.String, *SearchString_LCase.String,
                      EndLimit, PosInCode, *LayoutDefinition.string = 0, ActionFlags = #PBB_Action_Layout | #PBB_Action_Collect)
  ;
  ; ROLE & ARCHITECTURE:
  ; Core multifunctional parsing engine for PBBrowser. Analyzes a specific
  ; occurrence of an expression within the source code by tokenizing its
  ; local AST (Abstract Syntax Tree) and climbing the spatial hierarchy.
  ; Serves two primary, independent purposes controlled by 'ActionFlags':
  ;
  ; 1. UI LAYOUT GENERATION (#PBB_Action_Layout):
  ;    Translates the expression's logical role into standardized UI markers.
  ;    Builds a formatted string indicating the spatial scope (e.g., "Procedure X")
  ;    and the operational action (e.g., "& Set", "& Test", "Arg").
  ;
  ; 2. METADATA COLLECTION (#PBB_Action_Collect):
  ;    Extracts precise typological metadata (Species, SubSpecies, Type) and
  ;    physical coordinates (Position, Context Limits) to populate the global
  ;    'ContextFoundExpression()' dictionary. Delegates merging and trait
  ;    upgrades to 'InheritStrongSearchTraits()' for strict deduplication.
  ;
  ; SECONDARY BEHAVIORS:
  ; - #PBB_Action_CollectAll : Bypasses strict morphological filters (parentheses,
  ;                            types) and out-of-bounds checks to force extraction.
  ; - #PBB_Action_Reset      : Purges static caches and the shared dictionary
  ;                            before a brand new global search cycle.
  ;
  ; VARIABLE PREFIX CONVENTIONS (Architectural segregation):
  ; - 'LO_' : Layout Output. Variables exclusively dedicated to UI layout assembly.
  ; - 'CP_' : Context Parameters. Variables exclusively dedicated to metadata collection.
  ; - 'S_'  : Shared. Semantic properties and AST deductions required by both workflows.
  ;
  ; PARAMETERS:
  ; - *Filter             : Pointer to SearchDetails containing expected traits.
  ; - FileName$           : Full path of the currently parsed file.
  ; - *SearchString       : Pointer to the full PureBasic source code string.
  ; - *SearchString_LCase : Lowercase version of the source for fast matching.
  ; - EndLimit            : End of valid position in the source string.
  ; - PosInCode           : Absolute character position of the targeted expression.
  ; - *LayoutDefinition   : String pointer receiving the formatted UI text.
  ; - ActionFlags         : Bitmask directing the engine's execution path.
  ; ============================================================================
  ;
  Shared ContextFoundExpression()
  ;
  Static LO_LastLayoutPrefix$
  ;
  Protected ToSearch$, LenExp, TestChar, LO_Extended
  ;
  ; Initialize the static context and purge the shared memory list
  If ActionFlags & #PBB_Action_Reset
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowBuildLayout
        OPDReport("------ Reset in BuildLayout -------")
      EndIf
    CompilerEndIf
    LO_LastLayoutPrefix$ = ""
    ClearList(ContextFoundExpression())
    ManageProcMacroScopeShadowing(#PBB_Shadow_Reset, 0, "")
    ;
    If (ActionFlags & (#PBB_Action_Layout | #PBB_Action_Collect | #PBB_Action_CollectAll)) = 0
      ProcedureReturn #True
    EndIf
  EndIf
  ;
  If *LayoutDefinition : *LayoutDefinition\s = "" : EndIf
  ;
  If *SearchString = #Null Or *SearchString_LCase = #Null
    Alert("BuildLayout is called with a null *SearchString")
    ProcedureReturn #False
  EndIf
  ;
  If *Filter = 0 Or PosInCode = 0 Or *Filter\ElementName$ = ""
    Alert("BuildLayout is called with a null PosInCode or ElementName$")
    ProcedureReturn #False
  EndIf
  ;
  Protected PosInLine, RPExp, RPInCode, LineCtx
  Protected InComment, InQuote
  Protected LineOfCode$ = ""
  Protected ExpToken.LocalToken, *ExpToken.LocalToken
  ;
  ; Identify comment or string literal context first and enforce strict flag separation
  If *Filter\SearchInCom Or *Filter\SearchInQuotes Or (ActionFlags & #PBB_Action_CollectAll)
    LineCtx   = CheckCodeLineContext(*SearchString, PosInCode, #True)
    InComment = Bool(LineCtx = #CodeCtx_Comment)
    InQuote   = Bool(LineCtx = #CodeCtx_Quote)
    If InQuote And *Filter\NotAnElement = #False And
       (*Filter\TypeOfElement = #PBBImport Or (ActionFlags & #PBB_Action_CollectAll))
      TestChar = PeekC(PosToChar(@*SearchString\s, PosInCode))
      If TestChar = #DOUBLEQUOTE Or TestChar = '~'
        InQuote = #False
      EndIf
    EndIf
    If Not (ActionFlags & #PBB_Action_CollectAll)
      If (InComment And *Filter\SearchInCom = #False) Or (InQuote And *Filter\SearchInQuotes = #False)
        ProcedureReturn #False
      EndIf
    EndIf
  EndIf
  ;
  ; Out-of-context bounds check (applies strictly to active code occurrences, never to comments or quotes)
  If InComment = #False And InQuote = 0 And *Filter\SearchPurifiedOnly = #False And Not (ActionFlags & #PBB_Action_CollectAll)
    If ((*Filter\StartingContextLimit > 0 Or *Filter\EndingContextLimit > 0) And *Filter\ElementFile$ <> "" And FileName$ <> *Filter\ElementFile$) Or
       (*Filter\EndingContextLimit > 0 And PosInCode > *Filter\EndingContextLimit) Or
       (*Filter\StartingContextLimit > 0 And PosInCode < *Filter\StartingContextLimit)
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowBuildLayout
          OPDReport("OUT OF CONTEXT - Pos = " + Str(PosInCode) + " Limits : " + Str(*Filter\StartingContextLimit) + "/" + Str(*Filter\EndingContextLimit))
        EndIf
      CompilerEndIf
      ;
      ProcedureReturn #False
    EndIf
  EndIf
  ;
  If *Filter\NotAnElement = #False
    ToSearch$ = ReplaceString(*Filter\ElementPurifiedName$, #PBBJoker$, "")
    LenExp    = Len(ToSearch$)
    ;
    If FastMid(*SearchString_LCase\s, PosInCode, LenExp) <> ToSearch$
      Alert("BuildLayout is called with improper parameters." + #CR$ +
            "Given Expression: |" + ToSearch$ + "|" + #CR$ +
            "Expression at PosInCode: |" + Mid(*SearchString_LCase\s, PosInCode, LenExp) + "|")
      ProcedureReturn #False
    EndIf
    ;
    If *Filter\TypeOfElement <> #PBBField
      ; A structure field declared with an asterisk is subsequently used without an asterisk,
      ; so the following test is skipped for fields.
      RPExp = RecognizePointer(@ToSearch$, 1)
      If InComment = #False And InQuote = 0
        RPInCode = RecognizePointer(@*SearchString\s, PosInCode)
      Else
        ; Inside comments or string literals, treat '*Ident' as a pointer whenever '*' is not attached to a left operand
        If PeekC(@*SearchString\s + (PosInCode - 1) * SizeOf(Character)) = '*'
          TestChar = PosInCode - 1
        ElseIf PosInCode > 1 And PeekC(@*SearchString\s + (PosInCode - 2) * SizeOf(Character)) = '*'
          TestChar = PosInCode - 2
        Else
          TestChar = -1
        EndIf
        If TestChar = 0
          RPInCode = #True
        ElseIf TestChar > 0
          TestChar = PeekC(@*SearchString\s + (TestChar - 1) * SizeOf(Character))
          RPInCode = Bool(IsPBSeparator(TestChar) And TestChar <> ')' And TestChar <> ']' And TestChar <> '.')
        EndIf
      EndIf
      ; Do not confuse MyVar with *MyVar:
      If RPInCode <> RPExp
        ProcedureReturn #False
      EndIf
    EndIf
    ;
    ; If the original search was performed with the #PBBSearchPrefix flag,
    ; expand ToSearch$ to facilitate subsequent tokenization.
    If IsPBSeparator(Asc(Right(ToSearch$, 1))) = #False
      Repeat
        TestChar = Asc(FastMid(*SearchString_LCase\s, PosInCode + LenExp, 1))
        If IsPBSeparator(TestChar)
          Break
        Else
          ; The search was probably performed with #PBBSearchPrefix:
          LenExp + 1
          LO_Extended = #True
        EndIf
      ForEver
    EndIf
    ;
    If IsPBSeparator(Asc(Left(ToSearch$, 1))) = #False
      Repeat
        TestChar = Asc(FastMid(*SearchString_LCase\s, PosInCode - 1, 1))
        If IsPBSeparator(TestChar)
          Break
        Else
          ; The search was probably performed with #PBBSearchPrefix:
          PosInCode - 1
          LenExp + 1
          LO_Extended = #True
        EndIf
      ForEver
    EndIf
    ;
    ToSearch$ = FastMid(*SearchString\s, PosInCode, LenExp)
  EndIf
  ;
  If *Filter\NotAnElement Or InComment Or InQuote
    ; Complex expressions, comments, and quoted strings skip AST line tokenization
    PosInLine = 1
    *ExpToken = ExpToken
    If *Filter\NotAnElement
      *ExpToken\Value$ = *Filter\ElementName$
    Else
      *ExpToken\Value$ = ToSearch$
    EndIf
  Else
    LineOfCode$ = ExtractSectionAroundExpression(*SearchString, ToSearch$, PosInCode, 0, #CheckLeftAndRight | #CheckOneKeyWordOnly, EndLimit, @PosInLine)
    ;
    Protected LenLOC = Len(LineOfCode$)
    If LenLOC = 0 Or PosInLine = 0 Or PosInLine > (LenLOC - Len(ToSearch$) + 1) : ProcedureReturn #False : EndIf
    ;
    ; -------------------------------------------------------------------------
    ; 1. AST POINTERS RESOLUTION & SCANNING
    ; Role: Extracts the isolated line of code encompassing the target expression
    ;       and tokenizes it to build a local Abstract Syntax Tree (AST).
    ;       Scans sequentially to anchor functional pointers (*ExpToken) and
    ;       navigates backward to identify the caller function or preceding
    ;       assignments. Essential for both UI layout and context deduction.
    ; -------------------------------------------------------------------------
    Protected NewList TokenList.LocalToken()
    TokenizeCodeLine(@LineOfCode$, TokenList(), PosInLine)
    ;
    Protected *NextToken.LocalToken
    Protected *PrevToken.LocalToken
    Protected *PrevWord.LocalToken
    Protected *FirstWord.LocalToken
    Protected *LastToken.LocalToken
    ;
    Protected PosAfter = PosInLine + LenExp
    Protected S_FirstWord_LCase$, CP_DeterminedVarType$
    Protected S_FirstWord$, S_PrevWord$, S_PrevWord_LCase$
    Protected NextTokenValue$, PrevTokenValue$, OperatorTestSymbol$
    Protected S_IsFirstWordEqualPrevWord
    Protected AfterEqual, LockAfterEqual, BeforeEqualAssignation, LockBeforeEqual
    Protected InParenthesis, S_HasParentheses, S_HasTrailingBackslash, S_IsDotType, IsTest, ParenDepth
    Protected ProcOrMacroDeclaration
    Protected LO_PrevIsClosingParenthesis, LO_NextIsParenthesis
    Protected S_CallerFunction$, S_ArgPosition = 1
    Protected S_ExplicitModuleName$
    ;
    ; --- A. Token scanning
    ForEach TokenList()
      ; Normalize declaration keywords on unqualified tokens, or retain 'Mod::Ident' on module-qualified statements
      If *FirstWord = 0 And TokenList()\Type = #TokenType_Word
        *FirstWord = @TokenList()
        If *FirstWord\ModulePrefix$ <> ""
          S_FirstWord$ = *FirstWord\ModulePrefix$ + "::" + *FirstWord\Value$
        ElseIf (Left(*FirstWord\Value_LCase$, 9) = "procedure" And *FirstWord\Value_LCase$ <> "procedurereturn")
          S_FirstWord$ = "Procedure"
          ProcOrMacroDeclaration = #True
        ElseIf Left(*FirstWord\Value_LCase$, 7) = "declare" And *FirstWord\Value_LCase$ <> "declaremodule"
          S_FirstWord$ = "Declare"
          ProcOrMacroDeclaration = #True
        ElseIf *FirstWord\Value_LCase$ = "macro"
          S_FirstWord$ = "Macro"
          ProcOrMacroDeclaration = #True
        ElseIf Left(*FirstWord\Value_LCase$, 11) = "enumeration"
          S_FirstWord$ = "Enumeration"
        ElseIf Left(*FirstWord\Value_LCase$, 6) = "import"
          S_FirstWord$ = "Import"
        Else
          S_FirstWord$ = *FirstWord\Value$
          ;
          If FindString(LCase(#PBTestOperators$), "," + *FirstWord\Value_LCase$ + ",")
            IsTest = #True
          EndIf
        EndIf
      EndIf
      ;
      ; Test PosInLine - 1 in case of presence of "*" as first char of the token
      If TokenList()\PosInLine = PosInLine Or TokenList()\PosInLine = PosInLine - 1 Or
         (TokenList()\Type = #TokenType_StringInQuotes And ; Check if ToSearch$ is inside a quoted string
          FindString(TokenList()\Value$, ToSearch$) And
          TokenList()\PosInLine < PosInLine And
          (TokenList()\PosInLine + Len(TokenList()\Value$)) > PosInLine)
        *ExpToken = @TokenList()
        InParenthesis = *ExpToken\Level
        *PrevToken = *LastToken
      Else
        If *ExpToken = 0 And TokenList()\Type = #TokenType_Word And (*LastToken = 0 Or *LastToken\Value$ <> ".")
          *PrevWord = @TokenList()
        ElseIf *NextToken = 0 And TokenList()\PosInLine >= PosAfter
          *NextToken = @TokenList()
        EndIf
        ;
        If TokenList()\Value$ = "," And *ExpToken <> 0 And InParenthesis = TokenList()\Level
          ; Prevent BeforeEqualAssignation if "=" appears after a comma in a parameter list:
          LockBeforeEqual = #True
        EndIf
        ;
        If TokenList()\Value$ = "=" And *ExpToken <> 0 And InParenthesis = TokenList()\Level
          If InParenthesis = 0 And IsTest = 0
            ; At level 0, it is an assignment if the line does not start with a test operator
            BeforeEqualAssignation = #True
          ElseIf ProcOrMacroDeclaration And LockBeforeEqual = #False
            ; Inside parentheses, it is an assignment only for declarations with default values
            BeforeEqualAssignation = #True
          EndIf
        EndIf
      EndIf
      ;
      *LastToken = @TokenList()
    Next ; Next TokenList()
    ;
    If *ExpToken = 0
      Protected ErrorMessages$ = "Tokenization error in 'BuildLayout':" + #CR$ +
                                 "Key = " + Mid(LineOfCode$, PosInLine, LenExp) + " : " + Str(PosInLine) + #CR$ +
                                 "Tokens:"
      ForEach TokenList()
        ErrorMessages$ + #CR$ + TokenList()\Value$ + " : " + Str(TokenList()\PosInLine)
      Next
      Alert(ErrorMessages$) 
      ProcedureReturn #False
    EndIf
    ;
    ; --- B. Data Completion
    ;
    S_ExplicitModuleName$ = *ExpToken\ModulePrefix$
    If *PrevToken : PrevTokenValue$ = *PrevToken\Value$ : EndIf
    If *NextToken : NextTokenValue$ = *NextToken\Value$ : EndIf
    ;
    If *PrevWord = *FirstWord Or *PrevWord = 0
      S_IsFirstWordEqualPrevWord = #True
      S_PrevWord$ = S_FirstWord$
    ElseIf *PrevWord\ModulePrefix$ <> ""
      S_PrevWord$ = *PrevWord\ModulePrefix$ + "::" + *PrevWord\Value$
    Else
      S_PrevWord$ = *PrevWord\Value$
    EndIf
    ;
    S_FirstWord_LCase$ = LCase(S_FirstWord$)
    S_PrevWord_LCase$  = LCase(S_PrevWord$)
    ;
    ; --- C. Context variables update
    ;
    If NextTokenValue$ = "."
      ChangeCurrentElement(TokenList(), *NextToken)
      If NextElement(TokenList()) And TokenList()\Type = #TokenType_Word
        CP_DeterminedVarType$ = TokenList()\Value$
      EndIf
    Else
      ChangeCurrentElement(TokenList(), *ExpToken)
    EndIf
    ;
    If PrevTokenValue$ <> "." And NextElement(TokenList()) And TokenList()\Value$ = "("
      ;
      S_HasParentheses = 1
      ;
      ParenDepth = 0
      ; Traverse nested parentheses to find the matching closing parenthesis
      Repeat
        If TokenList()\Value$ = "("
          ParenDepth + 1
        ElseIf TokenList()\Value$ = ")"
          ParenDepth - 1
        EndIf
      Until Not NextElement(TokenList()) Or ParenDepth = 0
      ;
    EndIf
    ;
    ; Skip complete structure field dereference chains (\Field1\Field2...)
    While @TokenList() And TokenList()\Type = #TokenType_Symbol And TokenList()\Value$ = "\"
      ; Record that the target token is dereferenced as a structured container
      S_HasTrailingBackslash = #True
      ; Advance past both the backslash and its target field identifier
      If NextElement(TokenList()) And TokenList()\Type = #TokenType_Word
        NextElement(TokenList())
      Else
        ; Stop advancing if the trailing dereference is incomplete or malformed
        Break
      EndIf
    Wend
    ;
    ; Capture any trailing comparison or assignment operator
    If @TokenList() And TokenList()\Type = #TokenType_Symbol
      OperatorTestSymbol$ = TokenList()\Value$
    EndIf
    ;
    LO_NextIsParenthesis        = Bool(NextTokenValue$ = "(")
    LO_PrevIsClosingParenthesis = Bool(PrevTokenValue$ = ")")
    S_IsDotType                 = Bool(PrevTokenValue$ = "." And NextTokenValue$ <> "::")
    ;
    ; --- D. Backward scan
    ;
    Protected HasLeadingBackslash = -1, HasLeadingDot
    Protected ParentVar$, MainParent$, S_ResolvedEntityModule$, *MainParent.LocalToken, WasAWord = #False
    Protected HadParentheses, SkipDepth, EntityColonPos, CurrentWord$
    ;
    If *PrevToken And *PrevToken\Value$ = "\"
      ; Need to resolve Parent and MainParent for field:
      HasLeadingBackslash = 0
    EndIf
    ;
    ChangeCurrentElement(TokenList(), *ExpToken)
    While PreviousElement(TokenList())
      If TokenList()\Level = InParenthesis And TokenList()\Value$ = ","
        S_ArgPosition + 1
      EndIf
      ;
      If TokenList()\Level < InParenthesis And S_CallerFunction$ = ""
        If TokenList()\Value$ = "(" And Not(ProcOrMacroDeclaration)
          If PreviousElement(TokenList())
            If TokenList()\Type = #TokenType_Word
              If FindString(LCase(#PBOperatorAsWord$), "," + TokenList()\Value_LCase$ + ",") = 0
                S_CallerFunction$ = TokenList()\Value_LCase$
              EndIf
            EndIf
          EndIf
        EndIf
      EndIf
      ;
      If TokenList()\Level = InParenthesis
        If TokenList()\Value$ = "=" And LockAfterEqual = #False
          AfterEqual = #True
        ElseIf TokenList()\Value$ = ","
          LockAfterEqual = #True
        EndIf
      EndIf
      ;
      ; Resolve Parent and MainParent for Field
      If HasLeadingBackslash <> -1
        If TokenList()\Value$ = "\"
          HasLeadingBackslash = #True
          WasAWord            = #False
          MainParent$         = ""
        ElseIf TokenList()\Value$ = ")"
          ; Skip backward over container index or argument lists (e.g., List(), Array(i), Map(k))
          SkipDepth = 1
          While PreviousElement(TokenList())
            If TokenList()\Value$ = ")"
              SkipDepth + 1
            ElseIf TokenList()\Value$ = "("
              SkipDepth - 1
              If SkipDepth = 0 : Break : EndIf
            EndIf
          Wend
          HadParentheses = #True
          WasAWord       = #False
        ElseIf TokenList()\Value$ = "."
          HasLeadingDot = #True
          ; The last "parent" attributed to ParentVar$ was a DotType.
          ; Remove it:
          While ParentVar$ And Right(ParentVar$, 1) <> "\"
            ParentVar$ = Left(ParentVar$, Len(ParentVar$) - 1)
          Wend
          If Right(ParentVar$, 1) = "\"
            ParentVar$ = Left(ParentVar$, Len(ParentVar$) - 1)
          EndIf
          WasAWord = #False
        ElseIf TokenList()\Type = #TokenType_Symbol
          ; Any symbol other than "." and "\" is not allowed in a "Var\Field1\Field2" sequence
          HasLeadingBackslash = - 1 ; <-- Interrupt examination
        ElseIf HasLeadingBackslash And TokenList()\Type = #TokenType_Word
          ; Two consecutive words are not allowed in a "Var\Field1\Field2" sequence
          If WasAWord
            HasLeadingBackslash = - 1 ; <-- Interrupt examination
          Else
            CurrentWord$ = TokenList()\Value$
            If TokenList()\ModulePrefix$ <> ""
              CurrentWord$ = TokenList()\ModulePrefix$ + "::" + CurrentWord$
            EndIf
            If HadParentheses
              CurrentWord$ + "()"
              HadParentheses = #False
            EndIf
            ;
            If ParentVar$ = "" : ParentVar$ = CurrentWord$ : Else : ParentVar$ = CurrentWord$ + "\" + ParentVar$ : EndIf
            If HasLeadingDot And MainParent$
              MainParent$ = CurrentWord$ + "." + MainParent$
            Else
              MainParent$ = CurrentWord$
            EndIf
            *MainParent   = @TokenList()
            HasLeadingDot = #False
            WasAWord      = #True
          EndIf
        EndIf
      EndIf
    Wend ; Wend PreviousElement(TokenList()) (Backward scan)
    ;
    ; --- E. Declarator prefix type inheritance
    ; When an expression lacks an explicit postfix type (e.g., 'MyVar'), inherit the
    ; common prefix type defined on the scope declarator (e.g., 'Define.s MyVar').
    ; Variables bearing the native '$' suffix already possess explicit string typing.
    If CP_DeterminedVarType$ = "" And AfterEqual = #False And InParenthesis = 0 And
       S_IsDotType = #False And Right(*ExpToken\Value$, 1) <> "$"
      If *FirstWord And FindMapElement(PB_VariableDeclarators(), S_FirstWord_LCase$)
        ChangeCurrentElement(TokenList(), *FirstWord)
        ; Scan forward from the declarator keyword to capture the '.' separator and its type identifier
        If NextElement(TokenList()) And TokenList()\Value$ = "."
          If NextElement(TokenList()) And TokenList()\Type = #TokenType_Word
            ; Ensure the target expression is the declared variable, not the qualifying type token itself
            If *ExpToken <> @TokenList()
              CP_DeterminedVarType$ = TokenList()\Value$
            EndIf
          EndIf
        EndIf
      EndIf
    EndIf ; EndIf CP_DeterminedVarType$ = ""
    ;
    ; -------------------------------------------------------------------------
    ; 2. INITIAL SIMPLE COMPARISON (Fast-Path Exit)
    ; Role: Saves CPU cycles by exiting immediately if the local candidate
    ;       morphologically mismatches the filter (e.g., wrong parentheses
    ;       or explicit typing), UNLESS #PBB_Action_CollectAll overrides it.
    ;       'UndefineMacro' statements and parameterless macros bypass the parenthesis filter.
    ; -------------------------------------------------------------------------
    ;
    If Not (ActionFlags & #PBB_Action_CollectAll) And Not (*Filter\SearchPurifiedOnly)
      If S_FirstWord_LCase$ <> "undefinemacro" And
         Not (S_CallerFunction$ = "defined" And S_ArgPosition = 1) And
         Not (*Filter\TypeOfElement = #PBBMacro And *Filter\HasParentheses <= 0) And
         *Filter\HasParentheses <> -1 And S_HasParentheses <> -1 And *Filter\HasParentheses <> S_HasParentheses
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowBuildLayout
            OPDReport("  BuildLayout early exit 1 for filter " + PBBTypeNames$(*Filter\TypeOfElement) + " " + *Filter\ElementName$)
            OPDReport("  --> Parentheses don't match : *Filter\HasParentheses = " + Str(*Filter\HasParentheses))
            OPDReport("                                S_HasParentheses       = " + Str(S_HasParentheses))
            OPDReport("      Examined LineOfCode$: " + LineOfCode$)
          EndIf
        CompilerEndIf
        ProcedureReturn #False
      EndIf
      If *Filter\VarDotType$ <> "" And CP_DeterminedVarType$ <> "" And LCase(*Filter\VarDotType$) <> LCase(CP_DeterminedVarType$)
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowBuildLayout
            OPDReport("  BuildLayout early exit 1 for filter " + PBBTypeNames$(*Filter\TypeOfElement) + " " + *Filter\ElementName$)
            OPDReport("  --> VarDotType$ doesn't match : *Filter\VarDotType$   = " + *Filter\VarDotType$)
            OPDReport("                                  CP_DeterminedVarType$ = " + CP_DeterminedVarType$)
          EndIf
        CompilerEndIf
        ProcedureReturn #False
      EndIf
    EndIf ; EndIf Fast-Path Exit (Section 2)
    ;
    ; -------------------------------------------------------------------------
    ; 3. ROLE, TYPE & SCOPE DEDUCTION
    ; Role: Evaluates operator proximity, preceding keywords, and spatial
    ;       boundaries to classify the expression. Determines hierarchical roles
    ;       (Parameter, Argument, Definition, Mutation) and invokes GetEncompassingLimitsAndName()
    ;       to lock the exact geographical outer scope of the element.
    ; -------------------------------------------------------------------------
    ;
    Protected S_Parameter, LO_Argument, S_Definition, LO_VarTest, S_VarSet
    Protected S_LocalTypeOfElement = 0, S_MacroDeclPos = 0, S_MacroDeclFile.String
    Protected ctType
    Protected S_EncompassingParent$, S_FieldStructureName$
    Protected S_LocalSpecies$, S_LocalSubSpecies$
    Protected IsInsideProcOrMacro
    Protected LO_LayoutPrefix$, S_Prefix$, EncompassingElementData.EEDStruct
    Protected S_StartingContextLimit, S_EndingContextLimit
    Protected PrevTok_LCase$, IsTopLevelDefinition
    ;
    If S_IsDotType = #False
      ; Types are only conceptual definitions.
      ; They cannot be "Arg", "Set", or "Test".
      If InParenthesis > 0
        ;
        If ProcOrMacroDeclaration And AfterEqual = 0
          S_Parameter = #True
          ; If the current token is 'List', 'Map', or 'Array' and is immediately followed
          ; by the parameter identifier (e.g., 'Procedure Foo(List Bar())'), it acts strictly
          ; as a declarator, not an argument or parameter.
          If *NextToken And *NextToken\Type = #TokenType_Word
            If *ExpToken\Value_LCase$ = "list" Or *ExpToken\Value_LCase$ = "map" Or *ExpToken\Value_LCase$ = "array"
              S_Parameter = #False
              S_LocalTypeOfElement = #PBBBasicKeyword
              S_HasParentheses = 0
            EndIf
          EndIf
        EndIf
        ;
        If S_CallerFunction$ <> ""
          ; An argument exclusively exists inside a function, macro, or array call.
          ; It naturally excludes isolated math statements like: A = (B + 1)
          LO_Argument = #True
          ; If the element is the 1st argument of a Poke function, it represents the target memory address.
          ; It inherently undergoes a memory mutation, qualifying it as a 'Set'.
          If S_ArgPosition = 1 And Left(S_CallerFunction$, 4) = "poke"
            S_VarSet = #True
          EndIf
        EndIf
      EndIf
      ;
      ; Evaluate the overarching context of the instruction.
      If IsTest Or S_CallerFunction$ = "bool"
        LO_VarTest = #True
      Else
        If BeforeEqualAssignation
          S_VarSet = #True
        ElseIf AfterEqual = 0 And *PrevToken = #Null
          If FindString(#PBOperatorAsSign$, NextTokenValue$)
            S_VarSet = #True
          EndIf
        EndIf
      EndIf
      ;
      ; Check if the Field's main parent is assigned:
      If *PrevToken And *PrevToken\Value$ = "\" And LO_VarTest = #False And S_VarSet = #False
        If *MainParent And ChangeCurrentElement(TokenList(), *MainParent)
          If Not PreviousElement(TokenList()) And FindString(#PBOperatorAsSign$, NextTokenValue$)
            ; MainParent is the first token of the code line and the following token is an assignment symbol
            S_VarSet = #True
          EndIf
        EndIf
      EndIf
    EndIf ; EndIf S_IsDotType = #False
    ;
    If S_LocalTypeOfElement = 0 And *Filter\NotAnElement = #False And S_IsDotType
      ; A token immediately preceded by '.' is a Structure, Interface, or primitive type qualifier
      If S_ExplicitModuleName$ <> "" Or FindMapElement(PBBasicNativeTypes(), *ExpToken\Value_LCase$) = 0
        Protected ResolvedDotEntity$ = ResolveStructureOrInterface(Left(S_ExplicitModuleName$ + "::", Bool(S_ExplicitModuleName$ <> "") * (Len(S_ExplicitModuleName$) + 2)) + *ExpToken\Value$)
        EntityColonPos = FindString(ResolvedDotEntity$, "::")
        If EntityColonPos : S_ResolvedEntityModule$ = StringField(Left(ResolvedDotEntity$, EntityColonPos - 1), 2, " ") : EndIf
        If Left(ResolvedDotEntity$, Len(PBBTypeNames$(#PBBInterface))) = PBBTypeNames$(#PBBInterface)
          S_LocalTypeOfElement = #PBBInterface
        ElseIf ResolvedDotEntity$ <> ""
          S_LocalTypeOfElement = #PBBStructure
        EndIf
      EndIf
      S_HasParentheses = 0
    ElseIf S_LocalTypeOfElement = 0 And *Filter\NotAnElement = #False And S_IsDotType = #False
      ;
      If PrevTokenValue$ = "\"   : S_LocalTypeOfElement = #PBBField
      ElseIf FindString(",and,or,xor,not,to,step,", "," + *ExpToken\Value_LCase$ + ",") And PrevTokenValue$
        S_LocalTypeOfElement = #PBBBasicKeyword
      ElseIf Left(ToSearch$, 1) = "#" : S_LocalTypeOfElement = #PBBConstant
      ElseIf NextTokenValue$ = ":"
        ; Trailing colon directly anchors a label declaration statement
        S_LocalTypeOfElement = #PBBLabel
        S_Definition         = #True
      ElseIf PrevTokenValue$ = "?"
        ; Leading question mark resolves the memory address of an existing label
        S_LocalTypeOfElement = #PBBLabel
      ElseIf NextTokenValue$ = "::"  : S_LocalTypeOfElement = #PBBModule
      ElseIf PrevTokenValue$ = "@" And S_HasParentheses = 0 : S_LocalTypeOfElement = #PBBVariable
      ElseIf S_IsFirstWordEqualPrevWord And *ExpToken <> *FirstWord
        ;
        ; Structural definitions must be identified before evaluating unparenthesized variables
        For ctType = #PBBProcedure To (#EndEnumPBBElementTypes - 1)
          If S_FirstWord_LCase$ = LCase(PBBTypeNames$(ctType)) Or (S_FirstWord_LCase$ = "declare" And ctType = #PBBProcedure)
            S_LocalTypeOfElement = ctType : S_Definition = #True : Break
          EndIf
        Next
      EndIf
      ;
      If S_LocalTypeOfElement = 0
        If (S_FirstWord_LCase$ = "usemodule" Or S_FirstWord_LCase$ = "unusemodule" Or S_FirstWord_LCase$ = "declaremodule") And *ExpToken <> *FirstWord
          S_LocalTypeOfElement = #PBBModule
          If S_FirstWord_LCase$ = "declaremodule"
            S_Definition = #True
          EndIf
        ElseIf S_FirstWord_LCase$ = "undefinemacro" And *ExpToken <> *FirstWord
          ; 'UndefineMacro <Name>' targets the active macro being revoked regardless of parameter parentheses
          S_LocalTypeOfElement = #PBBMacro
          S_HasParentheses     = -1
        ElseIf (S_FirstWord_LCase$ = "goto" Or S_FirstWord_LCase$ = "gosub" Or S_FirstWord_LCase$ = "restore") And *ExpToken <> *FirstWord
          S_LocalTypeOfElement = #PBBLabel
        ElseIf *ExpToken <> *FirstWord And AfterEqual = #False And InParenthesis = 0 And FindMapElement(PB_VariableDeclarators(), S_FirstWord_LCase$)
          ; Handles explicit variable, array, list, and map declarations
          S_LocalTypeOfElement = #PBBVariable : S_Definition = #True
        ElseIf FindString(",for,foreach,swap,with,read,", "," + S_PrevWord_LCase$ + ",") And *ExpToken <> *FirstWord
          ; 'ForEach' strictly targets a List or Map variable
          S_LocalTypeOfElement = #PBBVariable
        ElseIf S_HasTrailingBackslash And S_CallerFunction$ <> "offsetof" And S_CallerFunction$ <> "sizeof"
          ; Any root identifier followed by '\Field' (outside OffsetOf/SizeOf) is a structured variable, list, map, or array
          S_LocalTypeOfElement = #PBBVariable
        ElseIf S_HasParentheses = 0
          ; Unparenthesized identifier in an expression or receiving an assignment is a variable
          If *ExpToken <> *FirstWord Or BeforeEqualAssignation Or (AfterEqual = 0 And FindString(#PBOperatorAsSign$, OperatorTestSymbol$))
            S_LocalTypeOfElement = #PBBVariable
          EndIf
        EndIf
      EndIf
      ;
      ; Resolve Macro lifecycle (Macro / UndefineMacro) vs PureBasic Native Functions at PosInCode.
      ; Never assigns #PBBNativeFunction here so GetInfosFromContext() Section 3 retains full control.
      If S_LocalTypeOfElement = #PBBMacro
        S_MacroDeclPos = CheckIfMacroOverridesNative(*ExpToken\Value_LCase$, FileName$, PosInCode, @S_MacroDeclFile, -1)
        If S_MacroDeclPos = 0 And S_Definition = #False
          S_LocalTypeOfElement = #PBBNotAnElement
        EndIf
      ElseIf (S_LocalTypeOfElement = 0 Or (S_LocalTypeOfElement = #PBBVariable And S_Definition = #False)) And
             (*Filter\TypeOfElement = #PBBMacro Or *Filter\TypeOfElement = #PBBNativeFunction Or
              (S_HasParentheses = 1 And FindMapElement(PBNativeFunctions(), *ExpToken\Value_LCase$)) Or
              (ExplorationState = #ExplorationState_Completed And ListSize(ElementList()) And (ActionFlags & #PBB_Action_Collect) And
               (ActionFlags & #PBB_Action_CollectAll) = 0))
        S_MacroDeclPos = CheckIfMacroOverridesNative(*ExpToken\Value_LCase$, FileName$, PosInCode, @S_MacroDeclFile, S_HasParentheses)
        If S_MacroDeclPos > 0
          S_LocalTypeOfElement = #PBBMacro
        ElseIf *Filter\TypeOfElement = #PBBMacro
          ; Macro is not active at PosInCode (precedes Macro declaration, follows UndefineMacro, or lacks required parentheses)
          S_LocalTypeOfElement = #PBBNotAnElement
        EndIf
      EndIf
    EndIf ; EndIf S_LocalTypeOfElement = 0
    ;
    ; -------------------------------------------------------------------------
    ; 4. EARLY COMPATIBILITY GATING (Fast-Path Exit)
    ; Role: Saves CPU cycles by exiting immediately if the deduced concrete type
    ;       directly contradicts the filter requirements.
    ; Bypass: Explicitly bypassed when #PBB_Action_CollectAll is active or when
    ;         searching purely by purified name (*Filter\SearchPurifiedOnly).
    ; -------------------------------------------------------------------------
    ;
    If Not (ActionFlags & #PBB_Action_CollectAll) And Not (*Filter\SearchPurifiedOnly)
      ;
      ; Strict concrete element type check:
      If *Filter\TypeOfElement <> 0 And S_LocalTypeOfElement <> 0 And *Filter\TypeOfElement <> S_LocalTypeOfElement
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowBuildLayout
            OPDReport("  BuildLayout early exit (4) for filter " + PBBTypeNames$(*Filter\TypeOfElement) + " " + *Filter\ElementName$)
            OPDReport("  --> TypeOfElement doesn't match : *Filter\TypeOfElement = " + Str(*Filter\TypeOfElement))
            OPDReport("                                    S_LocalTypeOfElement  = " + Str(S_LocalTypeOfElement))
          EndIf
        CompilerEndIf
        ProcedureReturn #False
      EndIf
      ;
      ; Inference by absence of type for immediately identifiable types.
      ; By checking *PrevToken for fields, we ensure that a field declaration
      ; within a structure is not rejected (in this case, the expression starts the line).
      If S_LocalTypeOfElement = 0 And Not (S_CallerFunction$ = "defined" And S_ArgPosition = 1) And
         ((*Filter\TypeOfElement = #PBBField And *PrevToken) Or
          *Filter\TypeOfElement = #PBBLabel Or
          *Filter\TypeOfElement = #PBBConstant)
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowBuildLayout
            OPDReport("  BuildLayout early exit (4) for filter " + PBBTypeNames$(*Filter\TypeOfElement) + " " + *Filter\ElementName$)
            OPDReport("  --> TypeOfElement doesn't match : *Filter\TypeOfElement = " + Str(*Filter\TypeOfElement))
            OPDReport("                                    S_LocalTypeOfElement  = " + Str(S_LocalTypeOfElement))
          EndIf
        CompilerEndIf
        ProcedureReturn #False
      EndIf
    EndIf ; EndIf Fast-Path Exit (Section 4)
  EndIf ; End of the Else section of If *Filter\NotAnElement Or InComment Or InQuote
  ;
  ; -------------------------------------------------------------------------
  ; 5. Populate S_EncompassingParent$ and Field's parent
  ;
  ; Only root module declarations are top-level and cannot be enclosed by an outer container.
  ; Procedures and macros can reside inside a Module, while executable usages and calls can be
  ; hosted inside an outer Procedure, Macro, or Module. When TypeOfElement = 0, all potential
  ; containers are scanned to infer where the occurrence lives.
  ;
  ; A.----- Identify Encompassing for mention " •  In Procedure XXXX", " •  In Structure XXXX", etc. -----
  ;
  Protected S_EncompassingModuleName$, S_EncompassingModulePrefix$
  ;
  ; Only root module declarations are truly top-level; procedures and macros can reside in modules
  IsTopLevelDefinition = Bool(S_Definition And S_LocalTypeOfElement = #PBBModule)
  If IsTopLevelDefinition = #False
    ;
    ; Iterate through potential outer containers in hierarchical order
    For ctType = #PBBProcedure To #PBBImport
      If ctType >= #PBBLabel And ctType <= #PBBVariable
        Continue
      EndIf
      ;
      ; An entity declaration cannot enclose an entity of its identical category
      If S_Definition And S_LocalTypeOfElement = ctType
        Continue
      ElseIf ctType = #PBBMacro And IsInsideProcOrMacro
        ; Procedures and macros cannot be nested inside one another in PureBasic
        Continue
      ElseIf S_Definition And (S_LocalTypeOfElement = #PBBProcedure Or S_LocalTypeOfElement = #PBBMacro) And ctType <> #PBBModule
        ; Procedures and macros can only be enclosed by outer modules, skipping inner blocks
        Continue
      ElseIf InComment = #False And InQuote = 0 And ctType >= #PBBEnumeration And ctType <= #PBBInterface
        ; Declarative blocks cannot enclose active procedural calls: skip to evaluate outer modules
        If *Filter\TypeOfElement = #PBBProcedure Or S_LocalTypeOfElement = #PBBProcedure Or
           *Filter\TypeOfElement = #PBBMacro Or S_LocalTypeOfElement = #PBBMacro
          Continue
        ; Variables and labels cannot reside as code inside declarative blocks
        ElseIf (*Filter\TypeOfElement <> 0 Or S_LocalTypeOfElement <> 0) And
               *Filter\TypeOfElement <> #PBBConstant And *Filter\TypeOfElement <> #PBBField And
               S_LocalTypeOfElement <> #PBBConstant And S_LocalTypeOfElement <> #PBBField
          Continue
        ; Named Constants can only be enclosed by Procedures, Macros, or Enumerations
        ElseIf (*Filter\TypeOfElement = #PBBConstant Or S_LocalTypeOfElement = #PBBConstant) And ctType > #PBBEnumeration
          Continue
        ; Structure fields belong exclusively to Structures or Interfaces, skipping Enumerations
        ElseIf (*Filter\TypeOfElement = #PBBField Or S_LocalTypeOfElement = #PBBField) And ctType = #PBBEnumeration
          Continue
        EndIf
      EndIf
      ;
      ; Prevent false parent lookups when PosInCode is located on the declaration line of the container itself
      If S_PrevWord_LCase$ <> LCase(PBBTypeNames$(ctType))
        S_Prefix$ = GetEncompassingLimitsAndName(ctType, FileName$, PosInCode, *SearchString, *SearchString_LCase, EndLimit, EncompassingElementData)
        ;
        If S_Prefix$
          ; Modules act as namespaces, leaving physical container attribution to procedures or structures
          If ctType <> #PBBModule
            S_EncompassingParent$ = EncompassingElementData\TypeOfElementName$ + #Space$ + EncompassingElementData\ElementName$
          EndIf
          ;
          ; Procedural boundaries define local execution frames and isolate local variables
          If ctType = #PBBProcedure Or ctType = #PBBMacro
            IsInsideProcOrMacro    = #True
            S_StartingContextLimit = EncompassingElementData\StartingPos
            S_EndingContextLimit   = EncompassingElementData\EndingPos
            LO_LayoutPrefix$       = S_Prefix$
          ElseIf ctType = #PBBModule
            ; Track enclosing module namespace for both scoping logic and layout formatting
            S_EncompassingModuleName$   = EncompassingElementData\ElementName$
            S_EncompassingModulePrefix$ = S_Prefix$
          Else
            ; Declarative blocks (Structures, Interfaces, or Imports) for categorizing internal declarations
            If (ctType = #PBBStructure Or ctType = #PBBInterface Or ctType = #PBBImport) And
               PrevTokenValue$ = "" And *Filter\NotAnElement = #False And InComment = #False And InQuote = 0
              If ctType = #PBBImport
                If *ExpToken = *FirstWord And S_FirstWord_LCase$ <> "import" And Left(S_FirstWord_LCase$, 8) <> "compiler"
                  S_LocalTypeOfElement = #PBBImported
                  S_Definition         = #True
                EndIf
              Else
                ; Any active code root identifier on a line inside a Structure/Interface block is a field definition
                S_LocalTypeOfElement  = #PBBField
                S_Definition          = #True
                S_FieldStructureName$ = S_EncompassingParent$
              EndIf
            EndIf
            ;
            ; If a Structure or Enumeration is nested inside a Procedure or Macro, build the composite prefix
            If IsInsideProcOrMacro
              LO_LayoutPrefix$ + #Space$ + S_EncompassingParent$
            Else
              LO_LayoutPrefix$ = S_Prefix$
            EndIf
          EndIf
        EndIf
      EndIf
    Next ; Next ctType
  EndIf  ; EndIf IsTopLevelDefinition = #False
  ;
  If S_IsDotType And S_LocalTypeOfElement <> 0 And S_ExplicitModuleName$ = ""
    ; Refine unqualified '.Type' resolution against the enclosing module and active 'UseModule' context
    ResolvedDotEntity$ = ResolveStructureOrInterface(*ExpToken\Value$, 0, 0, 0, #False, S_EncompassingModuleName$, FileName$, PosInCode)
    EntityColonPos     = FindString(ResolvedDotEntity$, "::")
    If EntityColonPos : S_ResolvedEntityModule$ = StringField(Left(ResolvedDotEntity$, EntityColonPos - 1), 2, " ") : Else : S_ResolvedEntityModule$ = "" : EndIf
    If Left(ResolvedDotEntity$, Len(PBBTypeNames$(#PBBInterface))) = PBBTypeNames$(#PBBInterface)
      S_LocalTypeOfElement = #PBBInterface
    ElseIf ResolvedDotEntity$ <> ""
      S_LocalTypeOfElement = #PBBStructure
    EndIf
  EndIf
  ;
  ; B. ----- Identify encompassing "With" for fields -----
  ;
  If S_LocalTypeOfElement = #PBBField
    If MainParent$ = "" And GetEncompassingLimitsAndName(#PBBWith, FileName$, PosInCode, *SearchString, *SearchString_LCase, EndLimit, EncompassingElementData)
      MainParent$ = EncompassingElementData\ElementName$
      If ParentVar$ = "" : ParentVar$ = MainParent$ : Else : ParentVar$ = MainParent$ + "\" + ParentVar$ : EndIf
    EndIf
    ;
    ; C. ----- Resolve S_FieldStructureName$ for #PBBField -----
    ;
    If S_FieldStructureName$ = "" And ParentVar$ <> ""
      If FindString(MainParent$, ".") = 0
        MainParent$ = S_EncompassingParent$
      EndIf
      S_FieldStructureName$ = GetStructureOrInterfaceNameFromVariable(ParentVar$, MainParent$, FileName$, PosInCode,
                                                                      S_StartingContextLimit, #False, -1, "", S_EncompassingModuleName$)
      EntityColonPos = FindString(S_FieldStructureName$, "::")
      If EntityColonPos
        S_ResolvedEntityModule$ = StringField(Left(S_FieldStructureName$, EntityColonPos - 1), 2, " ")
        S_FieldStructureName$   = StringField(S_FieldStructureName$, 1, " ") + #Space$ + Mid(S_FieldStructureName$, EntityColonPos + 2)
      EndIf
    EndIf
  EndIf ; EndIf S_LocalTypeOfElement = #PBBField
  ;
  ; D. ----- SUBSPECIES RESOLUTION (Declarative) -----
  ; Role: Identifies structured variable subtypes from explicit declarators.
  ;
  If (S_LocalTypeOfElement = 0 Or S_LocalTypeOfElement = #PBBVariable) And S_IsDotType = #False
    If *PrevToken And *PrevToken\Type = #TokenType_Word
      PrevTok_LCase$ = *PrevToken\Value_LCase$
    EndIf
    ;
    ; Check both surrounding syntax words and the immediate preceding token
    ; to detect parameter declarators inside procedure signatures.
    If S_PrevWord_LCase$ = "dim" Or S_PrevWord_LCase$ = "redim" Or S_PrevWord_LCase$ = "array" Or
       S_FirstWord_LCase$ = "array" Or PrevTok_LCase$ = "array"
      S_LocalSubSpecies$ = "Array"
    ElseIf S_PrevWord_LCase$ = "newlist" Or S_PrevWord_LCase$ = "list" Or S_FirstWord_LCase$ = "newlist" Or
           S_FirstWord_LCase$ = "list" Or PrevTok_LCase$ = "list"
      S_LocalSubSpecies$ = "List"
    ElseIf S_PrevWord_LCase$ = "newmap" Or S_PrevWord_LCase$ = "map" Or S_FirstWord_LCase$ = "newmap" Or
           S_FirstWord_LCase$ = "map" Or PrevTok_LCase$ = "map"
      S_LocalSubSpecies$ = "Map"
    EndIf
  EndIf
  ;
  ; E. ----- LOCAL SPECIES RESOLUTION (Declarative) -----
  ; Role: Detects explicit scope keywords directly on declaration statements.
  ; Shared early to allow fast-path filtering before layout/collect.
  ;
  If S_Parameter
    S_LocalSpecies$ = "Param" ; Tag explicit procedure parameter scope
  ElseIf S_Definition And FindMapElement(PB_VariableDeclarators(), S_FirstWord_LCase$)
    ; Validate that the declarator represents a scope boundary before assigning species
    If FindString(PBAreaDeclarator_LCase$, "," + S_FirstWord_LCase$ + ",")
      S_LocalSpecies$ = PB_VariableDeclarators() ; Assign normalized case species directly from declarator map
    EndIf
  EndIf
  ;
  ; -------------------------------------------------------------------------
  ; 6. Snippet Heuristics (Native Functions Inference)
  ; Role: Attempts to complete S_LocalTypeOfElement when the element is
  ;       involved in a PureBasic function call.
  ; -------------------------------------------------------------------------
  ;
  If S_Definition = 0 And (S_LocalTypeOfElement = 0 Or S_LocalTypeOfElement = #PBBVariable) And S_CallerFunction$ <> ""
    If FindString(#PBB_ListFunctions, "," + S_CallerFunction$ + ",")
      S_LocalTypeOfElement = #PBBVariable
      S_LocalSubSpecies$ = "List"
    ElseIf FindString(#PBB_MapFunctions, "," + S_CallerFunction$ + ",")
      S_LocalTypeOfElement = #PBBVariable
      S_LocalSubSpecies$ = "Map"
    ElseIf FindString(#PBB_ArrayFunctions, "," + S_CallerFunction$ + ",")
      S_LocalTypeOfElement = #PBBVariable
      S_LocalSubSpecies$ = "Array"
    ElseIf S_CallerFunction$ = "initializestructure" Or S_CallerFunction$ = "clearstructure" Or S_CallerFunction$ = "resetstructure"
      If S_ArgPosition = 1
        ; The first argument is the target structured variable or memory pointer
        S_LocalTypeOfElement = #PBBVariable
      ElseIf S_ArgPosition = 2
        S_LocalTypeOfElement = #PBBStructure
      EndIf
    ElseIf S_CallerFunction$ = "copystructure" Or S_CallerFunction$ = "comparestructure"
      If S_ArgPosition = 1 Or S_ArgPosition = 2
        ; The first two arguments are the source and destination structured variables
        S_LocalTypeOfElement = #PBBVariable
      ElseIf S_ArgPosition = 3
        S_LocalTypeOfElement = #PBBStructure
      EndIf
    ElseIf S_CallerFunction$ = "defined" And S_ArgPosition = 1
      ; Resolve target element category, parentheses, and subspecies from the 2nd argument of Defined()
      ChangeCurrentElement(TokenList(), *ExpToken)
      If NextElement(TokenList()) And TokenList()\Value$ = "," And NextElement(TokenList()) And TokenList()\Type = #TokenType_Word
        Protected DefinedSubStr.String
        S_LocalTypeOfElement = DecodePBDefinedType(UCase(TokenList()\Value_LCase$), @S_HasParentheses, @DefinedSubStr)
        If DefinedSubStr\s <> "Var" : S_LocalSubSpecies$ = DefinedSubStr\s : EndIf
        If S_LocalTypeOfElement = #PBBConstant And Left(*ExpToken\Value$, 1) <> "#"
          *ExpToken\Value$ = "#" + *ExpToken\Value$
        EndIf
      EndIf
    EndIf
  EndIf ; EndIf Native Functions Inference (Section 6)
  ;
  ; -------------------------------------------------------------------------
  ; 7. CREATION & INITIALIZATION OF COLLECTEDPARAMS
  ; -------------------------------------------------------------------------
  ;
  ; --- A. Initialization of CollectedParams ---
  ;
  Protected IsCompatible = #True
  Protected CollectedParams.SearchDetails
  If S_LocalTypeOfElement = #PBBField And Left(*ExpToken\Value$, 1) = "*"
    ; When declaring a field within a structure, the field name may begin with "*"
    ; to indicate that it is a pointer, but it is always referenced without the asterisk:
    *ExpToken\Value$ = FastMid(*ExpToken\Value$, 2)
  EndIf
  PrepareSearchDetailsElement(@CollectedParams, *ExpToken\Value$)
  ;
  If S_LocalTypeOfElement = #PBBMacro And S_MacroDeclPos > 0
    ; Anchor macro occurrences to their active 'Macro' declaration position
    CollectedParams\ElementPosition    = S_MacroDeclPos
    CollectedParams\ElementFile$       = S_MacroDeclFile\s
  Else
    CollectedParams\ElementPosition    = PosInCode
    CollectedParams\ElementFile$       = FileName$
  EndIf
  CollectedParams\HasParentheses     = S_HasParentheses
  CollectedParams\VarDotType$        = CP_DeterminedVarType$
  CollectedParams\TypeOfElement      = S_LocalTypeOfElement
  CollectedParams\ElementSubSpecies$ = S_LocalSubSpecies$
  CollectedParams\ElementSpecies$    = S_LocalSpecies$
  ;
  ; --- B. Subspecies & Species Resolution ---
  ;
  If CollectedParams\TypeOfElement = 0 Or CollectedParams\TypeOfElement = #PBBVariable
    ;
    If S_Definition And CollectedParams\ElementSpecies$ = "" And CollectedParams\ElementSubSpecies$ <> ""
      CollectedParams\ElementSpecies$ = "Define"
    ElseIf CollectedParams\ElementSpecies$ = "" And (CollectedParams\ElementSubSpecies$ <> "" Or S_VarSet)
      CollectedParams\ElementSpecies$ = "Implicit"
    EndIf
    ;
    If CollectedParams\TypeOfElement = #PBBVariable And CollectedParams\ElementSpecies$ = ""
      CollectedParams\ElementSpecies$ = "Implicit"
    EndIf
    ;
    ; Register or inherit active procedural species across sequential statements in the same procedure.
    ; Passing S_HasParentheses appends "()" to prevent a scalar variable 'MyVar' from colliding with 'MyVar()'.
    If *Filter\SearchPurifiedOnly = #False And S_StartingContextLimit > 0
      If CollectedParams\ElementSpecies$ = "Protected" Or
         CollectedParams\ElementSpecies$ = "Static" Or
         CollectedParams\ElementSpecies$ = "Param" Or
         CollectedParams\ElementSpecies$ = "Shared"
        ManageProcMacroScopeShadowing(#PBB_Shadow_Register, S_StartingContextLimit, CollectedParams\ElementPurifiedName$, S_HasParentheses, CollectedParams\ElementSpecies$)
      ElseIf CollectedParams\ElementSpecies$ = "" Or CollectedParams\ElementSpecies$ = "Implicit"
        Protected ActiveProcSpecies$ = ManageProcMacroScopeShadowing(#PBB_Shadow_Query, S_StartingContextLimit, CollectedParams\ElementPurifiedName$, S_HasParentheses)
        If ActiveProcSpecies$ <> ""
          CollectedParams\ElementSpecies$ = ActiveProcSpecies$
        EndIf
      EndIf
    EndIf
    ;
    If CollectedParams\ElementSpecies$ And CollectedParams\ElementSubSpecies$ = ""
      If S_HasParentheses = 0
        CollectedParams\ElementSubSpecies$ = "Var"
      EndIf
    EndIf
    ;
    ; Promote untyped tokens (such as parenthesized List/Map/Array usages) once their species is established
    If CollectedParams\ElementSpecies$ And CollectedParams\TypeOfElement = 0
      CollectedParams\TypeOfElement = #PBBVariable
    EndIf
    ;
    ; Bind the enclosing procedural or macro frame to the variable's parent attribute
    If CollectedParams\ElementSpecies$ <> "Global" And CollectedParams\ElementSpecies$ <> "Threaded"
      CollectedParams\ParentTypeAndName$ = S_EncompassingParent$
    EndIf
    ;
  ElseIf S_LocalTypeOfElement = #PBBField
    ; Enforce pure species identity and assign the resolved Structure or Interface type
    CollectedParams\ParentTypeAndName$ = S_FieldStructureName$
  ElseIf S_LocalTypeOfElement = #PBBImported And S_Definition
    CollectedParams\ParentTypeAndName$ = S_EncompassingParent$
  ElseIf S_LocalTypeOfElement = #PBBConstant And *ExpToken = *FirstWord And
         Left(S_EncompassingParent$, Len(PBBTypeNames$(#PBBEnumeration)) + 1) = PBBTypeNames$(#PBBEnumeration) + #Space$
    ; Bind the enclosing Enumeration only when the constant is being declared inside it
    CollectedParams\ParentTypeAndName$ = S_EncompassingParent$
  EndIf ; EndIf CollectedParams\TypeOfElement = 0 Or #PBBVariable
  ;
  ; Resolve the effective owning module (explicit 'Mod::', enclosing module, or active 'UseModule' import)
  If InComment = #False And InQuote = 0 And *Filter\NotAnElement = #False And *Filter\SearchPurifiedOnly = #False
    If CollectedParams\TypeOfElement = #PBBModule Or
       CollectedParams\TypeOfElement = #PBBNativeFunction Or
       CollectedParams\TypeOfElement = #PBBBasicKeyword Or
       CollectedParams\TypeOfElement = #PBBNotAnElement Or
       (S_ExplicitModuleName$ = "" And (ActionFlags & #PBB_Action_CollectAll) = 0 And
        (*Filter\TypeOfElement = #PBBNativeFunction Or *Filter\TypeOfElement = #PBBBasicKeyword))
      CollectedParams\ModuleName$ = ""
    ElseIf (CollectedParams\TypeOfElement = #PBBField And S_Definition = #False) Or S_IsDotType
      CollectedParams\ModuleName$ = S_ResolvedEntityModule$
      If S_IsDotType And S_ExplicitModuleName$ <> "" And
         LCase(S_ExplicitModuleName$) <> LCase(S_EncompassingModuleName$) And
         CheckIfDeclaredInPublicModule(S_ExplicitModuleName$, CollectedParams\ElementPurifiedName$) = #False
        CollectedParams\TypeOfElement = #PBBNotAnElement
      EndIf
    ElseIf S_ExplicitModuleName$ <> ""
      CollectedParams\ModuleName$ = S_ExplicitModuleName$
      If LCase(S_ExplicitModuleName$) <> LCase(S_EncompassingModuleName$) And
         (CollectedParams\TypeOfElement = #PBBLabel Or
          CheckIfDeclaredInPublicModule(S_ExplicitModuleName$, CollectedParams\ElementPurifiedName$) = #False)
        CollectedParams\TypeOfElement = #PBBNotAnElement
      EndIf
    Else
      CollectedParams\ModuleName$ = S_EncompassingModuleName$
      If S_Definition = #False And
         CollectedParams\TypeOfElement <> #PBBLabel And
         CollectedParams\ElementSpecies$ <> "Protected" And
         CollectedParams\ElementSpecies$ <> "Static" And
         CollectedParams\ElementSpecies$ <> "Param" And
         Not ((ActionFlags & #PBB_Action_CollectAll) And S_StartingContextLimit > 0 And
              (CollectedParams\TypeOfElement = 0 Or CollectedParams\TypeOfElement = #PBBVariable))
        Protected ImportedModule$ = CheckIfModuleIsUsed("", FileName$, PosInCode, CollectedParams\ElementPurifiedName$, S_EncompassingModuleName$)
        If ImportedModule$ <> ""
          CollectedParams\ModuleName$ = ImportedModule$
        EndIf
      EndIf
    EndIf
  EndIf
  ;
  ; --- C. Element Naming & Absolute Safety Net ---
  ;
  ; Append explicit dot-type only to non-string identifiers lacking type qualifiers
  If CollectedParams\VarDotType$ And FindString(CollectedParams\ElementName$, ".") = 0 And
     Right(CollectedParams\ElementName$, 1) <> "$"
    CollectedParams\ElementName$ + "." + CollectedParams\VarDotType$
  EndIf
  If CollectedParams\HasParentheses = 1 And FindString(CollectedParams\ElementName$, "(") = 0
    CollectedParams\ElementName$ + "()"
  EndIf
  CollectedParams\ElementPurifiedName$ = PurifyForSearch(CollectedParams\ElementName$)
  ;
  If CollectedParams\TypeOfElement = #PBBVariable And
     CollectedParams\ElementSpecies$ <> "Global" And
     CollectedParams\ElementSpecies$ <> "Threaded"
    CollectedParams\StartingContextLimit = S_StartingContextLimit
    CollectedParams\EndingContextLimit   = S_EndingContextLimit
  EndIf
  ;
  ; -------------------------------------------------------------------------
  ; 8. GATING COMPATIBILITY TEST WITH *FILTER's DATA
  ; -------------------------------------------------------------------------
  ;
  If *Filter\SearchPurifiedOnly = #False And InComment = #False And InQuote = 0
    ;
    If AreSearchDetailsCompatible(@CollectedParams, *Filter) = #False
      IsCompatible = #False
    EndIf
    ;
    CompilerIf #PBBReport
      If (OPDMod & #OPDFollowBuildLayout) And IsCompatible = #False And Not (ActionFlags & #PBB_Action_CollectAll)
        OPDReport("  BuildLayout Incompatibility Exit (8)")
        OPDReport("       " + LineOfCode$)
        OPDReport("       *Filter: " + PBBTypeNames$(*Filter\TypeOfElement) + " " + *Filter\ElementName$)
        OPDReport("       HasParentheses: " + Str(*Filter\HasParentheses) + "  VarDotType: " + *Filter\VarDotType$)
        ;
        OPDReport("       Species: " + *Filter\ElementSpecies$ + " " + *Filter\ElementSubSpecies$)
        OPDReport("       ContextLimit: " + Str(*Filter\StartingContextLimit) + " / " + Str(*Filter\EndingContextLimit))
        OPDReport("       ParentTypeAndName$: " + #DOUBLEQUOTE$ + *Filter\ParentTypeAndName$ + #DOUBLEQUOTE$)
        OPDReport("")
        OPDReport("       CollectedParams: " + PBBTypeNames$(CollectedParams\TypeOfElement) + " " + CollectedParams\ElementName$)
        OPDReport("       HasParentheses: " + Str(CollectedParams\HasParentheses) + "  VarDotType: " + CollectedParams\VarDotType$)
        OPDReport("       Species: " + CollectedParams\ElementSpecies$ + " " + CollectedParams\ElementSubSpecies$)
        OPDReport("       ContextLimit: " + Str(CollectedParams\StartingContextLimit) + " / " + Str(CollectedParams\EndingContextLimit))
        OPDReport("       ParentTypeAndName$: " + #DOUBLEQUOTE$ + CollectedParams\ParentTypeAndName$ + #DOUBLEQUOTE$)
      EndIf
    CompilerEndIf
  EndIf ; EndIf *Filter\SearchPurifiedOnly = #False (Section 8)
  ;
  ; -------------------------------------------------------------------------
  ; 9. ACTION MAPPING & LAYOUT ASSEMBLY (Isolated)
  ; Role: Triggered strictly by #PBB_Action_Layout. Translates the deduced
  ;       logical roles into standardized UI markers (Prefix + Core + Suffix)
  ;       and builds the final layout string appended with the file's line number.
  ; -------------------------------------------------------------------------
  ;
  If (ActionFlags & #PBB_Action_Layout) And *LayoutDefinition And IsCompatible
    Protected LO_LayoutCore$, LO_LayoutRoleSuffix$
    ;
    ; --- A. Spatial Context ---
    ; Prepend enclosing module namespace when element resides inside DeclareModule or Module
    If S_EncompassingModulePrefix$
      If LO_LayoutPrefix$
        LO_LayoutPrefix$ = S_EncompassingModulePrefix$ + #Space$ + Mid(LO_LayoutPrefix$, Len(PBB_LineBreak$) + 1)
      Else
        LO_LayoutPrefix$ = S_EncompassingModulePrefix$
      EndIf
    EndIf
    ;
    If LO_LayoutPrefix$ = ""
      ; Root-level declarations (Procedure through Module) define their own spatial frame
      ; and must not be marked with the '#OutOfElements' fallthrough label.
      If S_Definition And ((S_LocalTypeOfElement >= #PBBProcedure And S_LocalTypeOfElement <= #PBBModule) Or
                           S_LocalTypeOfElement = #PBBImport)
        LO_LayoutPrefix$ = PBB_LineBreak$
      Else
        ; True orphan instructions located in the global body outside any root container
        LO_LayoutPrefix$ = PBB_LineBreak$ + #OutOfElementsName
      EndIf
    EndIf
    ;
    ; Include FileName$ in the cached prefix key so transitioning to a new file
    ; always emits the leading line break (PBB_LineBreak$) and spatial scope header:
    If FileName$ + LO_LayoutPrefix$ <> LO_LastLayoutPrefix$
      LO_LastLayoutPrefix$ = FileName$ + LO_LayoutPrefix$
      LO_LayoutPrefix$ + #Space$
    Else
      LO_LayoutPrefix$ = ""
    EndIf
    ;
    ; --- B. Element Identity ---
    If InComment
      LO_LayoutCore$ = ";"
    ElseIf InQuote
      LO_LayoutCore$ = #DOUBLEQUOTE$ + "..." + #DOUBLEQUOTE$
    ElseIf S_Definition And S_FirstWord$ <> "" And S_LocalTypeOfElement = #PBBVariable
      LO_LayoutCore$ = S_FirstWord$
    ElseIf S_IsFirstWordEqualPrevWord And
           ((S_LocalTypeOfElement = #PBBModule And (S_FirstWord_LCase$ = "usemodule" Or S_FirstWord_LCase$ = "unusemodule")) Or
            (S_LocalTypeOfElement = #PBBMacro And S_FirstWord_LCase$ = "undefinemacro"))
      LO_LayoutCore$ = S_FirstWord$
    EndIf
    ;
    ; Restrict UI declaration markers strictly to genuine definition statements
    If S_Definition And (S_IsFirstWordEqualPrevWord Or S_LocalTypeOfElement = #PBBLabel)
      If S_LocalTypeOfElement = #PBBVariable
        AddStringElement(LO_LayoutCore$, #ItsHereChar$, #Space$)
      ElseIf S_FirstWord_LCase$ = "declare"
        AddStringElement(LO_LayoutCore$, "Declare" + #Space$ + #ItsHereChar$, #Space$)
      ElseIf S_FirstWord_LCase$ = "declaremodule"
        AddStringElement(LO_LayoutCore$, "DeclareModule" + #Space$ + #ItsHereChar$, #Space$)
      Else
        AddStringElement(LO_LayoutCore$, PBBTypeNames$(S_LocalTypeOfElement) + #Space$ + #ItsHereChar$, #Space$)
      EndIf
    EndIf
    ;
    If S_LocalTypeOfElement = #PBBField And S_Definition = #False
      ; Substitute syntactical keywords (e.g., 'If', 'While') with the resolved parent variable
      If ParentVar$ <> ""
        If *MainParent And *MainParent\ModulePrefix$ <> ""
          AddStringElement(LO_LayoutCore$, *MainParent\ModulePrefix$ + "::" + ParentVar$, #Space$)
        Else
          AddStringElement(LO_LayoutCore$, ParentVar$, #Space$)
        EndIf
        If S_HasParentheses = 1 Or LO_PrevIsClosingParenthesis
          LO_LayoutCore$ + "()"
        EndIf
        LO_LayoutCore$ + "\"
      Else
        ; Standalone backslash marker for orphan fields without a resolvable container
        AddStringElement(LO_LayoutCore$, "\", #Space$)
      EndIf
    ElseIf S_IsDotType
      If S_PrevWord$ <> ""
        AddStringElement(LO_LayoutCore$, S_PrevWord$, #Space$)
        If LO_NextIsParenthesis
          LO_LayoutCore$ + "()"
        EndIf
        LO_LayoutCore$ + "."
      EndIf
    EndIf
    ;
    ; --- C. Action Qualifiers ---
    Protected LO_ChainingOp$ = #NonBreakableSpace$ + "&" + #NonBreakableSpace$
    ;
    If S_FirstWord_LCase$ = "procedurereturn"
      AddStringElement(LO_LayoutRoleSuffix$, "Return", LO_ChainingOp$)
    EndIf
    If LO_VarTest
      AddStringElement(LO_LayoutRoleSuffix$, "Test", LO_ChainingOp$)
    EndIf
    If S_ArgPosition = 1 And (Left(S_CallerFunction$, 4) = "poke" Or Left(S_CallerFunction$, 4) = "peek") And Len(S_CallerFunction$) <= 5
      AddStringElement(LO_LayoutRoleSuffix$, UCase(Left(S_CallerFunction$, 1)) + FastMid(S_CallerFunction$, 2, 3) +
                                             UCase(Right(S_CallerFunction$, 1)), LO_ChainingOp$)
    ElseIf S_PrevWord_LCase$ = "read"
      AddStringElement(LO_LayoutRoleSuffix$, "Read", LO_ChainingOp$)
    ElseIf LO_Argument And S_Parameter = 0
      If S_CallerFunction$ = "clearlist" Or S_CallerFunction$ = "clearmap"
        AddStringElement(LO_LayoutRoleSuffix$, "Clear", LO_ChainingOp$)
      ElseIf S_CallerFunction$ = "freememory"
        AddStringElement(LO_LayoutRoleSuffix$, "FreeMemory", LO_ChainingOp$)
      Else
        AddStringElement(LO_LayoutRoleSuffix$, "Arg", LO_ChainingOp$)
      EndIf
    ElseIf S_VarSet
      AddStringElement(LO_LayoutRoleSuffix$, "Set", LO_ChainingOp$)
    EndIf
    ;
    If S_Parameter
      AddStringElement(LO_LayoutRoleSuffix$, "Param", LO_ChainingOp$)
    EndIf
    ;
    If LO_LayoutRoleSuffix$
      LO_LayoutRoleSuffix$ + #ThisPlaceChar$
    EndIf
    ;
    ; --- D. Final Layout String Assembly ---
    If *Filter\SearchInDeadZones And IsPosInDeadZone(FileName$, PosInCode)
      AddStringElement(LO_LayoutCore$, ReplaceString(GetTextFromCatalogPB("inactiveCompilerZone"), " ", #NonBreakableSpace$), #Space$)
    EndIf
    *LayoutDefinition\s = LO_LayoutCore$
    AddStringElement(*LayoutDefinition\s, LO_LayoutRoleSuffix$, #Space$)
    ;
    If *LayoutDefinition\s
      *LayoutDefinition\s + #Space$
    EndIf
    ;
    If LO_Extended
      ; The length of the found expression is greater than Len(*Filter\ElementPurifiedName$)
      ; (the search was probably performed with #PBBSearchPrefix).
      ; Specify the expanded found expression:
      *LayoutDefinition\s + "(" + ToSearch$ + ") "
    EndIf
    ;
    *LayoutDefinition\s + "[" + Str(GetNoLineFromStrAndPos(*SearchString, PosInCode)) + "]"
    ;
    ; Prepend the spatial container prefix for the first element on a line,
    ; or a comma separator when chaining subsequent occurrences on the same line.
    If LO_LayoutPrefix$
      *LayoutDefinition\s = LO_LayoutPrefix$ + *LayoutDefinition\s
    Else
      *LayoutDefinition\s = "," + #Space$ + *LayoutDefinition\s
    EndIf
  EndIf ; EndIf ActionFlags & #PBB_Action_Layout
  ;
  ; -------------------------------------------------------------------------
  ; 10. CONTEXT COLLECTION & DEDUPLICATION (Isolated)
  ; Role: Triggered strictly by #PBB_Action_Collect. Delegates merging to
  ;       InheritStrongSearchTraits() to maintain a deduplicated dictionary.
  ; -------------------------------------------------------------------------
  ;
  If ActionFlags & (#PBB_Action_Collect | #PBB_Action_CollectAll) And
     (IsCompatible Or ActionFlags & #PBB_Action_CollectAll) And
     ((InComment = #False And InQuote = 0) Or *Filter\SearchInCom Or *Filter\SearchInQuotes)
    ;
    Protected CP_IsNewContext = #True
    ;
    ; During a multi-line scope scan (#PBB_Action_CollectAll), encountering 'UndefineMacro'
    ; revokes any previously collected #PBBMacro definition so subsequent lines revert to native/normal scope:
    If (ActionFlags & #PBB_Action_CollectAll) And S_FirstWord_LCase$ = "undefinemacro"
      ForEach ContextFoundExpression()
        If ContextFoundExpression()\TypeOfElement = #PBBMacro
          DeleteElement(ContextFoundExpression())
        EndIf
      Next
    ElseIf CollectedParams\TypeOfElement <> 0 Or (ActionFlags & #PBB_Action_CollectAll)
      ;
      ForEach ContextFoundExpression()
        If InheritStrongSearchTraits(@CollectedParams, @ContextFoundExpression()) = #True
          CP_IsNewContext = #False
          Break
        EndIf
      Next ; Next ContextFoundExpression()
      ;
      If CP_IsNewContext
        AddElement(ContextFoundExpression())
        CopyStructure(@CollectedParams, @ContextFoundExpression(), SearchDetails)
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowBuildLayout
            OPDReport("Added: " + PBBTypeNames$(CollectedParams\TypeOfElement) + " " +
                      CollectedParams\ElementName$ + " - " + CollectedParams\ElementSpecies$ +
                      " " + CollectedParams\ElementSubSpecies$ + " " + CollectedParams\ParentTypeAndName$)
          EndIf
        CompilerEndIf
      EndIf
      ;
    EndIf
  EndIf ; EndIf ActionFlags & (#PBB_Action_Collect | #PBB_Action_CollectAll)
  ;
  ProcedureReturn Bool(IsCompatible Or (ActionFlags & #PBB_Action_CollectAll))
  ;
EndProcedure
;
; **********************************************************************************
;
;-                          Contextual scope deduction
;
; **********************************************************************************
;
Enumeration ScanContextReturnValues
  #PBBScan_None           = 0 ; No formal match found
  #PBBScan_HardMatch      = 1 ; Absolute local/global definition (Protected, Global, etc.)
  #PBBScan_SharedRedirect = 2 ; Soft match via 'Shared' redirector, requires a global root search
EndEnumeration
;
Procedure ScanContextRange(*Filter.SearchDetails, *BestMatch.SearchDetails, FileName$,
                           ScanStart, ScanEnd,
                           AcceptLocal, AcceptImplicit, *HasFallback.Integer)
  ;
  ; ROLE & ARCHITECTURE:
  ; Traverses a local procedural scope or global project sequences via
  ; FindInWholeProjectSequences() to resolve variable or structural definitions.
  ; Evaluates AST traits produced by BuildLayout() and categorizes matches
  ; into Hard Exits, Soft Exits, or Fallbacks.
  ;
  Shared ContextFoundExpression()
  ;
  Protected FIPS.FIPSPosStruct
  Protected ActionFlags, DoReset = #True
  ;
  If *Filter = 0 Or *Filter\ElementPurifiedName$ = ""
    ProcedureReturn #PBBScan_None
  EndIf
  ;
  ; Configure search boundaries: enforce strict local encapsulation when ScanStart > 0
  FIPS\FileName$           = FileName$
  FIPS\PosInCode           = ScanStart
  FIPS\IncludeDeadZones    = *Filter\SearchInDeadZones
  ;
  If AcceptLocal
    ;
    ; Enforce single-pass compilation: local declarations downstream of the target expression cannot qualify upstream usages
    If *Filter\ElementPosition > 0 And *Filter\ElementPosition < ScanEnd
      FIPS\EndLimit          = *Filter\ElementPosition
    Else
      FIPS\EndLimit          = ScanEnd
    EndIf
    FIPS\ScanMode            = #FIPSMode_LocalOnly
  Else
    FIPS\PosInCode           = 0
    FIPS\EndLimit            = ScanEnd
    FIPS\ScanMode            = #FIPSMode_Global
  EndIf
  ;
  FIPS\FIPSContinue = #False
  ;
  ; Iterate through all matching identifier occurrences across permitted sequences
  While FindInWholeProjectSequences(*Filter\ElementPurifiedName$, @FIPS, #CheckAll | #CheckOneKeyWordOnly)
    ;
    ActionFlags = #PBB_Action_Collect
    If DoReset
      ActionFlags | #PBB_Action_Reset
      DoReset = #False
    EndIf
    ;
    ; The GetEncompassingLimitsAndName() procedure called by BuildLayout() must have visibility
    ; extending beyond ScanEnd in order to detect the end of the element enclosing
    ; PosInCode. Thus, the end limit passed as a parameter to this procedure is the end
    ; of the code segment (FIPS\SegmentEnd) rather than ScanEnd.
    BuildLayout(*Filter, FIPS\FileName$, FIPS\SearchString, FIPS\SearchString_LCase,
                FIPS\SegmentEnd, FIPS\PosInCode, 0, ActionFlags)
    ;
    ; Evaluate candidates collected by BuildLayout() for the current line
    ForEach ContextFoundExpression()
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowBuildLayout
          OPDReport("  ScanContextrange is examining: " + PBBTypeNames$(ContextFoundExpression()\TypeOfElement) + " " + ContextFoundExpression()\ElementName$)
          OPDReport("       Species: " + ContextFoundExpression()\ElementSpecies$ + " " + ContextFoundExpression()\ElementSubSpecies$)
          OPDReport("       ContextLimit: " + Str(ContextFoundExpression()\StartingContextLimit) + " / " + Str(ContextFoundExpression()\EndingContextLimit))
          OPDReport("       ContextFoundExpression()\ParentTypeAndName$: " + #DOUBLEQUOTE$ + ContextFoundExpression()\ParentTypeAndName$ + #DOUBLEQUOTE$)
          OPDReport("       PosInCode: " + Str(FIPS\PosInCode) + "  /  ScanEnd: " + Str(ScanEnd))
        EndIf
      CompilerEndIf
      Protected CType = ContextFoundExpression()\TypeOfElement
      Protected CSp$  = ContextFoundExpression()\ElementSpecies$
      ;
      If CType = #PBBVariable
        Protected IsMatch = #False
        Protected ScanResult = #PBBScan_None
        ;
        ; Route variables based on scope declaration keywords
        If CSp$ = "Protected" Or CSp$ = "Static" Or CSp$ = "Param" Or CSp$ = "Shared"
          If AcceptLocal
            IsMatch = #True
            If CSp$ = "Shared"
              ScanResult = #PBBScan_SharedRedirect
            Else
              ScanResult = #PBBScan_HardMatch
            EndIf
          EndIf
          ;
        ElseIf CSp$ = "Global" Or CSp$ = "Define" Or CSp$ = "Threaded"
          IsMatch = #True
          ScanResult = #PBBScan_HardMatch
          ;
        ElseIf CSp$ = "Implicit" And AcceptImplicit
          IsMatch = #True
          *HasFallback\i = #True
        EndIf
        ;
        ; Inherit strong definition traits without overwriting physical anchor positions
        If IsMatch
          Protected IsInherited = #False
          If *BestMatch\ElementPosition = 0
            CopyStructure(@ContextFoundExpression(), *BestMatch, SearchDetails)
            IsInherited = InheritStrongSearchTraits(*Filter, *BestMatch)
          Else
            IsInherited = InheritStrongSearchTraits(@ContextFoundExpression(), *BestMatch)
          EndIf
          ;
          If IsInherited And ScanResult <> #PBBScan_None And IsSearchDetailsCompleted(*BestMatch, AcceptImplicit)
            CompilerIf #PBBReport
              If OPDMod & #OPDFollowBuildLayout
                If ScanResult = #PBBScan_HardMatch
                  OPDReport("       #PBBScan_HardMatch - 1")
                Else
                  OPDReport("       #PBBScan_SharedRedirect - 1")
                EndIf
              EndIf
            CompilerEndIf
            ProcedureReturn ScanResult
          EndIf
        EndIf
        ;
      Else
        ; Route structural declarations (Procedures, Macros, Structures, Fields, Interfaces)
        If *BestMatch\ElementPosition = 0
          CopyStructure(@ContextFoundExpression(), *BestMatch, SearchDetails)
          InheritStrongSearchTraits(*Filter, *BestMatch)
        Else
          InheritStrongSearchTraits(@ContextFoundExpression(), *BestMatch)
        EndIf
        ;
        ; Qualified field matches exit immediately, whereas unqualified fields register as fallbacks
        If CType = #PBBField
          If *Filter\ParentTypeAndName$ <> "" And *BestMatch\ParentTypeAndName$ = *Filter\ParentTypeAndName$
            CompilerIf #PBBReport
              If OPDMod & #OPDFollowBuildLayout
                OPDReport("       #PBBScan_HardMatch - 2")
              EndIf
            CompilerEndIf
            ProcedureReturn #PBBScan_HardMatch
          Else
            *HasFallback\i = #True
            CompilerIf #PBBReport
              If OPDMod & #OPDFollowBuildLayout
                OPDReport("       *HasFallback\i = #True")
              EndIf
            CompilerEndIf
          EndIf
        Else
          CompilerIf #PBBReport
            If OPDMod & #OPDFollowBuildLayout
              OPDReport("       #PBBScan_HardMatch - 3")
            EndIf
          CompilerEndIf
          ProcedureReturn #PBBScan_HardMatch
        EndIf
      EndIf
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowBuildLayout
          If ScanResult = #PBBScan_HardMatch
            OPDReport("       #PBBScan_HardMatch - 4")
          ElseIf ScanResult = #PBBScan_SharedRedirect
            OPDReport("       #PBBScan_SharedRedirect - 4")
          Else
            OPDReport("       #PBBScan_None - 4")
          EndIf
        EndIf
      CompilerEndIf
    Next
    ;
    FIPS\FIPSContinue = #True
  Wend
  ;
  ProcedureReturn #PBBScan_None
  ;
EndProcedure
;
CompilerIf #PBBReport
  Procedure OPDReportSearchDetails(Header$, *Item.SearchDetails, IncludePosition = #False)
    ;
    ; Role: Outputs a standardized diagnostic dump of '*Item' when #OPDFollowBuildLayout is active.
    ;
    If OPDMod & #OPDFollowBuildLayout
      OPDReport(Header$ + PBBTypeNames$(*Item\TypeOfElement) + " " + *Item\ElementName$)
      OPDReport("       Species: " + *Item\ElementSpecies$ + " " + *Item\ElementSubSpecies$)
      OPDReport("       ContextLimit: " + Str(*Item\StartingContextLimit) + " / " + Str(*Item\EndingContextLimit))
      OPDReport("       ParentTypeAndName$: " + #DOUBLEQUOTE$ + *Item\ParentTypeAndName$ + #DOUBLEQUOTE$)
      If IncludePosition
        OPDReport("       ElementPosition: " + Str(*Item\ElementPosition))
      EndIf
    EndIf
  EndProcedure
CompilerEndIf
;
Procedure GetInfosFromContext(*Filter.SearchDetails, FileName$, PosInCode)
  ;
  ; ROLE & DESCRIPTION:
  ; Resolves the concrete semantic type, species, and structural scope of an expression
  ; at 'PosInCode' within its immediate environment, guided by the incoming traits in '*Filter'.
  ;
  ; - Target Source: The search is first conducted in the code designated by 'FileName$'
  ;                  and then across the broader project if necessary.
  ; - Filter Scope: The content of '*Filter' can range from a simple raw identifier to a
  ;                 comprehensive specification of expected traits.
  ; - Precedence: The returned metadata reflects the closest valid definition matching '*Filter',
  ;               while strictly conforming to the physical reality discovered at 'PosInCode'.
  ;
  ; METHODOLOGY & ARCHITECTURE:
  ; 1. Target Proof Extraction (Context Peek):
  ;    Inspects 'PosInCode' and its immediate AST line via BuildLayout() with '#PBB_Action_CollectAll'
  ;    to extract undeniable syntactic traits (such as parentheses, dot types, or structure paths).
  ; 2. Short-Circuit Evaluation:
  ;    If the physical occurrence at 'PosInCode' is already a fully qualified declaration
  ;    (e.g., explicit declarator or procedure header), execution short-circuits immediately.
  ; 3. Working Filter Isolation:
  ;    Clones '*Filter' into a local 'WorkingFilter' so that exploratory mutations do not contaminate
  ;    the caller's original search parameters until definitive proof is established.
  ; 4. Local Procedural Scanning:
  ;    Scans the encompassing Procedure or Macro boundaries via ScanContextRange() to resolve
  ;    lexically isolated definitions (Protected, Static, Param) without cross-procedural leakage.
  ; 5. Global Scope Scanning:
  ;    Respecting PureBasic's single-pass rule: scans from line 0 up to the start of the current local
  ;    scope to identify outer definitions (Global, Define, Threaded) or resolve Shared redirections.
  ; 6. Final Reconciliation:
  ;    Applies fallback traits to promote qualified candidates without overwriting usage-level
  ;    context boundaries, ensuring deterministic classification.
  ;
  Shared ContextFoundExpression()
  ;
  Protected *SearchString.String
  Protected *SearchString_LCase.String
  Protected SStringLength, StartLimit, EndLimit, QuoteEndPos, QuoteStartPos, KwEndPos, KwStartPos
  Protected *QuoteScan.Character
  Protected PrecedingKw$
  ;
  ; Guard against empty queries or missing files:
  If *Filter = 0 Or *Filter\ElementName$ = "" Or FileName$ = ""
    ProcedureReturn 0
  EndIf
  ;
  *SearchString       = GetPointedCodeFromFile(FileName$)
  *SearchString_LCase = GetPointedCodeFromFile(FileName$, #True)
  If *SearchString = 0 Or *SearchString_LCase = 0 : ProcedureReturn 0 : EndIf
  ;
  Protected WasInDeadZone = IsPosInDeadZone(FileName$, PosInCode)
  ;
  SStringLength = Len(*SearchString\s)
  If WasInDeadZone
    *Filter\InDeadZone        = #True
    *Filter\SearchInDeadZones = #True
  EndIf
  ;
  ; Detect comments and string literals before PBBasicKeywordMap() or SearchPurifiedOnly exits:
  If PosInCode >= 1 And PosInCode <= SStringLength
    Select CheckCodeLineContext(*SearchString, PosInCode, #True, #False, @QuoteEndPos)
      Case #CodeCtx_Comment
        *Filter\SearchInCom = #True
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowBuildLayout
            OPDReport("*** GetInfosFromContext: PosInCode is inside a comment -> SearchInCom = #True")
          EndIf
        CompilerEndIf
        ProcedureReturn 0
      Case #CodeCtx_Quote
        ; If the cursor sits inside a quoted '.lib' target following 'Import' or 'ImportC',
        ; promote '*Filter' to the full quoted target so Section 2 resolves the #PBBImport element:
        PrecedingKw$ = ""
        If QuoteEndPos >= 7 And FastMid(*SearchString_LCase\s, QuoteEndPos - 4, 5) = ".lib" + #DOUBLEQUOTE$
          *QuoteScan = PosToChar(@*SearchString\s, QuoteEndPos - 1)
          While *QuoteScan > @*SearchString\s And *QuoteScan\c <> #CR And *QuoteScan\c <> #LF
            If *QuoteScan\c = #DOUBLEQUOTE And
               (*QuoteScan = @*SearchString\s Or PeekC(*QuoteScan - SizeOf(Character)) <> '\')
              Break
            EndIf
            *QuoteScan - SizeOf(Character)
          Wend
          If *QuoteScan\c = #DOUBLEQUOTE
            If *QuoteScan > @*SearchString\s And PeekC(*QuoteScan - SizeOf(Character)) = '~'
              *QuoteScan - SizeOf(Character)
            EndIf
            QuoteStartPos = ((*QuoteScan - @*SearchString\s) / SizeOf(Character)) + 1
            KwEndPos      = QuoteStartPos - 1
            While KwEndPos > 0 And (FastMid(*SearchString\s, KwEndPos, 1) = " " Or FastMid(*SearchString\s, KwEndPos, 1) = #TAB$)
              KwEndPos - 1
            Wend
            KwStartPos = KwEndPos
            While KwStartPos > 0 And IsPBSeparator(Asc(FastMid(*SearchString\s, KwStartPos, 1))) = #False
              KwStartPos - 1
            Wend
            PrecedingKw$ = FastMid(*SearchString_LCase\s, KwStartPos + 1, KwEndPos - KwStartPos)
          EndIf
        EndIf
        If PrecedingKw$ = "import" Or PrecedingKw$ = "importc"
          PosInCode = QuoteStartPos
          PrepareSearchDetailsElement(*Filter, FastMid(*SearchString\s, QuoteStartPos, QuoteEndPos - QuoteStartPos + 1))
        Else
          *Filter\SearchInQuotes = #True
          CompilerIf #PBBReport
            If OPDMod & #OPDFollowBuildLayout
              OPDReport("*** GetInfosFromContext: PosInCode is inside a string literal -> SearchInQuotes = #True")
            EndIf
          CompilerEndIf
          ProcedureReturn 0
        EndIf
    EndSelect
  EndIf
  ;
  If *Filter\SearchPurifiedOnly
    ProcedureReturn 0
  EndIf
  ;
  ; Initialize an isolated working filter by cloning the caller's initial hypothesis:
  Protected WorkingFilter.SearchDetails
  Protected HasExplicitModule = #False, *PrevChar.Character
  CopyStructure(*Filter, @WorkingFilter, SearchDetails)
  ;
  ; Detect whether PosInCode is explicitly qualified by a preceding '::' operator:
  If PosInCode > 2 And PosInCode <= SStringLength
    *PrevChar = PosToChar(@*SearchString\s, PosInCode - 1)
    While *PrevChar > @*SearchString\s And (*PrevChar\c = #Space Or *PrevChar\c = #NonBreakableSpace)
      *PrevChar - SizeOf(Character)
    Wend
    HasExplicitModule = Bool(*PrevChar > @*SearchString\s And *PrevChar\c = ':' And PeekC(*PrevChar - SizeOf(Character)) = ':')
  EndIf
  ;
  ; -------------------------------------------------------------------------
  ; 1. SEARCH IN PUREBASIC BASIC VOCABULARY
  ; Role: Check if *Filter\ElementName$ matches a PureBasic keyword.
  ; -------------------------------------------------------------------------
  ;{
  If HasExplicitModule = #False And IsSearchDetailsCompleted(WorkingFilter, #False) = #False And WorkingFilter\TypeOfElement = 0
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowBuildLayout
        OPDReport(#CR$ + "----- GetInfosFromContext: Search in PureBasic vocabulary -----")
        OPDReport("FileName$ = " + FileName$)
      EndIf
    CompilerEndIf
    ;
    ; Intercept reserved language keywords (e.g., 'If', 'While', 'ProcedureReturn'):
    If FindMapElement(PBBasicKeywordMap(), LCase(WorkingFilter\ElementName$))
      *Filter\ElementName$         = PBBasicKeywordMap()
      *Filter\ElementPurifiedName$ = MapKey(PBBasicKeywordMap())
      *Filter\TypeOfElement        = #PBBBasicKeyword
      *Filter\HasParentheses       = 0
      CompilerIf #PBBReport
        OPDReportSearchDetails("*** DEFINITIVE RESULT from Basic Vocabulary check: ", *Filter)
      CompilerEndIf
      ProcedureReturn *Filter\TypeOfElement
    EndIf
  EndIf ; EndIf Basic Vocabulary check (Section 1)
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowBuildLayout
      OPDReport(#CR$ + "-- NO DEFINITIVE RESULT from BASIC Vocabulary Scan --")
    EndIf
  CompilerEndIf
  ;}
  ; -------------------------------------------------------------------------
  ; 2. CONTEXT SCAN
  ; Role: Evaluates the exact target location before scanning to capture indisputable
  ;       structural proofs from the immediate environment.
  ; -------------------------------------------------------------------------
  ;{
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowBuildLayout
      OPDReport(#CR$ + "----- GetInfosFromContext: Context scan -----")
      OPDReport("FileName$ = " + FileName$)
    EndIf
  CompilerEndIf
  ;
  ; #PBB_Action_CollectAll forces data extraction, bypassing premature compatibility rejections;
  ; #PBB_Action_Reset clears the Shared list for this new scan cycle:
  BuildLayout(@WorkingFilter, FileName$, *SearchString, *SearchString_LCase,
              SStringLength, PosInCode, 0, #PBB_Action_Collect | #PBB_Action_CollectAll | #PBB_Action_Reset)
  ;
  If FirstElement(ContextFoundExpression()) ; Empty when PosInCode is inside a comment or quoted string.
    ; If the inspected token is an explicit declaration, terminate immediately without scanning:
    If IsSearchDetailsCompleted(ContextFoundExpression(), #False)
      CopyStructure(@ContextFoundExpression(), *Filter, SearchDetails)
      If WasInDeadZone
        *Filter\InDeadZone        = #True
        *Filter\SearchInDeadZones = #True
      EndIf
      CompilerIf #PBBReport
        OPDReportSearchDetails("*** DEFINITIVE RESULT from context Scan: ", *Filter)
      CompilerEndIf
      ProcedureReturn *Filter\TypeOfElement
    Else
      ; Enrich the working hypothesis with local AST proofs (e.g., dot types or dereferences):
      If InheritStrongSearchTraits(@ContextFoundExpression(), @WorkingFilter) = #False
        ; Context proof contradicts the initial *Filter hypothesis; retain the newly discovered context traits:
        CopyStructure(@ContextFoundExpression(), @WorkingFilter, SearchDetails)
        WorkingFilter\InDeadZone        = WasInDeadZone
        WorkingFilter\SearchInDeadZones = Bool(WasInDeadZone Or *Filter\SearchInDeadZones)
      EndIf
    EndIf
  Else
    ; No AST context proof could be extracted at PosInCode: exit so downstream GetInfosFromData()
    ; resolves the identifier across all project elements without single-pass position constraints.
    ProcedureReturn 0
  EndIf ; EndIf Context Scan (Section 2)
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowBuildLayout
      OPDReport(#CR$ + "-- NO DEFINITIVE RESULT from context Scan --")
    EndIf
  CompilerEndIf
  ;}
  ; -------------------------------------------------------------------------
  ; 3. SEARCH IN PUREBASIC NATIVE FUNCTIONS
  ; Role: Check if WorkingFilter\ElementPurifiedName$ matches a PureBasic function name.
  ; -------------------------------------------------------------------------
  ;{
  If HasExplicitModule = #False And
     IsSearchDetailsCompleted(WorkingFilter, #False) = #False And
     WorkingFilter\TypeOfElement = 0 And
     WorkingFilter\HasParentheses = 1
    ;
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowBuildLayout
        OPDReport(#CR$ + "----- GetInfosFromContext: Search in PureBasic function names -----")
        OPDReport("        FileName$ = " + FileName$)
        OPDReport("        PosInCode = " + Str(PosInCode))
      EndIf
    CompilerEndIf
    ;
    If FindMapElement(PBNativeFunctions(), WorkingFilter\ElementPurifiedName$)
      ; Determine whether a prior Macro declaration in the compilation timeline shadows this native function:
      Protected MacroOverridePos, MacroOverrideFile.String
      *Filter\ElementName$         = PBNativeFunctions() + "()"
      *Filter\ElementPurifiedName$ = MapKey(PBNativeFunctions())
      *Filter\TypeOfElement        = #PBBNativeFunction
      *Filter\HasParentheses       = 1
      *Filter\ModuleName$          = ""
      MacroOverridePos             = CheckIfMacroOverridesNative(*Filter\ElementPurifiedName$, FileName$, PosInCode, @MacroOverrideFile)
      If MacroOverridePos > 0
        *Filter\TypeOfElement      = #PBBMacro
        *Filter\ElementPosition    = MacroOverridePos
        *Filter\ElementFile$       = MacroOverrideFile\s
        *Filter\ModuleName$        = CheckIfModuleIsUsed("", FileName$, PosInCode, *Filter\ElementPurifiedName$, WorkingFilter\ModuleName$)
        If *Filter\ModuleName$ = ""
          *Filter\ModuleName$      = WorkingFilter\ModuleName$
        EndIf
      EndIf
      CompilerIf #PBBReport
        OPDReportSearchDetails("*** DEFINITIVE RESULT from Function Vocabulary check: ", *Filter)
      CompilerEndIf
      ProcedureReturn *Filter\TypeOfElement
    EndIf
  EndIf ; EndIf Native Functions check (Section 3)
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowBuildLayout
      OPDReport(#CR$ + "-- NO DEFINITIVE RESULT from FUNCTIONS Vocabulary Scan --")
    EndIf
  CompilerEndIf
  ;}
  ; -------------------------------------------------------------------------
  ; 4. LOCAL SCAN
  ; Role: Computes scope boundaries and checks for a match in the encompassing
  ;       procedure or macro (if any).
  ; -------------------------------------------------------------------------
  ;{
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowBuildLayout
      OPDReport(#CR$ + "----- GetInfosFromContext: Local Scan -----  PosInCode = " + Str(PosInCode))
    EndIf
  CompilerEndIf
  ;
  Protected EncompassingElementData.EEDStruct
  ;
  ; Locate the containing procedural frame to restrict local scanning.
  ; Omit Init = #True so the cache populated for FileName$ during Section 2 (Context Scan) is reused directly:
  GetEncompassingLimitsAndName(#PBBProcedure, FileName$, PosInCode, *SearchString, *SearchString_LCase, SStringLength, EncompassingElementData)
  If EncompassingElementData\EndingPos = 0
    GetEncompassingLimitsAndName(#PBBMacro, FileName$, PosInCode, *SearchString, *SearchString_LCase, SStringLength, EncompassingElementData)
  EndIf
  ;
  If EncompassingElementData\EndingPos
    StartLimit = EncompassingElementData\StartingPos
    EndLimit   = EncompassingElementData\EndingPos
  Else
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowBuildLayout
        OPDReport(#CR$ + "*** Encompassing element not found for local scan.")
      EndIf
    CompilerEndIf
  EndIf
  ;
  Protected BestMatch.SearchDetails
  Protected HasFallbackLocal, LocalScanStatus
  ;
  ; Seed BestMatch with usage coordinates only for variables or untyped searches to allow
  ; chronological evaluation of Protected statements; non-variables start with clean coordinates:
  If WorkingFilter\TypeOfElement = 0 Or WorkingFilter\TypeOfElement = #PBBVariable
    CopyStructure(@WorkingFilter, @BestMatch, SearchDetails)
  Else
    PrepareSearchDetailsElement(@BestMatch, WorkingFilter\ElementName$)
  EndIf
  ;
  If StartLimit > 0 And EndLimit > 0
    If WorkingFilter\TypeOfElement = 0 Or
       WorkingFilter\TypeOfElement = #PBBVariable Or
       WorkingFilter\TypeOfElement = #PBBField
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowBuildLayout
          OPDReport(#CR$ + "Local Scan: " + ReplaceString(FastMid(*SearchString\s, StartLimit, 40), #CR$, " "))
        EndIf
      CompilerEndIf
      ;
      LocalScanStatus = ScanContextRange(@WorkingFilter, @BestMatch, FileName$,
                                         StartLimit, EndLimit,
                                         #True, #True, @HasFallbackLocal)
      ;
      ; Only hard local declarations (Protected, Static, Param) terminate the search immediately.
      ; 'Shared' declarations return #PBBScan_SharedRedirect to fall through to the global scan:
      If LocalScanStatus = #PBBScan_HardMatch And IsSearchDetailsCompleted(BestMatch, #False)
        CopyStructure(@BestMatch, *Filter, SearchDetails)
        CompilerIf #PBBReport
          OPDReportSearchDetails("*** DEFINITIVE RESULT from local Scan: ", *Filter)
        CompilerEndIf
        ProcedureReturn *Filter\TypeOfElement
      EndIf
      ;
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowBuildLayout
          OPDReport(#CR$ + "-- NO DEFINITIVE RESULT from local Scan --")
        EndIf
      CompilerEndIf
    EndIf ; EndIf TypeOfElement eligible for local scan
  EndIf ; EndIf StartLimit > 0 And EndLimit > 0 (Local Scan)
  ;}
  ; -------------------------------------------------------------------------
  ; 5. GLOBAL SCAN
  ; If no absolute local match was found, trigger the global scan
  ; (handles both #PBBScan_None and #PBBScan_SharedRedirect).
  ; -------------------------------------------------------------------------
  ;{
  Protected HasFallbackGlobal
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowBuildLayout
      OPDReport(#CR$ + "----- GetInfosFromContext: Global Scan -----  PosInCode = " + Str(PosInCode))
    EndIf
  CompilerEndIf
  ;
  ; Clear procedural boundaries from working filter so global scan is not rejected as out-of-context:
  WorkingFilter\StartingContextLimit = 0
  WorkingFilter\EndingContextLimit   = 0
  ;
  ; When an unqualified token inside a procedure was not shadowed by a local variable, resolve any active 'UseModule' import:
  If StartLimit > 0 And HasExplicitModule = #False And WorkingFilter\TypeOfElement <> #PBBLabel And WorkingFilter\TypeOfElement <> #PBBField
    Protected ActiveUseMod$ = CheckIfModuleIsUsed("", FileName$, PosInCode, WorkingFilter\ElementPurifiedName$, WorkingFilter\ModuleName$)
    If ActiveUseMod$ <> ""
      WorkingFilter\ModuleName$ = ActiveUseMod$
      BestMatch\ModuleName$     = ActiveUseMod$
    EndIf
  EndIf
  ;
  ; Only accept implicit declarations in global scan if at program root or via Shared redirect:
  Protected GlobalAcceptImplicit = Bool(StartLimit = 0 Or LocalScanStatus = #PBBScan_SharedRedirect)
  ;
  ; Stop the global scan just before the current local scope begins to respect PB's single-pass rule:
  If StartLimit
    EndLimit = StartLimit - 1
  Else
    EndLimit = PosInCode
  EndIf
  ;
  If ScanContextRange(@WorkingFilter, @BestMatch, FileName$,
                      0, EndLimit,
                      #False, GlobalAcceptImplicit, @HasFallbackGlobal) <> #PBBScan_None Or (StartLimit = 0 And HasFallbackGlobal)
    If IsSearchDetailsCompleted(BestMatch, #True)
      CopyStructure(@BestMatch, *Filter, SearchDetails)
      CompilerIf #PBBReport
        OPDReportSearchDetails("*** RESULT from Global Scan: ", *Filter, #True)
      CompilerEndIf
      ProcedureReturn *Filter\TypeOfElement
    EndIf
  EndIf ; EndIf Global Scan match (Section 5)
  ;}
  ; -------------------------------------------------------------------------
  ; 6. RECONCILIATION
  ; -------------------------------------------------------------------------
  ;{
  ; Fallback promotion: validate local soft matches or preserve the physical context
  ; extracted at PosInCode (position, file, parentheses, and module qualification)
  ; so downstream searches do not lose scope constraints when a private target is rejected:
  If IsSearchDetailsCompleted(BestMatch, #True) Or BestMatch\ElementPosition > 0
    CopyStructure(@BestMatch, *Filter, SearchDetails)
  EndIf
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowBuildLayout
      OPDReport("*** NO RESULT from Global Scan ***")
      OPDReportSearchDetails("       Default Return: ", *Filter, #True)
    EndIf
  CompilerEndIf
  ;
  ProcedureReturn *Filter\TypeOfElement
  ;}
EndProcedure
;
; **********************************************************************************
;
;-                    Project-wide sequence match iterator
;
; **********************************************************************************
;
Procedure FindNextMatchInProject(*LastSearch.SearchDetails, *FIPS.FIPSPosStruct, Reset = #False,
                                 *LayoutDefinition.String = 0, *MatchLen.Integer = 0, TimeLimit = 0)
  ;
  ; Role: Advances '*FIPS' to the next valid occurrence matching '*LastSearch' across the project
  ;       (or within the active local scope), validating multi-token sequences and AST scope compatibility
  ;       via BuildLayout(). Shared by SetListFoundInFiles() and GUI_FR_ExecuteReplacement().
  ; Output:
  ;   - Returns the 1-based match position (> 0) when a valid occurrence is found (and populates
  ;     '*MatchLen\i' with its exact character span in the file buffer if '*MatchLen' is provided).
  ;   - Returns 0 when the compilation sequence traversal is completely exhausted.
  ;   - Returns -1 if 'TimeLimit' (> 0) is exceeded before a valid match is reached.
  ;
  ; Retain search configuration across sequential calls within the same search pass:
  Static CheckMode, CompleteExpressionToSearch$, ExpressionToSearch$
  ;
  Protected FoundPos, pos, PartIndex, Uncomplete, LineCtx, PartOfExp$
  ;
  If *LastSearch = 0 Or *FIPS = 0
    ProcedureReturn 0
  EndIf
  ;
  ; 1. Initialize iterator state, scope boundaries, and filter flags on reset
  If Reset
    ClearStructure(*FIPS, FIPSPosStruct)
    BuildLayout(0, "", 0, 0, 0, 0, 0, #PBB_Action_Reset)
    ;
    ; Restrict to #FIPSMode_LocalOnly only when comments, string literals, and global dead zones are excluded
    *FIPS\IncludeDeadZones = *LastSearch\SearchInDeadZones
    If *LastSearch\ElementFile$ <> "" And *LastSearch\EndingContextLimit > 0 And
       *LastSearch\SearchInCom = #False And *LastSearch\SearchInQuotes = #False
      *FIPS\ScanMode  = #FIPSMode_LocalOnly
      *FIPS\FileName$ = *LastSearch\ElementFile$
      *FIPS\PosInCode = *LastSearch\StartingContextLimit
      *FIPS\EndLimit  = *LastSearch\EndingContextLimit
    Else
      *FIPS\ScanMode  = #FIPSMode_Global
    EndIf
    ;
    CheckMode = #CheckOneKeyWordOnly
    CompleteExpressionToSearch$ = *LastSearch\ElementPurifiedName$
    If Left(CompleteExpressionToSearch$, 1) = #PBBJoker$
      CompleteExpressionToSearch$ = FastMid(CompleteExpressionToSearch$, 2)
    Else
      CheckMode | #CheckLeft
    EndIf
    If Right(CompleteExpressionToSearch$, 1) = #PBBJoker$
      CompleteExpressionToSearch$ = Left(CompleteExpressionToSearch$, Len(CompleteExpressionToSearch$) - 1)
    Else
      CheckMode | #CheckRight
    EndIf
    ;
    If *LastSearch\NotAnElement
      ExpressionToSearch$ = StringField(CompleteExpressionToSearch$, 1, #Space$)
    Else
      ExpressionToSearch$ = CompleteExpressionToSearch$
    EndIf
    If *LastSearch\NotAnElement And IsPBSeparator(Asc(Right(ExpressionToSearch$, 1)))
      ; First token is a standalone separator symbol: disable word-boundary checks for it
      CheckMode & ~(#CheckLeft | #CheckRight)
    EndIf
    ;
    If *LastSearch\SearchInCom = #False
      CheckMode | #CheckCom
    EndIf
    If *LastSearch\SearchInQuotes = #False
      CheckMode | #CheckQuote
    EndIf
  EndIf
  ;
  ; 2. Traverse sequences until a valid match is confirmed or the time slice expires
  Repeat
    FoundPos = FindInWholeProjectSequences(ExpressionToSearch$, *FIPS, CheckMode)
    If FoundPos = 0
      ProcedureReturn 0
    EndIf
    ;
    *FIPS\FIPSContinue = #True
    Uncomplete         = #False
    ;
    If *LastSearch\NotAnElement And ExpressionToSearch$ <> CompleteExpressionToSearch$
      ; Multi-token complex expression: verify that all subsequent tokens match in sequence,
      ; ignoring arbitrary intermediate spaces or line breaks in the source code.
      PartOfExp$ = ExpressionToSearch$
      PartIndex  = 1
      pos        = FoundPos
      Repeat
        pos + Len(PartOfExp$)
        PartIndex + 1
        PartOfExp$ = StringField(CompleteExpressionToSearch$, PartIndex, #Space$)
        If PartOfExp$
          While FastMid(*FIPS\SearchString_LCase\s, pos, 1) = #Space$ Or
                FastMid(*FIPS\SearchString_LCase\s, pos, 1) = #CR$
            pos + 1
          Wend
          If FastMid(*FIPS\SearchString_LCase\s, pos, Len(PartOfExp$)) <> PartOfExp$ Or
             (IsPBSeparator(Asc(Right(PartOfExp$, 1))) = #False And
              IsPBSeparator(Asc(FastMid(*FIPS\SearchString_LCase\s, pos + Len(PartOfExp$), 1))) = #False)
            Uncomplete = #True
            Break
          EndIf
        EndIf
      Until PartOfExp$ = ""
    Else
      pos = FoundPos + Len(CompleteExpressionToSearch$)
    EndIf
    ;
    If Uncomplete = #False And *LastSearch\EndingContextLimit > 0
      If *FIPS\FileName$ <> *LastSearch\ElementFile$ Or
         *FIPS\PosInCode < *LastSearch\StartingContextLimit Or
         *FIPS\PosInCode > *LastSearch\EndingContextLimit
        LineCtx = CheckCodeLineContext(*FIPS\SearchString, *FIPS\PosInCode, #True)
        If (LineCtx = #CodeCtx_Comment And *LastSearch\SearchInCom) = #False And
           (LineCtx = #CodeCtx_Quote And *LastSearch\SearchInQuotes) = #False
          Uncomplete = #True
        EndIf
      EndIf
    EndIf
    ;
    If Uncomplete = #False
      If BuildLayout(*LastSearch, *FIPS\FileName$, *FIPS\SearchString, *FIPS\SearchString_LCase,
                     *FIPS\SegmentEnd, *FIPS\PosInCode, *LayoutDefinition, #PBB_Action_Layout) <> #False
        If *MatchLen
          *MatchLen\i = pos - FoundPos
        EndIf
        ProcedureReturn FoundPos
      EndIf
    EndIf
    ;
    If TimeLimit > 0 And ElapsedMilliseconds() > TimeLimit
      ProcedureReturn -1
    EndIf
  ForEver
EndProcedure
;
; **********************************************************************************
;
;-                     Homonym collection & scope helpers
;
; **********************************************************************************
;
Define NewList SearchMenuChoices.SearchDetails()
;
Procedure FillFoundElement(Map SearchingMap.s(), Prefix.s,
                           TypeOfElement, HasParentheses, TypeOfSearch)
  ;
  ; Role: Populates 'SearchMenuChoices()' in a single pass with entries from 'SearchingMap()'
  ;       matching 'Prefix' (exact, prefix, or wildcard via #PBBJoker$), stopping immediately
  ;       once '#PBB_MaxNbrOfFound' is reached.
  ; Parameters:
  ;   TypeOfSearch - Bitwise execution flags (#PBBSearchPrefix enables prefix matching;
  ;                  #PBBUserWillChoose collects all matching candidates instead of stopping at the first match).
  ; Output: Returns the exact number of entries added to 'SearchMenuChoices()'.
  ;
  Shared SearchMenuChoices()
  Protected PrefixLower$ = LCase(Prefix)
  Protected PrefixLen    = Len(Prefix)
  Protected HasJoker     = Bool(FindString(PrefixLower$, #PBBJoker$) > 0)
  Protected AddedCount   = 0
  Protected Keyword$
  ;
  If PrefixLen = 0 Or ListSize(SearchMenuChoices()) >= #PBB_MaxNbrOfFound
    ProcedureReturn 0
  EndIf
  ;
  If HasJoker = #False And (TypeOfSearch & #PBBSearchPrefix) = 0
    ; Exact match lookup in O(1) when neither a wildcard nor the prefix flag is active:
    If FindMapElement(SearchingMap(), PrefixLower$)
      Keyword$ = SearchingMap()
      If Keyword$ <> ""
        AddElement(SearchMenuChoices())
        SearchMenuChoices()\ElementName$         = Keyword$
        SearchMenuChoices()\ElementPurifiedName$ = PurifyElementName(Keyword$)
        SearchMenuChoices()\TypeOfElement        = TypeOfElement
        SearchMenuChoices()\ElementSpecies$      = ""
        SearchMenuChoices()\HasParentheses       = HasParentheses
        AddedCount = 1
      EndIf
    EndIf
  Else
    ; Linear map scan supporting both wildcard patterns (#PBBJoker$) and prefix searches:
    ForEach SearchingMap()
      If (HasJoker And CompareExpressionsWithJoker(PrefixLower$, MapKey(SearchingMap()))) Or
         (HasJoker = #False And Left(MapKey(SearchingMap()), PrefixLen) = PrefixLower$)
        Keyword$ = SearchingMap()
        If Keyword$ <> ""
          AddElement(SearchMenuChoices())
          SearchMenuChoices()\ElementName$         = Keyword$
          SearchMenuChoices()\ElementPurifiedName$ = PurifyElementName(Keyword$)
          SearchMenuChoices()\TypeOfElement        = TypeOfElement
          SearchMenuChoices()\ElementSpecies$      = ""
          SearchMenuChoices()\HasParentheses       = HasParentheses
          AddedCount + 1
          If ListSize(SearchMenuChoices()) >= #PBB_MaxNbrOfFound Or (TypeOfSearch & #PBBUserWillChoose) = 0
            Break
          EndIf
        EndIf
      EndIf
    Next
  EndIf
  ;
  ProcedureReturn AddedCount
EndProcedure
;
Procedure AddCandidateToSearchChoices(*Candidate.SearchDetails, AllGroupKey$ = "", CoveredScopesCount = 0)
  ;
  ; Role: Appends a candidate SearchDetails record to SearchMenuChoices(), deduplicating
  ; same-scope variable entries or converting the record into a synthetic multi-scope
  ; choice (with FixedID reset to 0) when AllGroupKey$ is specified.
  ;
  Shared SearchMenuChoices()
  ;
  If AllGroupKey$ = "" And *Candidate\TypeOfElement = #PBBVariable
    ForEach SearchMenuChoices()
      If SearchMenuChoices()\DontCheckPosition = #False And
         SearchMenuChoices()\TypeOfElement = #PBBVariable And
         SearchMenuChoices()\ElementPurifiedName$ = *Candidate\ElementPurifiedName$ And
         SearchMenuChoices()\ElementFile$ = *Candidate\ElementFile$ And
         SearchMenuChoices()\StartingContextLimit = *Candidate\StartingContextLimit
        ProcedureReturn
      EndIf
    Next
  EndIf
  ;
  LastElement(SearchMenuChoices())
  AddElement(SearchMenuChoices())
  CopyStructure(*Candidate, @SearchMenuChoices(), SearchDetails)
  ;
  If AllGroupKey$ <> ""
    SearchMenuChoices()\FixedID                = 0
    SearchMenuChoices()\ParentTypeAndName$     = AllGroupKey$
    SearchMenuChoices()\VarDotType$            = ""
    SearchMenuChoices()\ModuleName$            = ""
    SearchMenuChoices()\ElementFile$           = ""
    SearchMenuChoices()\StartingContextLimit   = 0
    SearchMenuChoices()\EndingContextLimit     = 0
    SearchMenuChoices()\ElementPosition        = 0
    SearchMenuChoices()\DontCheckPosition      = #True
    SearchMenuChoices()\OtherScopesCount       = CoveredScopesCount
  EndIf
EndProcedure
;
Procedure.s CollectMapCandidateIDs(Map SourceMap.s(), SearchPattern$, HasJoker)
  ;
  ; Role: Extracts all comma-separated FixedID strings from SourceMap() matching SearchPattern$,
  ;       supporting both O(1) exact hash lookups and linear wildcard scans (#PBBJoker$).
  ;
  Protected ResultIDs$ ; Accumulates comma-separated FixedIDs across matching map keys.
  ;
  If HasJoker
    ForEach SourceMap()
      If CompareExpressionsWithJoker(SearchPattern$, MapKey(SourceMap()))
        ResultIDs$ + SourceMap()
      EndIf
    Next
  ElseIf FindMapElement(SourceMap(), SearchPattern$)
    ResultIDs$ = SourceMap()
  EndIf
  ;
  ProcedureReturn ResultIDs$
EndProcedure
;
Procedure AppendSyntheticScopeChoice(TypeOfElement, SpeciesFilter$, ContainerType$, RequireModule = #False)
  ;
  ; Role: Scans concrete candidates in SearchMenuChoices() matching TypeOfElement, SpeciesFilter$,
  ;       and ContainerType$. When at least two distinct scopes exist, appends a synthetic "All<ContainerType>s"
  ;       entry at the bottom of SearchMenuChoices().
  ;
  Shared SearchMenuChoices()
  Protected ContainerPrefix$ = ContainerType$ + #Space$
  Protected ScopeCount, HasModuleCandidate
  Protected *LastScopeChoice.SearchDetails
  Protected TemplateDetails.SearchDetails ; Working copy converted into the synthetic entry.
  ;
  ForEach SearchMenuChoices()
    If SearchMenuChoices()\DontCheckPosition = #False And SearchMenuChoices()\TypeOfElement = TypeOfElement
      If RequireModule
        If SearchMenuChoices()\StartingContextLimit = 0
          ScopeCount + 1
          If SearchMenuChoices()\ModuleName$ <> ""
            HasModuleCandidate = #True
          EndIf
          *LastScopeChoice = @SearchMenuChoices()
        EndIf
      ElseIf SpeciesFilter$ = "" Or SearchMenuChoices()\ElementSpecies$ = SpeciesFilter$
        If Left(SearchMenuChoices()\ParentTypeAndName$, Len(ContainerPrefix$)) = ContainerPrefix$ Or
           (ContainerType$ = PBBTypeNames$(#PBBProcedure) And SpeciesFilter$ = "Shared")
          ScopeCount + 1
          *LastScopeChoice = @SearchMenuChoices()
        EndIf
      EndIf
    EndIf
  Next
  ;
  If ScopeCount > 1 And *LastScopeChoice And (RequireModule = #False Or HasModuleCandidate)
    CopyStructure(*LastScopeChoice, @TemplateDetails, SearchDetails)
    AddCandidateToSearchChoices(@TemplateDetails, "All" + ContainerType$ + "s", ScopeCount)
  EndIf
EndProcedure
;
; **********************************************************************************
;
;-                    Interactive scope selection & lookup
;
; **********************************************************************************
;
Procedure SetSearchButtonState(ButtonMode = -1, IsEnabled = -1, CandidateCount = -1)
  ;
  ; Role: Synchronizes the operational mode (#PBBSearchBtn_Search or #PBBSearchBtn_LookUp),
  ;       localized label (including the "(N)" homonym counter), and enabled/disabled state
  ;       of the BSearchExpression button.
  ;       Passing -1 for any parameter preserves its previously recorded static state, allowing
  ;       language refreshes in GUI_SetPBBGadgetsTitles() to retain the active homonym count.
  ;
  ; Retain the button state across UI refreshes and language catalog reloads:
  Static SavedMode    = #PBBSearchBtn_Search
  Static SavedEnabled = #False
  Static SavedCount   = 0
  ;
  Protected ButtonLabel$ ; Final localized caption assigned to BSearchExpression.
  ;
  If ButtonMode <> -1
    SavedMode = ButtonMode
  EndIf
  If IsEnabled <> -1
    SavedEnabled = IsEnabled
  EndIf
  If CandidateCount <> -1
    SavedCount = CandidateCount
  EndIf
  ;
  SetGadgetData(GPBBGadgets\BSearchExpression, SavedMode)
  ;
  If SavedMode = #PBBSearchBtn_LookUp
    ButtonLabel$ = GetTextFromCatalogPB("LookUp")
    If SavedCount > 1
      ButtonLabel$ + " (" + Str(SavedCount) + ")"
    EndIf
  Else
    ButtonLabel$ = GetTextFromCatalogPB("Search")
  EndIf
  ;
  SetGadgetText(GPBBGadgets\BSearchExpression, ButtonLabel$)
  DisableGadget(GPBBGadgets\BSearchExpression, Bool(SavedEnabled = #False))
EndProcedure
;
Declare GetInfosFromData(*SearchParam.SearchDetails, TypeOfSearch = #PBBSearchExact)
;
; Forward declarations for procedures defined in SearchReplaceWindow.pb:
Declare.s GUI_FR_GetSearchExpression()
Declare   GUI_FR_RefreshScopes(*LastSearch.SearchDetails = 0, SelectedIndex = -1)
;
Procedure SelectSearchScopeChoice(ChoiceIndex, *LastSearch.SearchDetails, IsFromPopup = #False)
  ;
  ; Role: Applies the scope choice selected at 'ChoiceIndex' in SearchMenuChoices()
  ; (either from the disambiguation PopupMenu or from the Find/Replace Scopes ListViewGadget)
  ; to '*LastSearch' and schedules a refresh via ExecuteSearch.
  ; Cancels comment/string extensions when a specific or grouped scope is chosen from the popup,
  ; while letting the Find/Replace checkboxes govern when driven by the Find/Replace window.
  ;
  Shared SearchMenuChoices(), ExecuteSearch
  Protected SavedTotalCandidates, WasJokerSearch, SavedSearchInCom, SavedSearchInQuotes
  Protected SelectedChoice.SearchDetails
  ;
  If ChoiceIndex < 0 Or *LastSearch = 0
    ProcedureReturn
  EndIf
  ;
  If IsFromPopup And ChoiceIndex < ListSize(SearchMenuChoices())
    SavedSearchInCom    = #False
    SavedSearchInQuotes = #False
  ElseIf IsFromPopup = #False And FR_Gadgets\Window >= 0 And IsWindow(FR_Gadgets\Window)
    SavedSearchInCom    = Bool(GetGadgetState(FR_Gadgets\IgnoreComments) = 0)
    SavedSearchInQuotes = Bool(GetGadgetState(FR_Gadgets\IgnoreStrings) = 0)
  Else
    SavedSearchInCom    = *LastSearch\SearchInCom
    SavedSearchInQuotes = *LastSearch\SearchInQuotes
  EndIf
  ;
  If ChoiceIndex < ListSize(SearchMenuChoices())
    WasJokerSearch = Bool(FindString(*LastSearch\ElementPurifiedName$, #PBBJoker$) > 0)
    ;
    ; Restore fully resolved traits from the selected candidate
    SelectElement(SearchMenuChoices(), ChoiceIndex)
    CopyStructure(@SearchMenuChoices(), *LastSearch, SearchDetails)
    SetGadgetText(GPBBGadgets\SearchedExpression_gadget, *LastSearch\ElementName$)
    ;
    If IsFromPopup And WasJokerSearch
      ; Recompute SearchMenuChoices(), MultipleResults, and OtherScopesCount strictly for the selected element name
      CopyStructure(*LastSearch, @SelectedChoice, SearchDetails)
      GetInfosFromData(@SelectedChoice, #PBBSearchExact)
      *LastSearch\MultipleResults  = SelectedChoice\MultipleResults
      *LastSearch\OtherScopesCount = SelectedChoice\MultipleResults - 1
    EndIf
    ;
    If *LastSearch\DontCheckPosition
      ExecuteSearch = #PBBRefreshSearch | #PBBSearchInAllProcedures
    Else
      ExecuteSearch = #PBBRefreshSearch
    EndIf
  Else
    ; The trailing item maps to the raw text search option ('All')
    ExecuteSearch        = #PBBSearchRawText
    SavedTotalCandidates = *LastSearch\MultipleResults
    ;
    ForEach SearchMenuChoices()
      If SavedTotalCandidates = 0 And SearchMenuChoices()\MultipleResults > 0
        SavedTotalCandidates = SearchMenuChoices()\MultipleResults
      EndIf
      If Len(SearchMenuChoices()\ElementPurifiedName$) > Len(*LastSearch\ElementPurifiedName$)
        ExecuteSearch | #PBBSearchPrefix
      EndIf
    Next
    ;
    If FR_Gadgets\Window >= 0 And IsWindow(FR_Gadgets\Window)
      PrepareSearchDetailsElement(*LastSearch, GUI_FR_GetSearchExpression())
      SetGadgetText(GPBBGadgets\SearchedExpression_gadget, *LastSearch\ElementName$)
    Else
      PrepareSearchDetailsElement(*LastSearch, *LastSearch\ElementName$)
    EndIf
    *LastSearch\SearchPurifiedOnly = #True
    *LastSearch\SearchInDeadZones  = #True
    *LastSearch\MultipleResults    = SavedTotalCandidates
    *LastSearch\OtherScopesCount   = 0
    SetSearchButtonState(#PBBSearchBtn_LookUp, Bool(SavedTotalCandidates > 1), SavedTotalCandidates)
  EndIf
  ;
  *LastSearch\SearchInCom    = SavedSearchInCom
  *LastSearch\SearchInQuotes = SavedSearchInQuotes
EndProcedure
;
Procedure GetInfosFromData(*SearchParam.SearchDetails, TypeOfSearch = #PBBSearchExact)
  ;
  ; ROLE & 3-STEP ARCHITECTURE:
  ; 1. Unfiltered Homonym Collection:
  ;    Collects all project elements (ElementMapIndex), secondary declarations (\Continued),
  ;    formal procedure/macro parameters (ParamMapIndex), native functions, and PB keywords
  ;    matching the target name or wildcard pattern (#PBBJoker$) into SearchMenuChoices()
  ;    up to '#PBB_MaxNbrOfFound' without filtering by AreSearchDetailsCompatible().
  ; 2. Contextual Filtering:
  ;    Scans SearchMenuChoices() using AreSearchDetailsCompatible() to find the exact
  ;    context match while computing for each entry how many other scopes exist in the project.
  ; 3. Search Button State Synchronization:
  ;    Updates BSearchExpression immediately based on the total number of homonyms found:
  ;    - 0 results: "Search" (disabled)
  ;    - 1 result:  "LookUp" (disabled)
  ;    - >1 results (even if only 1 matched the context filter): "LookUp (N)" (enabled)
  ;
  Shared SearchMenuChoices()
  ;
  Protected ReturnValue = #PBBNotFound
  Protected CandidateIDs$, ParamCandidateIDs$, IDStr$, IDIndex, NbItem, KeptConcreteCount
  Protected CleanParams$, ParamToken$, ParamPurified$, ParamIdx
  Protected SearchPattern$, ElementNameLCase$ = *SearchParam\ElementPurifiedName$
  Protected HasJoker, TotalRealCandidates, CompatibleCount, FirstCompatibleIndex
  Protected *FirstCompatible.SearchDetails
  Protected SpeciesList$, CurrentSpecies$, SpeciesIdx, ContainerIdx
  Protected ItemDetails$, XMenu, YMenu
  Protected CandidateDetails.SearchDetails
  Protected EffectiveFilter.SearchDetails
  Protected NewList CEDList.CED_Structure()
  ;
  ClearList(SearchMenuChoices())
  ;
  If ElementNameLCase$
    ; Normalize prefix searches into right-wildcard patterns for unified matching
    If (TypeOfSearch & #PBBSearchPrefix) And Right(ElementNameLCase$, Len(#PBBJoker$)) <> #PBBJoker$
      SearchPattern$ = ElementNameLCase$ + #PBBJoker$
    Else
      SearchPattern$ = ElementNameLCase$
    EndIf
    HasJoker = Bool(FindString(SearchPattern$, #PBBJoker$) > 0)
    ;
    ; ======================================================================
    ; STEP 1: UNFILTERED HOMONYM COLLECTION INTO SearchMenuChoices()
    ; ======================================================================
    ;{
    ; 1A. Collect primary elements and unpack secondary declarations (\Continued)
    CandidateIDs$ = CollectMapCandidateIDs(ElementMapIndex(), SearchPattern$, HasJoker)
    While ListSize(SearchMenuChoices()) < #PBB_MaxNbrOfFound
      IDIndex + 1
      IDStr$ = StringField(CandidateIDs$, IDIndex, ",")
      If IDStr$ = "" : Break : EndIf
      If SelectPBBElement(Val(IDStr$))
        FillSearchDetailsFromElement(@ElementList(), @CandidateDetails)
        AddCandidateToSearchChoices(@CandidateDetails)
        If ElementList()\Type = #PBBVariable And ElementList()\Continued <> ""
          CED_StringToList(ElementList()\Continued, CEDList())
          ForEach CEDList()
            If ListSize(SearchMenuChoices()) >= #PBB_MaxNbrOfFound : Break 2 : EndIf
            FillSearchDetailsFromElement(@ElementList(), @CandidateDetails, @CEDList())
            AddCandidateToSearchChoices(@CandidateDetails)
          Next
        EndIf
      EndIf
    Wend
    ;
    ; 1B. Collect and unpack formal procedure/macro parameters from ParamMapIndex()
    ParamCandidateIDs$ = CollectMapCandidateIDs(ParamMapIndex(), SearchPattern$, HasJoker)
    IDIndex            = 0
    While ListSize(SearchMenuChoices()) < #PBB_MaxNbrOfFound
      IDIndex + 1
      IDStr$ = StringField(ParamCandidateIDs$, IDIndex, ",")
      If IDStr$ = "" : Break : EndIf
      If SelectPBBElement(Val(IDStr$)) And ElementList()\Parameters <> ""
        CleanParams$ = Trim(ElementList()\Parameters, #TAB$)
        ParamIdx     = 0
        While ListSize(SearchMenuChoices()) < #PBB_MaxNbrOfFound
          ParamIdx + 1
          ParamToken$ = StringField(CleanParams$, ParamIdx, #TAB$)
          If ParamToken$ = "" : Break : EndIf
          ParamPurified$ = PurifyElementName(ParamToken$)
          If ParamPurified$ <> "" And CompareExpressionsWithJoker(SearchPattern$, ParamPurified$)
            FillSearchDetailsFromElement(@ElementList(), @CandidateDetails, 0, ParamToken$)
            AddCandidateToSearchChoices(@CandidateDetails)
          EndIf
        Wend
      EndIf
    Wend
    ;
    ; 1C. Collect matching PureBasic native functions, keywords, and synthetic 'All...' scopes
    FillFoundElement(PBNativeFunctions(), SearchPattern$, #PBBNativeFunction, #True, TypeOfSearch)
    FillFoundElement(PBBasicKeywordMap(), SearchPattern$, #PBBBasicKeyword, #False, TypeOfSearch)
    If HasJoker = #False
      AppendSyntheticScopeChoice(#PBBField, "", PBBTypeNames$(#PBBStructure))
      AppendSyntheticScopeChoice(#PBBField, "", PBBTypeNames$(#PBBInterface))
      AppendSyntheticScopeChoice(#PBBConstant, "", PBBTypeNames$(#PBBEnumeration))
      SpeciesList$ = "Protected,Static,Define,Shared,Param"
      For SpeciesIdx = 1 To 5
        CurrentSpecies$ = StringField(SpeciesList$, SpeciesIdx, ",")
        AppendSyntheticScopeChoice(#PBBVariable, CurrentSpecies$, PBBTypeNames$(#PBBProcedure))
        AppendSyntheticScopeChoice(#PBBVariable, CurrentSpecies$, PBBTypeNames$(#PBBMacro))
      Next
      For ContainerIdx = #PBBProcedure To #PBBImported
        If ContainerIdx <> #PBBModule : AppendSyntheticScopeChoice(ContainerIdx, "", PBBTypeNames$(#PBBModule), #True) : EndIf
      Next
    EndIf
    ;}
    ;
    ; ======================================================================
    ; STEP 2: CONTEXTUAL MATCHING VIA AreSearchDetailsCompatible()
    ; ======================================================================
    ;{
    CopyStructure(*SearchParam, @EffectiveFilter, SearchDetails)
    EffectiveFilter\ElementPurifiedName$ = SearchPattern$
    FirstCompatibleIndex                 = ListSize(SearchMenuChoices())
    ;
    ForEach SearchMenuChoices()
      If SearchMenuChoices()\DontCheckPosition = #False
        If SearchMenuChoices()\ElementPurifiedName$ = ""
          SearchMenuChoices()\ElementPurifiedName$ = PurifyForSearch(SearchMenuChoices()\ElementName$)
        EndIf
        TotalRealCandidates + 1
        If AreSearchDetailsCompatible(@SearchMenuChoices(), @EffectiveFilter)
          CompatibleCount + 1
          If *FirstCompatible = 0
            *FirstCompatible     = @SearchMenuChoices()
            FirstCompatibleIndex = ListIndex(SearchMenuChoices())
          EndIf
        EndIf
      EndIf
    Next
    ;
    *SearchParam\MultipleResults = TotalRealCandidates
    If CompatibleCount = 0
      *SearchParam\OtherScopesCount = Bool(*SearchParam\SearchPurifiedOnly = #False) * TotalRealCandidates
    EndIf
    ;
    ForEach SearchMenuChoices()
      SearchMenuChoices()\MultipleResults = TotalRealCandidates
      If SearchMenuChoices()\DontCheckPosition
        SearchMenuChoices()\OtherScopesCount = TotalRealCandidates - SearchMenuChoices()\OtherScopesCount
      Else
        SearchMenuChoices()\OtherScopesCount = TotalRealCandidates - 1
      EndIf
    Next
    ;
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowSearchAndIdentification
        OPDReport("GetInfosFromData - " + Str(TotalRealCandidates) + " total homonyms (" + Str(CompatibleCount) +
                  " compatible) - Criteria: " + *SearchParam\ElementName$ + " " +
                  *SearchParam\ElementSpecies$ + " " + *SearchParam\ElementSubSpecies$)
      EndIf
    CompilerEndIf
    ;
    If CompatibleCount > 0
      If (TypeOfSearch & #PBBUserWillChoose) And CompatibleCount > 1 And
         ((FR_Gadgets\Window >= 0 And IsWindow(FR_Gadgets\Window)) Or
          ((TypeOfSearch & #PBBSearchFromIDE) And (EffectiveFilter\SearchInCom Or EffectiveFilter\SearchInQuotes)))
        ; Multiple scopes/homonyms when Find/Replace is open or from an IDE comment/string: select 'All' by default
        FirstCompatibleIndex = ListSize(SearchMenuChoices())
        SelectSearchScopeChoice(FirstCompatibleIndex, *SearchParam, #False)
        ReturnValue = #PBBNativeFunction
      ElseIf (FR_Gadgets\Window < 0 Or IsWindow(FR_Gadgets\Window) = #False) And
             (TypeOfSearch & #PBBUserWillChoose) And (CompatibleCount > 1 Or (*SearchParam\DontCheckPosition And CompatibleCount = 1))
        If IsMenu(GPBBGadgets\SearchResultMenu)
          FreeMenu(GPBBGadgets\SearchResultMenu)
        EndIf
        GPBBGadgets\SearchResultMenu = CreatePopupMenu(#PB_Any)
        ;
        ; Prune concrete candidates beyond #PBB_MaxNbrOfPopupLines while retaining synthetic 'All...' entries
        ; so menu item indices map 1-to-1 with SearchMenuChoices() elements in PBBrowserMainProc():
        If TotalRealCandidates > #PBB_MaxNbrOfPopupLines
          KeptConcreteCount = 0
          ForEach SearchMenuChoices()
            If SearchMenuChoices()\DontCheckPosition = #False
              KeptConcreteCount + 1
              If KeptConcreteCount > #PBB_MaxNbrOfPopupLines
                DeleteElement(SearchMenuChoices())
              EndIf
            EndIf
          Next
        EndIf
        ;
        NbItem = 0
        ForEach SearchMenuChoices()
          ItemDetails$ = ComputeExpressionDetails(SearchMenuChoices())
          MenuItem(#PBBMenu_SearchResultsOffset + NbItem, SearchMenuChoices()\ElementName$ + #SpecialSpace$ + ItemDetails$)
          NbItem + 1
        Next
        ;
        If TotalRealCandidates > #PBB_MaxNbrOfPopupLines
          MenuItem(#PBBMenu_SearchResultsOffset + NbItem, ReplaceString(GetTextFromCatalogPB("ResultsTruncated"), "%1", Str(#PBB_MaxNbrOfPopupLines)))
          DisableMenuItem(GPBBGadgets\SearchResultMenu, #PBBMenu_SearchResultsOffset + NbItem, #True)
          NbItem + 1
        EndIf
        ;
        MenuBar()
        MenuItem(#PBBMenu_SearchResultsOffset + NbItem, GetTextFromCatalogPB("All"))
        ApplyThemesToMenu(GPBBGadgets\SearchResultMenu)
        ;
        XMenu = DesktopScaledX(WindowX(GPBBGadgets\PBBWindow) + GadgetX(GPBBGadgets\SearchedExpression_gadget) + 7)
        YMenu = DesktopScaledY(WindowY(GPBBGadgets\PBBWindow) + GadgetY(GPBBGadgets\SearchedExpression_gadget) + 31)
        DisplayPopupMenu(GPBBGadgets\SearchResultMenu, WindowID(GPBBGadgets\PBBWindow), XMenu, YMenu)
        ;
        ReturnValue = #PBBNoType
        ;
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowSearchAndIdentification
            OPDReport("GetInfosFromData: " + Str(ListSize(SearchMenuChoices())) + " candidates found --> Interactive user choice displayed")
          EndIf
        CompilerEndIf
      Else
        ; Unambiguous contextual match (or non-interactive refresh of a synthetic 'All...' scope):
        If EffectiveFilter\DontCheckPosition And EffectiveFilter\ParentTypeAndName$ <> ""
          ForEach SearchMenuChoices()
            If SearchMenuChoices()\DontCheckPosition And
               SearchMenuChoices()\TypeOfElement = EffectiveFilter\TypeOfElement And
               SearchMenuChoices()\ParentTypeAndName$ = EffectiveFilter\ParentTypeAndName$ And
               (EffectiveFilter\ElementSpecies$ = "" Or SearchMenuChoices()\ElementSpecies$ = EffectiveFilter\ElementSpecies$)
              *FirstCompatible     = @SearchMenuChoices()
              FirstCompatibleIndex = ListIndex(SearchMenuChoices())
              Break
            EndIf
          Next
        ElseIf *FirstCompatible\TypeOfElement = #PBBVariable And *FirstCompatible\ElementSpecies$ = "Shared"
          Protected SharedFixedID = *FirstCompatible\FixedID, SharedMod$ = *FirstCompatible\ModuleName$
          ForEach SearchMenuChoices()
            If SearchMenuChoices()\DontCheckPosition And SearchMenuChoices()\ElementSpecies$ = "Shared"
              *FirstCompatible     = @SearchMenuChoices()
              FirstCompatibleIndex = ListIndex(SearchMenuChoices())
              Break
            EndIf
          Next
        EndIf
        ;
        ChangeCurrentElement(SearchMenuChoices(), *FirstCompatible)
        If SearchMenuChoices()\DontCheckPosition
          CopyStructure(@SearchMenuChoices(), *SearchParam, SearchDetails)
          If *SearchParam\ElementSpecies$ = "Shared" And *SearchParam\ParentTypeAndName$ <> "AllModules"
            *SearchParam\FixedID     = SharedFixedID
            *SearchParam\ModuleName$ = SharedMod$
          EndIf
        Else
          If HasJoker
            *SearchParam\ElementPurifiedName$ = ""
          EndIf
          InheritStrongSearchTraits(@SearchMenuChoices(), *SearchParam)
          *SearchParam\DontCheckPosition = SearchMenuChoices()\DontCheckPosition
          *SearchParam\MultipleResults   = SearchMenuChoices()\MultipleResults
          *SearchParam\OtherScopesCount  = SearchMenuChoices()\OtherScopesCount
        EndIf
        ;
        *SearchParam\SearchInCom    = EffectiveFilter\SearchInCom
        *SearchParam\SearchInQuotes = EffectiveFilter\SearchInQuotes
        If *SearchParam\DontCheckPosition = #False And *SearchParam\ElementFile$ <> "" And *SearchParam\ElementPosition > 0
          *SearchParam\InDeadZone = IsPosInDeadZone(*SearchParam\ElementFile$, *SearchParam\ElementPosition)
        Else
          *SearchParam\InDeadZone = #False
        EndIf
        *SearchParam\SearchInDeadZones = Bool(*SearchParam\SearchPurifiedOnly Or *SearchParam\InDeadZone)
        ReturnValue = *SearchParam\TypeOfElement
        ;
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowSearchAndIdentification
            OPDReport("GetInfosFromData: unique match retained - " + *SearchParam\ElementName$ + " " +
                      *SearchParam\ElementSpecies$ + " " + *SearchParam\ElementSubSpecies$)
          EndIf
        CompilerEndIf
      EndIf
    ElseIf *SearchParam\SearchPurifiedOnly And TotalRealCandidates > 0 And (TypeOfSearch & #PBBUserWillChoose) = 0
      ; Passive refresh of an active 'All' search: retain raw-text mode and update candidate count / scopes list
      FirstCompatibleIndex = ListSize(SearchMenuChoices())
      ReturnValue          = #PBBNativeFunction
    ElseIf *SearchParam\SearchPurifiedOnly = #False And *SearchParam\DontCheckPosition = #False And
           *SearchParam\ElementFile$ <> "" And *SearchParam\ElementPosition > 0 And
           IsPosInDeadZone(*SearchParam\ElementFile$, *SearchParam\ElementPosition)
      ; Target element resides inside an inactive CompilerIf zone (not indexed in ElementList()): resolve from file context
      ReturnValue = GetInfosFromContext(*SearchParam, *SearchParam\ElementFile$, *SearchParam\ElementPosition)
      If ReturnValue <= #PBBNoType Or ReturnValue = #PBBNotAnElement
        ReturnValue = #PBBNotFound
      EndIf
    EndIf
    ;}
  EndIf
  ;
  ; ======================================================================
  ; STEP 3: UPDATE SEARCH BUTTON LABEL AND STATE
  ; ======================================================================
  If TotalRealCandidates = 0
    ; Case A: No matching element exists anywhere in the project -> "Search" (disabled)
    SetSearchButtonState(#PBBSearchBtn_Search, #False, 0)
  ElseIf TotalRealCandidates = 1 And ReturnValue <> #PBBNotFound
    ; Case B: Single unambiguous element already matched in current scope -> "LookUp" (disabled)
    SetSearchButtonState(#PBBSearchBtn_LookUp, #False, 1)
  Else
    ; Case C: Other scopes or homonyms exist in SearchMenuChoices() -> "LookUp" (enabled)
    SetSearchButtonState(#PBBSearchBtn_LookUp, #True, TotalRealCandidates)
  EndIf
  ;
  CompilerIf #PBBReport
    If ReturnValue = #PBBNotFound And (OPDMod & #OPDFollowSearchAndIdentification)
      OPDReport("GetInfosFromData: Nothing found for: " + *SearchParam\ElementName$ + " (" +
                *SearchParam\ElementSpecies$ + " " + *SearchParam\ElementSubSpecies$ + ")")
    EndIf
  CompilerEndIf
  ;
  If FR_Gadgets\Window >= 0 And IsWindow(FR_Gadgets\Window)
    GUI_FR_RefreshScopes(*SearchParam, FirstCompatibleIndex)
  EndIf
  ;
  ProcedureReturn ReturnValue
  ;
EndProcedure
;
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 2111
; FirstLine = 1271
; Folding = AQAAAAAAAA+0
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb