; **********************************************************************************
;
;                          PBBrowser_CodeExploration.pb
;
;        Semantic Resolution, Single-Pass AST Indexing & Partial Code Updates
;                                  Nov. 20/2026
;
; **********************************************************************************
;
; ROLE & ARCHITECTURE:
; Core code-analysis engine of PBBrowser (independent of UI panels and search dialogs).
; Populates and maintains ElementList(), resolves identifier scopes and structure chains
; across both Text Mode and O(1) List Mode, inventories procedure-level variables, and
; updates project metadata incrementally whenever source files are edited on disk.
;
; FILE ORGANIZATION:
;   • Module Visibility & UseModule Timeline    - Tracks public exports and active UseModule scopes.
;   • Core Element, Structure & Field Lookup    - Resolves homonyms, structures, and '\Field' chains.
;   • Compile-Time Evaluator & Include Tree     - Constant/CompilerIf engine and project file tree builder.
;   • Token & Macro Usage Tracking              - Links caller/callee dependencies and expands macros.
;   • Shared Lexical Stream Engine              - Common tokenizer macros (OnePass & ListVariables).
;   • One-Pass Project Elements Exploration     - Indexes all project declarations and propagates usage.
;   • Incremental Partial Code Updates          - Re-parses modified file slices without full resets.
;   • Procedure Variable & Dependency Inventory - Builds local/global/unprotected/missing lists.
;
;
Define ExecuteSearch, RefreshEPanel
;
; **********************************************************************************
;
;-                   Module visibility & UseModule timeline
;
; **********************************************************************************
;
Structure PBB_ModuleUsageEvent
  FileName$
  PosInCode.i
  ModuleName$
  EnclosingModule$
  IsUnuse.b
EndStructure
;
NewList PBB_ModuleUsageEvents.PBB_ModuleUsageEvent()
;
Global PBB_ModuleCacheState = 0 ; 0 = unpopulated, 1 = populated (has UseModule/UnuseModule), -1 = none in project
;
Procedure PopulateModuleCaches()
  ;
  ; Role: Single-pass lazy initializer scanning project compilation sequences once to populate
  ; 'PBB_ModuleUsageEvents()' ('UseModule' / 'UnuseModule' timeline). Sets 'PBB_ModuleCacheState'
  ; to 1 if directives exist, or -1 if none.
  ;
  Shared PBB_ModuleUsageEvents(), PBB_CompilationSequence(), PBB_FileListing()
  ;
  Protected FIPS.FIPSPosStruct
  Protected EncompassingMod.EEDStruct
  Protected *Char.Character
  Protected ExtractedMod$
  Protected IsUnuse
  ;
  ClearList(PBB_ModuleUsageEvents())
  PushListPosition(ElementList())
  PushListPosition(PBB_CompilationSequence())
  PushListPosition(PBB_FileListing())
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
  CompilerEndIf
  ;
  FIPS\ScanMode = #FIPSMode_Global
  While FindInWholeProjectSequences("usemodule,unusemodule", @FIPS, #CheckAll)
    IsUnuse       = Bool(FastMid(FIPS\SearchString_LCase\s, FIPS\PosInCode, 11) = "unusemodule")
    *Char         = PosToChar(@FIPS\SearchString_LCase\s, FIPS\PosInCode + 9 + 2 * IsUnuse)
    ExtractedMod$ = ExtractOrSkipNextToken(@FIPS\SearchString_LCase\s, @*Char)
    If ExtractedMod$ <> ""
      AddElement(PBB_ModuleUsageEvents())
      PBB_ModuleUsageEvents()\FileName$   = FIPS\FileName$
      PBB_ModuleUsageEvents()\PosInCode   = FIPS\PosInCode
      PBB_ModuleUsageEvents()\ModuleName$ = ExtractedMod$
      PBB_ModuleUsageEvents()\IsUnuse     = IsUnuse
      If GetEncompassingLimitsAndName(#PBBModule, FIPS\FileName$, FIPS\PosInCode,
                                      FIPS\SearchString, FIPS\SearchString_LCase, FIPS\SegmentEnd, @EncompassingMod) <> ""
        PBB_ModuleUsageEvents()\EnclosingModule$ = LCase(EncompassingMod\ElementName$)
      EndIf
    EndIf
    FIPS\FIPSContinue = #True
  Wend
  ;
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
  CompilerEndIf
  PopListPosition(PBB_FileListing())
  PopListPosition(PBB_CompilationSequence())
  PopListPosition(ElementList())
  ;
  If ListSize(PBB_ModuleUsageEvents()) > 0 : PBB_ModuleCacheState = 1 : Else : PBB_ModuleCacheState = -1 : EndIf
EndProcedure
;
Procedure   CheckIfDeclaredInPublicModule(ModuleName$, ElementPurifiedName$)
  ;
  ; Role: Determines whether 'ElementPurifiedName$' is publicly exported inside the 'DeclareModule'
  ; block of 'ModuleName$' (excluding #PBBField members). Resolves in O(1) via ElementMapIndex() in
  ; List Mode, or via bounded on-demand scanning of the matching 'DeclareModule' block in Text Mode.
  ;
  Shared PBB_CompilationSequence(), PBB_FileListing()
  ;
  Protected CandidateIDs$, IDStr$, PrefixSlice$
  Protected *Elem.Elements, *Char.Character
  Protected FIPS.FIPSPosStruct
  Protected SearchState.FIPC_KeywordCache
  Protected IDIndex, BlockEndPos, FoundPos, PosInLine, IsPublic = #False
  ;
  If ModuleName$ = "" Or ElementPurifiedName$ = "" Or
     (ExplorationState = #ExplorationState_Completed And ListSize(ElementList()) And PBB_HasModules = #False)
    ProcedureReturn #False
  EndIf
  ;
  ModuleName$          = LCase(ModuleName$)
  ElementPurifiedName$ = LCase(ElementPurifiedName$)
  ;
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
    If OPDMod & #OPDBuildOutExternalProcedureTiming
      StartTimer(4, "CheckIfDeclaredInPublicModule")
    EndIf
  CompilerEndIf
  ;
  If ExplorationState = #ExplorationState_Completed And ListSize(ElementList())
    If FindMapElement(ElementMapIndex(), ElementPurifiedName$)
      CandidateIDs$ = ElementMapIndex()
      Repeat
        IDIndex + 1
        IDStr$ = StringField(CandidateIDs$, IDIndex, ",")
        If IDStr$ = "" : Break : EndIf
        *Elem = ElementFixedIndexes(Val(IDStr$))
        If *Elem And *Elem\StartingPos > 0 And *Elem\Type <> #PBBField And *Elem\IsDeclareModule And
           LCase(GetEnclosingModuleName(*Elem\EnclosingModuleID)) = ModuleName$
          IsPublic = #True
          Break
        EndIf
      ForEver
    EndIf
  Else
    PushListPosition(PBB_CompilationSequence())
    PushListPosition(PBB_FileListing())
    FIPS\ScanMode = #FIPSMode_Global
    While FindInWholeProjectSequences("declaremodule", @FIPS, #CheckAll)
      *Char = PosToChar(@FIPS\SearchString_LCase\s, FIPS\PosInCode + 13)
      If ExtractOrSkipNextToken(@FIPS\SearchString_LCase\s, @*Char) = ModuleName$
        FoundPos               = ComputeLenFromChars(@FIPS\SearchString_LCase\s, *Char) + 1
        BlockEndPos            = FindInPBPointedCode(FIPS\SearchString_LCase, "enddeclaremodule", FoundPos, #CheckAll)
        If BlockEndPos = 0 : BlockEndPos = Len(FIPS\SearchString_LCase\s) : EndIf
        SearchState\BufferAddr = 0
        While FoundPos < BlockEndPos
          FoundPos = FindInPBPointedCode(FIPS\SearchString_LCase, ElementPurifiedName$, FoundPos,
                                         #CheckAll | #CheckOneKeyWordOnly, #PB_String_CaseSensitive, BlockEndPos, 0, @SearchState)
          If FoundPos = 0 : Break : EndIf
          If GetEncompassingLimitsAndName(#PBBStructure, FIPS\FileName$, FoundPos, FIPS\SearchString, FIPS\SearchString_LCase, BlockEndPos) = "" And
             GetEncompassingLimitsAndName(#PBBInterface, FIPS\FileName$, FoundPos, FIPS\SearchString, FIPS\SearchString_LCase, BlockEndPos) = "" And
             GetEncompassingLimitsAndName(#PBBMacro,     FIPS\FileName$, FoundPos, FIPS\SearchString, FIPS\SearchString_LCase, BlockEndPos) = ""
            PrefixSlice$ = ExtractSectionAroundExpression(FIPS\SearchString_LCase, ElementPurifiedName$, FoundPos, 0,
                                                          #CheckAll | #CheckOneKeyWordOnly, BlockEndPos, @PosInLine)
            If PosInLine > 0
              PrefixSlice$ = FastLeft(PrefixSlice$, PosInLine - 1)
              If CountString(PrefixSlice$, "(") = CountString(PrefixSlice$, ")") And
                 FindString(StringField(PrefixSlice$, CountString(PrefixSlice$, ",") + 1, ","), "=") = 0 And
                 Right(PrefixSlice$, 1) <> "." And Right(PrefixSlice$, 1) <> ":" And
                 Left(PrefixSlice$, 9) <> "usemodule" And Left(PrefixSlice$, 11) <> "unusemodule" And Left(PrefixSlice$, 8) <> "compiler"
                IsPublic = #True
                Break 2
              EndIf
            EndIf
          EndIf
          FoundPos + Len(ElementPurifiedName$)
        Wend
      EndIf
      FIPS\FIPSContinue = #True
    Wend
    PopListPosition(PBB_FileListing())
    PopListPosition(PBB_CompilationSequence())
  EndIf
  ;
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
    If OPDMod & #OPDBuildOutExternalProcedureTiming
      ReapTimer(4)
    EndIf
  CompilerEndIf
  ProcedureReturn IsPublic
EndProcedure
;
Procedure.s CheckIfModuleIsUsed(ModuleName$, FileName$, PosInCode, ShadowedSymbol$ = "", EnclosingModule$ = "")
  ;
  ; Role: Evaluates active 'UseModule' / 'UnuseModule' directives within 'EnclosingModule$'
  ; (or global scope when empty) up to 'PosInCode' in 'FileName$':
  ; - Targeted mode (ModuleName$ <> ""): returns LCase(ModuleName$) if actively open via 'UseModule', or "".
  ; - Shadowing mode (ModuleName$ = "", ShadowedSymbol$ <> ""): returns the lowercase name of the
  ;   active 'UseModule' namespace that publicly exports 'ShadowedSymbol$', or "" otherwise.
  ; Preserves the caller's current ElementList() position across lazy cache population and queries.
  ;
  Shared PBB_ModuleUsageEvents()
  ;
  Protected ActiveModules$ = ",", ResultModule$, CandidateIDs$, IDStr$, CurrentMod$
  Protected IDIndex, ModIndex, HasCandidate
  Protected *Elem.Elements, *ModElem.Elements
  ;
  If (ModuleName$ = "" And ShadowedSymbol$ = "") Or (ShadowedSymbol$ <> "" And (FileName$ = "" Or PosInCode <= 0))
    ProcedureReturn ""
  EndIf
  ;
  ModuleName$      = LCase(ModuleName$)
  ShadowedSymbol$  = LCase(ShadowedSymbol$)
  EnclosingModule$ = LCase(EnclosingModule$)
  ;
  ; 1. O(1) fast-exit when no modules exist or when no candidate module has HasUseModule = #True
  If ExplorationState = #ExplorationState_Completed And ListSize(ElementList())
    If PBB_HasModules = #False : ProcedureReturn "" : EndIf
    If ModuleName$ <> "" : CandidateIDs$ = ModuleName$ : Else : CandidateIDs$ = ShadowedSymbol$ : EndIf
    If FindMapElement(ElementMapIndex(), CandidateIDs$) = 0
      ProcedureReturn ""
    EndIf
    CandidateIDs$ = ElementMapIndex()
    Repeat
      IDIndex + 1
      IDStr$ = StringField(CandidateIDs$, IDIndex, ",")
      If IDStr$ = "" : Break : EndIf
      *Elem = ElementFixedIndexes(Val(IDStr$))
      If *Elem
        If ModuleName$ <> ""
          If *Elem\Type = #PBBModule And *Elem\HasUseModule : HasCandidate = #True : Break : EndIf
        ElseIf *Elem\Type <> #PBBField And *Elem\IsDeclareModule And *Elem\EnclosingModuleID > 0 And *Elem\EnclosingModuleID < NextElementID
          *ModElem = ElementFixedIndexes(*Elem\EnclosingModuleID)
          If *ModElem And *ModElem\HasUseModule : HasCandidate = #True : Break : EndIf
        EndIf
      EndIf
    ForEver
    If HasCandidate = #False : ProcedureReturn "" : EndIf
  ElseIf PBB_ModuleCacheState < 0
    ProcedureReturn ""
  EndIf
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDBuildOutExternalProcedureTiming
      StartTimer(5, "CheckIfModuleIsUsed")
    EndIf
  CompilerEndIf
  ;
  If PBB_ModuleCacheState = 0
    PopulateModuleCaches()
  EndIf
  ;
  ; 2. Replay 'UseModule' / 'UnuseModule' events restricted to the same enclosing scope
  If ListSize(PBB_ModuleUsageEvents()) > 0
    ForEach PBB_ModuleUsageEvents()
      If FileName$ <> "" And PosInCode > 0 And
         ((PBB_ModuleUsageEvents()\FileName$ = FileName$ And PBB_ModuleUsageEvents()\PosInCode > PosInCode) Or
          (PBB_ModuleUsageEvents()\FileName$ <> FileName$ And
           IsPositionBefore(PBB_ModuleUsageEvents()\FileName$, PBB_ModuleUsageEvents()\PosInCode, FileName$, PosInCode) = #False))
        Break
      EndIf
      If PBB_ModuleUsageEvents()\EnclosingModule$ <> EnclosingModule$
        Continue
      EndIf
      If ModuleName$ <> ""
        If PBB_ModuleUsageEvents()\ModuleName$ = ModuleName$
          If PBB_ModuleUsageEvents()\IsUnuse : ResultModule$ = "" : Else : ResultModule$ = ModuleName$ : EndIf
        EndIf
      ElseIf PBB_ModuleUsageEvents()\IsUnuse
        ActiveModules$ = ReplaceString(ActiveModules$, "," + PBB_ModuleUsageEvents()\ModuleName$ + ",", ",")
      ElseIf FindString(ActiveModules$, "," + PBB_ModuleUsageEvents()\ModuleName$ + ",") = 0
        ActiveModules$ + PBB_ModuleUsageEvents()\ModuleName$ + ","
      EndIf
    Next
    ;
    If ShadowedSymbol$ <> ""
      ModIndex = 1
      Repeat
        ModIndex + 1
        CurrentMod$ = StringField(ActiveModules$, ModIndex, ",")
        If CurrentMod$ = "" : Break : EndIf
        If CheckIfDeclaredInPublicModule(CurrentMod$, ShadowedSymbol$)
          ResultModule$ = CurrentMod$
          Break
        EndIf
      ForEver
    EndIf
  EndIf
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDBuildOutExternalProcedureTiming
      ReapTimer(5)
    EndIf
  CompilerEndIf
  ProcedureReturn ResultModule$
EndProcedure
;
;
; **********************************************************************************
;
;-                   Core element, structure & field lookup
;
; **********************************************************************************
;
Procedure   FindElementInList(ElementNameLCase$, TypeOfElement = 0, PosInCode = 0, FileName$ = "",
                               MacroProcID = 0, Parentheses = -1, VarDotType_LCase$ = "",
                               TargetContainerID = 0, ModuleContainer = 0,
                               VariableSpecie$ = "", FieldOrVariableSubSpecie$ = "",
                               *PreloadedCandidateIDs.String = 0)
  ;
  ; Role: Core semantic and lexical resolution engine for ElementList().
  ; Queries ElementMapIndex() for 'ElementNameLCase$' and evaluates all homonymous candidate IDs
  ; against type, morphology, module visibility, container scope, and single-pass compilation chronology.
  ;
  ; WHY THIS PROCEDURE IS NECESSARY:
  ; In PureBasic, a purified identifier name alone is not enough to identify an element because
  ; multiple distinct entities can legitimately share the exact same name across a project:
  ; 1. Cross-Type Homonyms: A Procedure, a Structure, a Variable, and a Label can all share the
  ;    same identifier (e.g., 'Item()' as a procedure and 'Item.ItemStruct' as a variable).
  ; 2. Variable Subspecies & Parentheses: A scalar variable ('MyData.i') and a collection
  ;    ('Dim MyData(10)' or 'NewList MyData()') can coexist in the same scope and are distinguished
  ;    by the presence of parentheses ('HasParentheses') and subspecies ('Var', 'Array', 'List', 'Map').
  ; 3. Local Scope & Dot-Type Homonyms: Two procedures can each declare their own 'Protected Index.i',
  ;    or 'Define' can re-type a variable across separate scopes.
  ; 4. Structure & Interface Fields (#PBBField): Distinct structures often define fields with identical
  ;    names (e.g., '\Name' or '\Type'). Resolving the right field requires checking 'TargetContainerID'
  ;    as well as any parent structures inherited via 'Extends' ('\UsageChildren').
  ; 5. Module Isolation: Elements declared inside a 'Module' are invisible to the global scope unless
  ;    exported via 'DeclareModule' and accessed through 'ModuleName::' or an active 'UseModule'.
  ; 6. Single-Pass Precedence: When inside a procedure ('MacroProcID > 0'), a local variable ('Protected'
  ;    or 'Static') shadows a 'Global' homonym only after the line where the local declaration occurs.
  ;    By tracking 'GlobalFallbackID' and 'ModuleFallbackID', this procedure resolves local-first
  ;    precedence, global fallback, and 'UseModule' exports in a single loop pass.
  ;
  ; Parameters:
  ; - ElementNameLCase$         : Lowercase purified identifier (or "mod::ident" when explicitly qualified).
  ; - TypeOfElement             : 0 to accept any non-field element; #PBBField to strictly evaluate
  ;                               structure/interface fields; or a specific #PBB... constant to filter
  ;                               strictly by that element type.
  ; - PosInCode                 : 1-based character offset of the token in FileName$ to enforce
  ;                               single-pass chronological visibility, or 0 to bypass position checks.
  ; - FileName$                 : Full path of the file being analyzed (required when PosInCode > 0).
  ; - MacroProcID               : > 0 when inside a specific Procedure/Macro (prioritizes its local
  ;                               variables and rejects local variables of other procedures);
  ;                               -1 when explicitly in main out-of-procedure code (accepts only Global,
  ;                               Threaded, Define, or Shared variables);
  ;                               0 when procedural scope is unconstrained.
  ; - Parentheses               : 1 if the token is followed by '(', 0 if not, or -1 to ignore.
  ; - VarDotType_LCase$         : Expected lowercase type suffix (e.g., "s", "mystruct"), or "" to ignore.
  ; - TargetContainerID         : FixedID of the enclosing Structure/Interface (for #PBBField) or
  ;                               Enumeration (for #PBBConstant), or 0 if unconstrained.
  ; - ModuleContainer           : FixedID of the Module/DeclareModule enclosing the call site (-1 if global).
  ; - VariableSpecie$           : Optional variable species filter (e.g., "Global", "Protected"), or "".
  ; - FieldOrVariableSubSpecie$ : Optional subspecies filter (e.g., "Var", "Array", "List", "Map"), or "".
  ; - *PreloadedCandidateIDs    : Optional pointer to a pre-fetched ElementMapIndex() string to bypass a redundant map lookup.
  ;
  ; Output:
  ; Returns the FixedID of the winning element (leaving ElementList() positioned on it), or 0 if no match.
  ;
  Protected CandidateIDs$, IDStr$, StructChildren$, SearchTag$, RecordLine$, EvalFile$
  Protected IDIndex, CandType, ScopeStatus, FoundContainer, IsLocalMatch, EvalPos, ColonPos, ExplicitModuleID = -1
  Protected TagPos, CEDIndex, CEDContainer, ModuleFallbackIsStruct
  Protected IsUsed, GlobalFallbackID, StructFallbackID, ModuleFallbackID
  Protected *ModElem.Elements, *ProcElem.Elements
  ;
  ColonPos = FindString(ElementNameLCase$, "::")
  If ColonPos
    If PBB_HasModules = #False Or ColonPos = 1 : ProcedureReturn 0 : EndIf
    ExplicitModuleID  = FindElementInList(FastLeft(ElementNameLCase$, ColonPos - 1), #PBBModule)
    ElementNameLCase$ = FastMid(ElementNameLCase$, ColonPos + 2)
    If ExplicitModuleID = 0 Or FindString(ElementNameLCase$, "::") : ProcedureReturn 0 : EndIf
  EndIf
  If ElementNameLCase$ = ""
    ProcedureReturn 0
  ElseIf *PreloadedCandidateIDs And *PreloadedCandidateIDs\s <> ""
    CandidateIDs$ = *PreloadedCandidateIDs\s
  ElseIf FindMapElement(ElementMapIndex(), ElementNameLCase$)
    CandidateIDs$ = ElementMapIndex()
  Else
    ProcedureReturn 0
  EndIf
  ;
  ; Retrieve 'Extends' parent structure IDs when disambiguating a field on a specific structure
  If TypeOfElement = #PBBField And TargetContainerID > 0 And SelectPBBElement(TargetContainerID)
    StructChildren$ = ElementList()\UsageChildren
  EndIf
  If VariableSpecie$ <> "" And FieldOrVariableSubSpecie$ = ""
    FieldOrVariableSubSpecie$ = "Var"
  EndIf
  ;
 Repeat
    IDIndex + 1
    IDStr$ = StringField(CandidateIDs$, IDIndex, ",")
    If IDStr$ = ""
      Break
    EndIf
    ;
    ; Skip deleted IDs or elements temporarily invalidated (StartingPos = 0) during PartialUpdateOfCode()
    If SelectPBBElement(Val(IDStr$)) = 0 Or ElementList()\StartingPos = 0
      Continue
    EndIf
    ;
    CandType = ElementList()\Type
    ;
    ; -------------------------------------------------------------------------
    ; GATE 1: TYPE & MORPHOLOGY FILTER
    ; Ensures the candidate's syntactic nature matches the token at the call site.
    ; - When TypeOfElement = 0 (standard code token without a leading '\'), #PBBField
    ;   candidates are strictly excluded so structure fields never shadow variables or procedures.
    ; - When TypeOfElement = #PBBStructOrInterface, strictly restricts candidates to
    ;   #PBBStructure and #PBBInterface definitions (for '.Type' suffixes and 'Extends').
    ; - Distinguishes scalar variables ('MyVar', HasParentheses = 0) from arrays, lists,
    ;   maps, or procedures ('MyVar()', HasParentheses = 1).
    ; -------------------------------------------------------------------------
    If TypeOfElement = 0 Or TypeOfElement = #PBBTypeOrVar
      ; Standard expression tokens never target fields, enumerations, modules, or import blocks
      If CandType = #PBBField Or CandType = #PBBEnumeration Or CandType = #PBBModule Or CandType = #PBBImport
        Continue
      EndIf
    ElseIf TypeOfElement = #PBBStructOrInterface
      If CandType <> #PBBStructure And CandType <> #PBBInterface
        Continue
      EndIf
    ElseIf CandType <> TypeOfElement
      Continue
    EndIf
    ;
    If CandType = #PBBVariable
      ; Reject procedure-local variables immediately when explicitly qualified via 'Mod::' (ExplicitModuleID > 0),
      ; or when in main scope (MacroProcID = -1) with no secondary declarations in '\Continued', unless 'Shared'.
      If ((MacroProcID = -1 And ElementList()\Continued = "") Or ExplicitModuleID > 0) And
         ElementList()\LocalVarContainer > 0 And ElementList()\VariableSpecie <> "Shared"
        Continue
      EndIf
      If VarDotType_LCase$ <> "" And ElementList()\VarDotType_LCase <> VarDotType_LCase$ And
         (FindString(VarDotType_LCase$, "::") Or Right(ElementList()\VarDotType_LCase, Len(VarDotType_LCase$) + 2) <> "::" + VarDotType_LCase$)
        Continue
      EndIf
      If VariableSpecie$ <> "" And ElementList()\VariableSpecie <> VariableSpecie$
        Continue
      EndIf
      If FieldOrVariableSubSpecie$ <> "" And ElementList()\VariableSubSpecie <> FieldOrVariableSubSpecie$
        Continue
      EndIf
    ElseIf CandType = #PBBField
      If TargetContainerID > 0 And ElementList()\LocalVarContainer <> TargetContainerID And
         (StructChildren$ = "" Or FindString(StructChildren$, "/" + Str(ElementList()\LocalVarContainer) + "/") = 0)
        Continue
      EndIf
    ElseIf CandType = #PBBConstant And TargetContainerID > 0
      If ElementList()\LocalVarContainer <> TargetContainerID
        Continue
      EndIf
    EndIf
    ;
    If Parentheses <> -1 And ElementList()\HasParentheses <> -1 And ElementList()\HasParentheses <> Parentheses
      Continue
    EndIf
    ;
    ; -------------------------------------------------------------------------
    ; GATE 2: MODULE NAMESPACE VISIBILITY
    ; Evaluates PureBasic module encapsulation rules and assigns a visibility rank:
    ; - ScopeStatus = 1 (Direct match):
    ;   * The candidate is a #PBBModule container itself,
    ;   * Or the token is explicitly qualified ('Mod::Elem') and the candidate belongs
    ;     to 'ExplicitModuleID' (and is public in DeclareModule if called from outside),
    ;   * Or the token is unqualified and resides in the exact same module (or global scope).
    ; - ScopeStatus = 2 (Public fallback via active 'UseModule' or unconstrained lookup):
    ;   * ModuleContainer = 0 (unconstrained lookup), or the candidate is publicly exported
    ;     in a 'DeclareModule' block whose module is actively open via 'UseModule' in the
    ;     caller's scope (global or enclosing module) at the evaluation position.
    ; - ScopeStatus = 0 (Rejected):
    ;   * Private module elements accessed from outside, un-imported 'DeclareModule' elements,
    ;     or global elements accessed from inside a module.
    ; -------------------------------------------------------------------------
    ScopeStatus = 0
    If PBB_HasModules = #False Or CandType = #PBBModule
      ScopeStatus = 1
    ElseIf ExplicitModuleID > 0
      If ElementList()\EnclosingModuleID = ExplicitModuleID
        If ModuleContainer = 0 Or ModuleContainer = ExplicitModuleID Or ElementList()\IsDeclareModule
          ScopeStatus = 1
        EndIf
      EndIf
    ElseIf (ModuleContainer > 0 And ElementList()\EnclosingModuleID = ModuleContainer) Or
           (ModuleContainer <= 0 And ElementList()\EnclosingModuleID <= 0)
      ScopeStatus = 1
    ElseIf ModuleContainer = 0
      ScopeStatus = 2
    ElseIf ElementList()\EnclosingModuleID > 0 And ElementList()\IsDeclareModule
      If CandType = #PBBField
        ; Fields of structures/interfaces exported in a DeclareModule block are accessible via typed variables without 'UseModule'
        ScopeStatus = 2
      ElseIf ElementList()\EnclosingModuleID < NextElementID
        *ModElem = ElementFixedIndexes(ElementList()\EnclosingModuleID)
        If *ModElem And *ModElem\HasUseModule
          EvalFile$ = FileName$
          EvalPos   = PosInCode
          If EvalPos <= 0 And MacroProcID > 0 And MacroProcID < NextElementID
            *ProcElem = ElementFixedIndexes(MacroProcID)
            If *ProcElem
              EvalFile$ = *ProcElem\FileName
              EvalPos   = *ProcElem\EndingPos
            EndIf
          EndIf
          If CheckIfModuleIsUsed(*ModElem\PurifiedName, EvalFile$, EvalPos, "", GetEnclosingModuleName(ModuleContainer)) <> ""
            ScopeStatus = 2
          EndIf
        EndIf
      EndIf
    EndIf
    ;
    If ScopeStatus = 0
      Continue
    EndIf
    ;
    ; -------------------------------------------------------------------------
    ; GATE 3: ENCLOSING CONTAINER & SINGLE-PASS CHRONOLOGY
    ; IMPORTANT ARCHITECTURAL NOTE (DO NOT REPLACE 'FileName$ <> ElementList()\FileName' WITH IsPositionBefore()):
    ; Position checks intentionally test 'ElementList()\EndingPos < PosInCode Or FileName$ <> ElementList()\FileName'
    ; (calling IsPositionBefore() ONLY for #PBBMacro overriding native PureBasic functions) for three reasons:
    ; 1. During initial OnePass exploration, PBB_CompilationSequence() is already traversed in strict
    ;    chronological compiler order, so any element already indexed in another file is guaranteed upstream
    ;    without paying the cost of cross-file virtual coordinate lookups on every token.
    ; 2. Post-exploration UI callers (such as ResolveDetailPanelToken() in the Details panel) pass PosInCode > 0
    ;    specifically to enforce intra-procedure / intra-file shadowing (e.g., local variables after their
    ;    declaration line) while still expecting cross-file declarations to resolve directly.
    ; 3. Non-included project files listed in a .pbp project exist in PBB_FileListing() outside
    ;    PBB_CompilationSequence(), so enforcing cross-file virtual coordinates here would break lookups.
    ;
    ; Verifies that the candidate is both in an accessible container AND declared
    ; upstream of PosInCode (when PosInCode > 0):
    ; - For #PBBField: checks if the field belongs to TargetContainerID or to one of its
    ;   ancestor structures inherited via 'Extends' (recorded in StructChildren$).
    ; - For #PBBMacro: if the macro overrides a PureBasic native function (e.g., 'Macro FindString'),
    ;   verifies via IsPositionBefore() that the macro was declared before PosInCode.
    ; - For #PBBVariable (and segmented declarations): a variable may be declared in multiple
    ;   procedures via '\Continued' (e.g., 'Protected x' or 'Shared x' in ProcA and ProcB).
    ;   The SAME declaration record (either the primary ElementList() record or a specific
    ;   '\Continued' entry) must satisfy both container membership (MacroProcID) and upstream
    ;   position (EndingPos < PosInCode) to prevent a local declaration at the bottom of a
    ;   procedure from validating usages at the top of that same procedure.
    ; -------------------------------------------------------------------------
    FoundContainer = #False
    IsLocalMatch   = #False
    ;
    If CandType = #PBBField Or (CandType = #PBBConstant And TargetContainerID > 0)
      If PosInCode <= 0 Or ElementList()\EndingPos < PosInCode Or FileName$ <> ElementList()\FileName
        FoundContainer = #True
      EndIf
    ElseIf CandType = #PBBMacro
      If PosInCode <= 0 Or FindMapElement(PBNativeFunctions(), ElementList()\PurifiedName) = 0 Or
         IsPositionBefore(ElementList()\FileName, ElementList()\EndingPos, FileName$, PosInCode)
        FoundContainer = #True
      Else
        CompilerIf #PBBReport
          If OPDMod & #OPDIndicatesUsageDetection
            OPDReport("  --> This is a PB function and not a macro")
          EndIf
        CompilerEndIf
      EndIf
    Else
      ; Evaluate the primary element declaration first
      ; (#PBBLabel targets remain valid across the entire scope regardless of upstream/downstream line order)
      If PosInCode <= 0 Or CandType = #PBBLabel Or ElementList()\EndingPos < PosInCode Or FileName$ <> ElementList()\FileName
        If CandType <> #PBBVariable Or MacroProcID = 0 Or
           ElementList()\LocalVarContainer = 0 Or ElementList()\LocalVarContainer = MacroProcID Or
           (MacroProcID = -1 And ElementList()\VariableSpecie = "Shared")
          FoundContainer = #True
          If CandType = #PBBVariable And MacroProcID > 0 And ElementList()\LocalVarContainer = MacroProcID
            IsLocalMatch = #True
          EndIf
        EndIf
      EndIf
      ;
      ; Inspect secondary declarations in \Continued without building CEDList():
      ; - Fast path: inside a procedure (MacroProcID > 0), jump directly to '#TAB$ + MacroProcID + #TAB$'
      ;   using native C FindString() and verify column 5 (DirectContainer) to avoid false hits on positions.
      ; - Fallback: scan individual lines for forward-declared Procedures/Modules/Enumerations (DirectContainer = -1)
      ;   or unconstrained/main-scope continuations.
      If (FoundContainer = #False Or (CandType = #PBBVariable And MacroProcID > 0 And ExplicitModuleID <= 0 And IsLocalMatch = #False)) And
         ElementList()\Continued <> ""
        If CandType = #PBBVariable And MacroProcID > 0 And (FoundContainer = #True Or ElementList()\LocalVarContainer > 0)
          SearchTag$ = #TAB$ + Str(MacroProcID) + #TAB$
          TagPos     = 0
          Repeat
            TagPos = FindString(ElementList()\Continued, SearchTag$, TagPos + 1)
            If TagPos
              RecordLine$ = StringField(ElementList()\Continued, CountString(FastLeft(ElementList()\Continued, TagPos), #CR$) + 1, #CR$)
              If Val(StringField(RecordLine$, 5, #TAB$)) = MacroProcID
                If PosInCode <= 0 Or Val(StringField(RecordLine$, 3, #TAB$)) < PosInCode Or FileName$ <> StringField(RecordLine$, 1, #TAB$)
                  FoundContainer = #True
                  IsLocalMatch   = #True
                  Break
                EndIf
              EndIf
              TagPos + Len(SearchTag$) - 2
            EndIf
          Until TagPos = 0
        ElseIf FoundContainer = #False
          CEDIndex = 0
          Repeat
            CEDIndex + 1
            RecordLine$ = StringField(ElementList()\Continued, CEDIndex, #CR$)
            If RecordLine$ = ""
              Break
            EndIf
            If PosInCode <= 0 Or Val(StringField(RecordLine$, 3, #TAB$)) < PosInCode Or FileName$ <> StringField(RecordLine$, 1, #TAB$)
              CEDContainer = Val(StringField(RecordLine$, 5, #TAB$))
              If CandType <> #PBBVariable Or MacroProcID = 0 Or ElementList()\LocalVarContainer = 0 Or
                 CEDContainer = MacroProcID Or (MacroProcID = -1 And ElementList()\VariableSpecie = "Shared")
                FoundContainer = #True
                If CandType = #PBBVariable And MacroProcID > 0 And CEDContainer = MacroProcID
                  IsLocalMatch = #True
                EndIf
                If CandType <> #PBBVariable Or MacroProcID <= 0 Or IsLocalMatch
                  Break
                EndIf
              EndIf
            EndIf
          ForEver
        EndIf
      EndIf
    EndIf
    ;
    If FoundContainer = #False
      Continue
    EndIf
    ;
    ; -------------------------------------------------------------------------
    ; GATE 4: SINGLE-PASS PRIORITY RESOLUTION (LOCAL vs. GLOBAL vs. FOREIGN MODULE)
    ; Because a Global variable (declared at the top of a file) appears earlier in
    ; CandidateIDs$ than a homonymous Protected variable (declared inside a procedure),
    ; encountering a Global variable while inside a procedure (MacroProcID > 0) must not
    ; halt the loop prematurely:
    ; - Priority 1 (Immediate winner -> Break): Same-module candidate (ScopeStatus = 1)
    ;   that is either a non-variable, a variable evaluated outside a specific procedure
    ;   (MacroProcID <= 0), a module-qualified variable (ExplicitModuleID > 0), or a
    ;   variable explicitly local/Shared in MacroProcID (IsLocalMatch = #True).
    ; - Priority 2 (GlobalFallbackID): Same-module global variable encountered while inside
    ;   a procedure (retained in case no local variable shadows it further in CandidateIDs$).
    ; - Priority 3 (ModuleFallbackID): Public DeclareModule candidate from another module.
    ; -------------------------------------------------------------------------
    ; When TypeOfElement = 0, defer #PBBStructure and #PBBInterface to StructFallbackID so a homonymous
    ; variable or procedure (e.g. 'Global Config.Config') takes precedence over the structure definition,
    ; while standalone type expressions ('SizeOf(MyStruct)', 'AllocateStructure(MyStruct)') still resolve.
    If ScopeStatus = 1
      If TypeOfElement = 0 And (CandType = #PBBStructure Or CandType = #PBBInterface)
        If StructFallbackID = 0
          StructFallbackID = ElementList()\FixedID
        EndIf
      ElseIf CandType <> #PBBVariable Or MacroProcID <= 0 Or ExplicitModuleID > 0 Or IsLocalMatch
        IsUsed = ElementList()\FixedID
        Break
      ElseIf GlobalFallbackID = 0
        GlobalFallbackID = ElementList()\FixedID
      EndIf
    ElseIf ScopeStatus = 2
      If ModuleFallbackID = 0 Or
         (TypeOfElement = 0 And ModuleFallbackIsStruct And CandType <> #PBBStructure And CandType <> #PBBInterface) Or
         (TypeOfElement = #PBBTypeOrVar And ModuleFallbackIsStruct = #False And (CandType = #PBBStructure Or CandType = #PBBInterface))
        ModuleFallbackID       = ElementList()\FixedID
        ModuleFallbackIsStruct = Bool(CandType = #PBBStructure Or CandType = #PBBInterface)
      EndIf
    EndIf
    ;
  ForEver
  ;
  ; If no Priority 1 candidate was found, promote the highest-ranking fallback and select it
  If IsUsed = 0
    If GlobalFallbackID > 0
      IsUsed = GlobalFallbackID
    ElseIf StructFallbackID > 0
      IsUsed = StructFallbackID
    ElseIf ModuleFallbackID > 0
      IsUsed = ModuleFallbackID
    EndIf
    If IsUsed > 0 And ElementList()\FixedID <> IsUsed
      SelectPBBElement(IsUsed)
    EndIf
  EndIf
  ;
  ProcedureReturn IsUsed
EndProcedure
;
; ----------------------------------------------------------------------------------
;           ---- Dual-mode structure, interface & field typing ----
; ----------------------------------------------------------------------------------
;
;
Macro GetModuleID(ModName)
  ;
  ; Role: Resolves a module name to its FixedID in O(1) (short-circuiting immediately when PBB_HasModules = #False).
  ;
  (Bool(PBB_HasModules) * FindElementInList(PurifyElementName(ModName), #PBBModule))
EndMacro
;
Macro QualifyWithModule(ModName, ElemName)
  ;
  ; Role: Prepends 'ModName::' to 'ElemName' only when 'ModName' is non-empty.
  ;
  Left(ModName + "::", Bool(ModName <> "") * 999) + ElemName
EndMacro
;
Macro ResolveModContainerID(EnclosingMod, FileStr, CodePos)
  ;
  ; Role: Maps an enclosing module name and file/position anchor to a FindElementInList() ModuleContainer ID.
  ;
  If EnclosingMod <> ""
    ModContainer = GetModuleID(EnclosingMod)
  ElseIf FileStr <> "" And (CodePos) > 0
    ModContainer = -1
  Else
    ModContainer = 0
  EndIf
EndMacro
;
Procedure FindStructureOrInterfaceID(StructOrInterfaceName$, ModuleContainer = 0, PosInCode = 0, FileName$ = "")
  ;
  ; Role: Resolves a Structure or Interface name (including "Mod::Struct" or "Structure Mod::Struct")
  ;       to its FixedID in ElementList(), respecting module visibility and optional 'UseModule' chronology.
  ;
  If FindString(StructOrInterfaceName$, " ")
    StructOrInterfaceName$ = StringField(StructOrInterfaceName$, 2, " ")
  EndIf
  StructOrInterfaceName$ = PurifyElementName(StructOrInterfaceName$)
  If StructOrInterfaceName$ = "" : ProcedureReturn 0 : EndIf
  ProcedureReturn FindElementInList(StructOrInterfaceName$, #PBBStructOrInterface, PosInCode, FileName$, 0, -1, "", 0, ModuleContainer)
EndProcedure
;
Structure ResolvedStructCache
  Type.s
  File.s
  StartPos.i
  EndPos.i
EndStructure
;
Procedure.s ResolveStructureOrInterface(StructName$, *OutStartPos.Integer = 0, *OutEndPos.Integer = 0,
                                        *OutFile.String = 0, ForceListMode = #False,
                                        EnclosingModule$ = "", FileName$ = "", PosInCode = 0)
  ;
  ; Role: Resolves the concrete entity type and owning module ("Structure [Mod::]Name" or "Interface [Mod::]Name")
  ;       and physical source coordinates of 'StructName$' across O(1) List Mode and Text Mode.
  ;       Passing StructName$ = "" purges the static resolution cache.
  ;
  Shared PBB_CompilationSequence(), PBB_FileListing()
  Static NewMap StructCache.ResolvedStructCache()
  ;
  Protected StructName_LCase$, ExplicitMod_LCase$, CacheKey$, ImportedMod_LCase$
  Protected LineOfCode$, FirstWord$, ThirdWord$, ResolvedType$, DeclMod$, ResolvedFile$
  Protected PosStructStart, PosStructEnd, ColonPos, ModContainer = -1, IsScoped, IsDirectMatch, IsFallbackMatch, ListModeIsOK
  Protected FIPS.FIPSPosStruct
  Protected EncompassingMod.EEDStruct
  ;
  If *OutStartPos : *OutStartPos\i = 0  : EndIf
  If *OutEndPos   : *OutEndPos\i   = 0  : EndIf
  If *OutFile     : *OutFile\s     = "" : EndIf
  ;
  StructName$ = Trim(StructName$)
  If FindString(StructName$, " ") : StructName$ = Trim(StringField(StructName$, 2, " ")) : EndIf
  If StructName$ = ""
    ClearMap(StructCache())
    ProcedureReturn ""
  EndIf
  ;
  ColonPos = FindString(StructName$, "::")
  If ColonPos
    ExplicitMod_LCase$ = LCase(FastLeft(StructName$, ColonPos - 1))
    StructName$        = FastMid(StructName$, ColonPos + 2)
  EndIf
  StructName_LCase$ = LCase(StructName$)
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDBuildOutExternalProcedureTiming : StartTimer(3, "ResolveStructureOrInterface") : EndIf
  CompilerEndIf
  ;
  ; Fast-skip unqualified primitive types without traversing the project for scalar suffixes
  If ColonPos > 0 Or FindMapElement(PBBasicNativeTypes(), StructName_LCase$) = 0
    EnclosingModule$ = LCase(EnclosingModule$)
    IsScoped         = Bool(ColonPos > 0 Or EnclosingModule$ <> "" Or (FileName$ <> "" And PosInCode > 0))
    If ColonPos > 0
      CacheKey$ = "x:" + ExplicitMod_LCase$ + "::" + StructName_LCase$
    ElseIf IsScoped
      If FileName$ <> "" And PosInCode > 0
        ImportedMod_LCase$ = CheckIfModuleIsUsed("", FileName$, PosInCode, StructName_LCase$, EnclosingModule$)
      EndIf
      CacheKey$ = "s:" + EnclosingModule$ + ">" + ImportedMod_LCase$ + "::" + StructName_LCase$
    Else
      CacheKey$ = "*::" + StructName_LCase$
    EndIf
    ;
    If FindMapElement(StructCache(), CacheKey$)
      ResolvedType$  = StructCache()\Type
      PosStructStart = StructCache()\StartPos
      PosStructEnd   = StructCache()\EndPos
      ResolvedFile$  = StructCache()\File
    Else
      ListModeIsOK = Bool(ExplorationState = #ExplorationState_Completed And ListSize(ElementList()))
      If ListModeIsOK Or ForceListMode
        ResolveModContainerID(EnclosingModule$, FileName$, PosInCode)
        If (EnclosingModule$ = "" Or ModContainer > 0) And
           FindStructureOrInterfaceID(QualifyWithModule(ExplicitMod_LCase$, StructName_LCase$), ModContainer, PosInCode, FileName$)
          DeclMod$       = GetEnclosingModuleName(ElementList()\EnclosingModuleID)
          ResolvedType$  = PBBTypeNames$(ElementList()\Type) + #Space$ + QualifyWithModule(DeclMod$, ElementList()\Name)
          PosStructStart = ElementList()\StartingPos
          PosStructEnd   = ElementList()\EndingPos
          ResolvedFile$  = ElementList()\FileName
        EndIf
      ElseIf ColonPos <> 1
        ; Text Mode: traverse project compilation sequences while preserving caller cursors
        PushListPosition(PBB_CompilationSequence())
        PushListPosition(PBB_FileListing())
        FIPS\ScanMode = #FIPSMode_Global
        While FindInWholeProjectSequences(" " + StructName_LCase$, @FIPS, #CheckRightAndComAndQuote)
          LineOfCode$ = ExtractSectionAroundExpression(FIPS\SearchString, StructName$, FIPS\PosInCode)
          While FindString(LineOfCode$, "  ") : LineOfCode$ = ReplaceString(LineOfCode$, "  ", " ") : Wend
          FirstWord$ = LCase(StringField(LineOfCode$, 1, " "))
          ThirdWord$ = LCase(StringField(LineOfCode$, 3, " "))
          If (FirstWord$ = "structure" Or FirstWord$ = "interface") And LCase(StringField(LineOfCode$, 2, " ")) = StructName_LCase$ And
             (ThirdWord$ = "" Or ThirdWord$ = "extends" Or ThirdWord$ = "align")
            DeclMod$ = ""
            If GetEncompassingLimitsAndName(#PBBModule, FIPS\FileName$, FIPS\PosInCode,
                                            FIPS\SearchString, FIPS\SearchString_LCase, FIPS\SegmentEnd, @EncompassingMod) <> ""
              DeclMod$ = EncompassingMod\ElementName$
            EndIf
            If ColonPos > 0
              IsDirectMatch   = Bool(LCase(DeclMod$) = ExplicitMod_LCase$)
              IsFallbackMatch = #False
            ElseIf IsScoped
              IsDirectMatch   = Bool(LCase(DeclMod$) = EnclosingModule$)
              IsFallbackMatch = Bool(IsDirectMatch = #False And ImportedMod_LCase$ <> "" And
                                     LCase(DeclMod$) = ImportedMod_LCase$ And EncompassingMod\TypeOfElementName$ = "DeclareModule")
            Else
              IsDirectMatch   = Bool(DeclMod$ = "")
              IsFallbackMatch = Bool(DeclMod$ <> "")
            EndIf
            If IsDirectMatch Or (IsFallbackMatch And ResolvedType$ = "")
              PosStructStart = FIPS\PosInCode
              PosStructEnd   = FindInPBPointedCode(FIPS\SearchString_LCase, "endstructure,endinterface", PosStructStart, #CheckAll)
              If PosStructEnd : PosStructEnd + Len("endstructure") : Else : PosStructEnd = Len(FIPS\SearchString\s) : EndIf
              If FirstWord$ = "interface" : ResolvedType$ = PBBTypeNames$(#PBBInterface) : Else : ResolvedType$ = PBBTypeNames$(#PBBStructure) : EndIf
              ResolvedType$ + #Space$ + QualifyWithModule(DeclMod$, StructName$)
              ResolvedFile$ = FIPS\FileName$
              If IsDirectMatch : Break : EndIf
            EndIf
          EndIf
          FIPS\FIPSContinue = #True
        Wend
        PopListPosition(PBB_FileListing())
        PopListPosition(PBB_CompilationSequence())
      EndIf
      ;
      If ResolvedType$ = "" And ColonPos = 0 And (EnclosingModule$ = "" Or ModContainer <> 0)
        ResolvedType$ = PBBTypeNames$(#PBBStructure) + #Space$ + StructName$
      EndIf
      If PosStructStart > 0 Or ListModeIsOK Or ForceListMode = #False
        If AddMapElement(StructCache(), CacheKey$)
          StructCache()\Type     = ResolvedType$
          StructCache()\File     = ResolvedFile$
          StructCache()\StartPos = PosStructStart
          StructCache()\EndPos   = PosStructEnd
        EndIf
      EndIf
    EndIf
    ;
    If *OutStartPos : *OutStartPos\i = PosStructStart : EndIf
    If *OutEndPos   : *OutEndPos\i   = PosStructEnd   : EndIf
    If *OutFile     : *OutFile\s     = ResolvedFile$  : EndIf
  EndIf
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDBuildOutExternalProcedureTiming : ReapTimer(3) : EndIf
  CompilerEndIf
  ProcedureReturn ResolvedType$
EndProcedure
;
Procedure.s GetFieldTypeFromStructure(StructureName$, FieldName$, ForceListMode = #False,
                                      EnclosingModule$ = "", FileName$ = "", PosInCode = 0)
  ;
  ; Role: Resolves the concrete type name of 'FieldName$' within 'StructureName$' (including inherited
  ;       'Extends' parent structures in both O(1) List Mode and Text Mode, returning a module-canonical
  ;       "mod::type" when the field's structure resides in a module).
  ;
  Protected PosStructStart, PosStructEnd, FoundPos, FieldIdx, DotPos, StructID, ModContainer, ColonPos, Depth
  Protected StructFile.String
  Protected SearchState.FIPC_KeywordCache
  Protected TargetType$, ResolvedStruct$, StructMod$, LineOfCode$, ExamField$, HeaderToken$
  Protected *SearchString.String, *SearchString_LCase.String
  Protected *Char.Character
  ;
  If StructureName$ = "" Or FieldName$ = "" : ProcedureReturn "" : EndIf
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
    If OPDMod & #OPDBuildOutExternalProcedureTiming : StartTimer(2, "GetFieldTypeFromStructure") : EndIf
  CompilerEndIf
  ;
  ; 1. Fast-path: extract explicit '.Type' or '.Mod::Type' qualification if embedded directly in FieldName$
  DotPos = FindString(FieldName$, ".")
  If DotPos
    *Char       = PosToChar(@FieldName$, DotPos)
    TargetType$ = ExtractOptionalDotType(@FieldName$, @*Char)
    FieldName$  = Trim(Left(FieldName$, DotPos - 1))
  Else
    FieldName$  = Trim(FieldName$)
  EndIf
  ;
  If TargetType$ = ""
    While Left(FieldName$, 1) = "\" Or Left(FieldName$, 1) = "*" : FieldName$ = Mid(FieldName$, 2) : Wend
    FieldName$ = PurifyElementName(StringField(FieldName$, 1, "["))
    If FieldName$ <> ""
      If (ExplorationState = #ExplorationState_Completed And ListSize(ElementList())) Or ForceListMode
        ResolveModContainerID(EnclosingModule$, FileName$, PosInCode)
        If EnclosingModule$ = "" Or ModContainer > 0
          StructID = FindStructureOrInterfaceID(StructureName$, ModContainer, PosInCode, FileName$)
          If StructID > 0 And FindElementInList(FieldName$, #PBBField, 0, "", 0, -1, "", StructID)
            TargetType$ = ElementList()\VarDotType_LCase
          EndIf
        EndIf
      Else
        While StructureName$ <> "" And TargetType$ = "" And Depth < 8
          Depth + 1
          ResolvedStruct$ = ResolveStructureOrInterface(StructureName$, @PosStructStart, @PosStructEnd, @StructFile, #False,
                                                        EnclosingModule$, FileName$, PosInCode)
          StructureName$  = ""
          If ResolvedStruct$ <> "" And PosStructStart > 0
            ColonPos   = FindString(ResolvedStruct$, "::")
            StructMod$ = ""
            If ColonPos : StructMod$ = StringField(Left(ResolvedStruct$, ColonPos - 1), 2, " ") : EndIf
            *SearchString       = GetPointedCodeFromFile(StructFile\s)
            *SearchString_LCase = GetPointedCodeFromFile(StructFile\s, #True)
            If *SearchString And *SearchString_LCase
              FoundPos               = PosStructStart
              SearchState\BufferAddr = 0
              While FoundPos < PosStructEnd
                FoundPos = FindInPBPointedCode(*SearchString_LCase, FieldName$, FoundPos,
                                               #CheckAll, #PB_String_CaseSensitive, PosStructEnd, 0, @SearchState)
                If FoundPos = 0 : Break : EndIf
                LineOfCode$ = ParseArguments(ExtractSectionAroundExpression(*SearchString, FieldName$, FoundPos))
                FieldIdx    = 0
                Repeat
                  FieldIdx + 1
                  ExamField$ = Trim(StringField(LineOfCode$, FieldIdx, #TAB$))
                  If ExamField$ = "" : Break : EndIf
                  DotPos       = FindString(ExamField$, ".")
                  HeaderToken$ = ExamField$
                  If DotPos : HeaderToken$ = Trim(Left(ExamField$, DotPos - 1)) : EndIf
                  If FindString(HeaderToken$, " ")
                    HeaderToken$ = StringField(HeaderToken$, CountString(HeaderToken$, " ") + 1, " ")
                  EndIf
                  If Left(HeaderToken$, 1) = "*" : HeaderToken$ = Mid(HeaderToken$, 2) : EndIf
                  If PurifyElementName(StringField(HeaderToken$, 1, "[")) = FieldName$
                    If DotPos
                      *Char       = PosToChar(@ExamField$, DotPos)
                      TargetType$ = ExtractOptionalDotType(@ExamField$, @*Char)
                      If TargetType$ <> "" And FindString(TargetType$, "::") = 0 And StructMod$ <> "" And
                         FindMapElement(PBBasicNativeTypes(), LCase(TargetType$)) = 0
                        TargetType$ = StringField(ResolveStructureOrInterface(TargetType$, 0, 0, 0, #False,
                                                                              StructMod$, StructFile\s, FoundPos), 2, " ")
                      EndIf
                    EndIf
                    Break 3
                  EndIf
                ForEver
                FoundPos + Len(FieldName$)
              Wend
              ; Climb to any inherited 'Extends <Parent>' structure in Text Mode when field was not declared locally
              LineOfCode$ = ExtractSectionAroundExpression(*SearchString,
                                                           StringField(Mid(ResolvedStruct$, FindString(ResolvedStruct$, " ") + 1),
                                                                       1 + Bool(ColonPos > 0), "::"), PosStructStart)
              DotPos      = FindString(LCase(LineOfCode$), " extends ")
              If DotPos
                *Char           = PosToChar(@LineOfCode$, DotPos + 9)
                StructureName$  = ExtractOrSkipNextToken(@LineOfCode$, @*Char)
                If IsModuleUsage(@*Char) : StructureName$ + "::" + ExtractOrSkipNextToken(@LineOfCode$, @*Char) : EndIf
                EnclosingModule$ = StructMod$
                FileName$        = StructFile\s
                PosInCode        = PosStructStart
              EndIf
            EndIf
          EndIf
        Wend
      EndIf
    EndIf
  EndIf
  ;
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
    If OPDMod & #OPDBuildOutExternalProcedureTiming : ReapTimer(2) : EndIf
  CompilerEndIf
  ProcedureReturn TargetType$
EndProcedure
;
Procedure.s GetStructureOrInterfaceNameFromVariable(VarName$, MainParent$, FileName$, PosInCode,
                                                    StartLimit = 0, ForceListMode = #False,
                                                    MacroProcID = -1, MacroProcParams$ = "",
                                                    EnclosingModule$ = "")
  ;
  ; Role: Resolves the concrete entity type ("Structure [Mod::]X" or "Interface [Mod::]X") of a
  ;       structured variable or dereference chain across both O(1) List Mode and Text Mode.
  ;
  Shared PBB_CompilationSequence(), PBB_FileListing()
  Static LastFileName$, LastStartLimit, LastVarName$, LastMainParent$, LastEnclosingMod$, LastResult$
  ;
  Protected FIPS.FIPSPosStruct
  Protected EncompassingMod.EEDStruct
  Protected TargetVarDot$, StructName$, ResolvedEntity$, RootVar$, TypedRoot$, DeclMod$, ExplicitMod$
  Protected CleanRootVar_LCase$, BareRootVar_LCase$, ParentNamePurified$
  Protected Segments, SegmentIdx, DotPos, ColonPos, ParamPos, ParamEndPos, ModContainer
  Protected *ProcElem.Elements, *DotChar.Character
  ;
  If VarName$ = ""
    LastFileName$ = "" : LastStartLimit = 0 : LastVarName$ = "" : LastMainParent$ = "" : LastEnclosingMod$ = "" : LastResult$ = ""
    ProcedureReturn ""
  EndIf
  If ForceListMode = #False
    If FileName$ = LastFileName$ And StartLimit = LastStartLimit And VarName$ = LastVarName$ And
       MainParent$ = LastMainParent$ And EnclosingModule$ = LastEnclosingMod$
      ProcedureReturn LastResult$
    EndIf
    LastFileName$ = FileName$ : LastStartLimit = StartLimit : LastVarName$ = VarName$
    LastMainParent$ = MainParent$ : LastEnclosingMod$ = EnclosingModule$ : LastResult$ = ""
  EndIf
  ;
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
    If OPDMod & #OPDBuildOutExternalProcedureTiming : StartTimer(1, "GetStructureOrInterfaceNameFromVariable") : EndIf
  CompilerEndIf
  ;
  Segments = CountString(VarName$, "\") + 1
  RootVar$ = Trim(StringField(VarName$, 1, "\"))
  ColonPos = FindString(RootVar$, "::")
  If ColonPos > 1 And Left(RootVar$, 1) = "*"
    ; Normalize '*Mod::Ptr' into 'Mod::*Ptr' for both List Mode and Text Mode
    RootVar$ = Mid(RootVar$, 2, ColonPos - 2) + "::*" + Mid(RootVar$, ColonPos + 2)
  EndIf
  ;
  TypedRoot$ = Trim(StringField(MainParent$, 1, "\"))
  If TypedRoot$ = "" Or FindString(TypedRoot$, " ") Or FindString(TypedRoot$, ".") = 0
    If FindString(RootVar$, ".") : TypedRoot$ = RootVar$ : Else : TypedRoot$ = "" : EndIf
  EndIf
  ;
  ; 1. Base root variable typing (shared inline '.Type' fast-path, then List Mode vs. Text Mode)
  If TypedRoot$ <> ""
    DotPos      = FindString(TypedRoot$, ".")
    *DotChar    = PosToChar(@TypedRoot$, DotPos)
    StructName$ = ExtractOptionalDotType(@TypedRoot$, @*DotChar)
  ElseIf (ExplorationState = #ExplorationState_Completed And ListSize(ElementList())) Or ForceListMode
    CleanRootVar_LCase$ = PurifyElementName(RootVar$)
    ColonPos            = FindString(CleanRootVar_LCase$, "::")
    If ColonPos > 0 : BareRootVar_LCase$ = FastMid(CleanRootVar_LCase$, ColonPos + 2) : Else : BareRootVar_LCase$ = CleanRootVar_LCase$ : EndIf
    If EnclosingModule$ <> "" And ColonPos = 0
      ModContainer = GetModuleID(EnclosingModule$)
      If ModContainer = 0 : CleanRootVar_LCase$ = "" : EndIf
    EndIf
    If CleanRootVar_LCase$ <> ""
      If MacroProcID <= 0 And MainParent$ <> ""
        ParentNamePurified$ = PurifyElementName(Mid(MainParent$, FindString(MainParent$, " ") + 1))
        If FindElementInList(ParentNamePurified$, #PBBProcedure, 0, "", 0, -1, "", 0, ModContainer) Or
           FindElementInList(ParentNamePurified$, #PBBMacro, 0, "", 0, -1, "", 0, ModContainer)
          MacroProcID      = ElementList()\FixedID
          MacroProcParams$ = ElementList()\Parameters
        EndIf
      ElseIf MacroProcID > 0 And MacroProcParams$ = "" And SelectPBBElement(MacroProcID)
        MacroProcParams$ = ElementList()\Parameters
      EndIf
      If MacroProcID > 0 And MacroProcID < NextElementID
        *ProcElem = ElementFixedIndexes(MacroProcID)
        If *ProcElem And ColonPos = 0 And ModContainer = 0
          ModContainer     = *ProcElem\EnclosingModuleID
          EnclosingModule$ = GetEnclosingModuleName(ModContainer)
        EndIf
      ElseIf ColonPos = 0 And ModContainer = 0 And FileName$ <> ""
        ModContainer = -1
      EndIf
      If ColonPos = 0 And MacroProcParams$ <> ""
        ParamPos = FindString(LCase(MacroProcParams$), #TAB$ + BareRootVar_LCase$ + ".")
        If ParamPos > 0
          ParamPos + Len(#TAB$ + BareRootVar_LCase$ + ".")
          ParamEndPos = FindString(MacroProcParams$, #TAB$, ParamPos)
          If ParamEndPos > ParamPos : StructName$ = StringField(FastMid(MacroProcParams$, ParamPos, ParamEndPos - ParamPos), 1, "(") : EndIf
        EndIf
      EndIf
      If StructName$ = "" And FindElementInList(CleanRootVar_LCase$, 0, PosInCode, FileName$, MacroProcID,
                                                Bool(FindString(RootVar$, "(") > 0), "", 0, ModContainer)
        If ElementList()\Type = #PBBStructure Or ElementList()\Type = #PBBInterface
          StructName$ = QualifyWithModule(GetEnclosingModuleName(ElementList()\EnclosingModuleID), ElementList()\Name)
        ElseIf ElementList()\Type = #PBBVariable
          StructName$      = ElementList()\VarDotType_LCase
          EnclosingModule$ = GetEnclosingModuleName(ElementList()\EnclosingModuleID)
        EndIf
      EndIf
    EndIf
  Else
    ; Standalone text-scanning mode when background exploration is still pending
    ColonPos = FindString(RootVar$, "::")
    If ColonPos > 0
      ExplicitMod$ = LCase(Trim(Left(RootVar$, ColonPos - 1)))
      RootVar$     = Mid(RootVar$, ColonPos + 2)
    EndIf
    PushListPosition(PBB_CompilationSequence())
    PushListPosition(PBB_FileListing())
    TargetVarDot$  = PurifyForSearch(RootVar$) + "."
    FIPS\FileName$ = FileName$
    FIPS\EndLimit  = PosInCode
    If StartLimit > 0 And PosInCode > 0 And ExplicitMod$ = ""
      FIPS\PosInCode = StartLimit
      FIPS\ScanMode  = #FIPSMode_Hybrid
    Else
      FIPS\PosInCode = 0
      FIPS\ScanMode  = #FIPSMode_Global
    EndIf
    While FindInWholeProjectSequences(TargetVarDot$, @FIPS, #CheckLeftAndComAndQuote)
      DeclMod$ = ""
      If GetEncompassingLimitsAndName(#PBBModule, FIPS\FileName$, FIPS\PosInCode,
                                      FIPS\SearchString, FIPS\SearchString_LCase, FIPS\SegmentEnd, @EncompassingMod) <> ""
        DeclMod$ = EncompassingMod\ElementName$
      EndIf
      If (ExplicitMod$ <> "" And LCase(DeclMod$) <> ExplicitMod$) Or
         (ExplicitMod$ = "" And FIPS\IsLocalScan = #False And LCase(DeclMod$) <> LCase(EnclosingModule$) And
          (DeclMod$ = "" Or EncompassingMod\TypeOfElementName$ <> "DeclareModule" Or
           CheckIfModuleIsUsed(DeclMod$, FileName$, PosInCode, "", EnclosingModule$) = ""))
        FIPS\FIPSContinue = #True
        Continue
      EndIf
      *DotChar         = PosToChar(@FIPS\SearchString\s, FIPS\PosInCode + Len(TargetVarDot$) - 1)
      StructName$      = ExtractOptionalDotType(@FIPS\SearchString\s, @*DotChar)
      EnclosingModule$ = DeclMod$
      Break
    Wend
    PopListPosition(PBB_FileListing())
    PopListPosition(PBB_CompilationSequence())
  EndIf
  ;
  ; 2. Descend member hierarchy when resolving chained dereferences and resolve final entity species
  If StructName$ <> "" And Segments > 1
    For SegmentIdx = 2 To Segments
      StructName$ = GetFieldTypeFromStructure(StructName$, StringField(VarName$, SegmentIdx, "\"),
                                              ForceListMode, EnclosingModule$, FileName$, PosInCode)
      If StructName$ = "" : Break : EndIf
    Next SegmentIdx
  EndIf
  If StructName$ <> ""
    ResolvedEntity$ = ResolveStructureOrInterface(StructName$, 0, 0, 0, ForceListMode, EnclosingModule$, FileName$, PosInCode)
    If ForceListMode = #False : LastResult$ = ResolvedEntity$ : EndIf
  EndIf
  ;
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
    If OPDMod & #OPDBuildOutExternalProcedureTiming : ReapTimer(1) : EndIf
  CompilerEndIf
  ProcedureReturn ResolvedEntity$
EndProcedure
;
Procedure   ResolveFieldTargetStructID(CandidateIDs$, *BackslashChar.Character, *BufferStart,
                                       WithVar$, FileName$, PosInCode, StartLimit, MacroProcID,
                                       MacroProcParams$ = "", ModuleContainer = -1)
  ;
  ; Role: When multiple structures or interfaces define a homonymous field (#PBBField) in 'CandidateIDs$',
  ; extracts the variable chain preceding '\' and resolves the FixedID of its target Structure or Interface.
  ;
  Protected IDIndex, FieldCandidateCount
  Protected IDStr$, PrecedingChain$
  Protected *Elem.Elements
  ;
  If *BackslashChar >= *BufferStart And FindString(CandidateIDs$, ",", FindString(CandidateIDs$, ",") + 1)
    Repeat
      IDIndex + 1
      IDStr$ = StringField(CandidateIDs$, IDIndex, ",")
      If IDStr$ = "" : Break : EndIf
      *Elem = ElementFixedIndexes(Val(IDStr$))
      If *Elem And *Elem\StartingPos > 0 And *Elem\Type = #PBBField
        FieldCandidateCount + 1
        If FieldCandidateCount > 1
          PrecedingChain$ = ExtractPrecedingFieldChain(*BackslashChar, *BufferStart, WithVar$)
          If PrecedingChain$ <> ""
            ProcedureReturn FindStructureOrInterfaceID(GetStructureOrInterfaceNameFromVariable(PrecedingChain$, "", FileName$, PosInCode, StartLimit, #True,
                                                                                               MacroProcID, MacroProcParams$, GetEnclosingModuleName(ModuleContainer)))
          EndIf
          Break
        EndIf
      EndIf
    ForEver
  EndIf
  ProcedureReturn 0
EndProcedure
;
; **********************************************************************************
;
;-             Compile-time evaluator, CompilerIf & include tree
;
; **********************************************************************************
;
Structure ProcVarEntry
  DeclareLabel.s ; Declaration category (e.g., "Parameter", "Protected", "Global List", "Unprotected")
  DisplayName.s  ; Formatted identifier including optional '.Type' and '()' suffixes
  DotTypeLCase.s ; Lowercase dot-type without leading dot for O(1) structure field resolution
  IsUsed.b       ; #True once referenced in an expression or active ReDim/Define, #False otherwise
EndStructure
;
Structure CompilerEvalContext
  InListVariables.b
  InCompilerDirective.b
  InMacroDef.b
  MarkExprUsages.b
  Stage1Only.b
  TargetType.i
  ActiveOwnerID.i
  ProcContainer.i
  ModuleContainer.i
  PosInCode.i
  FileName.s
  ProcedureName.s
  Map ProcVarsMap.i()
EndStructure
;
Global EvalContext.CompilerEvalContext
;
Procedure RegisterElementUsage(OwnerID, TargetID, FileName$, PosInCode)
  ;
  ; Role: Records a main-scope reference (-1 and first OOE_UsagePos) when OwnerID <= 0,
  ;       or links OwnerID -> TargetID in O(1) via ElementFixedIndexes() without moving ElementList().
  ;
  Protected *Elem.Elements
  ;
  If TargetID <= 0 Or TargetID >= NextElementID : ProcedureReturn : EndIf
  If OwnerID <= 0
    *Elem = ElementFixedIndexes(TargetID)
    If *Elem
      *Elem\UsageParents = AddIndexToList(-1, *Elem\UsageParents)
      If *Elem\OOE_UsagePos = "" And FileName$ <> "" And PosInCode > 0
        *Elem\OOE_UsagePos = FileName$ + "," + Str(PosInCode) + #CR$
      EndIf
    EndIf
  ElseIf OwnerID <> TargetID
    LinkElementUsage(OwnerID, TargetID)
  EndIf
EndProcedure
;
Global EvError, EvIsString, EvResultStr$
Declare EvaluateExpression(Expression$)
;
Procedure.i DecodePBDefinedType(TypeConstUCase$, *OutHasParen.Integer = 0, *OutSubSpecie.String = 0)
  ;
  ; Role: Maps a PureBasic 'Defined()' second-argument constant ("#PB_VARIABLE", "#PB_LIST", "#PB_PROCEDURE"...)
  ;       to its PBBrowser element type (#PBB...), writing parenthesis and subspecies traits when requested.
  ;
  Protected Sub$ = "", HasParen = -1, ElemType = 0, ct
  ;
  Select TypeConstUCase$
    Case "#PB_VARIABLE"              : ElemType = #PBBVariable : Sub$ = "Var"   : HasParen = 0
    Case "#PB_ARRAY"                 : ElemType = #PBBVariable : Sub$ = "Array" : HasParen = 1
    Case "#PB_LIST", "#PB_LINKEDLIST": ElemType = #PBBVariable : Sub$ = "List"  : HasParen = 1
    Case "#PB_MAP"                   : ElemType = #PBBVariable : Sub$ = "Map"   : HasParen = 1
    Case "#PB_PROCEDURE"             : ElemType = #PBBProcedure                 : HasParen = 1
    Case "#PB_CONSTANT"              : ElemType = #PBBConstant                  : HasParen = 0
    Default
      If Left(TypeConstUCase$, 4) = "#PB_"
        For ct = #PBBMacro To #PBBLabel
          If UCase(PBBTypeNames$(ct)) = Mid(TypeConstUCase$, 5)
            ElemType = ct
            If ct <> #PBBMacro : HasParen = 0 : EndIf
            Break
          EndIf
        Next
      EndIf
  EndSelect
  If *OutHasParen  : *OutHasParen\i  = HasParen : EndIf
  If *OutSubSpecie : *OutSubSpecie\s = Sub$     : EndIf
  ProcedureReturn ElemType
EndProcedure
;
Procedure.s EE_UnknownResolver(IdentifierName$, Arguments$ = "")
  ;
  ; Role: Resolves PureBasic compile-time constants (#PB_Compiler_..., native constants, project #PBBConstant)
  ;       and compile-time functions (Defined, SizeOf, Chr, Asc) for EvaluateExpression() while preserving ElementList().
  ;
  Protected *NativeConst.ConstListStruct, *VarEntry.ProcVarEntry
  Protected SubSpecieStr.String
  Protected BareIdent$, IdentUCase$, ParsedArgs$, TargetLCase$, TypeConst$, Result$
  Protected ColonPos = FindString(IdentifierName$, "::"), IsQual, IsDefined, HasParen, ctType, SubVal
  ;
  PushListPosition(ElementList())
  If ColonPos > 0 : BareIdent$ = Mid(IdentifierName$, ColonPos + 2) : Else : BareIdent$ = IdentifierName$ : EndIf
  IdentUCase$ = UCase(BareIdent$)
  ;
  ; 1. Constant resolution (dynamic #PB_Compiler_..., native *PBNativeConstantMap(), and project #PBBConstant)
  If Left(BareIdent$, 1) = "#"
    If ColonPos = 0
      Select IdentUCase$
        Case "#TRUE"                         : Result$ = "1"
        Case "#FALSE", "#NULL"               : Result$ = "0"
        Case "#PB_COMPILER_ISMAINFILE"       : Result$ = Str(Bool(EvalContext\FileName = "" Or PBBFicPrincipalPB$ = "" Or LCase(EvalContext\FileName) = LCase(PBBFicPrincipalPB$)))
        Case "#PB_COMPILER_ISINCLUDEFILE"    : Result$ = Str(Bool(EvalContext\FileName <> "" And PBBFicPrincipalPB$ <> "" And LCase(EvalContext\FileName) <> LCase(PBBFicPrincipalPB$)))
        Case "#PB_COMPILER_OS"               : Result$ = Str(#PB_Compiler_OS)
        Case "#PB_COMPILER_PROCESSOR"        : Result$ = Str(#PB_Compiler_Processor)
        Case "#PB_COMPILER_BACKEND"          : Result$ = Str(#PB_Compiler_Backend)
        Case "#PB_COMPILER_EXECUTABLEFORMAT" : Result$ = Str(#PB_Compiler_ExecutableFormat)
        Case "#PB_COMPILER_VERSION"          : Result$ = Str(#PB_Compiler_Version)
        Case "#PB_COMPILER_64BIT"            : Result$ = Str(#PB_Compiler_64Bit)
        Case "#PB_COMPILER_32BIT"            : Result$ = Str(#PB_Compiler_32Bit)
        Case "#PB_COMPILER_UNICODE"          : Result$ = Str(#PB_Compiler_Unicode)
        Case "#PB_COMPILER_THREAD"           : Result$ = Str(PBB_ProjectThread)
        Case "#PB_COMPILER_DEBUGGER"         : Result$ = Str(PBB_ProjectDebugger)
        Case "#PB_COMPILER_DPIAWARE"         : Result$ = Str(#PB_Compiler_DPIAware)
        Case "#PB_COMPILER_LINE"             : If EvalContext\FileName <> "" And EvalContext\PosInCode > 0 : Result$ = Str(GetNoLineFromFilenameAndPos(EvalContext\FileName, EvalContext\PosInCode)) : Else : Result$ = "1" : EndIf
        Case "#PB_COMPILER_PROCEDURE"        : If EvalContext\Stage1Only = #False And EvalContext\InMacroDef = #False : Result$ = #DOUBLEQUOTE$ + EvalContext\ProcedureName + #DOUBLEQUOTE$ : EndIf
        Case "#PB_COMPILER_MODULE"           : If EvalContext\Stage1Only = #False And EvalContext\InMacroDef = #False : Result$ = #DOUBLEQUOTE$ + GetEnclosingModuleName(EvalContext\ModuleContainer) + #DOUBLEQUOTE$ : EndIf
        Case "#PB_COMPILER_FILE"             : Result$ = #DOUBLEQUOTE$ + EvalContext\FileName + #DOUBLEQUOTE$
        Case "#PB_COMPILER_FILEPATH"         : Result$ = #DOUBLEQUOTE$ + GetPathPart(EvalContext\FileName) + #DOUBLEQUOTE$
        Case "#PB_COMPILER_FILENAME"         : Result$ = #DOUBLEQUOTE$ + GetFilePart(EvalContext\FileName) + #DOUBLEQUOTE$
      EndSelect
      If Result$ = "" And FindMapElement(*PBNativeConstantMap(), IdentUCase$)
        *NativeConst = *PBNativeConstantMap()
        If *NativeConst\ConstantType = 2
          Result$ = #DOUBLEQUOTE$ + *NativeConst\ConstantValueStr + #DOUBLEQUOTE$
        ElseIf *NativeConst\ConstantType = 1
          Result$ = Str(Int(*NativeConst\ConstantValueFloat))
        Else
          Result$ = Str(*NativeConst\ConstantValueInt)
        EndIf
      EndIf
    EndIf
    ;
    If Result$ = ""
      If EvalContext\InCompilerDirective : PBB_CompilerGuardSymbols(LCase(BareIdent$)) = #True : EndIf
      If EvalContext\Stage1Only
        If ColonPos = 0 And FindMapElement(PBB_Stage1Constants(), LCase(BareIdent$)) : Result$ = PBB_Stage1Constants() : EndIf
      Else
        If EvalContext\MarkExprUsages And ColonPos > 1
          RegisterElementUsage(EvalContext\ActiveOwnerID, GetModuleID(Left(IdentifierName$, ColonPos - 1)), EvalContext\FileName, EvalContext\PosInCode)
        EndIf
        If FindElementInList(LCase(IdentifierName$), #PBBConstant, EvalContext\PosInCode, EvalContext\FileName, 0, -1, "", 0, EvalContext\ModuleContainer)
          If EvalContext\MarkExprUsages : RegisterElementUsage(EvalContext\ActiveOwnerID, ElementList()\FixedID, EvalContext\FileName, EvalContext\PosInCode) : EndIf
          If ElementList()\Value <> "" And ElementList()\Value <> "UnResolved" : Result$ = ElementList()\Value : EndIf
        EndIf
      EndIf
    EndIf
    ;
  ; 2. Compile-time functions (Defined, SizeOf, Chr, Asc) and inspected Macro formal parameters
  ElseIf ColonPos = 0
    Select IdentUCase$
      Case "DEFINED"
        ParsedArgs$  = ParseArguments(Arguments$)
        TargetLCase$ = PurifyElementName(Trim(StringField(ParsedArgs$, 1, #TAB$)))
        TypeConst$   = UCase(Trim(StringField(ParsedArgs$, 2, #TAB$)))
        If TargetLCase$ <> "" And Left(TypeConst$, 4) = "#PB_"
          IsQual = FindString(TargetLCase$, "::")
          If IsQual > 0 : BareIdent$ = Mid(TargetLCase$, IsQual + 2) : Else : BareIdent$ = TargetLCase$ : EndIf
          If TypeConst$ = "#PB_CONSTANT" And Left(BareIdent$, 1) <> "#"
            BareIdent$   = "#" + BareIdent$
            TargetLCase$ = Left(TargetLCase$, IsQual + Bool(IsQual > 0)) + BareIdent$
          EndIf
          If EvalContext\InCompilerDirective : PBB_CompilerGuardSymbols(BareIdent$) = #True : EndIf
          ;
          If EvalContext\InMacroDef = #False
            If TypeConst$ = "#PB_FUNCTION"
              Result$ = Str(Bool(IsQual = 0 And FindMapElement(PBNativeFunctions(), TargetLCase$) <> 0))
            ElseIf TypeConst$ = "#PB_OSFUNCTION"
              Result$ = Str(Bool(IsQual = 0 And Right(TargetLCase$, 1) = "_"))
            ElseIf TypeConst$ = "#PB_CONSTANT"
              IsDefined = Bool(IsQual = 0 And (TargetLCase$ = "#true" Or TargetLCase$ = "#false" Or TargetLCase$ = "#null" Or
                                               Left(TargetLCase$, 13) = "#pb_compiler_" Or FindMapElement(*PBNativeConstantMap(), UCase(TargetLCase$))))
              If IsDefined = #False
                If EvalContext\Stage1Only
                  If IsQual = 0 And FindMapElement(PBB_Stage1Constants(), TargetLCase$) : Result$ = "1" : EndIf
                Else
                  IsDefined = Bool(FindElementInList(TargetLCase$, #PBBConstant, EvalContext\PosInCode, EvalContext\FileName, 0, -1, "", 0, EvalContext\ModuleContainer) > 0)
                EndIf
              EndIf
              If IsDefined Or EvalContext\Stage1Only = #False : Result$ = Str(IsDefined) : EndIf
            ElseIf EvalContext\Stage1Only = #False
              ctType = DecodePBDefinedType(TypeConst$, @HasParen, @SubSpecieStr)
              If ctType = #PBBVariable
                If IsQual = 0 And EvalContext\InListVariables And FindMapElement(EvalContext\ProcVarsMap(), TargetLCase$ + Left("()", HasParen * 2))
                  *VarEntry = EvalContext\ProcVarsMap()
                  IsDefined = Bool(HasParen = 0 Or FindString(*VarEntry\DeclareLabel, SubSpecieStr\s) > 0)
                ElseIf IsQual = 0 And EvalContext\InListVariables And EvalContext\TargetType = #PBBProcedure
                  IsDefined = Bool(FindElementInList(TargetLCase$, #PBBVariable, EvalContext\PosInCode, EvalContext\FileName, -1, HasParen, "", 0, EvalContext\ModuleContainer, "Global", SubSpecieStr\s) > 0 Or
                                   FindElementInList(TargetLCase$, #PBBVariable, EvalContext\PosInCode, EvalContext\FileName, -1, HasParen, "", 0, EvalContext\ModuleContainer, "Threaded", SubSpecieStr\s) > 0)
                Else
                  IsDefined = Bool(FindElementInList(TargetLCase$, #PBBVariable, EvalContext\PosInCode, EvalContext\FileName, -1, HasParen, "", 0, EvalContext\ModuleContainer, "", SubSpecieStr\s) > 0)
                EndIf
                Result$ = Str(IsDefined)
              ElseIf ctType > 0
                If ctType = #PBBModule
                  IsDefined = Bool(IsQual = 0 And FindElementInList(TargetLCase$, #PBBModule, EvalContext\PosInCode, EvalContext\FileName) > 0)
                ElseIf ctType = #PBBLabel
                  IsDefined = Bool(FindElementInList(TargetLCase$, #PBBLabel, 0, EvalContext\FileName, EvalContext\ProcContainer, -1, "", 0, EvalContext\ModuleContainer) > 0)
                ElseIf FindElementInList(TargetLCase$, ctType, EvalContext\PosInCode, EvalContext\FileName, 0, -1, "", 0, EvalContext\ModuleContainer) > 0
                  IsDefined = Bool(ctType <> #PBBMacro Or EvalContext\PosInCode <= 0 Or
                                   IsPositionBefore(ElementList()\FileName, ElementList()\EndingPos, EvalContext\FileName, EvalContext\PosInCode))
                ElseIf ctType = #PBBStructure And IsQual = 0
                  IsDefined = Bool(FindMapElement(PBBasicNativeTypes(), TargetLCase$) <> 0)
                EndIf
                Result$ = Str(IsDefined)
              EndIf
            EndIf
          EndIf
        EndIf
      Case "SIZEOF"
        ; Intentional architectural boundary: only scalar native types and pointers are evaluated here;
        ; project structures/interfaces remain unresolved ("") while OnePass tracks their usage.
        TargetLCase$ = LCase(Trim(Arguments$))
        If Left(TargetLCase$, 1) = "*"
          Result$ = Str(SizeOf(Integer))
        Else
          Select TargetLCase$
            Case "b", "byte", "a", "ascii"   : Result$ = "1"
            Case "w", "word", "u", "unicode" : Result$ = "2"
            Case "c", "character"            : Result$ = Str(SizeOf(Character))
            Case "l", "long", "f", "float"   : Result$ = "4"
            Case "i", "integer"              : Result$ = Str(SizeOf(Integer))
            Case "q", "quad", "d", "double"  : Result$ = "8"
          EndSelect
        EndIf
      Case "CHR"
        If Arguments$ <> ""
          SubVal = EvaluateExpression(Arguments$)
          If EvError = 0 And SubVal >= 0 And SubVal <= $10FFFF : Result$ = #DOUBLEQUOTE$ + Chr(SubVal) + #DOUBLEQUOTE$ : EndIf
          EvError = 0
        EndIf
      Case "ASC"
        If Arguments$ <> ""
          EvaluateExpression(Arguments$)
          If EvError = 0 And EvIsString : Result$ = Str(Asc(EvResultStr$)) : EndIf
          EvError = 0
        EndIf
      Default
        If Arguments$ = "" And EvalContext\InListVariables And EvalContext\TargetType = #PBBMacro And
           FindMapElement(EvalContext\ProcVarsMap(), LCase(BareIdent$))
          *VarEntry = EvalContext\ProcVarsMap()
          If *VarEntry\DeclareLabel = "Parameter" : *VarEntry\IsUsed = #True : EndIf
        EndIf
    EndSelect
  EndIf
  ;
  PopListPosition(ElementList())
  ProcedureReturn Result$
EndProcedure
;
XIncludeFile "../../Zapman libraries/ExpressionEvaluator.pbi"
;
Procedure.s EvalConstExpr(RawExpr$, FileName$ = "", PosInCode = 0, ModuleContainer = -1, Stage1Only = #False, IsCompilerGuard = #True)
  ;
  ; Role: Evaluates a compile-time expression via EvaluateExpression() and returns its canonical literal:
  ;       - "" (empty) when undetermined or invalid (EvError <> 0),
  ;       - #DOUBLEQUOTE$ + Text$ + #DOUBLEQUOTE$ for string expressions (e.g., "" becomes Chr(34)+Chr(34)),
  ;       - Str(ValInt) for integer/boolean expressions.
  ;       Restores EvalContext\FileName, PosInCode, and ModuleContainer for nested ListVariables() calls.
  ;
  Protected SavedFile$ = EvalContext\FileName, SavedPos = EvalContext\PosInCode, SavedMod = EvalContext\ModuleContainer
  Protected ValInt
  Protected Result$
  ;
  RawExpr$ = Trim(RawExpr$)
  If RawExpr$ <> ""
    EvalContext\Stage1Only          = Stage1Only
    EvalContext\InCompilerDirective = IsCompilerGuard
    EvalContext\FileName            = FileName$
    EvalContext\PosInCode           = PosInCode
    EvalContext\ModuleContainer     = ModuleContainer
    ValInt                          = EvaluateExpression(RawExpr$)
    EvalContext\Stage1Only          = #False
    EvalContext\InCompilerDirective = #False
    EvalContext\FileName            = SavedFile$
    EvalContext\PosInCode           = SavedPos
    EvalContext\ModuleContainer     = SavedMod
    If EvError = 0
      If EvIsString
        Result$ = #DOUBLEQUOTE$ + EvResultStr$ + #DOUBLEQUOTE$
      Else
        Result$ = Str(ValInt)
      EndIf
    EndIf
  EndIf
  ProcedureReturn Result$
EndProcedure
;
Macro       ConstValToTriState(ValStr)
  ;
  ; Role: Maps a canonical EvalConstExpr() string to #Cond_Undetermined (-1), #Cond_False (0), or #Cond_True (1).
  ;
  (Bool((ValStr) = "") * #Cond_Undetermined + Bool((ValStr) <> "" And (ValStr) <> "0" And (ValStr) <> (#DOUBLEQUOTE$ + #DOUBLEQUOTE$)) * #Cond_True)
EndMacro
;
Procedure.i StepCompilerDirective(DirectiveLCase$, RawExpr$, Array Stack.CompilerStackLevel(1), *Depth.Integer,
                                  FileName$, TokStartPos, AfterExprPos, EvalPos = 0, ModuleContainer = -1,
                                  Stage1Only = #False, *DeadStartPos.Integer = 0, RecordDeadZone = #False)
  ;
  ; Role: Advances the unified CompilerIf / CompilerSelect tri-state stack by one directive,
  ;       centrally opens/closes '*DeadStartPos\i' intervals in PBB_FileListing()\DeadZones() so evaluated
  ;       directive headers are never marked dead while dead bodies and short-circuited branches are,
  ;       and returns #True when the resulting branch is active (or undetermined), #False when dead.
  ;
  Protected Depth = *Depth\i, ParentActive = #True, Evaluated = #False, IsSelectPair, IsActive
  Protected CondState, CaseIdx, ToPos, TargetInt, LowInt, HighInt, HasUndeterminedCase
  Protected Val$, LowVal$, HighVal$
  Protected CaseItem.String
  ;
  If EvalPos <= 0 : EvalPos = TokStartPos : EndIf
  If Depth > 0    : ParentActive = Stack(Depth)\BranchActive : EndIf
  IsSelectPair = Bool(DirectiveLCase$ = "compilerdefault" Or DirectiveLCase$ = "compilerendselect")
  ;
  Select DirectiveLCase$
    Case "compilerif"
      Evaluated = ParentActive
      If Depth < #PBB_MaxCompilerDepth
        Depth + 1
        ClearStructure(@Stack(Depth), CompilerStackLevel)
        Stack(Depth)\ParentActive = ParentActive
        If ParentActive = #False
          Stack(Depth)\HasMatched = #Cond_True
        Else
          Val$                      = EvalConstExpr(RawExpr$, FileName$, EvalPos, ModuleContainer, Stage1Only)
          CondState                 = ConstValToTriState(Val$)
          Stack(Depth)\BranchActive = Bool(CondState <> #Cond_False)
          Stack(Depth)\HasMatched   = CondState
        EndIf
      EndIf
      ;
    Case "compilerelseif"
      If Depth > 0 And Stack(Depth)\IsSelect = #False
        Evaluated = Bool(Stack(Depth)\ParentActive And Stack(Depth)\HasMatched <> #Cond_True)
        If Evaluated = #False
          Stack(Depth)\BranchActive = #False
        Else
          Val$                      = EvalConstExpr(RawExpr$, FileName$, EvalPos, ModuleContainer, Stage1Only)
          CondState                 = ConstValToTriState(Val$)
          Stack(Depth)\BranchActive = Bool(CondState <> #Cond_False)
          If CondState = #Cond_Undetermined Or (CondState = #Cond_True And Stack(Depth)\HasMatched = #Cond_False)
            Stack(Depth)\HasMatched = CondState
          EndIf
        EndIf
      EndIf
      ;
    Case "compilerselect"
      Evaluated = ParentActive
      If Depth < #PBB_MaxCompilerDepth
        Depth + 1
        ClearStructure(@Stack(Depth), CompilerStackLevel)
        Stack(Depth)\IsSelect     = #True
        Stack(Depth)\ParentActive = ParentActive
        If ParentActive
          Stack(Depth)\SelectVal = EvalConstExpr(RawExpr$, FileName$, EvalPos, ModuleContainer, Stage1Only)
        EndIf
      EndIf
      ;
    Case "compilercase"
      If Depth > 0 And Stack(Depth)\IsSelect
        Evaluated                 = Bool(Stack(Depth)\ParentActive And Stack(Depth)\HasMatched <> #Cond_True)
        Stack(Depth)\BranchActive = #False
        If Evaluated
          If Stack(Depth)\SelectVal = ""
            Stack(Depth)\BranchActive = #True
            Stack(Depth)\HasMatched   = #Cond_Undetermined
          Else
            RawExpr$  = ParseArguments(RawExpr$)
            TargetInt = Val(Stack(Depth)\SelectVal)
            Repeat
              CaseIdx + 1
              CaseItem\s = Trim(StringField(RawExpr$, CaseIdx, #TAB$))
              If CaseItem\s = "" : Break : EndIf
              ; Support both 'CompilerCase A' and range 'CompilerCase A To B'
              ToPos = FindInPBPointedCode(@CaseItem, "to", 1, #CheckAll | #CheckOneKeyWordOnly, #PB_String_NoCase)
              If ToPos > 1
                LowVal$  = EvalConstExpr(Left(CaseItem\s, ToPos - 1), FileName$, EvalPos, ModuleContainer, Stage1Only)
                HighVal$ = EvalConstExpr(Mid(CaseItem\s, ToPos + 2), FileName$, EvalPos, ModuleContainer, Stage1Only)
                If LowVal$ = "" Or HighVal$ = ""
                  HasUndeterminedCase = #True
                ElseIf Left(Stack(Depth)\SelectVal, 1) <> #DOUBLEQUOTE$ And Left(LowVal$, 1) <> #DOUBLEQUOTE$ And Left(HighVal$, 1) <> #DOUBLEQUOTE$
                  LowInt  = Val(LowVal$)
                  HighInt = Val(HighVal$)
                  If (LowInt <= HighInt And TargetInt >= LowInt And TargetInt <= HighInt) Or
                     (LowInt > HighInt And TargetInt >= HighInt And TargetInt <= LowInt)
                    Stack(Depth)\BranchActive = #True
                    If Stack(Depth)\HasMatched = #Cond_False : Stack(Depth)\HasMatched = #Cond_True : EndIf
                    Break
                  EndIf
                EndIf
              Else
                Val$ = EvalConstExpr(CaseItem\s, FileName$, EvalPos, ModuleContainer, Stage1Only)
                If Val$ = ""
                  HasUndeterminedCase = #True
                ElseIf Val$ = Stack(Depth)\SelectVal
                  Stack(Depth)\BranchActive = #True
                  If Stack(Depth)\HasMatched = #Cond_False : Stack(Depth)\HasMatched = #Cond_True : EndIf
                  Break
                EndIf
              EndIf
            ForEver
            If Stack(Depth)\BranchActive = #False And HasUndeterminedCase
              Stack(Depth)\BranchActive = #True
              Stack(Depth)\HasMatched   = #Cond_Undetermined
            EndIf
          EndIf
        EndIf
      EndIf
      ;
    Case "compilerelse", "compilerdefault"
      If Depth > 0 And Stack(Depth)\IsSelect = IsSelectPair
        Evaluated                 = Bool(Stack(Depth)\ParentActive And Stack(Depth)\HasMatched <> #Cond_True)
        Stack(Depth)\BranchActive = Evaluated
        Stack(Depth)\HasMatched   = #Cond_True
      EndIf
      ;
    Case "compilerendif", "compilerendselect"
      If Depth > 0 And Stack(Depth)\IsSelect = IsSelectPair
        Evaluated = Stack(Depth)\ParentActive
        Depth - 1
      EndIf
  EndSelect
  ;
  *Depth\i = Depth
  If Depth > 0 : IsActive = Stack(Depth)\BranchActive : Else : IsActive = #True : EndIf
  ;
  ; Centralized DeadZones() boundary tracking across all callers (Stage 1 & OnePass):
  If *DeadStartPos
    If Evaluated
      If *DeadStartPos\i > 0
        If RecordDeadZone And IsFileInMemory(FileName$)
          AddDeadZoneRange(PBB_FileListing()\DeadZones(), *DeadStartPos\i, TokStartPos - 1)
        EndIf
        *DeadStartPos\i = 0
      EndIf
    ElseIf *DeadStartPos\i = 0
      *DeadStartPos\i = TokStartPos
    EndIf
    If IsActive = #False And *DeadStartPos\i = 0 And DirectiveLCase$ <> "compilerselect"
      *DeadStartPos\i = AfterExprPos
    EndIf
  EndIf
  ;
  ProcedureReturn IsActive
EndProcedure
;
Procedure.i ProcessCompilerDirectiveAtStatement(DirectiveLCase$, *TokenStartChar.Character, *StringStartingAdr, *PtrToChar,
                                                Array Stack.CompilerStackLevel(1), *Depth.Integer, MinFloorDepth,
                                                FileName$, EvalPos, ModuleContainer, ActiveOwnerID = -1,
                                                RecordDeadZone = #False, MarkExprUsages = #False, *LimitChar.Character = 0)
  ;
  ; Role: Single-loop evaluator and dead-branch fast-forwarder for OnePass and ListVariables.
  ;       Evaluates active conditions (marking referenced constants/modules when MarkExprUsages = #True),
  ;       skips dead branches (ignoring inner 'Macro ... EndMacro' blocks and stopping cleanly at '*LimitChar'),
  ;       and delegates DeadZones() span recording to StepCompilerDirective().
  ;
  Protected *Char.Character = PeekI(*PtrToChar), *CurTokStart.Character = *TokenStartChar, *NextTokStart.Character
  Protected CurDirective$ = DirectiveLCase$, RawExpr$, NextTokLCase$
  Protected HasArg, IsActive, DeadStartPos, MacroDepth, FirstCharLow, TokStartPos
  ;
  If IsCompilerDirective(CurDirective$) = #False
    ProcedureReturn #False
  EndIf
  ;
  RecordDeadZone             = Bool(RecordDeadZone And FileName$ <> "")
  EvalContext\MarkExprUsages = MarkExprUsages
  EvalContext\ActiveOwnerID  = ActiveOwnerID
  ;
  Repeat
    TokStartPos = CharToPos(*CurTokStart)
    If (CurDirective$ = "compilerendif" Or CurDirective$ = "compilerendselect") And *Depth\i <= MinFloorDepth
      SkipToEndOfStatement(@*Char, *LimitChar, #False)
      If RecordDeadZone And DeadStartPos > 0 And IsFileInMemory(FileName$)
        AddDeadZoneRange(PBB_FileListing()\DeadZones(), DeadStartPos, TokStartPos - 1)
      EndIf
      Break
    EndIf
    HasArg   = Bool(CurDirective$ = "compilerif" Or CurDirective$ = "compilerelseif" Or
                    CurDirective$ = "compilerselect" Or CurDirective$ = "compilercase")
    RawExpr$ = ExtractCompilerDirectiveExpr(*StringStartingAdr, @*Char, HasArg)
    IsActive = StepCompilerDirective(CurDirective$, RawExpr$, Stack(), *Depth, FileName$,
                                     TokStartPos, CharToPos(*Char) + 1, EvalPos, ModuleContainer,
                                     #False, @DeadStartPos, RecordDeadZone)
    If IsActive : Break : EndIf
    ;
    ; Fast-forward statement by statement until the next top-level (MacroDepth = 0) conditional directive:
    MacroDepth = 0
    Repeat
      SkipToEndOfStatement(@*Char, *LimitChar, #True)
      SkipSpaces(*Char)
      If (*LimitChar And *Char >= *LimitChar) Or *Char\c = 0
        If RecordDeadZone And DeadStartPos > 0 And IsFileInMemory(FileName$)
          AddDeadZoneRange(PBB_FileListing()\DeadZones(), DeadStartPos, CharToPos(*Char) - 1)
        EndIf
        Break 2
      EndIf
      FirstCharLow = *Char\c | 32
      If FirstCharLow = 'm' Or FirstCharLow = 'e' Or (MacroDepth = 0 And FirstCharLow = 'c')
        *NextTokStart = *Char
        NextTokLCase$ = LCase(ExtractOrSkipNextToken(*StringStartingAdr, @*Char))
        If NextTokLCase$ = "macro"
          MacroDepth + 1
        ElseIf NextTokLCase$ = "endmacro" And MacroDepth > 0
          MacroDepth - 1
        ElseIf MacroDepth = 0 And IsCompilerDirective(NextTokLCase$)
          CurDirective$ = NextTokLCase$
          *CurTokStart  = *NextTokStart
          Break
        EndIf
      EndIf
    ForEver
  ForEver
  ;
  EvalContext\MarkExprUsages = #False
  EvalContext\ActiveOwnerID  = -1
  PokeI(*PtrToChar, *Char)
  ProcedureReturn #True
EndProcedure
;
; ----------------------------------------------------------------------------------
;       ---- Include tree construction (IncludeFile / XIncludeFile / .pbp) ----
; ----------------------------------------------------------------------------------
;
Define HighestParent$
;
Procedure.i SliceTouchesGuardSymbol(*LowerSlice.String)
  ;
  ; Role: Single-pass O(N) pointer scanner checking whether a lowercase code slice contains
  ;       any structural directive or CompilerIf/Defined() guard symbol in PBB_CompilerGuardSymbols().
  ;
  Protected *Char.Character, Found
  Protected Tok$
  ;
  If *LowerSlice = 0 Or *LowerSlice\s = "" : ProcedureReturn #False : EndIf
  CompilerIf #PBBReport : Protected SavedOPDTokenLine$ = OPDTokenLine$ : CompilerEndIf
  *Char = @*LowerSlice\s
  While *Char\c <> 0 And Found = #False
    If IsStringOrCharLiteralStart(*Char) : ExtractOrSkipNextString(@*Char, *Char\c, #True)
    ElseIf *Char\c = ';'                 : GoToEOL_AndExtractComment(@*Char)
    ElseIf *Char\c = '#' Or *Char\c = '*' Or IsPBSeparator(*Char\c) = #False
      Tok$ = ExtractOrSkipNextToken(@*LowerSlice\s, @*Char)
      If Tok$ = "" : GoNextChar(*Char)
      Else         : Found = Bool(FindMapElement(PBB_CompilerGuardSymbols(), Tok$) Or FindMapElement(PBB_CompilerGuardSymbols(), "#" + Tok$))
      EndIf
    Else : GoNextChar(*Char)
    EndIf
  Wend
  CompilerIf #PBBReport : OPDTokenLine$ = SavedOPDTokenLine$ : CompilerEndIf
  ProcedureReturn Found
EndProcedure
;
Procedure   RegisterFileInListing(List TargetList.FileListingStruct(), FileName$, ParentName$, PosInParent,
                                  KeepActiveParent = #False, IsExcluded = #False)
  ;
  ; Role: Ensures 'FileName$' exists in 'TargetList()' (case-insensitive), marks it active in the
  ;       current inclusion tree, and manages its parent link and CompilerIf exclusion state:
  ;       - An active inclusion (IsExcluded = #False) always clears ExcludedByCompilerIf and overrides
  ;         any prior excluded parent link.
  ;       - An orphan .pbp registration (KeepActiveParent = #True) never reactivates a file already
  ;         flagged as ExcludedByCompilerIf by a dead IncludeFile branch.
  ;
  Protected Found
  ;
  ForEach TargetList()
    If LCase(TargetList()\FileName$) = LCase(FileName$)
      Found = #True
      Break
    EndIf
  Next
  ;
  If Found = #False
    AddElement(TargetList())
    TargetList()\FileName$            = FileName$
    TargetList()\PosInParent          = -1
    TargetList()\ExcludedByCompilerIf = IsExcluded
  ElseIf KeepActiveParent And TargetList()\ExcludedByCompilerIf
    TargetList()\IsActiveInTree = #True
    ProcedureReturn
  ElseIf IsExcluded = #False
    If TargetList()\ExcludedByCompilerIf
      TargetList()\ExcludedByCompilerIf = #False
      ClearList(TargetList()\DeadZones())
      TargetList()\ParentFileName = ParentName$
      TargetList()\PosInParent    = PosInParent
    EndIf
  ElseIf TargetList()\IsActiveInTree
    ; Already registered by an active branch earlier in the tree: keep it active
    ProcedureReturn
  Else
    TargetList()\ExcludedByCompilerIf = #True
  EndIf
  ;
  If KeepActiveParent
    If TargetList()\IsActiveInTree = #False
      TargetList()\ParentFileName = ParentName$
      TargetList()\PosInParent    = PosInParent
    EndIf
  ElseIf TargetList()\ParentFileName = "" Or TargetList()\PosInParent = -1
    TargetList()\ParentFileName = ParentName$
    TargetList()\PosInParent    = PosInParent
  EndIf
  TargetList()\IsActiveInTree = #True
EndProcedure
;
Procedure   DetectProjectCompilerOptions(ProjectFile$, MainSourceFile$)
  ;
  ; Role: Updates PBB_ProjectDebugger and PBB_ProjectThread from the default target '<options .../>'
  ;       of 'ProjectFile$' (.pbp) when present, or from the trailing '; IDE Options' of 'MainSourceFile$'.
  ;
  Protected *Content.String
  Protected Descr$, TargetBlock$
  Protected p, pf
  ;
  PBB_ProjectDebugger = 1
  PBB_ProjectThread   = 0
  ;
  If ProjectFile$ <> "" And FileSize(ProjectFile$) > 1
    Descr$ = FileToText(ProjectFile$, #PB_UTF8, 1)
    p      = FindString(Descr$, "default=" + #DOUBLEQUOTE$ + "1" + #DOUBLEQUOTE$, 1, #PB_String_NoCase)
    If p = 0 : p = FindString(Descr$, "<target ", 1, #PB_String_NoCase) : EndIf
    If p > 0
      pf = FindString(Descr$, "</target>", p, #PB_String_NoCase)
      If pf = 0 : pf = Len(Descr$) + 1 : EndIf
      TargetBlock$        = FastMid(Descr$, p, pf - p)
      PBB_ProjectThread   = Bool(FindString(TargetBlock$, "thread=" + #DOUBLEQUOTE$ + "1" + #DOUBLEQUOTE$, 1, #PB_String_NoCase) > 0)
      PBB_ProjectDebugger = Bool(FindString(TargetBlock$, "debug=" + #DOUBLEQUOTE$ + "0" + #DOUBLEQUOTE$, 1, #PB_String_NoCase) = 0)
    EndIf
  ElseIf MainSourceFile$ <> ""
    *Content = GetPointedCodeFromFile(MainSourceFile$, #False)
    If *Content And *Content\s <> ""
      Descr$              = FastMid(*Content\s, FindIDENotations(*Content))
      PBB_ProjectThread   = Bool(FindString(Descr$, "EnableThread", 1, #PB_String_NoCase) > 0)
      PBB_ProjectDebugger = Bool(FindString(Descr$, "DisableDebugger", 1, #PB_String_NoCase) = 0)
    EndIf
  EndIf
EndProcedure
;
Procedure.s CompleteListOfFiles(FileName$, ParentFileName$ = "", PosInParent = -1,
                                PriorityMode = #WorkInBackGround, ErrorWhileReadingFile$ = "",
                                IsExcluded = #False)
  ;
  ; Role: Single-pass pointer scanner that discovers active and excluded inclusions (IncludeFile,
  ;       XIncludeFile, IncludeBinary) in 'FileName$', evaluates upstream '#Const = Expr' definitions
  ;       and CompilerIf/CompilerSelect directives via StepCompilerDirective() to populate Stage-1 DeadZones(),
  ;       and recurses in 'IsExcluded = #True' mode on dead branches so project backups retain all platform files.
  ;
  Shared HighestParent$
  Protected *SearchString.String, *StringStartingAdr, *Char.Character, *TokStart.Character, *LimitChar.Character
  Protected TokLCase$, RawExpr$, Val$, Included$, ActiveIncPath$, DeadIncPath$
  Protected FoundInList, TokPos, StackDepth, MacroDepth, DeadStartPos, HasArg, FirstLow
  Protected IsActive = Bool(IsExcluded = #False)
  Protected Dim CStack.CompilerStackLevel(#PBB_MaxCompilerDepth)
  ;
  If ParentFileName$ = "" And IsExcluded = #False
    ClearMap(PBB_Stage1Constants())
    ClearList(PBB_BinaryFileListing())
  EndIf
  ;
  If FileName$ <> ""
    If FileSize(FileName$) < 2
      If IsExcluded = #False And FindString(ErrorWhileReadingFile$, FileName$ + #CR$) = 0
        ErrorWhileReadingFile$ + FileName$
        If ParentFileName$ <> ""
          ErrorWhileReadingFile$ + #CR$ + "  -> Included in: " + ParentFileName$ + " at pos " + Str(PosInParent)
        EndIf
        ErrorWhileReadingFile$ + #CR$
      EndIf
    Else
      ForEach PBB_FileListing()
        If LCase(PBB_FileListing()\FileName$) = LCase(FileName$)
          FoundInList = #True
          Break
        EndIf
      Next
      ;
      ; Skip re-scanning if already active in tree, or if already registered while scanning an excluded branch:
      If FoundInList And PBB_FileListing()\IsActiveInTree And (PBB_FileListing()\ExcludedByCompilerIf = #False Or IsExcluded)
        If IsExcluded = #False And ParentFileName$ <> "" And PosInParent <> -1 And
           (PBB_FileListing()\ParentFileName = "" Or PBB_FileListing()\PosInParent = -1)
          PBB_FileListing()\ParentFileName = ParentFileName$
          PBB_FileListing()\PosInParent    = PosInParent
          If HighestParent$ = "" Or LCase(HighestParent$) = LCase(FileName$)
            HighestParent$ = ParentFileName$
          EndIf
        EndIf
      Else
        *SearchString = GetPointedCodeFromFile(FileName$, #False)
        GetPointedCodeFromFile(FileName$, #True)
        RegisterFileInListing(PBB_FileListing(), FileName$, ParentFileName$, PosInParent, #False, IsExcluded)
        ClearList(PBB_FileListing()\DeadZones())
        ;
        If *SearchString And *SearchString\s <> ""
          *StringStartingAdr = @*SearchString\s
          *LimitChar         = PosToChar(*StringStartingAdr, FindIDENotations(*SearchString))
          *Char              = *StringStartingAdr
          ;
          While *Char < *LimitChar And *Char\c <> 0
            SkipSpaces(*Char)
            If *Char >= *LimitChar Or *Char\c = 0 : Break : EndIf
            ;
            FirstLow = *Char\c | 32
            If *Char\c = '#' Or FirstLow = 'm' Or FirstLow = 'e' Or FirstLow = 'c' Or FirstLow = 'i' Or FirstLow = 'x'
              *TokStart = *Char
              TokLCase$ = LCase(ExtractOrSkipNextToken(*StringStartingAdr, @*Char))
              TokPos    = CharToPos(*TokStart)
              ;
              If TokLCase$ = "macro"
                MacroDepth + 1
              ElseIf TokLCase$ = "endmacro"
                If MacroDepth > 0 : MacroDepth - 1 : EndIf
              ElseIf MacroDepth = 0
                If IsExcluded = #False And IsCompilerDirective(TokLCase$)
                  HasArg   = Bool(TokLCase$ = "compilerif" Or TokLCase$ = "compilerelseif" Or
                                  TokLCase$ = "compilerselect" Or TokLCase$ = "compilercase")
                  RawExpr$ = ExtractCompilerDirectiveExpr(*StringStartingAdr, @*Char, HasArg)
                  IsActive = StepCompilerDirective(TokLCase$, RawExpr$, CStack(), @StackDepth,
                                                   FileName$, TokPos, CharToPos(*Char) + 1,
                                                   TokPos, -1, #True, @DeadStartPos, #True)
                  ;
                ElseIf IsActive And Left(TokLCase$, 1) = "#" And Len(TokLCase$) > 1
                  SkipSpaces(*Char)
                  If *Char\c = '='
                    GoNextChar(*Char)
                    Val$ = EvalConstExpr(ExtractCompilerDirectiveExpr(*StringStartingAdr, @*Char, #True), FileName$, TokPos, -1, #True, #False)
                    If Val$ <> ""
                      PBB_Stage1Constants(TokLCase$) = Val$
                    EndIf
                  EndIf
                  ;
                ElseIf TokLCase$ = "includepath" Or TokLCase$ = "includefile" Or
                       TokLCase$ = "xincludefile" Or TokLCase$ = "includebinary"
                  Val$ = EvalConstExpr(ExtractCompilerDirectiveExpr(*StringStartingAdr, @*Char, #True), FileName$, TokPos, -1, #True, #False)
                  If Left(Val$, 1) = #DOUBLEQUOTE$
                    Val$ = Mid(Val$, 2, Len(Val$) - 2)
                    If TokLCase$ = "includepath"
                      If Val$ <> "" And Right(Val$, 1) <> "/" And Right(Val$, 1) <> "\" : Val$ + "\" : EndIf
                      If IsActive : ActiveIncPath$ = Val$ : EndIf
                      DeadIncPath$ = Val$
                    Else
                      If IsActive
                        Included$ = RelativeAddressToFullAddress(ActiveIncPath$ + Val$, FileName$)
                      Else
                        Included$ = RelativeAddressToFullAddress(DeadIncPath$ + Val$, FileName$)
                      EndIf
                      If TokLCase$ = "includebinary"
                        RegisterFileInListing(PBB_BinaryFileListing(), Included$, FileName$, TokPos, #False, Bool(IsActive = #False))
                      ElseIf IsActive
                        ErrorWhileReadingFile$ = CompleteListOfFiles(Included$, FileName$, TokPos, PriorityMode, ErrorWhileReadingFile$, #False)
                      ElseIf FileSize(Included$) >= 2
                        ErrorWhileReadingFile$ = CompleteListOfFiles(Included$, FileName$, TokPos, PriorityMode, ErrorWhileReadingFile$, #True)
                      EndIf
                    EndIf
                  EndIf
                EndIf
              EndIf
            EndIf
            SkipToEndOfStatement(@*Char, *LimitChar, #True)
          Wend
          ;
          ; Close any unclosed trailing dead zone at the IDE Options boundary
          If DeadStartPos > 0 And IsActive = #False And IsFileInMemory(FileName$)
            AddDeadZoneRange(PBB_FileListing()\DeadZones(), DeadStartPos, CharToPos(*LimitChar) - 1)
          EndIf
        EndIf
      EndIf
    EndIf
  EndIf
  ;
  If ParentFileName$ = "" And IsExcluded = #False
    ClearMap(PBB_Stage1Constants())
  EndIf
  ProcedureReturn ErrorWhileReadingFile$
EndProcedure
;
Procedure.s SetListOfFiles(FileName$, PriorityMode = #WorkInBackGround, *HasChanged.Integer = #Null)
  ;
  ; Role: Determines the dependency tree of source files (.pb, .pbi) and binary assets
  ;       for a specified PureBasic source file or project file (.pbp), resolving PBBFicPrincipalPB$
  ;       and compiler options upfront so Stage-1 CompilerIf/CompilerSelect evaluation is exact.
  ;
  Shared HighestParent$
  Protected ErrorWhileReadingFile$, StartingFile$, Ext$, PrevMainFile$ = PBBFicPrincipalPB$
  Protected *FirstFileContent.String, LocalHasChanged, ActiveCount
  Protected OriginalFileName$, TargetFileName$, ResolvedMainFile$, FName$, ProjectFiles$, ProjectDescr$
  Protected p, pf, PosInFileList, OriginalInTree
  Protected NewMap PreviouslyActive.i()
  ;
  PBBCurrentProject$ = ""
  ClearMap(PBB_CompilerGuardSymbols())
  FName$ = "includefile,xincludefile,includepath,includebinary,module,endmodule,declaremodule,enddeclaremodule,usemodule,unusemodule," +
           "macro,endmacro,undefinemacro,compilerif,compilerelseif,compilerelse,compilerendif,compilerselect,compilercase,compilerdefault,compilerendselect"
  For p = 1 To 21 : PBB_CompilerGuardSymbols(StringField(FName$, p, ",")) = #True : Next
  ;
  ; 1. Back up previously active tree files and reset tree activation states
  ForEach PBB_FileListing()
    If PBB_FileListing()\IsActiveInTree
      PreviouslyActive(LCase(PBB_FileListing()\FileName$) + "|" + Str(PBB_FileListing()\ExcludedByCompilerIf)) = #True
    EndIf
    PBB_FileListing()\IsActiveInTree       = #False
    PBB_FileListing()\ExcludedByCompilerIf = #False
    ClearList(PBB_FileListing()\DeadZones())
    PBB_FileListing()\ParentFileName       = ""
    PBB_FileListing()\PosInParent          = -1
  Next
  ForEach PBB_BinaryFileListing()
    PBB_BinaryFileListing()\IsActiveInTree = #False
    PBB_BinaryFileListing()\ParentFileName = ""
    PBB_BinaryFileListing()\PosInParent    = -1
  Next
  ;
  If FileName$
    Ext$ = LCase(GetExtensionPart(FileName$))
    If Ext$ = "pb" Or Ext$ = "pbi"
      ; Check whether the opened source file belongs to an active project open in the PB IDE
      PBBCurrentProject$ = GetPBOpenProject(PureBasicProgAddr$)
      If FileSize(PBBCurrentProject$) > 1
        ProjectFiles$ = ListProjectsFiles(PBBCurrentProject$)
        Repeat
          PosInFileList + 1
          FName$ = StringField(ProjectFiles$, PosInFileList, #CR$)
          If FName$ <> "" And LCase(FName$) = LCase(FileName$)
            If LCase(PBBEditorSourceFile$) = LCase(FileName$)
              PBBEditorSourceFile$ = FName$
            EndIf
            FileName$ = PBBCurrentProject$
            Ext$      = "pbp"
            Break
          EndIf
        Until FName$ = ""
      EndIf
      PBBCurrentProject$ = ""
    EndIf
    ;
    If Ext$ = "pbp"
      PBBCurrentProject$ = FileName$
      ProjectFiles$      = ListProjectsFiles(FileName$)
      StartingFile$      = GetFilePart(FileName$)
      HighestParent$     = ""
      ;
      ; Identify the main target file declared in the project configuration BEFORE running CompleteListOfFiles()
      ProjectDescr$ = FileToText(FileName$)
      p = FindString(ProjectDescr$, "<section name=" + #DOUBLEQUOTE$ + "targets")
      If p
        p = FindString(ProjectDescr$, "<inputfile value=" + #DOUBLEQUOTE$, p)
        If p
          p + Len("<inputfile value=" + #DOUBLEQUOTE$)
          pf              = FindString(ProjectDescr$, #DOUBLEQUOTE$, p)
          TargetFileName$ = RelativeAddressToFullAddress(Mid(ProjectDescr$, p, pf - p), FileName$)
          If TargetFileName$ <> ""
            PBBFicPrincipalPB$ = TargetFileName$
          EndIf
          DetectProjectCompilerOptions(PBBCurrentProject$, TargetFileName$)
          ErrorWhileReadingFile$ = CompleteListOfFiles(TargetFileName$, "", -1, PriorityMode, ErrorWhileReadingFile$)
        EndIf
      EndIf
      If p = 0 Or TargetFileName$ = ""
        TargetFileName$ = HighestParent$
      EndIf
      FileName$          = TargetFileName$
      PBBFicPrincipalPB$ = TargetFileName$
      ;
      ; Register remaining project files not reachable via the main file inclusion hierarchy
      ; (files marked ExcludedByCompilerIf during CompleteListOfFiles() retain their excluded status):
      PosInFileList = 0
      Repeat
        PosInFileList + 1
        FName$ = StringField(ProjectFiles$, PosInFileList, #CR$)
        Ext$   = LCase(GetExtensionPart(FName$))
        If Ext$ = "pb" Or Ext$ = "pbi"
          RegisterFileInListing(PBB_FileListing(), FName$, PBBCurrentProject$, -1, #True)
        ElseIf FName$ <> ""
          RegisterFileInListing(PBB_BinaryFileListing(), FName$, PBBCurrentProject$, -1, #True)
        EndIf
      Until FName$ = ""
      ;
      If ErrorWhileReadingFile$
        ErrorWhileReadingFile$ + "  --> Listed in project " + StartingFile$ + #CR$
      EndIf
    Else
      ; Single source file mode (.pb / .pbi)
      *FirstFileContent = GetPointedCodeFromFile(FileName$, #False)
      If *FirstFileContent = 0 Or *FirstFileContent\s = ""
        PBBFicPrincipalPB$ = PrevMainFile$
        SetGadgetText(GPBBGadgets\Adr_gadget, "")
        ProcedureReturn FileName$
      EndIf
      ;
      OriginalFileName$ = FileName$
      TargetFileName$   = FileName$
      ;
      ; Check for IDE 'UseMainFile' configuration directive in trailing source comments BEFORE CompleteListOfFiles()
      p = FindString(*FirstFileContent\s, #CR$ + "; UseMainFile = ")
      If p
        p + Len("; UseMainFile = ")
        pf = FindString(*FirstFileContent\s, #CR$, p)
        If pf = 0 : pf = Len(*FirstFileContent\s) + 1 : EndIf
        ResolvedMainFile$ = RelativeAddressToFullAddress(Trim(FastMid(*FirstFileContent\s, p, pf - p)), FileName$)
        If FileSize(ResolvedMainFile$) > 2
          TargetFileName$ = ResolvedMainFile$
        EndIf
      EndIf
      ;
      PBBFicPrincipalPB$ = TargetFileName$
      DetectProjectCompilerOptions("", TargetFileName$)
      ErrorWhileReadingFile$ = CompleteListOfFiles(TargetFileName$, "", -1, PriorityMode)
      ;
      ; If the designated main file did not include the initial file (nor exclude it via CompilerIf), parse it as well
      ForEach PBB_FileListing()
        If LCase(PBB_FileListing()\FileName$) = LCase(OriginalFileName$) And PBB_FileListing()\IsActiveInTree
          OriginalInTree = #True
          If LCase(PBBEditorSourceFile$) = LCase(OriginalFileName$)
            PBBEditorSourceFile$ = PBB_FileListing()\FileName$
          EndIf
          Break
        EndIf
      Next
      If OriginalInTree = #False
        ErrorWhileReadingFile$ = CompleteListOfFiles(OriginalFileName$, "", -1, PriorityMode, ErrorWhileReadingFile$)
      EndIf
      ;
      FileName$     = TargetFileName$
      StartingFile$ = GetFilePart(FileName$)
    EndIf
  EndIf
  ;
  If ErrorWhileReadingFile$
    AlertInPBBWindow(GetTextFromCatalogPB("ErrorWhileReadingFile") + #CR$ + ErrorWhileReadingFile$, GetTextFromCatalogPB("IncludedFileError"), 1, 0, 1)
  EndIf
  ;
  ; 2. Purge inactive files and evaluate UI update needs (including CompilerIf exclusion changes)
  ForEach PBB_FileListing()
    If PBB_FileListing()\IsActiveInTree = #False
      If FindMapElement(PreviouslyActive(), LCase(PBB_FileListing()\FileName$) + "|0") Or
         FindMapElement(PreviouslyActive(), LCase(PBB_FileListing()\FileName$) + "|1")
        LocalHasChanged = #True
      EndIf
      DeleteElement(PBB_FileListing())
    Else
      ActiveCount + 1
      If FindMapElement(PreviouslyActive(), LCase(PBB_FileListing()\FileName$) + "|" + Str(PBB_FileListing()\ExcludedByCompilerIf)) = 0
        LocalHasChanged = #True
      EndIf
    EndIf
  Next
  ;
  If ActiveCount <> MapSize(PreviouslyActive())
    LocalHasChanged = #True
  ElseIf ActiveCount > 0 And IsGadget(GPBBGadgets\Adr_gadget) And GetGadgetText(GPBBGadgets\Adr_gadget) <> StartingFile$
    LocalHasChanged = #True
  EndIf
  ;
  ; 3. Compute virtual positions, sort files, and rebuild compilation sequence
  UpdateVirtualPositionsAndSequence(TargetFileName$, Bool(LocalHasChanged Or ListSize(PBB_CompilationSequence()) = 0))
  ;
  ; 4. Update GPBBGadgets\Adr_gadget
  If ActiveCount = 0
    SetGadgetText(GPBBGadgets\Adr_gadget, "")
  ElseIf LocalHasChanged
    SetGadgetText(GPBBGadgets\Adr_gadget, StartingFile$)
  EndIf
  ;
  If *HasChanged
    *HasChanged\i = LocalHasChanged
  EndIf
  ProcedureReturn FileName$
EndProcedure
;
;
; **********************************************************************************
;
;-                        Token & macro usage tracking
;
; **********************************************************************************
;
Procedure CheckAndMarkTokenUsage(*StringStartingAdr, *PtrToChar, FileName$, TokenLCase$, MacroProcEndingPos = 0,
                                 MacroProcID = -1, MacroProcParams$ = "", TargetType = 0,
                                 WithVar$ = "", ModuleContainer = -1,
                                 Parentheses = -1, VarDotType_LCase$ = "",
                                 *PreloadedCandidateIDs.String = 0, OwnerID = 0)
  ;
  ; Role:
  ; Checks if a purified lowercase token (or "mod::ident" qualified token) corresponds to an existing element
  ; and marks both any 'mod::' prefix and the target element as used, respecting module, parameter-shadowing,
  ; and container scopes.
  ;
  ; Parameters:
  ; *PtrToChar             - Pointer to a character pointer in the code stream (used as a positional anchor).
  ; FileName$              - Full path of the file where the token is referenced.
  ; TokenLCase$            - Lowercase purified identifier name (or "mod::ident").
  ; MacroProcEndingPos     - Character offset marking the end of the active container.
  ; MacroProcID            - ID of the active lexical container (-1 if outside any container).
  ; MacroProcParams$       - Tab-delimited formal parameter list (for parameter shadowing and '\Field' typing).
  ; TargetType             - 0 for standard expressions, #PBBField after '\', #PBBLabel for labels,
  ;                          #PBBStructure / #PBBStructOrInterface / #PBBTypeOrVar for structure and type contexts,
  ;                          or #PBBModule for 'Mod::' / 'UseModule'.
  ; WithVar$               - Active 'With' block expression used to resolve orphan '\Field' accesses.
  ; ModuleContainer        - FixedID of the enclosing Module/DeclareModule (-1 when in global scope).
  ; Parentheses            - 1 if followed by '(', 0 if not, or -1 when unconstrained (e.g. '@Proc' or types).
  ; VarDotType_LCase$      - Optional lowercase structure/type suffix attached to the variable token.
  ; *PreloadedCandidateIDs - Optional pointer to a pre-fetched ElementMapIndex() string to bypass a redundant map lookup.
  ; OwnerID                - FixedID receiving the usage dependency link when distinct from MacroProcID (0 defaults to MacroProcID).
  ;
  ; Output:
  ; Returns the ID of the matched element (leaving ElementList() selected on it), or 0 if unmatched.
  ;
  Protected *Char.Character, *BackslashChar.Character
  Protected BareTokenLCase$, PosInCode, ColonPos, TargetStructID, IsUsed
  Protected LocalCandidateIDs.String
  ;
  If MacroProcID = 0 : MacroProcID = -1 : EndIf
  If TokenLCase$ = "" Or *PtrToChar = 0 : ProcedureReturn 0 : EndIf
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDIndicatesUsageDetection
      OPDReport("• Examining " + TokenLCase$ + " for usage")
    EndIf
  CompilerEndIf
  ;
  *Char     = PeekI(*PtrToChar)
  PosInCode = CharToPos(*Char)
  ;
  ; Reset container context if token lies past the active container's ending boundary:
  If PosInCode >= MacroProcEndingPos And MacroProcEndingPos > 0
    MacroProcEndingPos = 0
    MacroProcID        = -1
    MacroProcParams$   = #TAB$
  EndIf
  If OwnerID = 0 : OwnerID = MacroProcID : EndIf
  ;
  ; Resolve and automatically mark any 'mod::' prefix as used before looking up the target member:
  ColonPos = FindString(TokenLCase$, "::")
  If ColonPos
    If ColonPos = 1 Or CheckAndMarkTokenUsage(*StringStartingAdr, *PtrToChar, FileName$, FastLeft(TokenLCase$, ColonPos - 1),
                                              MacroProcEndingPos, MacroProcID, "", #PBBModule, "", ModuleContainer, -1, "", 0, OwnerID) = 0
      ProcedureReturn 0
    EndIf
    BareTokenLCase$ = FastMid(TokenLCase$, ColonPos + 2)
    If BareTokenLCase$ = "" Or FindString(BareTokenLCase$, "::") : ProcedureReturn 0 : EndIf
  Else
    BareTokenLCase$ = TokenLCase$
  EndIf
  ;
  ; Retrieve candidates from map index and shield formal parameters on unqualified tokens:
  If *PreloadedCandidateIDs = 0 Or *PreloadedCandidateIDs\s = ""
    If FindMapElement(ElementMapIndex(), BareTokenLCase$) = 0
      CompilerIf #PBBReport
        If OPDMod & #OPDIndicatesUsageDetection : OPDReport("  --> Not an existing element") : EndIf
      CompilerEndIf
      ProcedureReturn 0
    ElseIf TargetType = 0 And ColonPos = 0 And IsParam(BareTokenLCase$, MacroProcParams$, #False)
      CompilerIf #PBBReport
        If OPDMod & #OPDIndicatesUsageDetection : OPDReport("  --> Argument. Not a real element") : EndIf
      CompilerEndIf
      ProcedureReturn 0
    EndIf
    LocalCandidateIDs\s    = ElementMapIndex()
    *PreloadedCandidateIDs = @LocalCandidateIDs
  EndIf
  ;
  If TargetType = #PBBField And *Char > *StringStartingAdr
    *BackslashChar = *Char - SizeOf(Character)
    ; Rewind across optional spaces separating '\' from the field name so ResolveFieldTargetStructID() receives the '\' anchor
    SkipSpacesBackward(*BackslashChar, *StringStartingAdr + SizeOf(Character))
    TargetStructID = ResolveFieldTargetStructID(*PreloadedCandidateIDs\s, *BackslashChar, *StringStartingAdr, WithVar$, FileName$, PosInCode, 0, MacroProcID, MacroProcParams$, ModuleContainer)
  EndIf
  ;
  IsUsed = FindElementInList(TokenLCase$, TargetType, PosInCode, FileName$, MacroProcID, Parentheses, VarDotType_LCase$, TargetStructID, ModuleContainer, "", "", *PreloadedCandidateIDs)
  If IsUsed > 0 And TargetType = #PBBTypeOrVar And ElementList()\Type = #PBBVariable And
     ColonPos = 0 And IsParam(BareTokenLCase$, MacroProcParams$, #False)
    ; In SizeOf()/TypeOf(), a homonymous Structure/Interface takes precedence over a parameter,
    ; whereas a variable fallback remains shielded when shadowed by a formal parameter:
    CompilerIf #PBBReport
      If OPDMod & #OPDIndicatesUsageDetection : OPDReport("  --> Argument. Not a real element") : EndIf
    CompilerEndIf
    ProcedureReturn 0
  EndIf
  ;
  If IsUsed > 0
    RegisterElementUsage(OwnerID, IsUsed, FileName$, PosInCode)
  EndIf
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDIndicatesUsageDetection
      If IsUsed > 0
        OPDReport("  --> Used")
      Else
        OPDReport("  --> X Unused X")
        PushListPosition(ElementList())
        Protected IDIndex, IDStr$
        Repeat
          IDIndex + 1
          IDStr$ = StringField(*PreloadedCandidateIDs\s, IDIndex, ",")
          If IDStr$ = "" : Break : EndIf
          SelectPBBElement(Val(IDStr$))
          OPDReport("      Closest existing element: " + ElementList()\Name + " " + ElementList()\VariableSpecie)
        ForEver
        PopListPosition(ElementList())
      EndIf
    EndIf
  CompilerEndIf
  ;
  ProcedureReturn IsUsed
EndProcedure
;
Macro       AssignSharedCEDDotType(TargetStr, IsSharedCond, DotTypeStr, IsAddedFlag)
  ;
  ; Role: Formats the 6th column of a CED_Structure record ("None" or explicit dot-type for 'Shared' variables, "" otherwise).
  ;
  If IsSharedCond
    If DotTypeStr <> "" And IsAddedFlag = #False : TargetStr = DotTypeStr : Else : TargetStr = "None" : EndIf
  Else
    TargetStr = ""
  EndIf
EndMacro
;
Procedure   ResolveMacroTokens(MacroID, ProcContainer, FileName$, *StringStartingAdr, *PtrToChar, Depth = 0, ModuleContainer = -1)
  ;
  ; Role:
  ; Iterates through pre-extracted tokens of a macro (stored as TName$ + #TAB$ + TID) and resolves them.
  ; Distinguishes between element creations (TID <> -1) and simple usages (TID = -1, carrying "mod::ident" when qualified).
  ;
  ; Parameters:
  ; ProcContainer   - FixedID of the calling container (procedure or -1 for main scope) for context resolution.
  ; *PtrToChar      - Pointer to a character pointer anchored on the start of the macro call identifier.
  ; Depth           - Recursion depth tracker (max 16).
  ; ModuleContainer - FixedID of the active Module/DeclareModule scope (-1 if in global scope).
  ;
  ; Output:
  ; Resolves and marks tokens (including 'mod::' prefixes), recording local variable declarations in
  ; '\Continued' at the start of the macro call and evaluating usages at the end of the macro call token.
  ;
  Protected *CallStartChar.Character, *CallEndChar.Character
  Protected MacroTokens$, Line$, TName$, SharedDotType$
  Protected TIndex, TID, MatchedID, TargetType, CallStartPos, ScopeAlreadyRegistered, CEDIdx
  ;
  PushListPosition(ElementList())
  If *PtrToChar = 0 Or Depth > 16 Or SelectPBBElement(MacroID) = #False Or ElementList()\Type <> #PBBMacro
    PopListPosition(ElementList())
    ProcedureReturn
  EndIf
  ;
  ; Anchor injected declarations at the start of the macro call token (*CallStartChar) and evaluate
  ; expanded usages at the end of the macro call token (*CallEndChar) so EndingPos < PosInCode holds naturally:
  *CallStartChar = PeekI(*PtrToChar)
  *CallEndChar   = *CallStartChar
  While IsPBSeparator(*CallEndChar\c) = #False
    GoNextChar(*CallEndChar)
  Wend
  If *CallEndChar = *CallStartChar And *CallEndChar\c <> 0
    GoNextChar(*CallEndChar)
  EndIf
  CallStartPos = CharToPos(*CallStartChar)
  MacroTokens$ = ElementList()\MacroTokens
  ;
  Repeat
    TIndex + 1
    Line$ = StringField(MacroTokens$, TIndex, #CR$)
    If Line$ = "" : Break : EndIf
    ;
    TName$ = StringField(Line$, 1, #TAB$)
    TID    = Val(StringField(Line$, 2, #TAB$))
    ;
    If TID = -1
      ; Case 1: Simple usage token (TID = -1) with pre-parsed target role, parenthesis constraint, and optional '.Type'
      TargetType = Val(StringField(Line$, 3, #TAB$))
      MatchedID  = CheckAndMarkTokenUsage(*StringStartingAdr, @*CallEndChar, FileName$, TName$, 0, ProcContainer, "",
                                          TargetType, "", ModuleContainer, Val(StringField(Line$, 4, #TAB$)), StringField(Line$, 5, #TAB$))
      If MatchedID > 0
        If TargetType = #PBBModule
          ElementList()\HasUseModule = #True
        ElseIf MatchedID <> MacroID And ElementList()\Type = #PBBMacro
          ResolveMacroTokens(MatchedID, ProcContainer, FileName$, *StringStartingAdr, *PtrToChar, Depth + 1, ModuleContainer)
        EndIf
      EndIf
      ;
    ElseIf SelectPBBElement(TID) And ElementList()\Type = #PBBVariable And
           ElementList()\VariableSpecie <> "Global" And ElementList()\VariableSpecie <> "Threaded"
      ; Case 2: Element creation/declaration inside the macro (valid TID).
      ; Append a secondary declaration in '\Continued' anchored at CallStartPos for the calling ProcContainer.
      ScopeAlreadyRegistered = Bool(ElementList()\LocalVarContainer = ProcContainer And
                                    (ElementList()\FileName <> FileName$ Or ElementList()\EndingPos <= CallStartPos))
      If ScopeAlreadyRegistered = #False And ElementList()\Continued <> "" And
         FindString(ElementList()\Continued, #TAB$ + Str(ProcContainer) + #TAB$)
        CEDIdx = 0
        Repeat
          CEDIdx + 1
          Line$ = StringField(ElementList()\Continued, CEDIdx, #CR$)
          If Line$ = "" : Break : EndIf
          If Val(StringField(Line$, 5, #TAB$)) = ProcContainer And
             (StringField(Line$, 1, #TAB$) <> FileName$ Or Val(StringField(Line$, 3, #TAB$)) <= CallStartPos)
            ScopeAlreadyRegistered = #True
            Break
          EndIf
        ForEver
      EndIf
      ;
      If ScopeAlreadyRegistered = #False
        AssignSharedCEDDotType(SharedDotType$, (ElementList()\VariableSpecie = "Shared"), ElementList()\VarDotType_LCase, ElementList()\VarAddedDotType)
        Line$ = CED_CreateRecord(FileName$, CallStartPos, CallStartPos, ElementList()\Declaration, ProcContainer, SharedDotType$)
        If FindString(ElementList()\Continued, Line$) = 0
          ElementList()\Continued + Line$
          CED_SortBasedOnPositions()
        EndIf
      EndIf
    EndIf
  ForEver
  ;
  PopListPosition(ElementList())
EndProcedure
;
;
; **********************************************************************************
;
;-                        Shared lexical stream engine
;
; **********************************************************************************
;
Macro AddUniqueListItem(ListStr, Item)
  ;
  ; Role: Appends 'Item' to a comma-separated display string ('ListStr') if not already present (case-insensitive).
  ; Prerequisites: 'ResolvedString$' must be declared as a string variable in the caller's scope.
  ;
  ResolvedString$ = Item
  If Item <> "" And FindString(", " + LCase(ListStr), ", " + LCase(ResolvedString$) + ", ") = 0
    ListStr + ResolvedString$ + ", "
  EndIf
EndMacro
;
Macro ResetParsingModeFlags
  ;
  ; Role: Resets all statement-level parsing flags and temporary buffers before analyzing a new statement.
  ;
  TypeOfElement                = #PBBNotAnElement
  CompilerIf Defined(Declaration$, #PB_Variable)
    Declaration$               = ""
  CompilerEndIf
  LastComment$                 = ""
  CompilerIf Defined(CommentStr, #PB_Variable)
    CommentStr\s               = ""
  CompilerEndIf
  VariableSpecie$              = ""
  FieldOrVariableSubSpecie$    = ""
  VariableDotType$             = ""
  LastToken$                   = ""
  LastSeparator                = 0
  NextTokenIsPotentialVariable = #False
  OpenParentheses              = 0
  StructFuncParenLevel         = 0
  IsInsideDefinition           = #False
  IsInsideDefaultValue         = #False
  IsQualified                  = #False
  ExpectedTargetType           = 0
  UsageTargetType              = 0
  UsageParentheses             = -1
EndMacro
;
Macro GetNextToken(ParsingMode)
  ;
  ; Role: Advances the lexical stream to the next meaningful token, statement boundary, or end of buffer,
  ; absorbing whitespace, literals, comments, standalone separators, 'Mod::Identifier' qualifications
  ; (emitted directly as Token$ = "Mod::Ident" / TokenLCase$ = "mod::ident" with IsQualified = #True),
  ; and role-defining PureBasic keywords ('Extends', 'Goto'/'Gosub'/'Restore', 'UseModule'/'UnuseModule',
  ; 'Defined(...)', 'SizeOf'/'TypeOf', 'OffsetOf', 'AllocateStructure', 'ClearStructure'/...) in a single step,
  ; populating 'UsageTargetType' and 'UsageParentheses' for the returned token.
  ; Parameters:
  ;   ParsingMode - Compile-time constant (#ParsingMode_OnePass or #ParsingMode_ListVariables) selecting
  ;                 zero-overhead behavior specific to file indexing or procedure variable analysis.
  ; Prerequisites:
  ;   Caller must initialize ParsingState = #Parsing_Start, IsStartOfStatement = #True, and ResetParsingModeFlags
  ;   before entering the main parsing loop. On #Parsing_EndOfLine, *Char remains positioned on the delimiter
  ;   (#CR, ':', or 0) until the next invocation of GetNextToken() so the caller can compute EndingPos.
  ;
  If ParsingState = #Parsing_NewToken
    ; Hot path (~85% of iterations): clear statement-start and previous token's separator/qualification/role state
    IsStartOfStatement = #False
    LastSeparator      = 0
    IsQualified        = #False
    UsageTargetType    = 0
  ElseIf ParsingState = #Parsing_EndOfLine
    ; Finalize the previous logical statement when not inside a multi-line continuation
    If LineToBeContinued = #False
      CompilerIf #PBBReport And ParsingMode = #ParsingMode_OnePass
        If OPDTokenLine$ And (OPDMod & #OPDFollowsTokens)
          If (ProcEndingPos Or MacroEndingPos Or BlockEndingPos) And
             EndingPos <> ProcEndingPos And EndingPos <> MacroEndingPos And EndingPos <> BlockEndingPos
            OPDTokenLine$ = "   " + OPDTokenLine$
          EndIf
          OPDReport(OPDTokenLine$)
          OPDTokenLine$ = ""
        EndIf
      CompilerEndIf
      ;
      CompilerIf ParsingMode = #ParsingMode_OnePass
        ActiveParams$ = #TAB$
      CompilerEndIf
      ;
      IsStartOfStatement = #True
      ResetParsingModeFlags
    EndIf
    ;
    ; Transition to #Parsing_EndOfFile once the null terminator has been processed as an EOL boundary
    If *Char = 0 Or *Char\c = 0
      ParsingState = #Parsing_EndOfFile
    Else
      ; Memorize the delimiter (#CR vs ':') before advancing so physical line tracking stays accurate
      LastSeparator = *Char\c
      GoNextChar(*Char)
      SkipSpaces(*Char)
      ;
      CompilerIf ParsingMode = #ParsingMode_OnePass
        If LastSeparator = #CR
          StartingPhysicalLinePos = CharToPos(*Char)
        EndIf
        If LineToBeContinued = #False
          StartingCodeLinePos = CharToPos(*Char)
          ; Enforce range boundaries and cooperative background timeslicing strictly between complete statements/blocks
          If CharLimitEnd And *Char > CharLimitEnd
            ParsingState = #Parsing_EndOfFile
          ElseIf ProcEndingPos = 0 And MacroEndingPos = 0 And BlockEndingPos = 0 And ModuleEndingPos = 0 And
                 PriorityMode <> #FinishCompletionNow And (ElapsedMilliseconds() - CLE_StartTime) > 200
            *MemPosChar  = *Char
            ParsingState = #Parsing_EndOfFile
          EndIf
        EndIf
      CompilerEndIf
      LastSeparator   = 0
      IsQualified     = #False
      UsageTargetType = 0
    EndIf
  EndIf
  ;
  If ParsingState <> #Parsing_EndOfFile
    Repeat
      Select *Char\c
        ;
        Case #CR, ':', 0
          ; Handle malformed orphan '::' with no preceding module identifier (e.g., '::Orphan' -> "::orphan")
          If IsModuleUsage(@*Char)
            IsStartOfStatement = #False
            ExpectedTargetType = 0
            LastToken$         = ""
            *TokenPosChar      = *Char
            Token$             = ExtractOrSkipNextToken(*StringStartingAdr, @*Char)
            If Token$ <> ""
              FirstChar        = Asc(Token$)
              Token$           = "::" + Token$
              TokenLCase$      = LCase(Token$)
              LastToken$       = TokenLCase$
              IsQualified      = #True
              UsageTargetType  = 0
              UsageParentheses = -1
              ParsingState     = #Parsing_NewToken
              Break
            EndIf
            Continue
          EndIf
          ;
          ; Detect multi-line statement continuation when a physical line ends after a continuation operator.
          ; Keep LastToken$ intact so subsequent comment-only or blank lines preserve the continuation state.
          LineToBeContinued = #False
          If *Char\c = #CR Or *Char\c = 0
            LastSeparator = *Char\c
            If *Char\c = #CR And LastToken$ <> "" And FindString(#PBEndOfSegmentedLine$, #TAB$ + LastToken$ + #TAB$)
              LineToBeContinued  = #True
              IsStartOfStatement = #False
            EndIf
          EndIf
          ParsingState = #Parsing_EndOfLine
          Break
          ;
        Case #Space, #NonBreakableSpace
          SkipSpaces(*Char)
          ;
        Case '('
          IsStartOfStatement = #False
          OpenParentheses + 1
          SkipToken
          ;
        Case ')'
          IsStartOfStatement = #False
          ExpectedTargetType = 0
          If StructFuncParenLevel > 0 And OpenParentheses <= StructFuncParenLevel
            StructFuncParenLevel = 0
          EndIf
          OpenParentheses - 1
          If OpenParentheses <= 0
            OpenParentheses = 0
            If IsInsideDefinition
              IsInsideDefinition           = #False
              IsInsideDefaultValue         = #False
              VariableSpecie$              = ""
              FieldOrVariableSubSpecie$    = ""
              NextTokenIsPotentialVariable = #False
              CompilerIf #PBBReport And ParsingMode = #ParsingMode_OnePass
                If OPDTokenLine$ And (OPDMod & #OPDFollowsTokens) And (OPDMod & #OPDFollowsSeparators) = 0
                  OPDTokenLine$ + ")"
                EndIf
              CompilerEndIf
            EndIf
          EndIf
          SkipToken
          ;
        Case #SIMPLEQUOTE, #DOUBLEQUOTE, '~'
          IsStartOfStatement = #False
          If *Char\c = '~' And PeekC(*Char + SizeOf(Character)) <> #DOUBLEQUOTE
            SkipToken
          Else
            LastSeparator = 0
            LastToken$    = ""
            ExtractOrSkipNextString(@*Char, *Char\c, #True)
          EndIf
          ;
        Case ';'
          LastComment$ = GoToEOL_AndExtractComment(@*Char)
          CompilerIf ParsingMode = #ParsingMode_OnePass
            If *Char >= IDEOptionsCharPos
              ; Reached PureBasic IDE options footer (logical end of source file)
              *Char        = 0
              ParsingState = #Parsing_EndOfFile
              Break
            EndIf
          CompilerEndIf
          ;
        Case '='
          IsStartOfStatement = #False
          If IsInsideDefinition And OpenParentheses = 1
            ; Entering a formal parameter default value expression (e.g., Param = #DEFAULT_CONST)
            IsInsideDefaultValue = #True
          EndIf
          NextTokenIsPotentialVariable = #False
          SkipToken
          ;
        Case ','
          IsStartOfStatement = #False
          ExpectedTargetType = 0
          If IsInsideDefinition And OpenParentheses = 1
            ; Next formal parameter begins: exit default value mode and clear per-parameter container qualifier
            IsInsideDefaultValue         = #False
            FieldOrVariableSubSpecie$    = ""
            NextTokenIsPotentialVariable = #True
          EndIf
          If (TypeOfElement = #PBBVariable Or TypeOfElement = #PBBField) And OpenParentheses = 0
            NextTokenIsPotentialVariable = #True
            If TypeOfElement = #PBBField
              FieldOrVariableSubSpecie$ = ""
            EndIf
          EndIf
          SkipToken
          ;
        Default
          ; Fast-path: consume standalone separators ('.', '\', '+', '-', '@', '?', etc.) without calling ExtractOrSkipNextToken().
          ; Only '*' among IsPBSeparator() characters can initiate a valid identifier (*Pointer).
          If IsPBSeparator(*Char\c) And *Char\c <> '*'
            IsStartOfStatement = #False
            SkipToken
          Else
            *TokenPosChar = *Char
            Token$        = ExtractOrSkipNextToken(*StringStartingAdr, @*Char)
            If Token$ = ""
              ; Standalone '*' operator (multiplication) not forming a pointer name
              IsStartOfStatement = #False
              SkipToken
            Else
              TokenLCase$ = LCase(Token$)
              FirstChar   = Asc(TokenLCase$)
              *TempChar   = *Char
              SkipSpaces(*TempChar)
              If IsModuleUsage(@*TempChar)
                ; Module scope prefix ('Mod::'): fuse directly into a self-contained "Mod::Ident" token
                IsStartOfStatement = #False
                IsQualified        = #True
                *Char              = *TempChar
                CompilerIf #PBBReport
                  If OPDMod & #OPDFollowsSeparators
                    OPDTokenLine$ + "::"
                  EndIf
                CompilerEndIf
                If IsValidIdentifierStart(FirstChar)
                  LastToken$ = Token$ + "::"
                Else
                  LastToken$ = "::"
                EndIf
                SkipSpaces(*Char)
                *TokenPosChar = *Char
                Token$        = ExtractOrSkipNextToken(*StringStartingAdr, @*Char)
                If Token$ = ""
                  ; Malformed trailing 'Mod::' with no identifier: discard and continue scanning
                  IsQualified = #False
                  Continue
                EndIf
                ; Reject chained scope operators ('Mod1::Mod2::Elem')
                *TempChar = *Char
                SkipSpaces(*TempChar)
                If IsModuleUsage(@*TempChar)
                  While IsModuleUsage(@*TempChar)
                    *Char = *TempChar
                    ExtractOrSkipNextToken(*StringStartingAdr, @*Char)
                    *TempChar = *Char
                    SkipSpaces(*TempChar)
                  Wend
                  IsQualified = #False
                  Continue
                EndIf
                FirstChar   = Asc(LCase(Token$))
                Token$      = LastToken$ + Token$
                TokenLCase$ = LCase(Token$)
              EndIf
              ;
              ; Promote any target role armed by the preceding keyword and intercept variable/container/expression keywords
              ; (*TempChar already points to the first non-space character following Token$)
              UsageTargetType    = ExpectedTargetType
              ExpectedTargetType = 0
              LastToken$         = ""
              ;
              If LastSeparator = '.' And NextTokenIsPotentialVariable And IsValidIdentifierStart(FirstChar)
                ; Capture statement-level default type (e.g., 'Global.Mod::MyStruct' or 'Protected NewList.MyStruct')
                VariableDotType$ = Token$
                IsQualified      = #False
                CompilerIf ParsingMode = #ParsingMode_OnePass
                  LastToken$ = TokenLCase$
                CompilerEndIf
              ElseIf UsageTargetType = 0 And IsQualified = #False And FirstChar >= 'a' And FirstChar <= 'z'
                If LastSeparator = 0 And OpenParentheses = 0 And InterfaceOrStructContainer <= 0 And
                   IsStartOfStatement And FindString(",global,define,protected,static,shared,threaded,", "," + TokenLCase$ + ",")
                  TypeOfElement                = #PBBVariable
                  VariableSpecie$              = UCase(Left(Token$, 1)) + LCase(Mid(Token$, 2))
                  FieldOrVariableSubSpecie$    = ""
                  NextTokenIsPotentialVariable = #True
                  LastToken$                   = TokenLCase$
                ElseIf (LastSeparator = 0 Or LastSeparator = ',') And OpenParentheses = 0 And InterfaceOrStructContainer <= 0 And
                       (IsStartOfStatement Or NextTokenIsPotentialVariable) And
                       FindString(",newlist,newmap,dim,redim,", "," + TokenLCase$ + ",")
                  TypeOfElement = #PBBVariable
                  If VariableSpecie$ = "" : VariableSpecie$ = "Define" : EndIf
                  If TokenLCase$ = "newlist"    : FieldOrVariableSubSpecie$ = "List"
                  ElseIf TokenLCase$ = "newmap" : FieldOrVariableSubSpecie$ = "Map"
                  Else                          : FieldOrVariableSubSpecie$ = "Array"
                  EndIf
                  NextTokenIsPotentialVariable = #True
                  LastToken$                   = TokenLCase$
                ElseIf (TokenLCase$ = "array" Or TokenLCase$ = "list" Or TokenLCase$ = "map") And
                       ((OpenParentheses = 0 And (LastSeparator = 0 Or LastSeparator = ',') And
                         InterfaceOrStructContainer > 0 And TypeOfElement <> #PBBStructure And
                         (IsStartOfStatement Or NextTokenIsPotentialVariable) And *TempChar\c <> '(' And *TempChar\c <> '.') Or
                        (IsInsideDefinition And OpenParentheses = 1 And IsInsideDefaultValue = #False And
                         (LastSeparator = '(' Or LastSeparator = ',') And
                         *TempChar\c <> ',' And *TempChar\c <> ')' And *TempChar\c <> '='))
                  If OpenParentheses = 0 : TypeOfElement = #PBBField : EndIf
                  FieldOrVariableSubSpecie$    = UCase(Left(Token$, 1)) + LCase(Mid(Token$, 2))
                  NextTokenIsPotentialVariable = #True
                  LastToken$                   = TokenLCase$
                ElseIf LastSeparator <> '.' And LastSeparator <> '\' And LastSeparator <> '?' And
                       NextTokenIsPotentialVariable = #False And (IsInsideDefinition = #False Or IsInsideDefaultValue = #True)
                  If (TypeOfElement = #PBBStructure Or TypeOfElement = #PBBInterface) And TokenLCase$ = "extends"
                    ExpectedTargetType = #PBBStructOrInterface
                  CompilerIf ParsingMode = #ParsingMode_OnePass
                  ElseIf (InterfaceOrStructContainer > 0 Or TypeOfElement = #PBBStructure Or TypeOfElement = #PBBInterface) And
                         FindString(",align,structureunion,endstructureunion,endstructure,endinterface,", "," + TokenLCase$ + ",")
                    LastToken$ = TokenLCase$
                  CompilerEndIf
                  ElseIf InterfaceOrStructContainer <= 0 Or IsStartOfStatement = #False
                    Select TokenLCase$
                      Case "goto", "gosub", "restore"
                        ExpectedTargetType = #PBBLabel
                      Case "usemodule", "unusemodule"
                        ExpectedTargetType = #PBBModule
                      Case "undefinemacro"
                        CompilerIf ParsingMode = #ParsingMode_OnePass
                          PBB_HasUndefineMacro = #True
                        CompilerEndIf
                      Case "defined"
                        If *TempChar\c = '('
                          *Char = *TempChar
                          CompilerIf ParsingMode = #ParsingMode_OnePass
                            SkipBalancedParentheses(@*Char, IDEOptionsCharPos)
                            While *TempChar < *Char
                              If *TempChar\c = #CR
                                If TypeOfElement <> #PBBNotAnElement : AppendDeclarationSegment(#True) : EndIf
                                GoNextChar(*TempChar) : SkipSpaces(*TempChar)
                                StartingPhysicalLinePos = CharToPos(*TempChar)
                                Continue
                              EndIf
                              GoNextChar(*TempChar)
                            Wend
                          CompilerElse
                            SkipBalancedParentheses(@*Char)
                          CompilerEndIf
                          LastToken$ = ")"
                        EndIf
                      Case "sizeof", "typeof"
                        If *TempChar\c = '(' : ExpectedTargetType = #PBBTypeOrVar : EndIf
                      Case "offsetof"
                        If *TempChar\c = '(' : ExpectedTargetType = #PBBStructOrInterface : EndIf
                      Case "allocatestructure"
                        If *TempChar\c = '(' : ExpectedTargetType = #PBBStructure : EndIf
                      Case "clearstructure", "initializestructure", "resetstructure", "copystructure", "comparestructure"
                        If *TempChar\c = '('
                          StructFuncParenLevel = OpenParentheses + 1
                          LastToken$           = TokenLCase$
                        EndIf
                    EndSelect
                  EndIf
                  If ExpectedTargetType <> 0 : LastToken$ = TokenLCase$ : EndIf
                EndIf
              EndIf
              If LastToken$ <> ""
                IsStartOfStatement = #False
                LastSeparator      = 0
                Continue
              EndIf
              ;
              ; Assign syntactic target category and parenthesis constraint for the returned token:
              If StructFuncParenLevel > 0 And OpenParentheses = StructFuncParenLevel And
                 UsageTargetType = 0 And LastSeparator = ',' And IsValidIdentifierStart(FirstChar) And *TempChar\c = ')'
                UsageTargetType      = #PBBStructure
                StructFuncParenLevel = 0
              EndIf
              If UsageTargetType <> 0
                UsageParentheses = -1 + Bool(UsageTargetType = #PBBTypeOrVar) * (1 + Bool(*TempChar\c = '('))
              ElseIf LastSeparator = '?'
                UsageTargetType  = #PBBLabel
                UsageParentheses = -1
              ElseIf LastSeparator = '\'
                UsageTargetType  = #PBBField
                UsageParentheses = Bool(*TempChar\c = '(')
              ElseIf LastSeparator = '.' And IsValidIdentifierStart(FirstChar)
                UsageTargetType  = #PBBStructOrInterface
                UsageParentheses = -1
              Else
                UsageParentheses = Bool(*TempChar\c = '(') - Bool(*TempChar\c <> '(' And LastSeparator = '@')
              EndIf
              ;
              LastToken$   = TokenLCase$
              ParsingState = #Parsing_NewToken
              Break
            EndIf
          EndIf
          ;
      EndSelect
    ForEver
  EndIf
EndMacro
;
;
; **********************************************************************************
;
;-                     One-pass project elements exploration
;
; **********************************************************************************
;
Structure HeaderDeferredToken
  TokenLCase.s
  *TokenPosChar.Character
  TargetType.i
  Parentheses.i
  DotTypeLCase.s
EndStructure
;
Structure PendingLabelUsage
  TokenLCase.s
  *TokenPosChar.Character
  ContainerID.i
  ModuleContainer.i
EndStructure
;
Structure PendingVarOrField
  Type.i
  Name.s
  StartingPos.i
  VariableSpecie.s
  VariableSubSpecie.s
  HasParentheses.i
EndStructure
;
; ----------------------------------------------------------------------------------
;                  Helpers for the one pass Elements identification
; ----------------------------------------------------------------------------------
;
Procedure   AddPBBElementFromDeclaration(FileName$, TypeOfElement, ElementName$, StartingPos, EndingPos, Declaration$,
                                         CommentStr$, VariableSpecie$, FieldOrVariableSubSpecie$, HasParentheses,
                                         DirectContainer, Parameters$ = "", EnclosingModuleID = -1, IsDeclareModule = #False)
  ;
  ; Role: Inserts a new element into ElementList() or refreshes an existing/invalidated record in place
  ;       while preserving FixedID and UsageParents, leaving ElementList() positioned on the target element.
  ;
  Protected Found, Refresh = #True, AddedDotType, IDIndex
  Protected ElementNameLCase$, CandidateIDs$, IDStr$, ContinuedLine$, CED_SharedDotType_LCase$
  Protected VarDotType.String
  Protected NewList CEDList.CED_Structure()
  ;
  If TypeOfElement = #PBBField And Left(ElementName$, 1) = "*"
    ElementName$= FastMid(ElementName$, 2)
  EndIf
  ElementNameLCase$= PurifyElementName(ElementName$, 0, @VarDotType)
  VarDotType\s      = LCase(VarDotType\s)
  ;
  If TypeOfElement = #PBBVariable Or TypeOfElement = #PBBField Or (TypeOfElement = #PBBImported And HasParentheses = #False)
    If FieldOrVariableSubSpecie$ = "" And (TypeOfElement = #PBBVariable Or HasParentheses = #False)
      FieldOrVariableSubSpecie$ = "Var"
    EndIf
    If VarDotType\s = "" And Right(ElementName$, 1) <> "$" And (TypeOfElement <> #PBBField Or FieldOrVariableSubSpecie$ <> "")
      ElementName$ + ".i"
      VarDotType\s = "i"
      AddedDotType = #True
    EndIf
  EndIf
  If HasParentheses : ElementName$ + "()" : EndIf
  ;
  If FindMapElement(ElementMapIndex(), ElementNameLCase$)
    CandidateIDs$ = ElementMapIndex()
    Repeat
      IDIndex + 1
      IDStr$= StringField(CandidateIDs$, IDIndex, ",")
      If IDStr$ = "" : Break : EndIf
      If SelectPBBElement(Val(IDStr$)) = 0 : Continue : EndIf
      If TypeOfElement <> #PBBModule And ElementList()\Type <> #PBBModule And ElementList()\EnclosingModuleID <> EnclosingModuleID
        Continue
      EndIf
      ;
      If ElementList()\FileName = FileName$ And ElementList()\StartingPos = StartingPos
        Found = ElementList()\FixedID : Break
      ElseIf ElementList()\Continued <> "" And FindString(#CR$ + ElementList()\Continued, #CR$+ FileName$ + #TAB$ + Str(StartingPos) + #TAB$)
        CED_StringToList(ElementList()\Continued, CEDList())
        ForEach CEDList()
          If CEDList()\Filename = FileName$ And CEDList()\StartingPos = StartingPos
            Found                   = ElementList()\FixedID
            CEDList()\Declaration   = Declaration$
            Refresh                 = #False
            ElementList()\Continued = CED_ListToString(CEDList())
            Break 2
          EndIf
        Next
      EndIf
      ;
      If (ElementList()\Type = #PBBDeclare And TypeOfElement = #PBBProcedure) Or
         (ElementList()\Type = #PBBModule And ElementList()\IsDeclareModule And TypeOfElement = #PBBModule And IsDeclareModule = #False) Or
         (ElementList()\Type = #PBBEnumeration And TypeOfElement = #PBBEnumeration And ElementList()\StartingPos > 0 And
          IsPositionBefore(FileName$, StartingPos, ElementList()\FileName, ElementList()\StartingPos))
        If ElementList()\StartingPos > 0
          ElementList()\Continued = CED_CreateRecord(ElementList()\FileName, ElementList()\StartingPos, ElementList()\EndingPos, ElementList()\Declaration) + ElementList()\Continued
        EndIf
        Found = ElementList()\FixedID : Break
        ;
      ElseIf (ElementList()\Type = #PBBProcedure And TypeOfElement = #PBBDeclare) Or
             ((TypeOfElement = #PBBEnumeration Or TypeOfElement = #PBBModule Or TypeOfElement = #PBBImport) And ElementList()\Type = TypeOfElement)
        If ElementList()\StartingPos > 0 Or ElementList()\Type <> TypeOfElement
          ContinuedLine$ = CED_CreateRecord(FileName$, StartingPos, EndingPos, Declaration$)
          Refresh        = #False
        EndIf
        Found = ElementList()\FixedID : Break
        ;
      ElseIf ElementList()\Type = TypeOfElement
        If TypeOfElement = #PBBVariable
          If VariableSpecie$ = "Shared" Or ElementList()\VariableSpecie = "Shared" Or
             ((ElementList()\VarDotType_LCase = VarDotType\s Or
               (FindString(VarDotType\s, "::") = 0 And Right(ElementList()\VarDotType_LCase, Len(VarDotType\s) + 2) = "::" + VarDotType\s)) And
              VariableSpecie$= ElementList()\VariableSpecie And FieldOrVariableSubSpecie$ = ElementList()\VariableSubSpecie)
            If ElementList()\StartingPos = 0
              Found = ElementList()\FixedID
            ElseIf ElementList()\VariableSpecie = "Shared" And VariableSpecie$ <> "Shared"
              AssignSharedCEDDotType(CED_SharedDotType_LCase$, #True, ElementList()\VarDotType_LCase, ElementList()\VarAddedDotType)
              If VariableSpecie$ = "Define" : VariableSpecie$ = "Shared" : EndIf
              ElementList()\Continued + CED_CreateRecord(ElementList()\FileName, ElementList()\StartingPos, ElementList()\EndingPos,
                                                         ElementList()\Declaration, ElementList()\LocalVarContainer, CED_SharedDotType_LCase$)
              Found = ElementList()\FixedID
            Else
              AssignSharedCEDDotType(CED_SharedDotType_LCase$, (VariableSpecie$ = "Shared"), VarDotType\s, AddedDotType)
              If ElementList()\VariableSpecie = "Define" And VariableSpecie$ = "Shared" : ElementList()\VariableSpecie = "Shared" : EndIf
              ContinuedLine$= CED_CreateRecord(FileName$, StartingPos, EndingPos, Declaration$, DirectContainer, CED_SharedDotType_LCase$)
              Found          = ElementList()\FixedID
              Refresh        = #False
            EndIf
            Break
          EndIf
        ElseIf TypeOfElement = #PBBField Or TypeOfElement = #PBBConstant
          If TypeOfElement = #PBBConstant Or ElementList()\LocalVarContainer = DirectContainer
            If ElementList()\StartingPos > 0
              ContinuedLine$ = CED_CreateRecord(FileName$, StartingPos, EndingPos, Declaration$, DirectContainer)
              Refresh        = #False
            EndIf
            Found = ElementList()\FixedID : Break
          EndIf
        Else
          Found = ElementList()\FixedID : Break
        EndIf
      EndIf
    ForEver
  EndIf
  ;
  If Found = #False
    AddPBBElement(TypeOfElement, ElementName$, ElementNameLCase$)
  ElseIf ContinuedLine$ <> "" And FindString(ElementList()\Continued, ContinuedLine$) = 0
    ElementList()\Continued + ContinuedLine$
  EndIf
  ;
  If Refresh
    If FindString(ElementName$, "::")
      ElementList()\Name = StringField(ElementName$, 1, ".") + "." + FastMid(ElementName$, FindString(ElementName$, "::") + 2)
    Else
      ElementList()\Name = ElementName$
    EndIf
    ElementList()\VarDotType_LCase  = VarDotType\s
    ElementList()\VarAddedDotType   = AddedDotType
    ElementList()\PurifiedName      = ElementNameLCase$
    ElementList()\Type              = TypeOfElement
    ElementList()\VariableSpecie    = VariableSpecie$
    ElementList()\VariableSubSpecie = FieldOrVariableSubSpecie$
    ElementList()\FileName          = FileName$
    ElementList()\Declaration       = Declaration$
    If (TypeOfElement = #PBBVariable And FindString(#PBBLocalVariables$, VariableSpecie$)) Or
       TypeOfElement = #PBBField Or ((TypeOfElement = #PBBConstant Or TypeOfElement = #PBBImported) And DirectContainer > 0)
      ElementList()\LocalVarContainer = DirectContainer
    Else
      ElementList()\LocalVarContainer = 0
    EndIf
    If TypeOfElement = #PBBMacro
      ElementList()\HasParentheses  = Bool(FindString(ElementList()\Declaration, "(") > 0)
      ElementList()\MacroVirtualPos = ComputePositionInVirtualWholeProjectFile(FileName$, StartingPos)
    Else
      ElementList()\HasParentheses  = HasParentheses
    EndIf
    ElementList()\StartingPos = StartingPos
    ElementList()\EndingPos   = EndingPos
    If TypeOfElement = #PBBMacro Or TypeOfElement = #PBBProcedure Or TypeOfElement = #PBBDeclare Or TypeOfElement = #PBBImported
      If ElementList()\Parameters <> "" : RemoveParametersFromMapIndex(ElementList()\Parameters, ElementList()\FixedID) : EndIf
      ElementList()\Parameters = Parameters$
      If TypeOfElement = #PBBMacro Or TypeOfElement = #PBBProcedure
        UpdateParametersInMapIndex(ElementList()\Parameters, ElementList()\FixedID, #False)
      EndIf
    EndIf
    ElementList()\Comment           = CommentStr$
    ElementList()\EnclosingModuleID = EnclosingModuleID
    If TypeOfElement = #PBBModule
      ElementList()\IsDeclareModule = IsDeclareModule
    Else
      ElementList()\IsDeclareModule = Bool(ElementList()\IsDeclareModule Or IsDeclareModule)
    EndIf
  EndIf
  ;
  If ElementList()\Continued <> "" And FindString(ElementList()\Continued, #CR$, FindString(ElementList()\Continued, #CR$) + 1)
    CED_SortBasedOnPositions()
  EndIf
  CompilerIf #PBBReport
    If OPDMod & #OPDElementCreation
      If Found = #False
        OPDReport("--> Creation of " + PBBTypeNames$(TypeOfElement) + " " + StringField(ElementList()\Name, 1, ")") +
                  ReplaceString(Trim(ElementList()\Parameters, #TAB$), #TAB$, ", ") + ")")
      ElseIf Refresh
        OPDReport("--> Update of " + PBBTypeNames$(TypeOfElement) + " " + ElementList()\Name)
      EndIf
    EndIf
  CompilerEndIf
EndProcedure
;
Procedure.s ComputeConstantValue(Declaration$, ModuleContainer = -1, FileName$ = "", PosInCode = 0)
  ;
  ; Role: Evaluates the right-hand side of a constant declaration via EvalConstExpr(),
  ;       resolving global, module, 'UseModule', and 'Mod::#Const' references, or returns "UnResolved".
  ;
  Protected *Char.Character = @Declaration$
  Protected Val$
  ;
  ; Locate the assignment '=' before any trailing comment (no quotes can appear before '=' in a constant header)
  While *Char\c <> 0 And *Char\c <> ';' And *Char\c <> '=' : GoNextChar(*Char) : Wend
  If *Char\c = '='
    Val$ = EvalConstExpr(PeekS(*Char + SizeOf(Character)), FileName$, PosInCode, ModuleContainer, #False, #False)
    If Val$ <> ""
      ProcedureReturn Val$
    EndIf
  EndIf
  ProcedureReturn "UnResolved"
EndProcedure
;
Procedure   UsagePropagation()
  ;
  ; Role: Propagates the 'Used' status from initially used elements (referenced in main scope)
  ; to all their referenced children (UsageChildren) across the call graph.
  ;
  Protected IChildPos, TypeOfElement
  Protected ChildrenList$, IChild$
  Protected *Elem.Elements, *Child.Elements
  Protected NewList ToProcess.i()
  ;
  For TypeOfElement = #PBBProcedure To #EndEnumPBBElementTypes - 1
    ListOfUsedElementsNbr(TypeOfElement) = 0
  Next
  ;
  ; 1. Seed the queue with all elements referenced directly in main out-of-element code
  ForEach ElementList()
    If ElementList()\OOE_UsagePos = ""
      ElementList()\Used = #False
    Else
      ElementList()\Used = #True
      If ElementList()\Type < #EndEnumPBBElementTypes
        ListOfUsedElementsNbr(ElementList()\Type) + 1
      EndIf
      If ElementList()\UsageChildren <> ""
        AddElement(ToProcess())
        ToProcess() = ElementList()\FixedID
      EndIf
    EndIf
  Next
  ;
  ; 2. Propagate usage through the call graph in O(1) per edge via ElementFixedIndexes()
  While FirstElement(ToProcess())
    *Elem = ElementFixedIndexes(ToProcess())
    DeleteElement(ToProcess())
    If *Elem And *Elem\UsageChildren <> ""
      ChildrenList$ = *Elem\UsageChildren
      IChildPos     = 1
      Repeat
        IChildPos + 1
        IChild$ = StringField(ChildrenList$, IChildPos, "/")
        If IChild$ = "" : Break : EndIf
        *Child = ElementFixedIndexes(Val(IChild$))
        If *Child And *Child\Used = #False
          *Child\Used = #True
          If *Child\Type < #EndEnumPBBElementTypes
            ListOfUsedElementsNbr(*Child\Type) + 1
          EndIf
          If *Child\UsageChildren <> ""
            AddElement(ToProcess())
            ToProcess() = *Child\FixedID
          EndIf
        EndIf
      ForEver
    EndIf
  Wend
EndProcedure
;
Procedure   ResetOneElementLineage(IncludingParents = #False)
  ;
  ; Role: Prepares an update of an element's lineage (IncludingParents = #False)
  ; or prepares the complete deletion of an element (IncludingParents = #True).
  ;
  Protected FixedID       = ElementList()\FixedID
  Protected ContainerTag$ = #TAB$ + Str(FixedID) + #TAB$
  Protected RelList$, RelID$
  Protected RelPos, CEDModified
  Protected *RelElem.Elements
  Protected NewList CEDList.CED_Structure()
  ;
  ; 1. Notify all children: remove this element from their UsageParents field
  RelList$ = ElementList()\UsageChildren
  If RelList$ <> ""
    RelPos = 1
    Repeat
      RelPos + 1
      RelID$= StringField(RelList$, RelPos, "/")
      If RelID$ = "" : Break : EndIf
      *RelElem = ElementFixedIndexes(Val(RelID$))
      If *RelElem
        *RelElem\UsageParents = RemoveIndexFromList(FixedID, *RelElem\UsageParents)
      EndIf
    ForEver
    ElementList()\UsageChildren = ""
  EndIf
  ;
  ; 2. Erase 'Continued' declarations referencing this procedural container (fast-filtered via ContainerTag$)
  If ElementList()\Type = #PBBProcedure Or ElementList()\Type = #PBBMacro
    PushListPosition(ElementList())
    ForEach ElementList()
      If ElementList()\Continued <> "" And FindString(ElementList()\Continued, ContainerTag$)
        CEDModified = #False
        CED_StringToList(ElementList()\Continued, CEDList())
        ForEach CEDList()
          If CEDList()\DirectContainer = FixedID
            DeleteElement(CEDList())
            CEDModified = #True
          EndIf
        Next
        If CEDModified
          ElementList()\Continued = CED_ListToString(CEDList())
        EndIf
      EndIf
    Next
    PopListPosition(ElementList())
  EndIf
  ;
  ; 3. Notify all parents (only when completely deleting the element)
  If IncludingParents
    RelList$ = ElementList()\UsageParents
    If RelList$ <> ""
      RelPos = 1
      Repeat
        RelPos + 1
        RelID$= StringField(RelList$, RelPos, "/")
        If RelID$ = "" : Break : EndIf
        If RelID$ <> "-1"
          *RelElem = ElementFixedIndexes(Val(RelID$))
          If *RelElem
            *RelElem\UsageChildren = RemoveIndexFromList(FixedID, *RelElem\UsageChildren)
          EndIf
        EndIf
      ForEver
      ElementList()\UsageParents = ""
    EndIf
  EndIf
  ;
  ProcedureReturn #True
EndProcedure
;
Macro       QueueHeaderDeferredToken(TokLCase, TokPos, Target, Paren, DotType = "")
  ;
  ; Role: Appends a header token (carrying "mod::ident" when qualified) to HeaderDeferredTokens()
  ;       so its usage link is deferred until the owning container element FixedID is created at end-of-statement.
  ;
  AddElement(HeaderDeferredTokens())
  HeaderDeferredTokens()\TokenLCase    = TokLCase
  HeaderDeferredTokens()\TokenPosChar  = TokPos
  HeaderDeferredTokens()\TargetType    = Target
  HeaderDeferredTokens()\Parentheses   = Paren
  HeaderDeferredTokens()\DotTypeLCase  = DotType
EndMacro
;
Macro       AppendDeclarationSegment(ScanToEOL)
  ;
  ; Role: Extracts any trailing comment and appends the current physical statement slice
  ;       (stripped of comments and trailing spaces) to Declaration$ in OnePassSetElements_SegmentExam().
  ;
  If StartingPhysicalLinePos > StartingCodeLinePos
    Pos = StartingPhysicalLinePos
  Else
    Pos = StartingCodeLinePos
  EndIf
  CompilerIf ScanToEOL
    *ScanEOLChar = PosToChar(*StringStartingAdr, Pos)
    LastComment$ = GoToEOL_AndExtractComment(@*ScanEOLChar)
    EndingPos    = CharToPos(*ScanEOLChar) - 1
  CompilerElse
    EndingPos    = CharToPos(*Char) - 1
  CompilerEndIf
  If ScanToEOL Or *Char\c = #CR Or *Char\c = 0
    If LastComment$
      EndingPos - Len(LastComment$)
      If Trim(LastComment$) <> ";" : CommentStr\s = LastComment$ : EndIf
      LastComment$ = ""
    EndIf
  ElseIf CommentStr\s = ""
    *ScanEOLChar = *Char
    CommentStr\s = GoToEOL_AndExtractComment(@*ScanEOLChar)
  EndIf
  LenAddDeclaration = EndingPos - Pos + 1
  If LenAddDeclaration > 0
    If Declaration$ : Declaration$ + " " : EndIf
    Declaration$ + RTrim(FastMid(*SearchString\s, Pos, LenAddDeclaration))
  EndIf
EndMacro
;
Macro       QueuePendingLabel(LabelsList, TokLCase, TokPos, ActiveCont, ModCont)
  ;
  ; Role: Queues an unresolved label reference in PendingLabels() so forward label targets
  ;       declared further down in the same file are linked at end-of-file.
  ;
  AddElement(LabelsList)
  LabelsList\TokenLCase      = TokLCase
  LabelsList\TokenPosChar    = TokPos
  LabelsList\ContainerID     = ActiveCont
  LabelsList\ModuleContainer = ModCont
EndMacro
;
Procedure   CommitPendingVarsAndFields(List FieldsOrVariablesToCreate.PendingVarOrField(),
                                       FileName$, *StringStartingAdr, *AnchorChar.Character,
                                       EndingPos, Declaration$, CommentStr$,
                                       InterfaceOrStructContainer, MacroContainer, ProcContainer,
                                       ModuleContainer, IsDeclareModule, *MacroTokens.String)
  ;
  ; Role: Finalizes all variables or structure/interface fields buffered during the current statement,
  ; registers their enclosing container and '.MyStruct' / '.Mod::MyStruct' type dependencies,
  ; and returns the FixedID of the last created element (leaving ElementList() positioned on it).
  ;
  Protected TargetContainer, ActiveContainer
  Protected UsageDotType$
  ;
  While FirstElement(FieldsOrVariablesToCreate())
    If FieldsOrVariablesToCreate()\Type = #PBBField
      TargetContainer = InterfaceOrStructContainer
    ElseIf MacroContainer > 0
      TargetContainer = MacroContainer
    Else
      TargetContainer = ProcContainer
    EndIf
    ;
    AddPBBElementFromDeclaration(FileName$, FieldsOrVariablesToCreate()\Type,
                                 FieldsOrVariablesToCreate()\Name, FieldsOrVariablesToCreate()\StartingPos, EndingPos,
                                 Declaration$, CommentStr$, FieldsOrVariablesToCreate()\VariableSpecie,
                                 FieldsOrVariablesToCreate()\VariableSubSpecie, FieldsOrVariablesToCreate()\HasParentheses,
                                 TargetContainer, "", ModuleContainer, IsDeclareModule)
    ActiveContainer = ElementList()\FixedID
    If MacroContainer > 0 And *MacroTokens
      *MacroTokens\s + LCase(ElementList()\Name) + #TAB$ + Str(ActiveContainer) + #CR$
    EndIf
    If FieldsOrVariablesToCreate()\Type = #PBBField And InterfaceOrStructContainer > 0
      LinkElementUsage(ActiveContainer, InterfaceOrStructContainer)
    EndIf
    ;
    ; Link custom '.MyStruct' or '.Mod::MyStruct' type and canonicalize VarDotType_LCase to the target structure's true module:
    UsageDotType$ = ElementList()\VarDotType_LCase
    If UsageDotType$ <> "" And ElementList()\VarAddedDotType = #False And
       (FindString(UsageDotType$, "::") Or FindMapElement(PBBasicNativeTypes(), UsageDotType$) = 0)
      If CheckAndMarkTokenUsage(*StringStartingAdr, @*AnchorChar, FileName$, UsageDotType$, 0, ActiveContainer, "", #PBBStructOrInterface, "", ModuleContainer) And
         ElementList()\EnclosingModuleID > 0 And FindString(UsageDotType$, "::") = 0
        UsageDotType$ = LCase(GetEnclosingModuleName(ElementList()\EnclosingModuleID)) + "::" + UsageDotType$
      EndIf
      SelectPBBElement(ActiveContainer)
      ElementList()\VarDotType_LCase = UsageDotType$
    EndIf
    ;
    DeleteElement(FieldsOrVariablesToCreate(), #True)
  Wend
  ;
  ProcedureReturn ActiveContainer
EndProcedure
;
Procedure.s FindPreviousEnumerationValue(EnumContainer, FileName$, StartingCodeLinePos)
  ;
  ; Role: For multi-part named enumerations, scans constants already linked to 'EnumContainer'
  ; and returns the evaluated value of the latest valid constant declared chronologically before
  ; 'StartingCodeLinePos', restoring ElementList() selection to 'EnumContainer' before returning.
  ;
  Protected EnumUsageParents$ = ElementList()\UsageParents
  Protected BestPrevFile$, EnumValue$
  Protected BestPrevPos, Pos = 1, ConstID
  ;
  If EnumUsageParents$ <> ""
    Repeat
      Pos + 1
      ConstID = Val(StringField(EnumUsageParents$, Pos, "/"))
      If ConstID = 0 : Break : EndIf
      If SelectPBBElement(ConstID) And ElementList()\Type = #PBBConstant And
         ElementList()\LocalVarContainer = EnumContainer And ElementList()\StartingPos > 0
        If (ElementList()\FileName <> FileName$ Or ElementList()\StartingPos < StartingCodeLinePos) And
           IsPositionBefore(ElementList()\FileName, ElementList()\StartingPos, FileName$, StartingCodeLinePos)
          If BestPrevFile$ = "" Or IsPositionBefore(BestPrevFile$, BestPrevPos, ElementList()\FileName, ElementList()\StartingPos)
            BestPrevFile$ = ElementList()\FileName
            BestPrevPos   = ElementList()\StartingPos
            EnumValue$    = ElementList()\Value
          EndIf
        EndIf
      EndIf
    ForEver
    SelectPBBElement(EnumContainer)
  EndIf
  ;
  ProcedureReturn EnumValue$
EndProcedure
;
Procedure   FlushHeaderDeferredTokens(List HeaderDeferredTokens.HeaderDeferredToken(),
                                      List PendingLabels.PendingLabelUsage(),
                                      *StringStartingAdr, FileName$,
                                      ActiveContainer, ScopeContainer, HeaderParams$,
                                      WithVar$, ModuleContainer)
  ;
  ; Role: Resolves and links all tokens deferred while parsing a declaration or container header
  ; (return '.Type', 'Extends' parent, formal parameter '.MyStruct' types, parameter default values,
  ; Structure 'Align #Const', Enumeration bounds, and '#Const = ...' expressions) directly to
  ; 'ActiveContainer' while evaluating lexical visibility within 'ScopeContainer', then clears HeaderDeferredTokens().
  ;
  Protected *TokPosChar.Character
  Protected MatchedElementID
  ;
  If ActiveContainer > 0
    ForEach HeaderDeferredTokens()
      If HeaderDeferredTokens()\TokenLCase <> ""
        *TokPosChar      = HeaderDeferredTokens()\TokenPosChar
        MatchedElementID = CheckAndMarkTokenUsage(*StringStartingAdr, @*TokPosChar, FileName$,
                                                  HeaderDeferredTokens()\TokenLCase, 0, ScopeContainer, HeaderParams$,
                                                  HeaderDeferredTokens()\TargetType, WithVar$, ModuleContainer,
                                                  HeaderDeferredTokens()\Parentheses, HeaderDeferredTokens()\DotTypeLCase,
                                                  0, ActiveContainer)
        If MatchedElementID = 0 And HeaderDeferredTokens()\TargetType = #PBBLabel
          QueuePendingLabel(PendingLabels(), HeaderDeferredTokens()\TokenLCase, *TokPosChar, ActiveContainer, ModuleContainer)
        ElseIf MatchedElementID > 0 And MatchedElementID <> ActiveContainer And ElementList()\Type = #PBBMacro
          ResolveMacroTokens(MatchedElementID, ActiveContainer, FileName$, *StringStartingAdr, @*TokPosChar, 0, ModuleContainer)
        EndIf
      EndIf
    Next
    SelectPBBElement(ActiveContainer)
  EndIf
  ClearList(HeaderDeferredTokens())
EndProcedure
;
; ----------------------------------------------------------------------------------
;                   ---- Master one-pass exploration procedures ----
; ----------------------------------------------------------------------------------
;
Declare ResetPBBAnalysisCaches()
Declare ClearAllElementsAndData()
;
Procedure OnePassSetElements_SegmentExam(FileName$, PriorityMode = #WorkInBackGround, CLE_StartTime = 0, *Limits.Point = 0)
  ;
  ; Role: Single-pass master procedure that identifies, registers, and tracks all elements within a source file.
  ; Parameters:
  ;   FileName$     - Full path of the file to parse.
  ;   PriorityMode  - Execution priority (#WorkInBackGround or #FinishCompletionNow).
  ;   CLE_StartTime - Timestamp marking the start of the current timeslice.
  ;   *Limits       - Optional pointer to a Point structure defining (X=StartPos, Y=EndPos) range boundaries.
  ; Prerequisites: PBB_FileListing() must be populated and file contents accessible.
  ; Output: Returns #ExplorationState_Completed when finished, or #ExplorationState_Pending if execution was paused.
  ;
  Static mFileName$ = "", *MemPosChar.Character = 0
  Static NewList PendingLabels.PendingLabelUsage()
  ;
  ; --------------------------------------------------------------------
  ;                             Declarations
  ; --------------------------------------------------------------------
  ;{
  Protected Token$, TokenLCase$, LastToken$, FirstChar,                       ; Data for the currently parsed token
            LastSeparator, ParsingState
  Protected CurrentMacroTokens.String
  Protected TypeOfElement, ElementName$, Declaration$, LenAddDeclaration,     ; Field data for newly created elements
            CommentStr.String, LastComment$
  Protected VariableSpecie$, FieldOrVariableSubSpecie$, VariableDotType$      ; Descriptors for variable creation
  Protected OpenParentheses, StructFuncParenLevel, NextTokenIsPotentialVariable ; Flags specifying
  Protected IsStartOfStatement, LineToBeContinued,                              ;   a particular parsing state (mode)
            HasParentheses, IsInsideDefinition, IsInsideDefaultValue,
            IsDeclareModule, IsQualified, ColonPos, ExpectedTargetType
  Protected ActiveParams$, ProcParams$, MacroParams$, MacroContainer, MacroStartCDepth, ProcContainer, ; Management of Procedure, Macro, Structure,
            InterfaceOrStructContainer, EnumContainer, ImportContainer, ModuleContainer, ;   Enumeration, Import, Module, and other block
            ProcEndingPos, MacroEndingPos, ActiveEndingPos, BlockEndingPos, ModuleEndingPos,
            EnumValue$, IsBinaryEnum                                                     ;   declarations.
  Protected *Char.Character, *TempChar.Character, *TokenPosChar.Character,    ; String navigators
            *ScanEOLChar.Character
  Protected StartingCodeLinePos, StartingPhysicalLinePos, Pos, EndingPos,     ; Position offsets inside the source code
            CharLimitEnd, TempPos
  Protected NewList FieldsOrVariablesToCreate.PendingVarOrField()             ; Buffers variables or structure/interface fields
                                                                              ;   declared in the current statement
  Protected NewList HeaderDeferredTokens.HeaderDeferredToken()                ; Stores header '.Type', 'Extends', and default-value
                                                                              ;   tokens until the owning element FixedID exists
  Protected *SearchString.String           = GetPointedCodeFromFile(FileName$, #False)  ; Pre-filtered code buffers (stripped of #LF and #TAB)
  Protected *SearchString_LCase            = GetPointedCodeFromFile(FileName$, #True)
  Protected ReturnTypeToken$, BareTokenLCase$, WithVar$, MatchedElementID,
            UsageDotType$, UsageParentheses, UsageTargetType, ActiveContainer,
            InElementMap, CandidateIDs.String, EndExpression$, HeaderParams$
  Protected EncompassingModule.EEDStruct
  Protected *FileEntry.FileListingStruct
  ;
  If *SearchString = 0 Or *SearchString\s = ""
    ; Empty file name or unreadable buffer: reset static state and mark as completed
    mFileName$  = ""
    *MemPosChar = 0
    ClearList(PendingLabels())
    ProcedureReturn #ExplorationState_Completed
  Else
    Protected *StringStartingAdr = @*SearchString\s
    Protected IDEOptionsCharPos  = PosToChar(*StringStartingAdr, FindIDENotations(*SearchString))
    IsFileInMemory(FileName$)
    *FileEntry = @PBB_FileListing()
    If ExplorationState = #ExplorationState_Completed
      *FileEntry\OnePassCDepth = 0
    EndIf
  EndIf
  ;}
  ;
  ; --------------------------------------------------------------------
  ;                            Initialization
  ; --------------------------------------------------------------------
  ;{
  If mFileName$ <> FileName$ Or *MemPosChar = 0
    ClearList(PendingLabels())
    mFileName$ = FileName$
  EndIf
  ;
  ProcContainer              = -1
  InterfaceOrStructContainer = -1
  ModuleContainer            = -1
  ActiveParams$              = #TAB$
  ProcParams$                = #TAB$
  MacroParams$               = #TAB$
  *Char                      = *MemPosChar
  ;
  If *Limits And *Limits\x
    ; Partial scanning requested within specific boundary positions
    If *Char = 0 : *Char = PosToChar(*StringStartingAdr, *Limits\x) : EndIf
    ;
    ; Ensure at least one full physical line is examined:
    If *Limits\y < *Limits\x + 1 : *Limits\y = *Limits\x + 1 : EndIf
    *Limits\y = FindString(*SearchString\s, #CR$, *Limits\y)
    If *Limits\y = 0 : *Limits\y = Len(*SearchString\s) + 1 : EndIf
    CharLimitEnd = PosToChar(*StringStartingAdr, *Limits\y)
    ;
    ; When starting a partial or segmented scan inside an already-indexed Module or DeclareModule block,
    ; restore the enclosing module context so newly parsed elements inherit the proper EnclosingModuleID:
    If *MemPosChar = 0 And ListSize(ElementList()) > 0
      GetEncompassingLimitsAndName_ListMode(#PBBModule, FileName$, *Limits\x, *SearchString, @EncompassingModule)
      If EncompassingModule\ElementName$ <> "" And EncompassingModule\EndingPos > 0 And
         GetModuleID(EncompassingModule\ElementName$)
        ModuleContainer = ElementList()\FixedID
        ModuleEndingPos = EncompassingModule\EndingPos
        IsDeclareModule = Bool(EncompassingModule\TypeOfElementName$ = "DeclareModule")
      EndIf
    EndIf
  ElseIf *Char = 0
    ; Full scanning of the file from the beginning of the buffer
    *Char = *StringStartingAdr
  EndIf
  ;
  SkipSpaces(*Char) ; Skip leading spaces at the beginning of the line
  ParsingState            = #Parsing_Start
  StartingPhysicalLinePos = CharToPos(*Char)
  StartingCodeLinePos     = CharToPos(*Char)
  IsStartOfStatement      = #True
  ;
  ResetParsingModeFlags
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDIndicatesFile
      OPDReport("----------------------------------------------------------------------------------")
      OPDReport("Analyzing " + GetFilePart(FileName$) + "  -  *Limits = " + Str(*Limits))
      If *Limits
        OPDReport("    *Limits\x : from line " + Str(GetNoLineFromStrAndPos(*SearchString, *Limits\x)) +
                  "  to line " + Str(GetNoLineFromStrAndPos(*SearchString, *Limits\y)))
        Protected sp = *Limits\x
      Else
        sp = 1
      EndIf
      Protected t = FindString(*SearchString\s, #CR$, sp)
      OPDReport("    First line = " + Mid(*SearchString\s, sp, t - sp))
      OPDReport("----------------------------------------------------------------------------------")
    EndIf
  CompilerEndIf
  ;}
  ;
  ; --------------------------------------------------------------------
  ;                               Scanning
  ; --------------------------------------------------------------------
  ;
  Repeat
    GetNextToken(#ParsingMode_OnePass)
    ;
    If ParsingState = #Parsing_NewToken
      ;
      CompilerIf #PBBReport
        If OPDMod & #OPDOnePassTiming
          StartTimer(3, "Elements recognition")
        EndIf
      CompilerEndIf
      ;
      ; ------------- Conditional compilation directives (CompilerIf / CompilerSelect) -------------
      If IsStartOfStatement And Left(TokenLCase$, 8) = "compiler"
        EvalContext\ProcedureName = ""
        EvalContext\InMacroDef    = Bool(MacroContainer > 0)
        If ProcContainer > 0 And SelectPBBElement(ProcContainer)
          EvalContext\ProcedureName = ElementList()\Name
        EndIf
        If ProcessCompilerDirectiveAtStatement(TokenLCase$, *TokenPosChar, *StringStartingAdr, @*Char,
                                               *FileEntry\OnePassCStack(), @*FileEntry\OnePassCDepth,
                                               Bool(MacroContainer > 0) * MacroStartCDepth,
                                               FileName$, 0, ModuleContainer,
                                               Bool(MacroContainer > 0) * MacroContainer + Bool(MacroContainer = 0) * ProcContainer,
                                               Bool(ExplorationState <> #ExplorationState_Completed), #True, IDEOptionsCharPos)
          EvalContext\ProcedureName = ""
          EvalContext\InMacroDef    = #False
          CompilerIf #PBBReport
            If OPDMod & #OPDOnePassTiming
              ReapTimer(3)
            EndIf
          CompilerEndIf
          Continue
        EndIf
        EvalContext\ProcedureName = ""
        EvalContext\InMacroDef    = #False
      EndIf
      ;
      ; ------------------- Label declarations ----------------------
      ;{
      If IsStartOfStatement And InterfaceOrStructContainer <= 0 And IsStandaloneLabel(TokenLCase$, *Char)
        ; Token is confirmed to be a standalone code label
        HasParentheses = #False
        TypeOfElement  = #PBBLabel
        ElementName$   = Token$
        ;
        ; Advance to end of line (no executable statements can follow a label on the same segment):
        GoNextChar(*Char)
        LastComment$ = GoToEOL_AndExtractComment(@*Char)
        Continue
      EndIf
      ;}
      ;
      If *TokenPosChar\c = '#' And IsStartOfStatement
        ;
        ; ------------------ Constant declaration statement --------------------
        ;{
        ElementName$   = Token$
        TypeOfElement  = #PBBConstant
        HasParentheses = #False
        Continue
        ;}
      ElseIf OpenParentheses = 0
        ;
        ; -------------------- Imported function / variable declarations --------------------
        If ImportContainer > 0
          If IsStartOfStatement And TokenLCase$ <> "endimport" And Left(TokenLCase$, 8) <> "compiler"
            TypeOfElement    = #PBBImported
            ElementName$     = Token$
            ReturnTypeToken$ = ExtractOptionalDotType(*StringStartingAdr, @*Char)
            HasParentheses   = CheckForParenthesis(@*Char)
            If ReturnTypeToken$ <> ""
              UsageDotType$ = LCase(ReturnTypeToken$)
              If FindString(UsageDotType$, "::") Or FindMapElement(PBBasicNativeTypes(), UsageDotType$) = 0
                QueueHeaderDeferredToken(UsageDotType$, *Char, #PBBStructOrInterface, -1)
              EndIf
              If HasParentheses = #False And Right(ElementName$, 1) <> "$"
                ElementName$ + "." + ReturnTypeToken$
              EndIf
              ReturnTypeToken$ = ""
            EndIf
            If HasParentheses
              ActiveParams$        = #TAB$
              IsInsideDefinition   = #True
              IsInsideDefaultValue = #False
            EndIf
            Continue
          ElseIf TypeOfElement = #PBBImported And TokenLCase$ = "as"
            Continue
          EndIf
        EndIf
        ;
        ; --------------- Variable and Structure/Interface Field declarations ---------------
        ;{
        If InterfaceOrStructContainer > 0 And IsStartOfStatement And Left(TokenLCase$, 8) <> "compiler"
          ; Regular structure field or interface method declaration ('Array'/'List'/'Map' and layout keywords are absorbed by GetNextToken)
          TypeOfElement                = #PBBField
          NextTokenIsPotentialVariable = #True
        EndIf
        ;
        If NextTokenIsPotentialVariable And (TypeOfElement = #PBBVariable Or TypeOfElement = #PBBField)
          ; Pre-record variable or structure/interface field instance:
          ; Queue a new element in FieldsOrVariablesToCreate() and fill its primary fields.
          ; Final attributes (EndingPos, Declaration$, Comment) will be populated at the end of the statement.
          NextTokenIsPotentialVariable = #False
          ElementName$                 = Token$
          ReturnTypeToken$             = ExtractOptionalDotType(*StringStartingAdr, @*Char, VariableDotType$)
          If Right(ElementName$, 1) = "$"
            ReturnTypeToken$ = ""
          ElseIf ReturnTypeToken$ <> ""
            ElementName$ + "." + ReturnTypeToken$
          EndIf
          ;
          HasParentheses = CheckForParenthesis(@*Char)
          If FieldOrVariableSubSpecie$
            If HasParentheses = #False
              ; NewList, NewMap, Dim, or structure List/Map/Array detected without required parentheses:
              ClearList(FieldsOrVariablesToCreate())
              GoToEOL_AndExtractComment(@*Char) ; Go to the end of the physical line,
              ResetParsingModeFlags             ; reset all parsing flags,
              Continue                          ; and discard this malformed declaration.
            EndIf
          ElseIf TypeOfElement = #PBBField And HasParentheses
            ; Interface method declaration with formal parameters inside parentheses:
            IsInsideDefinition   = #True
            IsInsideDefaultValue = #False
          EndIf
          ;
          ; When VariableSpecie$ is "Define" (bare Dim, ReDim, NewList, NewMap, or Define without Protected/Static),
          ; check whether the variable already exists as a Global/Threaded or same-scope variable.
          ; If so, do not create a duplicate element and let execution fall through to usage tracking:
          If TypeOfElement <> #PBBVariable Or VariableSpecie$ <> "Define" Or
             FindElementInList(TokenLCase$, #PBBVariable, StartingCodeLinePos, FileName$, ProcContainer, HasParentheses, LCase(ReturnTypeToken$), 0, ModuleContainer) = 0 Or
             (ProcContainer = -1 And ElementList()\VariableSpecie = "Shared" And ElementList()\LocalVarContainer > 0)
            AddElement(FieldsOrVariablesToCreate())
            FieldsOrVariablesToCreate()\Type              = TypeOfElement
            FieldsOrVariablesToCreate()\Name              = ElementName$
            FieldsOrVariablesToCreate()\StartingPos       = StartingCodeLinePos
            FieldsOrVariablesToCreate()\VariableSpecie    = VariableSpecie$
            FieldsOrVariablesToCreate()\VariableSubSpecie = FieldOrVariableSubSpecie$
            FieldsOrVariablesToCreate()\HasParentheses    = HasParentheses
            ReturnTypeToken$ = ""
            Continue
          EndIf
          ReturnTypeToken$ = ""
          ;
        EndIf
        ;}
        ;
        If IsStartOfStatement = #True And TypeOfElement = #PBBNotAnElement
          ;
          ; Track active 'With ... EndWith' target expression for orphan '\Field' resolution
          If TokenLCase$ = "with"
            WithVar$ = ExtractWithTarget(*Char)
            Continue
          ElseIf TokenLCase$ = "endwith"
            WithVar$ = ""
            Continue
          EndIf
          ;
          ; ------------------ Block element declarations --------------------
          If FindMapElement(PB_BlockDeclarators(), TokenLCase$)
            ; Token is a block element declarator keyword (Procedure, Macro, Structure, etc.)
            ;{
            If Left(TokenLCase$, Len("procedure")) = "procedure" Or
               (Left(TokenLCase$, Len("declare")) = "declare" And TokenLCase$ <> "declaremodule")
              ; If the declaration specifies a return structure/interface type (e.g., "Procedure.MyStruct" or "Procedure.Mod::MyStruct"),
              ; queue it in HeaderDeferredTokens() to link it once the Procedure/Declare element is created:
              UsageDotType$ = LCase(ExtractOptionalDotType(*StringStartingAdr, @*Char))
              If UsageDotType$ <> "" And (FindString(UsageDotType$, "::") Or FindMapElement(PBBasicNativeTypes(), UsageDotType$) = 0)
                QueueHeaderDeferredToken(UsageDotType$, *Char, #PBBStructOrInterface, -1)
              EndIf
            EndIf
            ;
            If Left(TokenLCase$, 6) = "import"
              ElementName$ = ExtractWithTarget(*Char)
              If ElementName$ = #DOUBLEQUOTE$ + #DOUBLEQUOTE$
                ElementName$ = ""
              ElseIf Left(ElementName$, 1) = #DOUBLEQUOTE$ Or Left(ElementName$, 2) = "~" + #DOUBLEQUOTE$
                SkipSpaces(*Char)
                ExtractOrSkipNextString(@*Char, *Char\c, #True)
              EndIf
            Else
              ElementName$ = ExtractOrSkipNextToken(*StringStartingAdr, @*Char)
            EndIf
            If ElementName$ = "" ; Handle anonymous enumerations or empty 'Import ""' blocks:
              ElementName$ = "UnNamed - " + GetFilePart(FileName$) + "/Line " +
                             Str(GetNoLineFromStrAndPos(*SearchString, CharToPos(*Char) - 1))
            EndIf
            ;
            If Left(TokenLCase$, Len("procedure"))       = "procedure"
              TypeOfElement = #PBBProcedure
            ElseIf TokenLCase$                           = "macro"
              TypeOfElement = #PBBMacro
            ElseIf Left(TokenLCase$, Len("enumeration")) = "enumeration"
              TypeOfElement = #PBBEnumeration
            ElseIf TokenLCase$                           = "structure"
              TypeOfElement = #PBBStructure
            ElseIf TokenLCase$                           = "interface"
              TypeOfElement = #PBBInterface
            ElseIf TokenLCase$ = "module" Or TokenLCase$ = "declaremodule"
              ; Match 'declaremodule' before 'Left(..., "declare")' and record public interface state
              TypeOfElement   = #PBBModule
              IsDeclareModule = Bool(TokenLCase$ = "declaremodule")
            ElseIf Left(TokenLCase$, Len("import"))      = "import"
              TypeOfElement = #PBBImport
            ElseIf Left(TokenLCase$, Len("declare"))     = "declare"
              TypeOfElement = #PBBDeclare
            EndIf
            ;
            HasParentheses = CheckForParenthesis(@*Char)
            ;
            If (TypeOfElement = #PBBProcedure Or TypeOfElement = #PBBMacro Or TypeOfElement = #PBBDeclare)
              If (TypeOfElement = #PBBProcedure Or TypeOfElement = #PBBDeclare) And HasParentheses = #False
                ; Procedure or Declare keyword detected without required opening parenthesis:
                ClearList(HeaderDeferredTokens())
                GoToEOL_AndExtractComment(@*Char) ; Go to the end of the physical line,
                ResetParsingModeFlags             ; reset all parsing flags,
                Continue                          ; and discard this malformed declaration.
              EndIf
              ;
              ActiveParams$        = #TAB$
              IsInsideDefinition   = #True
              IsInsideDefaultValue = #False
              CompilerIf #PBBReport
                If OPDTokenLine$ And (OPDMod & #OPDFollowsTokens) And (OPDMod & #OPDFollowsSeparators) = 0
                  OPDTokenLine$ + " ("
                EndIf
              CompilerEndIf
            EndIf
            ;
            Continue
            ;}
          EndIf ; If FindMapElement(PB_BlockDeclarators(), TokenLCase$)
        EndIf ; If IsStartOfStatement = #True And TypeOfElement = #PBBNotAnElement
        ;
      EndIf ; ElseIf OpenParentheses = 0
      ;
      CompilerIf #PBBReport
        If OPDMod & #OPDOnePassTiming
          ReapTimer(3)
          StartTimer(4, "Elements Usage")
        EndIf
      CompilerEndIf
      ;
      ; ------------------------------ Element usage -----------------------------
      ;{
      MatchedElementID = 0
      ;
      ; Formal parameters of procedures, declarations, macros, and interface methods represent local signatures
      ; rather than project elements; record them in ActiveParams$ so downstream body usages are shielded,
      ; while queueing any custom '.MyStruct' parameter types in HeaderDeferredTokens() for deferred linking.
      If IsInsideDefinition And IsInsideDefaultValue = #False And LastSeparator <> '.'
        If FirstChar = '*' Or IsValidIdentifierStart(FirstChar)
          UsageDotType$ = ExtractOptionalDotType(*StringStartingAdr, @*Char)
          If UsageDotType$ <> "" And Right(Token$, 1) <> "$"
            Token$ + "." + UsageDotType$
            UsageDotType$ = LCase(UsageDotType$)
            If FindString(UsageDotType$, "::") Or FindMapElement(PBBasicNativeTypes(), UsageDotType$) = 0
              QueueHeaderDeferredToken(UsageDotType$, *Char, #PBBStructOrInterface, -1)
            EndIf
          EndIf
          If CheckForParenthesis(@*Char)
            Token$ + "()"
          EndIf
          If InterfaceOrStructContainer <= 0
            ActiveParams$ + Token$ + #TAB$
          EndIf
        EndIf
      Else
        ; Re-sync *TempChar in case a 'Define' / 'ReDim' homonym consumed '.Type' above before falling through
        *TempChar = *Char
        SkipSpaces(*TempChar)
        If UsageTargetType = 0 And *TempChar\c = '('
          UsageParentheses = 1
        EndIf
        If IsQualified
          ColonPos        = FindString(TokenLCase$, "::")
          BareTokenLCase$ = FastMid(TokenLCase$, ColonPos + 2)
        Else
          ColonPos        = 0
          BareTokenLCase$ = TokenLCase$
        EndIf
        ;
        ; Reject malformed '::Orphan' tokens, numeric/hex literals, and native scalar types upfront,
        ; then query ElementMapIndex() once without allocating strings on unindexed tokens:
        If ColonPos <> 1 And (*TokenPosChar\c > '9' Or *TokenPosChar\c = '*' Or *TokenPosChar\c = '#') And
           ((UsageTargetType <> #PBBStructOrInterface And UsageTargetType <> #PBBStructure And UsageTargetType <> #PBBTypeOrVar) Or
            IsQualified Or FindMapElement(PBBasicNativeTypes(), BareTokenLCase$) = 0)
          InElementMap = FindMapElement(ElementMapIndex(), BareTokenLCase$)
          ;
          ; Select the active signature parameters and container boundary without cross-contaminating Macro and Procedure scopes:
          If IsInsideDefinition
            HeaderParams$   = ActiveParams$
            ActiveEndingPos = 0
          ElseIf MacroContainer > 0
            HeaderParams$   = MacroParams$
            ActiveEndingPos = MacroEndingPos
          Else
            HeaderParams$   = ProcParams$
            ActiveEndingPos = ProcEndingPos
          EndIf
          ;
          ; Process only indexed elements, forward-declarable labels, or capturable macro tokens
          If InElementMap Or UsageTargetType = #PBBLabel Or
             (MacroContainer > 0 And UsageTargetType <> #PBBField And IsCapturableMacroToken(BareTokenLCase$, IsQualified))
            ; Shield formal parameters once when active parameters exist in the current scope
            If HeaderParams$ <> #TAB$
              If IsQualified
                MatchedElementID = -Bool(MacroContainer > 0 And IsParam(FastLeft(TokenLCase$, ColonPos - 1), HeaderParams$))
              ElseIf UsageTargetType = 0 Or (MacroContainer > 0 And UsageTargetType <> #PBBField)
                MatchedElementID = -Bool(IsParam(BareTokenLCase$, HeaderParams$))
              EndIf
            EndIf
            ;
            If MatchedElementID = 0
              ; Perform a single non-destructive lookahead for optional '.Type' and '(' suffixes
              UsageDotType$ = ""
              If UsageTargetType = 0 And *TempChar\c = '.' And Right(BareTokenLCase$, 1) <> "$"
                CompilerIf #PBBReport
                  Protected mOPDTokenLine$ = OPDTokenLine$
                CompilerEndIf
                UsageDotType$ = LCase(ExtractOptionalDotType(*StringStartingAdr, @*TempChar))
                CompilerIf #PBBReport
                  OPDTokenLine$ = mOPDTokenLine$
                CompilerEndIf
                SkipSpaces(*TempChar)
                If *TempChar\c = '('
                  UsageParentheses = 1
                EndIf
              EndIf
              ;
              ; Macro bodies may reference identifiers not yet declared in ElementMapIndex();
              ; store self-contained TokenLCase$ ("mod::ident"), target role, parenthesis constraint, and '.Type'.
              If MacroContainer > 0 And UsageTargetType <> #PBBField
                CurrentMacroTokens\s + TokenLCase$ + #TAB$ + "-1" + #TAB$ + Str(UsageTargetType) + #TAB$ +
                                       Str(UsageParentheses) + #TAB$ + UsageDotType$ + #CR$
              EndIf
              ;
              ; Resolve indexed elements immediately (or queue #PBBLabel targets for EOF replay and headers for EOL linking)
              If InElementMap Or UsageTargetType = #PBBLabel
                If IsInsideDefinition Or TypeOfElement = #PBBConstant Or TypeOfElement = #PBBImport Or
                   TypeOfElement = #PBBStructure Or TypeOfElement = #PBBInterface Or TypeOfElement = #PBBEnumeration
                  QueueHeaderDeferredToken(TokenLCase$, *TokenPosChar, UsageTargetType, UsageParentheses, UsageDotType$)
                Else
                  If InElementMap
                    CandidateIDs\s = ElementMapIndex()
                  EndIf
                  If InterfaceOrStructContainer > 0
                    ActiveContainer = InterfaceOrStructContainer
                  ElseIf MacroContainer > 0
                    ActiveContainer = MacroContainer
                  Else
                    ActiveContainer = ProcContainer
                  EndIf
                  ;
                  If InElementMap
                    MatchedElementID = CheckAndMarkTokenUsage(*StringStartingAdr, @*TokenPosChar, FileName$, TokenLCase$,
                                                              ActiveEndingPos, ActiveContainer, HeaderParams$,
                                                              UsageTargetType, WithVar$, ModuleContainer,
                                                              UsageParentheses, UsageDotType$, @CandidateIDs)
                  EndIf
                  ;
                  If MatchedElementID = 0 And UsageTargetType = #PBBLabel
                    QueuePendingLabel(PendingLabels(), TokenLCase$, *TokenPosChar, ActiveContainer, ModuleContainer)
                  ElseIf MatchedElementID > 0
                    If UsageTargetType = #PBBModule
                      ElementList()\HasUseModule = #True
                    ElseIf MatchedElementID <> MacroContainer And ElementList()\Type = #PBBMacro
                      ResolveMacroTokens(MatchedElementID, ActiveContainer, FileName$, *StringStartingAdr, @*TokenPosChar, 0, ModuleContainer)
                    EndIf
                  EndIf
                EndIf
              EndIf
            EndIf
          EndIf
        EndIf
      EndIf
      ;
      CompilerIf #PBBReport
        If OPDMod & #OPDOnePassTiming
          ReapTimer(4)
        EndIf
      CompilerEndIf
      ;}
      ;
    ElseIf ParsingState = #Parsing_EndOfLine
      ;
      CompilerIf #PBBReport
        If OPDMod & #OPDOnePassTiming
          StartTimer(2, "Elements creation")
        EndIf
      CompilerEndIf
      ;
      If LineToBeContinued = #False
        ; Statement ended, check whether we exited any active Macro, Procedure, Structure, or Module scopes independently
        If MacroEndingPos And CharToPos(*Char) >= MacroEndingPos
          If MacroContainer And SelectPBBElement(MacroContainer)
            ElementList()\MacroTokens = CurrentMacroTokens\s
            CompilerIf #PBBReport
              If OPDMod & #OPDElementCreation
                OPDReport("--> MacroToken for " + ElementList()\Name + " : " +
                          ReplaceString(ReplaceString(CurrentMacroTokens\s, #TAB$, ", "), #CR$, " / "))
              EndIf
            CompilerEndIf
          EndIf
          *FileEntry\OnePassCDepth = MacroStartCDepth
          MacroStartCDepth     = 0
          MacroEndingPos       = 0
          MacroContainer       = 0
          MacroParams$         = #TAB$
          CurrentMacroTokens\s = ""
        EndIf
        If ProcEndingPos And CharToPos(*Char) >= ProcEndingPos
          ProcEndingPos = 0
          ProcContainer = -1
          ProcParams$   = #TAB$
        EndIf
        If BlockEndingPos And CharToPos(*Char) >= BlockEndingPos
          BlockEndingPos             =  0
          InterfaceOrStructContainer = -1
          EnumContainer              =  0
          ImportContainer            =  0
        EndIf
        If ModuleEndingPos And CharToPos(*Char) >= ModuleEndingPos
          ModuleEndingPos =  0
          ModuleContainer = -1
          IsDeclareModule = #False
        EndIf
      EndIf
      ;
      ; -------------------- Creation of elements ---------------------
      ;
      If TypeOfElement And TypeOfElement <> #PBBNotAnElement
        ; A new element was detected and must now be stored.
        ;{
        ; Append the current physical or ':'-delimited segment to Declaration$ and capture any trailing comment
        AppendDeclarationSegment(#False)
        ;
        If LineToBeContinued = #False
          ; Reached the logical end of the complete statement: finalize the created element(s)
          ActiveContainer = 0
          HeaderParams$   = ""
          Select TypeOfElement
            Case #PBBVariable, #PBBField
              ActiveContainer = CommitPendingVarsAndFields(FieldsOrVariablesToCreate(), FileName$, *StringStartingAdr, *Char,
                                                           EndingPos, Declaration$, CommentStr\s,
                                                           InterfaceOrStructContainer, MacroContainer, ProcContainer,
                                                           ModuleContainer, IsDeclareModule, CurrentMacroTokens)
              ;
            Default
              If TypeOfElement <= #PBBModule Or TypeOfElement = #PBBImport
                ; Block element creation: locate the closing 'EndBlock' keyword (EndDeclareModule for DeclareModule)
                If TypeOfElement = #PBBModule And IsDeclareModule
                  EndExpression$ = "enddeclaremodule"
                Else
                  EndExpression$ = "end" + LCase(PBBTypeNames$(TypeOfElement))
                EndIf
                TempPos = FindInPBPointedCode(*SearchString_LCase, EndExpression$, EndingPos + 1)
                If TempPos = 0
                  ClearList(HeaderDeferredTokens())
                  GoToEOL_AndExtractComment(@*Char)
                  ResetParsingModeFlags
                  Continue
                EndIf
                EndingPos = TempPos + Len(EndExpression$)
                TempPos   = ProcContainer
              ElseIf TypeOfElement = #PBBImported
                TempPos = ImportContainer
              Else
                TempPos = EnumContainer
              EndIf
              ;
              AddPBBElementFromDeclaration(FileName$, TypeOfElement, ElementName$,
                                           StartingCodeLinePos, EndingPos, Declaration$, CommentStr\s,
                                           VariableSpecie$, FieldOrVariableSubSpecie$, HasParentheses,
                                           TempPos, ActiveParams$, ModuleContainer, IsDeclareModule)
              ActiveContainer = ElementList()\FixedID
              HeaderParams$   = ActiveParams$
              If MacroContainer
                CurrentMacroTokens\s + LCase(ElementList()\Name) + #TAB$ + Str(ActiveContainer) + #CR$
              EndIf
              ;
              Select TypeOfElement
                Case #PBBImported
                  If ImportContainer > 0 : LinkElementUsage(ActiveContainer, ImportContainer) : EndIf
                Case #PBBConstant
                  ; Evaluate explicit constant assignment or auto-increment within the active Enumeration block
                  If FindString(Declaration$, "=")
                    ElementList()\Value = ComputeConstantValue(Declaration$, ModuleContainer, FileName$, StartingCodeLinePos)
                    If EnumContainer > 0 : EnumValue$ = ElementList()\Value : EndIf
                  ElseIf EnumContainer > 0
                    If EnumValue$ <> "UnResolved"
                      If IsBinaryEnum
                        If EnumValue$ = "" : EnumValue$ = "1" : Else : EnumValue$ = Str(Val(EnumValue$) * 2) : EndIf
                      Else
                        If EnumValue$ = "" : EnumValue$ = "0" : Else : EnumValue$ = Str(Val(EnumValue$) + 1) : EndIf
                      EndIf
                    EndIf
                    ElementList()\Value = EnumValue$
                  EndIf
                  If EnumContainer > 0 : LinkElementUsage(ActiveContainer, EnumContainer) : EndIf
                Case #PBBMacro
                  MacroContainer       = ElementList()\FixedID
                  MacroStartCDepth     = *FileEntry\OnePassCDepth
                  MacroEndingPos       = EndingPos
                  MacroParams$         = ActiveParams$
                  CurrentMacroTokens\s = ""
                Case #PBBProcedure
                  ProcContainer        = ElementList()\FixedID
                  ProcEndingPos        = EndingPos
                  ProcParams$          = ActiveParams$
                Case #PBBStructure, #PBBInterface
                  InterfaceOrStructContainer = ElementList()\FixedID
                  BlockEndingPos             = EndingPos
                Case #PBBEnumeration
                  EnumContainer  = ElementList()\FixedID
                  BlockEndingPos = EndingPos
                  IsBinaryEnum   = Bool(Left(LCase(Declaration$), 17) = "enumerationbinary")
                  EnumValue$     = FindPreviousEnumerationValue(EnumContainer, FileName$, StartingCodeLinePos)
                Case #PBBImport
                  ImportContainer = ElementList()\FixedID
                  BlockEndingPos  = EndingPos
                Case #PBBModule
                  ModuleContainer = ElementList()\FixedID
                  ModuleEndingPos = EndingPos
              EndSelect
              ;
          EndSelect ; Select TypeOfElement
          ;
          If ListSize(HeaderDeferredTokens()) > 0
            If HeaderParams$ = "" Or HeaderParams$ = #TAB$
              If MacroContainer > 0 : HeaderParams$ = MacroParams$ : Else : HeaderParams$ = ProcParams$ : EndIf
            EndIf
            FlushHeaderDeferredTokens(HeaderDeferredTokens(), PendingLabels(), *StringStartingAdr, FileName$, ActiveContainer,
                                      Bool(MacroContainer > 0) * MacroContainer + Bool(MacroContainer = 0) * ProcContainer,
                                      HeaderParams$, WithVar$, ModuleContainer)
          EndIf
          ;
        EndIf ; If LineToBeContinued = #False
        ;}
      ElseIf LineToBeContinued = #False And ListSize(HeaderDeferredTokens()) > 0
        ClearList(HeaderDeferredTokens())
      EndIf ; If TypeOfElement And TypeOfElement <> #PBBNotAnElement
      ;
      CompilerIf #PBBReport
        If OPDMod & #OPDOnePassTiming
          ReapTimer(2)
        EndIf
      CompilerEndIf
      ;
    Else ; #Parsing_EndOfFile
      Break
    EndIf
  ForEver
  ;
  ; Restore the position of PBB_FileListing()
  IsFileInMemory(FileName$)
  ;
  If *Char = 0 Or *Char\c = 0 Or (CharLimitEnd And *Char > CharLimitEnd)
    ; Replay any forward label references ('?Label', 'Goto', 'Gosub', 'Restore') whose target
    ; label was declared further down in the same file during this single-pass scan:
    If ListSize(PendingLabels()) > 0
      CompilerIf #PBBReport
        If OPDMod & #OPDOnePassTiming
          StartTimer(4, "Elements Usage")
        EndIf
      CompilerEndIf
      ForEach PendingLabels()
        *TokenPosChar = PendingLabels()\TokenPosChar
        CheckAndMarkTokenUsage(*StringStartingAdr, @*TokenPosChar, FileName$,
                               PendingLabels()\TokenLCase, 0, PendingLabels()\ContainerID, "",
                               #PBBLabel, "", PendingLabels()\ModuleContainer, -1)
      Next
      ClearList(PendingLabels())
      CompilerIf #PBBReport
        If OPDMod & #OPDOnePassTiming
          ReapTimer(4)
        EndIf
      CompilerEndIf
    EndIf
    ;
    If ExplorationState <> #ExplorationState_Completed
      If *Char = 0 Or *Char\c = 0
        *FileEntry\OnePassResumePos = Len(*SearchString\s) + 1
      Else
        *FileEntry\OnePassResumePos = CharToPos(*Char)
      EndIf
    EndIf
    ;
    *MemPosChar = 0
    ;
    ProcedureReturn #ExplorationState_Completed
  Else
    ProcedureReturn #ExplorationState_Pending
  EndIf
  ;
EndProcedure
;
Procedure OnePassSetElements(PriorityMode = #WorkInBackGround, OneFileOnly$ = "", *Limits.Point = 0, ResetExploration = #False)
  ;
  ; Role: Orchestrates the single-pass code exploration across all project compilation sequences.
  ;
  Shared ExecuteSearch, RefreshEPanel
  Protected TypeOfElement, CLE_StartTime, InclusionPosInParent
  Protected InclusionParentFile$
  Protected SegmentLimits.Point
  Static MustInitialize = #True, ExplorationStep = 0
  ;
  If ResetExploration
    ; Force a clean initialization state before the next exploration pass begins
    MustInitialize   = #True
    ExplorationStep  = 0
    ExplorationState = #ExplorationState_Pending
    CompilerIf #PBBReport
      If OPDMod & #OPDOnePassTiming
        OPDReport("************* Reset OnePass ****************")
        InitAllTimers()
        StartTimer(1, "Global OnePass")
      EndIf
    CompilerEndIf
    ProcedureReturn ExplorationState
  EndIf
  ;
  If ListSize(PBB_CompilationSequence()) = 0 And ListSize(PBB_FileListing()) = 0
    ExplorationState = #ExplorationState_Completed
    ProcedureReturn ExplorationState
  EndIf
  ;
  CLE_StartTime = ElapsedMilliseconds()
  ;
  If OneFileOnly$
    OnePassSetElements_SegmentExam(OneFileOnly$, #FinishCompletionNow, CLE_StartTime, *Limits)
  ElseIf ExplorationState <> #ExplorationState_Completed
    If MustInitialize
      ExplorationStep = 0
      ForEach PBB_FileListing()
        GetPointedCodeFromFile(PBB_FileListing()\FileName$, #False)
        PBB_FileListing()\OnePassStarted   = #False
        PBB_FileListing()\OnePassResumePos = 0
        PBB_FileListing()\OnePassCDepth    = 0
      Next
      ForEach PBB_CompilationSequence()
        PBB_CompilationSequence()\SegmentIsExplored = #ExplorationState_Pending
      Next
      ;
      ; Purge the master list, map indexes, pointer lookup table, and analysis caches if not already cleared
      If ListSize(ElementList()) > 0
        ClearAllElementsAndData()
      EndIf
      MustInitialize = #False
    EndIf
    ;
    ExplorationState = #ExplorationState_Pending
    ;
    ForEach PBB_CompilationSequence()
      If PBB_CompilationSequence()\SegmentIsExplored = #ExplorationState_Pending
        ; Set segment boundaries, propagate newly discovered parent exclusions, and resume past already-consumed directives:
        SegmentLimits\x = PBB_CompilationSequence()\LocalStartPos
        SegmentLimits\y = PBB_CompilationSequence()\LocalEndPos
        If IsFileInMemory(PBB_CompilationSequence()\FileName$)
          If PBB_FileListing()\ExcludedByCompilerIf
            SegmentLimits\x = SegmentLimits\y + 1
          ElseIf PBB_FileListing()\ParentFileName <> "" And PBB_FileListing()\PosInParent > 0
            InclusionParentFile$ = PBB_FileListing()\ParentFileName
            InclusionPosInParent = PBB_FileListing()\PosInParent
            If IsFileInMemory(InclusionParentFile$) And
               (PBB_FileListing()\ExcludedByCompilerIf Or SkipDeadZoneInCurrentFile(InclusionPosInParent) > InclusionPosInParent)
              ; Parent IncludeFile directive fell inside a newly resolved DeadZone (or excluded parent): mark file dead and refresh UI
              IsFileInMemory(PBB_CompilationSequence()\FileName$)
              PBB_FileListing()\ExcludedByCompilerIf = #True
              ClearList(PBB_FileListing()\DeadZones())
              RefreshEPanel = #True
              If ExecuteSearch = #PBBNoSearch : ExecuteSearch = #PBBRefreshSearch : EndIf
              SegmentLimits\x = SegmentLimits\y + 1
            Else
              IsFileInMemory(PBB_CompilationSequence()\FileName$)
            EndIf
          EndIf
          If SegmentLimits\x <= SegmentLimits\y
            If PBB_FileListing()\OnePassStarted = #False
              ; Replace Stage-1 provisional DeadZones() just-in-time as OnePass begins indexing this active file
              ClearList(PBB_FileListing()\DeadZones())
              PBB_FileListing()\OnePassStarted = #True
            ElseIf PBB_FileListing()\OnePassResumePos > SegmentLimits\x
              SegmentLimits\x = PBB_FileListing()\OnePassResumePos
            EndIf
          EndIf
        EndIf
        If SegmentLimits\x > SegmentLimits\y
          PBB_CompilationSequence()\SegmentIsExplored = #ExplorationState_Completed
          ExplorationStep + 1
          Continue
        EndIf
        ;
        SetGadgetState(GPBBGadgets\EProgressBar, 100 * ExplorationStep / ListSize(PBB_CompilationSequence()))
        ;
        ; Execute the exploration strictly within the segment boundaries
        PBB_CompilationSequence()\SegmentIsExplored = OnePassSetElements_SegmentExam(PBB_CompilationSequence()\FileName$,
                                                                                     PriorityMode, CLE_StartTime, @SegmentLimits)
        If PBB_CompilationSequence()\SegmentIsExplored = #ExplorationState_Completed
          ExplorationStep + 1
        Else
          ; The process has hit the time limit and needs to yield
          ProcedureReturn ExplorationState
        EndIf
      EndIf
    Next
    ;
    MustInitialize = #True
  EndIf
  ;
  If ListSize(ElementList()) > 0
    ; Clean orphan "Declare" elements and purge invalidated elements (StartingPos = 0) deleted during a partial update:
    ForEach ElementList()
      If ElementList()\StartingPos = 0 Or ElementList()\Type = #PBBDeclare
        ResetOneElementLineage(#True)
        DeletePBBElement()
      ElseIf ElementList()\Type = #PBBVariable And ElementList()\VarDotType_LCase = "" And Right(ElementList()\PurifiedName, 1) <> "$"
        ; Do not append a default ".i" suffix to string variables, arrays, lists, or maps
        ElementList()\VarDotType_LCase = "i"
        ElementList()\Name + ".i"
        ElementList()\VarAddedDotType = #True
      EndIf
    Next
    ;
    UsagePropagation()
    ;
    For TypeOfElement = 0 To (#EndEnumPBBElementTypes - 1)
      ListCompletion(TypeOfElement) = #ListCompletion_UnPrinted
    Next
  EndIf
  ;
  ExplorationState = #ExplorationState_Completed
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDOnePassTiming
      ReapTimer(1)
      DebugAllTimers()
    EndIf
  CompilerEndIf
  ;
  ProcedureReturn ExplorationState
EndProcedure
;
;
; **********************************************************************************
;
;-                        Incremental partial code updates
;
; **********************************************************************************
;
; ----------------------------------------------------------------------------------
;               ---- Helpers for partial updates of element lists ----
; ----------------------------------------------------------------------------------
;
Procedure.s AddDependentConstantsToUpdateList(UsageParents$, MustBeUpdated$)
  ;
  ; Role: Scans a slash-delimited 'UsageParents$' list for any dependent '#PBBConstant' elements
  ; and appends either their enclosing '#PBBEnumeration' ID (when inside an enumeration)
  ; or their own FixedID (when standalone) to 'MustBeUpdated$'.
  ;
  Protected ParentPos = 1
  Protected ParentIDStr$
  Protected *Elem.Elements
  ;
  If UsageParents$ <> ""
    Repeat
      ParentPos + 1
      ParentIDStr$= StringField(UsageParents$, ParentPos, "/")
      If ParentIDStr$ = "" : Break : EndIf
      If ParentIDStr$ <> "-1"
        *Elem = ElementFixedIndexes(Val(ParentIDStr$))
        If *Elem And *Elem\Type = #PBBConstant
          If *Elem\LocalVarContainer > 0
            MustBeUpdated$= AddIndexToList(*Elem\LocalVarContainer, MustBeUpdated$)
          Else
            MustBeUpdated$= AddIndexToList(*Elem\FixedID, MustBeUpdated$)
          EndIf
        EndIf
      EndIf
    ForEver
  EndIf
  ;
  ProcedureReturn MustBeUpdated$
EndProcedure
;
Macro DeleteOneElementForPartialUpdate()
  ;
  ; Role: Invalidates or demotes the current element during a partial code update without
  ; destroying its FixedID or incoming UsageParents from other unmodified procedures.
  ; - If a secondary declaration in '\Continued' exists outside [OldZoneStart, OldZoneEnd],
  ;   promotes it to become the main declaration and schedules it for re-examination.
  ; - If the element is a constant, schedules any dependent constants or enumerations in MustBeUpdated$.
  ; - Otherwise, marks the element as pending refresh (StartingPos = 0, EndingPos = $7FFFFFFF)
  ;   so AddPBBElementFromDeclaration() can update it in place if it still exists in the modified code,
  ;   or OnePassSetElements() can purge it afterward if it was genuinely deleted.
  ;
  CEDIsModified = #False
  ;
  If ElementList()\Type = #PBBConstant And ElementList()\UsageParents <> ""
    MustBeUpdated$ = AddDependentConstantsToUpdateList(ElementList()\UsageParents, MustBeUpdated$)
  EndIf
  ;
  If ElementList()\Continued
    CED_StringToList(ElementList()\Continued, CEDList())
    ForEach CEDList()
      If CEDList()\Filename <> FileName$ Or CEDList()\EndingPos < OldZoneStart Or CEDList()\StartingPos > OldZoneEnd
        ; Promote surviving declaration outside the modified zone and schedule it for re-examination
        MustBeUpdated$ = AddIndexToList(ElementList()\FixedID, MustBeUpdated$)
        CED_ToMainDeclaration(CEDList())
        CEDIsModified = #True
        Break
      EndIf
    Next
  EndIf
  ;
  ; Reset outgoing child links (UsageChildren) while preserving incoming UsageParents
  ResetOneElementLineage(#False)
  ;
  If CEDIsModified
    ElementList()\Continued = CED_ListToString(CEDList())
  Else
    ; Invalidate coordinates in place so AddPBBElementFromDeclaration() reuses this FixedID if re-encountered
    If ElementList()\Parameters <> ""
      RemoveParametersFromMapIndex(ElementList()\Parameters, ElementList()\FixedID)
      ElementList()\Parameters = ""
    EndIf
    ElementList()\StartingPos = 0
    ElementList()\EndingPos   = $7FFFFFFF
  EndIf
  ;
EndMacro
;
; Macros to eliminate redundant boundary checks and re-parsing sequence calls
;
Macro IsOutsideModifiedZone(mFileName, mStart, mEnd)
  ;
  ; Role: Evaluates whether a code element resides outside the bounds of the modified zone.
  ;
  Bool(mStart > 0 And (mFileName <> FileName$ Or BoundariesVersusLimits(mStart, mEnd, OldZoneStart, NewZoneEnd) <> #InsideLimits))
EndMacro
;
Macro ReParseElementIfOutside(mFileName, mStart, mEnd, mRestoreID)
  ;
  ; Role: Triggers a localized re-analysis of elements located outside the active edit range.
  ;
  If IsOutsideModifiedZone(mFileName, mStart, mEnd)
    Limits\x = mStart
    Limits\y = mEnd
    OnePassSetElements_SegmentExam(mFileName, #FinishCompletionNow, ElapsedMilliseconds(), @Limits)
    If mRestoreID : SelectPBBElement(mRestoreID) : EndIf
  EndIf
EndMacro
;
; ----------------------------------------------------------------------------------
;                  ---- Master partial updates procedures ----
; ----------------------------------------------------------------------------------
;
Procedure PartialUpdateOfCode(*OldContent.String, *NewContent.String, FileName$)
  ;
  ; Role: Isolates the modified section between '*OldContent' and '*NewContent' in 'FileName$',
  ; computes the exact encompassing element envelope in '*OldContent' coordinates, shifts downstream
  ; positions by 'POffset', invalidates affected elements in place (preserving FixedID and UsageParents),
  ; and re-parses only the affected code range.
  ;
  ; Variables to determine which part of the code has been modified:
  Protected ModStartPos, EndExamOld, EndExamNew, MatchingEndChars
  Protected ModOldEndPos, ModNewEndPos, OldZoneStart, OldZoneEnd, NewZoneEnd
  Protected *OldPtr.Character, *Char.Character
  Protected CharSize = SizeOf(Character)
  ;
  Shared PBB_CompilationSequence(), PBB_FileListing(), PBB_BinaryFileListing()
  ;
  ; Variables for updating elements based on the detected change:
  Protected MustBeUpdated$, ClearedOOE$, FilePrefix$, ZoneExpanded, IsConcerned, ctType, CatchupModID, StartSeqElem
  Protected *Elem.Elements
  Protected pc, mpc, pcd, PosInOutOfElement
  Protected POffset, CEDIsModified
  Protected Limits.Point, MBU_Index, MBU_FixedIndex
  Protected ModifiedSlice.String
  Protected FIPS.FIPSPosStruct
  Protected EncompassingData.EEDStruct, EncompassingMod.EEDStruct
  Protected NewList CEDList.CED_Structure()
  ;
  If ExplorationState <> #ExplorationState_Completed
    ; Exploration is still in progress: restart full exploration from the beginning
    ExplorationState = #ExplorationState_MustRestart
    ProcedureReturn #False
  EndIf
  ;
  ; Determine which part of the file has been modified:
  EndExamOld = FindIDENotations(*OldContent) - 1
  EndExamNew = FindIDENotations(*NewContent) - 1
  ;
  ModStartPos = FindStartOfDifference(*OldContent, *NewContent, min(EndExamOld, EndExamNew))
  If ModStartPos = 0 And EndExamOld <> EndExamNew
    ModStartPos = min(EndExamOld, EndExamNew) + 1
  EndIf
  If ModStartPos
    ;
    ; Determine the end of the modified section:
    ModOldEndPos = EndExamOld
    ModNewEndPos = EndExamNew
    ;
    ; Search for the end of differences between the two buffers (backwards from the end):
    MatchingEndChars = FindEndOfDifference(*OldContent, *NewContent, ModStartPos, EndExamOld, EndExamNew)
    If MatchingEndChars > 0
      ModOldEndPos = EndExamOld - MatchingEndChars + 1
      ModNewEndPos = EndExamNew - MatchingEndChars + 1
    EndIf
    ;
    ; Expand the modified section to cover at least one full physical line in *OldContent:
    *OldPtr = @*OldContent\s
    ;
    *Char = *OldPtr + (ModStartPos - 2) * CharSize
    While ModStartPos > 1 And *Char\c <> #CR : ModStartPos - 1 : *Char - CharSize : Wend
    ;
    *Char = *OldPtr + (ModOldEndPos - 1) * CharSize
    While ModOldEndPos < EndExamOld And *Char\c <> #CR : ModOldEndPos + 1 : ModNewEndPos + 1 : *Char + CharSize : Wend
    ; Position right before the carriage return when ending on a line break:
    If ModOldEndPos > 0 And PeekC(*OldPtr + (ModOldEndPos - 1) * CharSize) = #CR : ModOldEndPos - 1 : EndIf
    If ModNewEndPos > 0 And PeekC(@*NewContent\s + (ModNewEndPos - 1) * CharSize) = #CR : ModNewEndPos - 1 : EndIf
    ;
    If ModOldEndPos >= ModStartPos Or ModNewEndPos >= ModStartPos
      ; Escalate to full re-exploration when include, module, macro, or compiler directives were added, removed, or edited:
      ModifiedSlice\s = ""
      If ModOldEndPos >= ModStartPos
        ModifiedSlice\s = LCase(FastMid(*OldContent\s, ModStartPos, ModOldEndPos - ModStartPos + 1))
      EndIf
      If ModNewEndPos >= ModStartPos
        ModifiedSlice\s + #CR$ + LCase(FastMid(*NewContent\s, ModStartPos, ModNewEndPos - ModStartPos + 1))
      EndIf
      If SliceTouchesGuardSymbol(@ModifiedSlice)
        ExplorationState = #ExplorationState_MustRestart
        ProcedureReturn #True
      EndIf
      ;
      OldZoneStart = ModStartPos
      OldZoneEnd   = ModOldEndPos
      POffset      = ModNewEndPos - ModOldEndPos
      FilePrefix$  = #CR$ + FileName$ + #TAB$
      ;
      ; -----------------------------------------------------------------------------------
      ; PHASE 1: Compute the full envelope [OldZoneStart, OldZoneEnd] in *OldContent coordinates.
      ; Repeats until no further overlapping container (Procedure, Structure, Enumeration, etc.)
      ; expands the boundaries, regardless of element order in ElementList().
      ; -----------------------------------------------------------------------------------
      Repeat
        ZoneExpanded = #False
        ForEach ElementList()
          If ElementList()\FileName = FileName$ And ElementList()\StartingPos > 0
            ; Do not expand to an entire Module when the edit is strictly inside the module body
            If ElementList()\Type = #PBBModule And ElementList()\StartingPos < OldZoneStart And ElementList()\EndingPos > OldZoneEnd
              Continue
            EndIf
            If BoundariesVersusLimits(ElementList()\StartingPos, ElementList()\EndingPos, OldZoneStart, OldZoneEnd) = #InsideLimits
              If ElementList()\Type = #PBBMacro
                ; Editing inside an existing Macro invalidates inlined MacroTokens across callers: trigger full restart
                ExplorationState = #ExplorationState_MustRestart
                ProcedureReturn #True
              EndIf
              If ElementList()\StartingPos < OldZoneStart
                OldZoneStart = ElementList()\StartingPos
                ZoneExpanded = #True
              EndIf
              If ElementList()\EndingPos > OldZoneEnd
                OldZoneEnd   = ElementList()\EndingPos
                ZoneExpanded = #True
              EndIf
            EndIf
          EndIf
          ;
          ; Expand the modified zone envelope to any overlapping secondary declaration in '\Continued':
          If ElementList()\Continued <> "" And FindString(#CR$ + ElementList()\Continued, FilePrefix$)
            CED_StringToList(ElementList()\Continued, CEDList())
            ForEach CEDList()
              If CEDList()\Filename = FileName$ And
                 BoundariesVersusLimits(CEDList()\StartingPos, CEDList()\EndingPos, OldZoneStart, OldZoneEnd) = #InsideLimits
                If CEDList()\StartingPos < OldZoneStart
                  OldZoneStart = CEDList()\StartingPos
                  ZoneExpanded = #True
                EndIf
                If CEDList()\EndingPos > OldZoneEnd
                  OldZoneEnd   = CEDList()\EndingPos
                  ZoneExpanded = #True
                EndIf
              EndIf
            Next
          EndIf
        Next
      Until ZoneExpanded = #False
      ;
      NewZoneEnd = OldZoneEnd + POffset
      ;
      ; -----------------------------------------------------------------------------------
      ; PHASE 2: Update OOE_UsagePos, Continued records, and ElementList() against [OldZoneStart, OldZoneEnd]
      ; -----------------------------------------------------------------------------------
      ForEach ElementList()
        ;
        ; 1. Adjust usage references in out-of-element code (\OOE_UsagePos):
        pc = 0
        Repeat
          pc = FindString(ElementList()\OOE_UsagePos, FileName$ + ",", pc + 1)
          If pc
            mpc = pc
            pc + Len(FileName$ + ",")
            pcd = FindString(ElementList()\OOE_UsagePos, #CR$, pc)
            If pcd = 0 : pcd = Len(ElementList()\OOE_UsagePos) + 1 : EndIf
            PosInOutOfElement = Val(FastMid(ElementList()\OOE_UsagePos, pc, pcd - pc))
            If PosInOutOfElement > OldZoneEnd
              ; Strictly after the modified zone: shift character position by POffset
              ElementList()\OOE_UsagePos = FastLeft(ElementList()\OOE_UsagePos, pc - 1) + Str(PosInOutOfElement + POffset) +
                                           FastMid(ElementList()\OOE_UsagePos, pcd)
            ElseIf PosInOutOfElement >= OldZoneStart
              ; Inside the modified zone: remove reference (will be re-evaluated during re-scan)
              ElementList()\OOE_UsagePos = FastLeft(ElementList()\OOE_UsagePos, mpc - 1) + FastMid(ElementList()\OOE_UsagePos, pcd + 1)
              If ElementList()\OOE_UsagePos = ""
                ClearedOOE$                = AddIndexToList(ElementList()\FixedID, ClearedOOE$)
                ElementList()\UsageParents = RemoveIndexFromList(-1, ElementList()\UsageParents)
              EndIf
              pc = mpc - 1
            EndIf
          EndIf
        Until pc = 0
        ;
        ; 2. Adjust 'Continued' positions for enumerations, variables, procedures, or modules:
        CEDIsModified = #False
        If ElementList()\Continued <> "" And FindString(#CR$ + ElementList()\Continued, FilePrefix$)
          CED_StringToList(ElementList()\Continued, CEDList())
          ForEach CEDList()
            If CEDList()\Filename = FileName$
              If CEDList()\EndingPos < OldZoneStart
                ; Entirely before the modified zone: keep intact
              ElseIf CEDList()\StartingPos > OldZoneEnd
                ; Entirely after the modified zone: shift positions by POffset
                CEDList()\StartingPos + POffset
                CEDList()\EndingPos   + POffset
                CEDIsModified = #True
              ElseIf ElementList()\Type = #PBBModule And CEDList()\StartingPos < OldZoneStart And CEDList()\EndingPos > OldZoneEnd
                ; The modified zone is strictly inside this secondary Module/DeclareModule block: shift its end boundary
                CEDList()\EndingPos + POffset
                CEDIsModified = #True
              Else
                ; Inside the modified zone: remove segment (will be re-created if still present)
                If ElementList()\Type = #PBBEnumeration
                  MustBeUpdated$ = AddIndexToList(ElementList()\FixedID, MustBeUpdated$)
                EndIf
                DeleteElement(CEDList())
                CEDIsModified = #True
              EndIf
            EndIf
          Next
          ;
          If CEDIsModified
            ElementList()\Continued = CED_ListToString(CEDList())
          EndIf
        EndIf
        ;
        ; 3. Shift or invalidate the primary element declaration if located in FileName$:
        If ElementList()\FileName = FileName$
          IsConcerned = BoundariesVersusLimits(ElementList()\StartingPos, ElementList()\EndingPos, OldZoneStart, OldZoneEnd)
          If IsConcerned = #AfterLimits
            ; Entirely after the modified zone: shift positions by POffset
            ElementList()\StartingPos + POffset
            ElementList()\EndingPos   + POffset
          ElseIf IsConcerned = #InsideLimits
            If ElementList()\Type = #PBBModule And ElementList()\StartingPos < OldZoneStart And ElementList()\EndingPos > OldZoneEnd
              ; The modified zone is strictly inside the module body: only adjust the module's ending position
              ElementList()\EndingPos + POffset
            Else
              ; Invalidate in place (or promote a surviving '\Continued' entry outside the zone)
              DeleteOneElementForPartialUpdate()
            EndIf
          EndIf
        EndIf
        ;
      Next
      ;
      ; -----------------------------------------------------------------------------------
      ; PHASE 3: Update inclusion offsets, re-parse the modified zone, then update dependents
      ; -----------------------------------------------------------------------------------
      If POffset <> 0
        PushListPosition(PBB_FileListing())
        ; Shift downstream dead zones in FileName$ and inclusion offsets in child source/binary files
        ForEach PBB_FileListing()
          If PBB_FileListing()\FileName$ = FileName$
            ForEach PBB_FileListing()\DeadZones()
              If PBB_FileListing()\DeadZones()\StartingPos > ModOldEndPos
                PBB_FileListing()\DeadZones()\StartingPos + POffset
              EndIf
              If PBB_FileListing()\DeadZones()\EndingPos > ModOldEndPos
                PBB_FileListing()\DeadZones()\EndingPos + POffset
              EndIf
            Next
          ElseIf PBB_FileListing()\ParentFileName = FileName$ And PBB_FileListing()\PosInParent > ModOldEndPos
            PBB_FileListing()\PosInParent + POffset
          EndIf
        Next
        ForEach PBB_BinaryFileListing()
          If PBB_BinaryFileListing()\ParentFileName = FileName$ And PBB_BinaryFileListing()\PosInParent > ModOldEndPos
            PBB_BinaryFileListing()\PosInParent + POffset
          EndIf
        Next
        UpdateVirtualPositionsAndSequence(PBBFicPrincipalPB$, #True)
        PopListPosition(PBB_FileListing())
      EndIf
      ;
      ResetPBBAnalysisCaches()
      ;
      ; 1. Re-examine the modified zone in *NewContent (unless the edit occurred strictly inside an inactive DeadZone):
      If NewZoneEnd < OldZoneStart
        NewZoneEnd = OldZoneStart
      EndIf
      If OldZoneStart > EndExamNew And EndExamNew > 1
        OldZoneStart = EndExamNew - 1
        NewZoneEnd   = EndExamNew
      EndIf
      If IsPosInDeadZone(FileName$, OldZoneStart) = #False
        Limits\x = OldZoneStart
        Limits\y = NewZoneEnd
        OnePassSetElements_SegmentExam(FileName$, #FinishCompletionNow, ElapsedMilliseconds(), @Limits)
        NewZoneEnd = Limits\y
      EndIf
      ;
      ; If a deleted/renamed element had callers and still shares a homonym in ElementMapIndex(),
      ; trigger a full restart so those callers rebind to the surviving homonym:
      ForEach ElementList()
        If ElementList()\StartingPos = 0 And (ElementList()\UsageParents <> "" Or ElementList()\OOE_UsagePos <> "")
          If FindMapElement(ElementMapIndex(), ElementList()\PurifiedName) And
             FindString(ElementMapIndex(), ",", FindString(ElementMapIndex(), ",") + 1)
            ExplorationState = #ExplorationState_MustRestart
            ProcedureReturn #True
          EndIf
        EndIf
      Next
      ;
      ; 2. Locate compilation sequence anchor at NewZoneEnd for downstream catch-up scans:
      ForEach PBB_CompilationSequence()
        If PBB_CompilationSequence()\FileName$ = FileName$ And
           NewZoneEnd >= PBB_CompilationSequence()\LocalStartPos And
           NewZoneEnd <= PBB_CompilationSequence()\LocalEndPos
          StartSeqElem = @PBB_CompilationSequence()
          Break
        EndIf
      Next
      ;
      ; 3. Catch-up pass: locate downstream usages of unlinked edited elements (or elements whose
      ;    OOE_UsagePos was cleared) and delegate re-parsing to MustBeUpdated$ / OnePassSetElements_SegmentExam():
      ForEach ElementList()
        If ElementList()\StartingPos > 0 And
           Not (ElementList()\Type = #PBBVariable And
                (ElementList()\VariableSpecie = "Protected" Or ElementList()\VariableSpecie = "Static")) And
           ((ElementList()\FileName = FileName$ And ElementList()\StartingPos >= OldZoneStart And
             ElementList()\StartingPos <= NewZoneEnd And ElementList()\OOE_UsagePos = "" And ElementList()\UsageParents = "") Or
            (ElementList()\OOE_UsagePos = "" And ClearedOOE$ <> "" And
             FindString(ClearedOOE$, "/" + Str(ElementList()\FixedID) + "/")))
          ;
          ClearStructure(@FIPS, FIPSPosStruct)
          If ElementList()\Type = #PBBLabel Or StartSeqElem = 0
            ; Labels can be referenced upstream in the same file; loose project files are scanned locally
            FIPS\ScanMode  = #FIPSMode_LocalOnly
            FIPS\FileName$ = FileName$
            If ElementList()\Type = #PBBLabel
              FIPS\PosInCode = 1
            Else
              FIPS\PosInCode = NewZoneEnd
            EndIf
            FIPS\EndLimit = EndExamNew
          Else
            ; Resume global chronological sequence scan strictly from NewZoneEnd onward
            FIPS\ScanMode          = #FIPSMode_Global
            FIPS\FIPSContinue      = #True
            FIPS\FileName$         = FileName$
            FIPS\PosInCode         = NewZoneEnd - 1
            FIPS\CurrentSeqElement = StartSeqElem
          EndIf
          ;
          FilePrefix$ = ElementList()\PurifiedName
          PushListPosition(ElementList())
          While FindInWholeProjectSequences(FilePrefix$, @FIPS, #CheckAll | #CheckOneKeyWordOnly)
            FIPS\FIPSContinue = #True
            If FIPS\FileName$ = FileName$ And FIPS\PosInCode >= OldZoneStart And FIPS\PosInCode <= NewZoneEnd
              Continue
            EndIf
            ;
            ; Check if the occurrence resides inside an encompassing block container (excluding #PBBModule)
            MBU_FixedIndex = 0
            For ctType = #PBBProcedure To #PBBImport
              If ctType <> #PBBModule And ctType <> #PBBLabel And ctType <> #PBBConstant And ctType <> #PBBVariable
                GetEncompassingLimitsAndName(ctType, FIPS\FileName$, FIPS\PosInCode,
                                             FIPS\SearchString, FIPS\SearchString_LCase, FIPS\SegmentEnd, @EncompassingData)
                If EncompassingData\EndingPos > 0
                  If ctType = #PBBMacro
                    PopListPosition(ElementList())
                    ExplorationState = #ExplorationState_MustRestart
                    ProcedureReturn #True
                  EndIf
                  CatchupModID = -1
                  If GetEncompassingLimitsAndName(#PBBModule, FIPS\FileName$, FIPS\PosInCode,
                                                  FIPS\SearchString, FIPS\SearchString_LCase, FIPS\SegmentEnd, @EncompassingMod) <> ""
                    CatchupModID = GetModuleID(EncompassingMod\ElementName$)
                  EndIf
                  If CatchupModID <> 0
                    MBU_FixedIndex = FindElementInList(PurifyElementName(EncompassingData\ElementName$), ctType, 0, "", 0, -1, "", 0, CatchupModID)
                  EndIf
                  If MBU_FixedIndex > 0
                    MustBeUpdated$ = AddIndexToList(MBU_FixedIndex, MustBeUpdated$)
                    FIPS\PosInCode = EncompassingData\EndingPos
                  EndIf
                  Break
                EndIf
              EndIf
            Next
            ;
            ; Out-of-container occurrence: re-parse the enclosing 'With' block or complete logical statement
            If MBU_FixedIndex = 0
              GetEncompassingLimitsAndName(#PBBWith, FIPS\FileName$, FIPS\PosInCode,
                                           FIPS\SearchString, FIPS\SearchString_LCase, FIPS\SegmentEnd, @EncompassingData)
              If EncompassingData\EndingPos > 0
                Limits\x = EncompassingData\StartingPos
                Limits\y = EncompassingData\EndingPos
              Else
                pc = FIPS\PosInCode
                FastFindPrecReturn(FIPS\SearchString\s, pc)
                While pc > 1
                  mpc = IsLineContinuedAtEnd(FIPS\SearchString, pc)
                  If mpc : pc = mpc : Else : Break : EndIf
                Wend
                If PeekC(@FIPS\SearchString\s + (pc - 1) * CharSize) = #CR : pc + 1 : EndIf
                Limits\x = pc
                Limits\y = FIPS\PosInCode
              EndIf
              OnePassSetElements_SegmentExam(FIPS\FileName$, #FinishCompletionNow, ElapsedMilliseconds(), @Limits)
              FIPS\PosInCode = Limits\y
            EndIf
          Wend
          PopListPosition(ElementList())
        EndIf
      Next
      ;
      ; 4. Re-examine queued containers, promoted continuations, multi-part enumerations, and dependent constants:
      If MustBeUpdated$
        MBU_Index = 1 ; Skip the leading "/" maintained by AddIndexToList()
        Repeat
          MBU_Index + 1
          MBU_FixedIndex = Val(StringField(MustBeUpdated$, MBU_Index, "/"))
          If MBU_FixedIndex And SelectPBBElement(MBU_FixedIndex)
            ;
            If ElementList()\Type = #PBBEnumeration
              ; Re-parse primary block and any promoted \Continued secondary blocks located outside the modified zone
              ReParseElementIfOutside(ElementList()\FileName, ElementList()\StartingPos, ElementList()\EndingPos, MBU_FixedIndex)
              If ElementList()\Continued <> ""
                CED_StringToList(ElementList()\Continued, CEDList())
                ForEach CEDList()
                  ReParseElementIfOutside(CEDList()\Filename, CEDList()\StartingPos, CEDList()\EndingPos, MBU_FixedIndex)
                Next
              EndIf
              ;
              ; Queue any external constants that depend on the constants of this enumeration
              pc = 1
              Repeat
                pc + 1
                mpc = Val(StringField(ElementList()\UsageParents, pc, "/"))
                If mpc = 0 : Break : EndIf
                If mpc > 0
                  *Elem = ElementFixedIndexes(mpc)
                  If *Elem And *Elem\Type = #PBBConstant And *Elem\LocalVarContainer = MBU_FixedIndex
                    MustBeUpdated$ = AddDependentConstantsToUpdateList(*Elem\UsageParents, MustBeUpdated$)
                  EndIf
                EndIf
              ForEver
              ;
            ElseIf ElementList()\Type = #PBBConstant
              ; Propagate to any transitive dependent constants and re-evaluate in memory if outside the modified zone
              MustBeUpdated$ = AddDependentConstantsToUpdateList(ElementList()\UsageParents, MustBeUpdated$)
              If IsOutsideModifiedZone(ElementList()\FileName, ElementList()\StartingPos, ElementList()\EndingPos)
                ElementList()\Value = ComputeConstantValue(ElementList()\Declaration, ElementList()\EnclosingModuleID,
                                                           ElementList()\FileName, ElementList()\StartingPos)
              EndIf
              ;
            Else
              ; For local variables, re-parse their enclosing container (Procedure/Macro) instead of the variable itself
              If ElementList()\LocalVarContainer > 0 And ElementList()\Type = #PBBVariable
                SelectPBBElement(ElementList()\LocalVarContainer)
                MBU_FixedIndex = ElementList()\FixedID
              EndIf
              If IsOutsideModifiedZone(ElementList()\FileName, ElementList()\StartingPos, ElementList()\EndingPos)
                ResetOneElementLineage(#False)
                ReParseElementIfOutside(ElementList()\FileName, ElementList()\StartingPos, ElementList()\EndingPos, MBU_FixedIndex)
              EndIf
            EndIf
          EndIf
        Until MBU_FixedIndex = 0
      EndIf
      ;
      ; 5. Purge any remaining invalidated elements (StartingPos = 0) and run UsagePropagation():
      OnePassSetElements(#FinishCompletionNow)
      ;
      ProcedureReturn #True
    EndIf
  EndIf
  ;
  ProcedureReturn #False
EndProcedure
;
Procedure UpdateOneFileInMemory(FileName$, TestOnly = #False)
  ;
  ; Role: Checks if the file 'FileName$' has been modified on disk since its last read.
  ; If so, reloads both normal and lowercase in-memory code buffers in place (preserving DeadZones(),
  ; ExcludedByCompilerIf, and inclusion tree metadata) and delegates updates to 'PartialUpdateOfCode()'.
  ;
  Shared ExecuteSearch, RefreshEPanel
  Protected OldContent.String
  ;
  If IsFileInMemory(FileName$)
    If GetFileDate(FileName$, #PB_Date_Modified) = PBB_FileListing()\FileLastMod
      ProcedureReturn #FileWasUpToDate
    ElseIf TestOnly
      ProcedureReturn #FileMustBeUpdated
    Else
      OldContent\s = PBB_FileListing()\FileContentInList\s
      If ReadCodeFromFile(FileName$, PBB_FileListing()\FileContentInList) = #False
        DeleteElement(PBB_FileListing(), 1)
        ProcedureReturn #FileDoesntExist
      EndIf
      PBB_FileListing()\FileLastMod              = GetFileDate(FileName$, #PB_Date_Modified)
      PBB_FileListing()\FileLCaseContentInList\s = LCase(PBB_FileListing()\FileContentInList\s)
      EraseObsoleteElementsContents(FileName$)
      ;
      If PartialUpdateOfCode(OldContent, PBB_FileListing()\FileContentInList, FileName$)
        ; Refresh panels immediately only if partial update completed in place
        ; (when ExplorationState = #ExplorationState_MustRestart, ManagePBBBackgroundTasks() rebuilds everything):
        If ExplorationState = #ExplorationState_Completed
          If ExecuteSearch = #PBBNoSearch : ExecuteSearch = #PBBRefreshSearch : EndIf
          RefreshEPanel = #True
          ResetPBBAnalysisCaches()
        EndIf
        ProcedureReturn #FileIsUpdated
      EndIf
      ProcedureReturn #FileWasUpToDate
    EndIf
  EndIf
  ProcedureReturn #FileDoesntExist
EndProcedure
;
;
; **********************************************************************************
;
;-                  Procedure variable & dependency inventory
;
; **********************************************************************************
;
Define ListOfMacro$
;
Structure VariableListings
  ProcedureName.s
  CompleteList.s
  Unused.s
  Unprotected.s
  CalledProcedures.s
  MissingElements.s
EndStructure
;
Define VList.VariableListings
;
Structure MacroExpandEntry
  MacroID.i
  SkipUntilPos.i
EndStructure
;
Macro CollectCalledMacros(ChildrenStr, BlankLocalDefs)
  ;
  ; Role: Appends valid upstream or local #PBBMacro IDs from a slash-delimited UsageChildren list into CalledMacros(),
  ;       blanking out local 'Macro ... EndMacro' definitions in-place during the initial Procedure pass.
  ;       Copies 'ChildrenStr' upfront so SelectPBBElement() inside the loop never mutates the source field.
  ;
  ChildrenList$ = ChildrenStr
  If ChildrenList$ <> ""
    ChildIdx = 1
    Repeat
      ChildIdx + 1
      ChildID = Val(StringField(ChildrenList$, ChildIdx, "/"))
      If ChildID = 0 : Break : EndIf
      If ChildID <> ElementFixedID And FindString(KnownMacros$, "/" + Str(ChildID) + "/") = 0 And
         SelectPBBElement(ChildID) And ElementList()\Type = #PBBMacro
        If ElementList()\FileName = ProcFile$ And ElementList()\StartingPos >= ProcStartPos And ElementList()\EndingPos <= ProcEndPos
          CompilerIf BlankLocalDefs
            If IsProc
              LocalPos = ElementList()\StartingPos - ProcStartPos + 1
              LocalLen = ElementList()\EndingPos - ElementList()\StartingPos + 1
              *ProcOrMacroCode\s = FastLeft(*ProcOrMacroCode\s, LocalPos - 1) + Space(LocalLen) + FastMid(*ProcOrMacroCode\s, LocalPos + LocalLen)
            EndIf
          CompilerEndIf
          AddElement(CalledMacros())
          CalledMacros()\MacroID = ChildID
          KnownMacros$ + Str(ChildID) + "/"
        ElseIf Not IsPositionBefore(ProcFile$, ProcStartPos, ElementList()\FileName, ElementList()\StartingPos)
          AddElement(CalledMacros())
          CalledMacros()\MacroID = ChildID
          KnownMacros$ + Str(ChildID) + "/"
        EndIf
      EndIf
    ForEver
  EndIf
EndMacro
;
Procedure   HasMacroArguments(MacroDeclaration$)
  ;
  ; Role: Examines a macro declaration header to determine if it requires parameters.
  ; Output: Returns 1 if mandatory parameters exist, -1 if the first parameter is optional,
  ;         or 0 if the macro has no arguments (or empty parentheses).
  ;
  Protected DeclBuf.String, ListArg$
  Protected ParenPos = FindString(MacroDeclaration$, "(")
  ;
  If ParenPos
    DeclBuf\s = MacroDeclaration$
    ListArg$  = Trim(ExtractBalancedCallParams(@DeclBuf, ParenPos, 0), #TAB$)
    If ListArg$ <> ""
      If FindString(StringField(ListArg$, 1, #TAB$), "=")
        ProcedureReturn -1
      EndIf
      ProcedureReturn 1
    EndIf
  EndIf
  ProcedureReturn 0
EndProcedure
;
Procedure   ReplaceMacroCallsByMacroCodes(ElementFixedID, *ProcOrMacroCode.String)
  ;
  ; Role: Extracts the source code of the Procedure or Macro identified by 'ElementFixedID'
  ;       and iteratively expands all referenced project macros (including nested macro arguments,
  ;       sub-macros, single-line ':' macros, empty macros, and '#' token concatenations) in a
  ;       single non-recursive loop using per-macro 'SkipUntilPos' horizons to prevent self-expansion.
  ;
  Shared ListOfMacro$
  ;
  Protected *SearchString.String, ProcOrMacroCode_LCase.String, MacroCode.String, MacroCode_LCase.String
  Protected *Char.Character, *Start.Character, *SegStart.Character, *ActiveMacro.MacroExpandEntry
  Protected ProcFile$, KnownMacros$ = "/", ChildrenList$, ElementName$, ElementNameLCase$, MacroDecl$
  Protected CallingParamList$, MacroParamList$, RawParam$, SubToken$, ReplacedSeg$, NewBody$
  Protected ProcStartPos, ProcEndPos, PosInCode, mPosInCode, SearchFromPos, EOFLP, EOFLM, EOFLM2
  Protected ChildIdx, ChildID, IsProc, LocalPos, LocalLen, HasMacroArgs, HasMacroParens, Idx, EqPos, DeltaLen
  Protected NewList CalledMacros.MacroExpandEntry()
  Protected NewMap ParamMap.s()
  ;
  If *ProcOrMacroCode = 0 : ProcedureReturn : EndIf
  If ElementFixedID <= 0 Or SelectPBBElement(ElementFixedID) = 0
    *ProcOrMacroCode\s = "Error"
    ProcedureReturn
  EndIf
  IsProc = Bool(ElementList()\Type = #PBBProcedure)
  If IsProc : ListOfMacro$ = "" : EndIf
  ;
  ProcFile$          = ElementList()\FileName
  ProcStartPos       = ElementList()\StartingPos
  ProcEndPos         = ElementList()\EndingPos
  *SearchString      = GetPointedElementContentFromFile(ProcFile$, ProcStartPos, ProcEndPos)
  *ProcOrMacroCode\s = *SearchString\s
  ;
  ; Start scanning strictly past the container's own header identifier (or closing ')' for a Macro)
  If IsProc = #False
    PosInCode = GetProcedureHeaderEndPos(*ProcOrMacroCode, 1, Len(*ProcOrMacroCode\s)) + 1
  EndIf
  If PosInCode <= 1
    PosInCode = FindInPBPointedCode(*ProcOrMacroCode, ElementList()\PurifiedName, 1, #CheckAll | #CheckOneKeyWordOnly, #PB_String_NoCase)
    If PosInCode : PosInCode + Len(ElementList()\PurifiedName) : Else : PosInCode = 1 : EndIf
  EndIf
  ;
  CollectCalledMacros(ElementList()\UsageChildren, #True)
  If ListSize(CalledMacros()) = 0 : ProcedureReturn : EndIf
  ;
  ProcOrMacroCode_LCase\s = LCase(*ProcOrMacroCode\s)
  Repeat
    mPosInCode   = 0
    *ActiveMacro = 0
    ForEach CalledMacros()
      If SelectPBBElement(CalledMacros()\MacroID)
        If CalledMacros()\SkipUntilPos > PosInCode : SearchFromPos = CalledMacros()\SkipUntilPos : Else : SearchFromPos = PosInCode : EndIf
        SearchFromPos = FindInPBPointedCode(ProcOrMacroCode_LCase, ElementList()\PurifiedName, SearchFromPos, #CheckAll | #CheckOneKeyWordOnly, #PB_String_CaseSensitive)
        If SearchFromPos And (mPosInCode = 0 Or SearchFromPos < mPosInCode)
          mPosInCode   = SearchFromPos
          *ActiveMacro = @CalledMacros()
        EndIf
      EndIf
    Next
    If mPosInCode = 0 : Break : EndIf
    PosInCode = mPosInCode
    ;
    ChangeCurrentElement(CalledMacros(), *ActiveMacro)
    SelectPBBElement(CalledMacros()\MacroID)
    ElementName$      = ElementList()\Name
    ElementNameLCase$ = ElementList()\PurifiedName
    MacroDecl$        = ElementList()\Declaration
    HasMacroArgs      = HasMacroArguments(MacroDecl$)
    HasMacroParens    = Bool(FindString(MacroDecl$, "(") > 0)
    ;
    If FindString(#TAB$ + ListOfMacro$, #TAB$ + ElementName$ + #TAB$) = 0
      ListOfMacro$ + ElementName$ + #TAB$
    EndIf
    ;
    EOFLP = PosInCode + Len(ElementNameLCase$)
    If HasMacroParens : CallingParamList$ = ExtractBalancedCallParams(*ProcOrMacroCode, EOFLP, @EOFLP) : Else : CallingParamList$ = "" : EndIf
    ;
    If CallingParamList$ <> "" And HasMacroArgs = 0
      AlertInPBBWindow("In procedure ReplaceMacroCallsByMacroCodes(), calling of " + ElementName$ + " is done with parameters, but this macro has no parameters!")
      PosInCode + 1
      Continue
    ElseIf CallingParamList$ = "" And HasMacroArgs = 1
      AlertInPBBWindow("In procedure ReplaceMacroCallsByMacroCodes(), calling of " + ElementName$ + " is done without parameters, but this macro has parameters!")
      PosInCode + 1
      Continue
    EndIf
    ;
    *SearchString = GetPointedElementContentFromFile(ElementList()\FileName, ElementList()\StartingPos, ElementList()\EndingPos)
    MacroCode\s   = *SearchString\s
    DeleteCommentsAndSpacesFromPBCode(MacroCode, #True, #True)
    EOFLM = Len("Macro ") + 1 + Len(ElementNameLCase$)
    If HasMacroParens : MacroParamList$ = ExtractBalancedCallParams(MacroCode, EOFLM, @EOFLM) : Else : MacroParamList$ = "" : EndIf
    ;
    MacroCode_LCase\s = LCase(MacroCode\s)
    EOFLM2            = FindInPBPointedCode(MacroCode_LCase, "endmacro", EOFLM, #CheckAll, #PB_String_CaseSensitive) - 1
    If EOFLM2 <= 0
      AlertInPBBWindow("In procedure ReplaceMacroCallsByMacroCodes(), 'EndMacro' is missing in the code of " + ElementName$ + "!")
      *ProcOrMacroCode\s = "Error"
      ProcedureReturn
    EndIf
    ;
    ; Strip leading/trailing spaces and single-line ':' or #CR$ macro boundary delimiters
    While EOFLM <= EOFLM2 And FastMid(MacroCode\s, EOFLM, 1) = " " : EOFLM + 1 : Wend
    If EOFLM <= EOFLM2 And (FastMid(MacroCode\s, EOFLM, 1) = ":" Or FastMid(MacroCode\s, EOFLM, 1) = #CR$) : EOFLM + 1 : EndIf
    While EOFLM2 >= EOFLM And FastMid(MacroCode\s, EOFLM2, 1) = " " : EOFLM2 - 1 : Wend
    If EOFLM2 >= EOFLM And (FastMid(MacroCode\s, EOFLM2, 1) = ":" Or FastMid(MacroCode\s, EOFLM2, 1) = #CR$) : EOFLM2 - 1 : EndIf
    ;
    If EOFLM2 >= EOFLM
      MacroCode\s = Trim(FastMid(MacroCode\s, EOFLM, EOFLM2 - EOFLM + 1))
      ;
      ; Simultaneous single-pass parameter substitution and '#' token concatenation
      ClearMap(ParamMap())
      If MacroParamList$ <> ""
        Idx = 0
        Repeat
          Idx + 1
          RawParam$ = Trim(StringField(MacroParamList$, Idx, #TAB$))
          If RawParam$ = "" : Break : EndIf
          EqPos = FindString(RawParam$, "=")
          If EqPos : ParamMap(LCase(Trim(FastLeft(RawParam$, EqPos - 1)))) = Trim(FastMid(RawParam$, EqPos + 1)) : Else : ParamMap(LCase(RawParam$)) = "" : EndIf
          If CallingParamList$ <> ""
            ReplacedSeg$ = Trim(StringField(CallingParamList$, Idx, #TAB$))
            If ReplacedSeg$ <> "" : ParamMap() = ReplacedSeg$ : EndIf
          EndIf
        ForEver
      EndIf
      ;
      If MapSize(ParamMap()) > 0 Or FindString(MacroCode\s, "#")
        NewBody$ = ""
        *Char    = @MacroCode\s
        *Start   = *Char
        While *Char\c <> 0
          If IsStringOrCharLiteralStart(*Char)
            ExtractOrSkipNextString(@*Char, *Char\c, #True)
          ElseIf *Char\c = '#'
            *SegStart = *Char + SizeOf(Character)
            While *SegStart\c <> 0 And *SegStart\c <> '#' And IsPBSeparator(*SegStart\c) = #False : GoNextChar(*SegStart) : Wend
            SubToken$ = PeekS(*Char + SizeOf(Character), ComputeLenFromChars(*Char + SizeOf(Character), *SegStart))
            ; Strip PureBasic '#' concatenation operator when attached to a preceding identifier or followed by a formal parameter:
            If PeekC(*Char + SizeOf(Character)) <> 0 And PeekC(*Char + SizeOf(Character)) <> #Space And PeekC(*Char + SizeOf(Character)) <> #CR And
               ((*Char > @MacroCode\s And IsPBSeparator(PeekC(*Char - SizeOf(Character))) = #False And PeekC(*Char - SizeOf(Character)) <> '#') Or
                (SubToken$ <> "" And FindMapElement(ParamMap(), LCase(SubToken$))))
              NewBody$ + PeekS(*Start, ComputeLenFromChars(*Start, *Char))
              GoNextChar(*Char)
              *Start = *Char
            Else
              ; Preserve standalone '#Constant' token verbatim so its root name is never mistaken for a formal parameter
              *Char = *SegStart
            EndIf
          ElseIf IsPBSeparator(*Char\c) = #False
            *SegStart = *Char
            While *Char\c <> 0 And *Char\c <> '#' And IsPBSeparator(*Char\c) = #False : GoNextChar(*Char) : Wend
            SubToken$ = PeekS(*SegStart, ComputeLenFromChars(*SegStart, *Char))
            If FindMapElement(ParamMap(), LCase(SubToken$))
              NewBody$ + PeekS(*Start, ComputeLenFromChars(*Start, *SegStart)) + ParamMap()
              *Start = *Char
            EndIf
          Else
            GoNextChar(*Char)
          EndIf
        Wend
        If *Char > *Start : NewBody$ + PeekS(*Start, ComputeLenFromChars(*Start, *Char)) : EndIf
        MacroCode\s = NewBody$
      EndIf
      ;
      ; Register any inner macros referenced by the expanded macro itself
      SelectPBBElement(*ActiveMacro\MacroID)
      CollectCalledMacros(ElementList()\UsageChildren, #False)
    Else
      MacroCode\s = ""
    EndIf
    ;
    ; Splice expanded body, shift active 'SkipUntilPos' horizons by DeltaLen, and shield this macro over its own expansion
    DeltaLen = Len(MacroCode\s) - (EOFLP - PosInCode)
    ForEach CalledMacros()
      If CalledMacros()\SkipUntilPos >= EOFLP
        CalledMacros()\SkipUntilPos + DeltaLen
      EndIf
    Next
    *ActiveMacro\SkipUntilPos = PosInCode + Len(MacroCode\s)
    *ProcOrMacroCode\s        = FastLeft(*ProcOrMacroCode\s, PosInCode - 1) + MacroCode\s + FastMid(*ProcOrMacroCode\s, EOFLP)
    ProcOrMacroCode_LCase\s   = LCase(*ProcOrMacroCode\s)
  ForEver
EndProcedure
;
Procedure RegisterOrUpdateProcVariable(*Result.VariableListings, List ProcVarsList.ProcVarEntry(), Map ProcVarsMap.i(),
                                       Token$, TokenLCase$, *StringStartingAdr, *NextChar.Character,
                                       VariableSpecie$, SubSpecie$, DefaultDotType$,
                                       ProcFixedID, ProcFileName$, ProcEndingPos,
                                       ModuleContainer, LastSeparator = 0)
  ;
  ; Role: Evaluates and records a declared variable/parameter (when VariableSpecie$ <> "") or an expression
  ; identifier (when VariableSpecie$ = "", including "Mod::Ident") in O(1) time without advancing the caller's pointer.
  ; Module-qualified tokens carry their "mod::" prefix directly in ProcVarsMap() so 'Mod::Var' never collides
  ; with or marks a homonymous local variable as used.
  ;
  Protected TypeToken$, DotTypeLCase$, DeclareLabel$, DisplayName$, MapKey$, BareTokenLCase$
  Protected HasParentheses, IsUsed, MatchedID, ColonPos
  Protected *VarEntry.ProcVarEntry
  Protected ResolvedString$ ; Required by macro AddUniqueListItem()
  ;
  If TokenLCase$ = "" : ProcedureReturn : EndIf
  ColonPos = FindString(TokenLCase$, "::")
  If ColonPos
    BareTokenLCase$ = FastMid(TokenLCase$, ColonPos + 2)
  Else
    BareTokenLCase$ = TokenLCase$
  EndIf
  ;
  ; 1. Read-only lookahead for optional '.Type' and '()' (*NextChar is a local value copy of caller's *Char)
  SkipSpaces(*NextChar)
  If *NextChar\c = '.' Or DefaultDotType$ <> ""
    TypeToken$ = ExtractOptionalDotType(*StringStartingAdr, @*NextChar, DefaultDotType$)
    SkipSpaces(*NextChar)
  EndIf
  HasParentheses = Bool(*NextChar\c = '(')
  ;
  ; 2. Single O(1) lookup in ProcVarsMap() with natural "mod::var" key; zero-allocation fast-exit on repeated usages
  MapKey$ = TokenLCase$ + Left("()", HasParentheses * 2)
  If FindMapElement(ProcVarsMap(), MapKey$)
    *VarEntry = ProcVarsMap()
    If VariableSpecie$ = ""
      *VarEntry\IsUsed = #True
      If TypeToken$ = "" Or *VarEntry\DotTypeLCase <> "" Or Right(Token$, 1) = "$"
        ProcedureReturn
      EndIf
      DeclareLabel$ = *VarEntry\DeclareLabel
    EndIf
  EndIf
  ;
  DotTypeLCase$ = LCase(TypeToken$)
  IsUsed        = Bool(VariableSpecie$ = "")
  If IsUsed = #False
    DeclareLabel$ = VariableSpecie$
    If SubSpecie$ <> "" And HasParentheses
      DeclareLabel$ + " " + SubSpecie$
    EndIf
  EndIf
  ;
  ; 3. Query ElementList() for 'Shared'/'Define' declarations (restricted to #PBBVariable) or first-time non-local usages
  If (VariableSpecie$ = "Shared" Or VariableSpecie$ = "Define" Or (IsUsed And *VarEntry = 0)) And FindMapElement(ElementMapIndex(), BareTokenLCase$)
    MatchedID = FindElementInList(TokenLCase$, Bool(IsUsed = #False) * #PBBVariable, ProcEndingPos, ProcFileName$, ProcFixedID, HasParentheses, DotTypeLCase$, 0, ModuleContainer)
    If MatchedID = 0 And IsUsed
      MatchedID = FindElementInList(TokenLCase$, 0, 0, ProcFileName$, ProcFixedID, HasParentheses, DotTypeLCase$, 0, ModuleContainer)
    EndIf
    If MatchedID > 0
      If ElementList()\Type = #PBBVariable
        If VariableSpecie$ = "Define" And (ElementList()\LocalVarContainer <= 0 Or *VarEntry <> 0)
          IsUsed = #True
        EndIf
        If TypeToken$ = "" Or IsUsed = #False
          DisplayName$ = FastLeft(Token$, ColonPos + Bool(ColonPos > 0)) + ElementList()\Name
        EndIf
        If DotTypeLCase$ = "" : DotTypeLCase$ = ElementList()\VarDotType_LCase : EndIf
        DeclareLabel$ = ElementList()\VariableSpecie
        If ElementList()\VariableSubSpecie <> "" And ElementList()\VariableSubSpecie <> "Var"
          DeclareLabel$ + " " + ElementList()\VariableSubSpecie
        EndIf
      ElseIf ElementList()\Type = #PBBImported And ElementList()\HasParentheses = #False
        If TypeToken$ = ""
          DisplayName$ = FastLeft(Token$, ColonPos + Bool(ColonPos > 0)) + ElementList()\Name
        EndIf
        If DotTypeLCase$ = "" : DotTypeLCase$ = ElementList()\VarDotType_LCase : EndIf
        DeclareLabel$ = "Imported"
      Else
        If ElementList()\Type = #PBBProcedure Or (ElementList()\Type = #PBBImported And ElementList()\HasParentheses)
          AddUniqueListItem(*Result\CalledProcedures, ElementList()\Name)
        EndIf
        ProcedureReturn
      EndIf
    EndIf
  EndIf
  ;
  ; 4. Filter out PureBasic keywords/native functions/types on unqualified usages, or classify as Unprotected
  If IsUsed And DeclareLabel$ = ""
    If ColonPos = 0 And
       (FindMapElement(PBBasicKeywordMap(), TokenLCase$) Or
        FindMapElement(PBNativeFunctions(), TokenLCase$) Or
        (HasParentheses And Right(TokenLCase$, 1) = "_") Or
        FindMapElement(PB_BlockDeclarators(), TokenLCase$) Or
        FindMapElement(PBBasicNativeTypes(), TokenLCase$))
      ProcedureReturn
    ElseIf HasParentheses = #False And ColonPos = 0
      DeclareLabel$ = "Unprotected"
    EndIf
  EndIf
  ;
  ; 5. Unified DisplayName$ formatting for missing elements, late-typed usages, and variable entries
  If DisplayName$ = ""
    DisplayName$ = Token$
    If Right(DisplayName$, 1) <> "$"
      If TypeToken$ <> ""
        DisplayName$ + "." + TypeToken$
      ElseIf DeclareLabel$ <> "" And DeclareLabel$ <> "Unprotected" And VariableSpecie$ <> "Parameter" And VariableSpecie$ <> "Shared"
        DisplayName$ + ".i"
        DotTypeLCase$ = "i"
      EndIf
    EndIf
    If HasParentheses : DisplayName$ + "()" : EndIf
  EndIf
  ;
  If DeclareLabel$ = ""
    If HasParentheses And LastSeparator = '@'
      DisplayName$ = "@" + DisplayName$
    EndIf
    AddUniqueListItem(*Result\MissingElements, DisplayName$)
    ProcedureReturn
  EndIf
  ;
  If *VarEntry = 0
    AddElement(ProcVarsList())
    ProcVarsList()\DeclareLabel = DeclareLabel$
    ProcVarsList()\DisplayName  = DisplayName$
    ProcVarsList()\DotTypeLCase = DotTypeLCase$
    ProcVarsList()\IsUsed       = IsUsed
    ProcVarsMap(MapKey$)        = @ProcVarsList()
  Else
    If IsUsed : *VarEntry\IsUsed = #True : EndIf
    If *VarEntry\DeclareLabel = "Unprotected" And DeclareLabel$ <> "Unprotected"
      *VarEntry\DeclareLabel = DeclareLabel$
    EndIf
    If DotTypeLCase$ <> "" And *VarEntry\DotTypeLCase = ""
      *VarEntry\DotTypeLCase = DotTypeLCase$
      *VarEntry\DisplayName  = DisplayName$
    EndIf
  EndIf
EndProcedure
;
Procedure ListVariables(ProcOrMacroName$, *Result.VariableListings, TargetElementType = #PBBProcedure, ModuleName$ = "")
  ;
  ; Role: Single-pass lexical analyzer that builds the complete inventory of variables
  ; (parameters, explicitly declared local variables, referenced Global/Threaded/Shared variables,
  ; and implicit unprotected variables), called procedures, and missing project elements inside 'ProcOrMacroName$'.
  ; Architecture: Skips the container's own header prefix upfront, then shares GetNextToken(#ParsingMode_ListVariables)
  ; with OnePassSetElements_SegmentExam(), resolving "mod::ident" tokens through FindElementInList() and local variables via ProcVarsMap().
  ;
  Shared ListOfMacro$, PBNativeConstantMap()
  ;
  Protected ProcCode.String, RootType.String
  Protected *StringStartingAdr, *Char.Character, *TokenPosChar.Character, *TempChar.Character
  Protected ParsingState, LineToBeContinued, IsStartOfStatement, OpenParentheses, StructFuncParenLevel
  Protected IsInsideDefinition, IsInsideDefaultValue, NextTokenIsPotentialVariable, HasParentheses, IsQualified, ColonPos
  Protected TypeOfElement, VariableSpecie$, FieldOrVariableSubSpecie$, VariableDotType$
  Protected LastSeparator, LastComment$, Token$, TokenLCase$, LastToken$, WithVar$, PrecedingChain$, ResolvedString$
  Protected FileName$, ProcContainer = -1, MacroProcEndingPos, InterfaceOrStructContainer = -1
  Protected ModuleContainer = -1, ExpectedTargetType, UsageTargetType, UsageParentheses, FirstChar, CDepth
  Protected NewList ProcVarsList.ProcVarEntry()
  Protected Dim CStack.CompilerStackLevel(#PBB_MaxCompilerDepth)
  ;
  ClearStructure(*Result, VariableListings)
  *Result\ProcedureName = ProcOrMacroName$
  If ExplorationState <> #ExplorationState_Completed : ProcedureReturn #False : EndIf
  ;
  ; 1. Retrieve target Procedure or Macro metadata and inline any referenced project macros
  If ModuleName$ <> "" : ModuleContainer = GetModuleID(ModuleName$) : EndIf
  ProcContainer = FindElementInList(PurifyElementName(ProcOrMacroName$), TargetElementType, 0, "", 0, -1, "", 0, ModuleContainer)
  If ProcContainer = 0 : ProcedureReturn #False : EndIf
  FileName$          = ElementList()\FileName
  MacroProcEndingPos = ElementList()\EndingPos
  ModuleContainer    = ElementList()\EnclosingModuleID
  ;
  ListOfMacro$ = ""
  ReplaceMacroCallsByMacroCodes(ProcContainer, ProcCode)
  If ProcCode\s = "Error" Or ProcCode\s = "" : ProcedureReturn #False : EndIf
  ProcCode\s + #CR$
  ;
  ClearMap(EvalContext\ProcVarsMap())
  EvalContext\InListVariables = #True
  EvalContext\TargetType      = TargetElementType
  EvalContext\ProcContainer   = ProcContainer
  EvalContext\ModuleContainer = ModuleContainer
  EvalContext\PosInCode       = MacroProcEndingPos
  EvalContext\FileName        = FileName$
  EvalContext\ProcedureName   = ProcOrMacroName$
  ;
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
  CompilerEndIf
  ;
  ; 2. Initialize single-pass parser state and advance past the Procedure/Macro header identifier upfront
  *StringStartingAdr = @ProcCode\s
  *Char              = *StringStartingAdr
  SkipSpaces(*Char)
  ParsingState       = #Parsing_Start
  IsStartOfStatement = #True
  ResetParsingModeFlags
  ;
  ExtractOrSkipNextToken(*StringStartingAdr, @*Char) ; Consume 'Procedure...' or 'Macro' keyword
  Token$ = ExtractOptionalDotType(*StringStartingAdr, @*Char)
  If Token$ <> ""
    ColonPos = FindString(Token$, "::")
    If ColonPos > 0 And GetModuleID(FastLeft(Token$, ColonPos - 1)) = 0
      AddUniqueListItem(*Result\MissingElements, FastLeft(Token$, ColonPos + 1))
    ElseIf (ColonPos > 0 Or FindMapElement(PBBasicNativeTypes(), LCase(Token$)) = 0) And
           FindElementInList(LCase(Token$), #PBBStructOrInterface, MacroProcEndingPos, FileName$, 0, -1, "", 0, ModuleContainer) = 0
      AddUniqueListItem(*Result\MissingElements, "." + Token$)
    EndIf
  EndIf
  ExtractOrSkipNextToken(*StringStartingAdr, @*Char) ; Consume container's own identifier so it is never parsed as a call
  If CheckForParenthesis(@*Char)
    IsInsideDefinition           = #True
    VariableSpecie$              = "Parameter"
    NextTokenIsPotentialVariable = #True
  EndIf
  ;
  Repeat
    GetNextToken(#ParsingMode_ListVariables)
    If ParsingState = #Parsing_EndOfFile : Break : EndIf
    If ParsingState <> #Parsing_NewToken : Continue : EndIf
    ;
    ; ---------------- 1. Conditional compilation, local blocks, 'With' context, and labels ----------------
    If IsStartOfStatement
      If Left(TokenLCase$, 8) = "compiler" And
         ProcessCompilerDirectiveAtStatement(TokenLCase$, *TokenPosChar, *StringStartingAdr, @*Char,
                                             CStack(), @CDepth, 0, FileName$, MacroProcEndingPos, ModuleContainer)
        Continue
      ElseIf TokenLCase$ = "structure" Or TokenLCase$ = "interface" Or TokenLCase$ = "import" Or TokenLCase$ = "importc"
        InterfaceOrStructContainer = 1
        If TokenLCase$ = "import" Or TokenLCase$ = "importc"
          GoToEOL_AndExtractComment(@*Char)
        Else
          TypeOfElement = #PBBStructure
          ExtractOrSkipNextToken(*StringStartingAdr, @*Char) ; Skip local structure/interface name
        EndIf
        Continue
      ElseIf TokenLCase$ = "endstructure" Or TokenLCase$ = "endinterface" Or TokenLCase$ = "endimport"
        InterfaceOrStructContainer = -1
        Continue
      ElseIf InterfaceOrStructContainer <= 0
        If TokenLCase$ = "with"    : WithVar$ = ExtractWithTarget(*Char) : Continue : EndIf
        If TokenLCase$ = "endwith" : WithVar$ = ""                       : Continue : EndIf
        If IsStandaloneLabel(TokenLCase$, *Char)
          GoNextChar(*Char) : GoToEOL_AndExtractComment(@*Char)
          Continue
        EndIf
      EndIf
    EndIf
    If InterfaceOrStructContainer > 0 And TypeOfElement <> #PBBStructure : Continue : EndIf
    ;
    ; ---------------- 2. Formal parameter & local/global variable registration ----------------
    If UsageTargetType = 0 And LastSeparator <> '.' And IsQualified = #False And
       OpenParentheses = Bool(IsInsideDefinition) And IsInsideDefaultValue = #False And
       NextTokenIsPotentialVariable And VariableSpecie$ <> ""
      NextTokenIsPotentialVariable = #False
      If FirstChar = '*' Or IsValidIdentifierStart(FirstChar)
        RegisterOrUpdateProcVariable(*Result, ProcVarsList(), EvalContext\ProcVarsMap(), Token$, TokenLCase$,
                                     *StringStartingAdr, *Char, VariableSpecie$, FieldOrVariableSubSpecie$,
                                     VariableDotType$, ProcContainer, FileName$, MacroProcEndingPos, ModuleContainer)
      EndIf
      Continue
    EndIf
    ;
    ; -------- 3. Central 'Mod::' validation, typed contexts, fields, constants, and usages --------
    If IsQualified
      ColonPos = FindString(TokenLCase$, "::")
      If ColonPos <= 1 Or GetModuleID(FastLeft(TokenLCase$, ColonPos - 1)) = 0
        If ColonPos > 1 : AddUniqueListItem(*Result\MissingElements, FastLeft(Token$, ColonPos + 1)) : EndIf
        Continue
      EndIf
    EndIf
    ;
    If FirstChar = '#' And UsageTargetType = 0
      UsageTargetType = #PBBConstant
    EndIf
    If UsageTargetType = #PBBField
      If FindElementInList(TokenLCase$, #PBBField) = 0 And FindMapElement(PBBasicNativeTypes(), TokenLCase$) = 0
        *TempChar = *TokenPosChar - SizeOf(Character)
        SkipSpacesBackward(*TempChar, *StringStartingAdr + SizeOf(Character))
        PrecedingChain$ = ExtractPrecedingFieldChain(*TempChar, *StringStartingAdr, WithVar$)
        RootType\s      = ""
        If FindString(PrecedingChain$, "\") = 0
          PrecedingChain$ = PurifyElementName(PrecedingChain$, @HasParentheses, @RootType)
          If RootType\s = "" And FindMapElement(EvalContext\ProcVarsMap(), PrecedingChain$ + Left("()", HasParentheses * 2))
            ChangeCurrentElement(ProcVarsList(), EvalContext\ProcVarsMap())
            RootType\s = ProcVarsList()\DotTypeLCase
          EndIf
        EndIf
        If RootType\s = "" And PrecedingChain$ <> ""
          RootType\s = StringField(GetStructureOrInterfaceNameFromVariable(PrecedingChain$, "", FileName$, MacroProcEndingPos, 0, #True,
                                                                           ProcContainer, "", GetEnclosingModuleName(ModuleContainer)), 2, " ")
        EndIf
        If RootType\s = "" Or FindString(RootType\s, "::") Or
           FindString("," + LCase(#PBPredefinedStructures$) + ",", "," + LCase(RootType\s) + ",") = 0
          AddUniqueListItem(*Result\MissingElements, "\" + Token$)
        EndIf
      EndIf
      Continue
    ElseIf UsageTargetType <> 0
      Select UsageTargetType
        Case #PBBModule
          If GetModuleID(TokenLCase$) = 0 : AddUniqueListItem(*Result\MissingElements, Token$ + "::") : EndIf
          Continue
        Case #PBBLabel, #PBBConstant
          If (UsageTargetType = #PBBConstant And IsQualified = #False And FindMapElement(PBNativeConstantMap(), UCase(Token$))) Or
             FindElementInList(TokenLCase$, UsageTargetType, 0, FileName$, ProcContainer, -1, "", 0, ModuleContainer)
            Continue
          EndIf
          AddUniqueListItem(*Result\MissingElements, Left("?", Bool(LastSeparator = '?')) + Token$)
          Continue
        Case #PBBStructure, #PBBStructOrInterface, #PBBTypeOrVar
          If (IsQualified = #False And FindMapElement(PBBasicNativeTypes(), TokenLCase$)) Or
             FindElementInList(TokenLCase$, Bool(UsageTargetType = #PBBTypeOrVar) * #PBBStructOrInterface +
                                            Bool(UsageTargetType <> #PBBTypeOrVar) * UsageTargetType,
                               MacroProcEndingPos, FileName$, 0, -1, "", 0, ModuleContainer)
            Continue
          EndIf
          If UsageTargetType <> #PBBTypeOrVar Or
             (FirstChar <> '#' And FirstChar <> '*' And Right(Token$, 1) <> "$" And
              FindMapElement(EvalContext\ProcVarsMap(), TokenLCase$ + Left("()", UsageParentheses * 2)) = 0 And
              FindElementInList(TokenLCase$, #PBBVariable, 0, FileName$, ProcContainer, UsageParentheses, "", 0, ModuleContainer) = 0)
            AddUniqueListItem(*Result\MissingElements, "." + Token$)
            Continue
          EndIf
      EndSelect
    EndIf
    ;
    ; General identifier usage (and SizeOf/TypeOf fallback on local/global variables)
    If FirstChar = '*' Or IsValidIdentifierStart(FirstChar)
      RegisterOrUpdateProcVariable(*Result, ProcVarsList(), EvalContext\ProcVarsMap(), Token$, TokenLCase$,
                                   *StringStartingAdr, *Char, "", "", "", ProcContainer, FileName$,
                                   MacroProcEndingPos, ModuleContainer, LastSeparator)
    EndIf
  ForEver
  ;
  EvalContext\InListVariables = #False
  EvalContext\FileName        = ""
  EvalContext\ProcedureName   = ""
  EvalContext\PosInCode       = 0
  ClearMap(EvalContext\ProcVarsMap())
  ;
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
  CompilerEndIf
  ;
  ; 4. Build CompleteList, Unused, Unprotected, CalledProcedures, and MissingElements:
  ForEach ProcVarsList()
    *Result\CompleteList + ProcVarsList()\DeclareLabel + #TAB$ + ProcVarsList()\DisplayName + #TAB$ +
                           Mid("UnusedUsed", 1 + Bool(ProcVarsList()\IsUsed) * 6, 6 - Bool(ProcVarsList()\IsUsed) * 2) + #CR$
    If ProcVarsList()\IsUsed = #False              : *Result\Unused      + ProcVarsList()\DisplayName + ", " : EndIf
    If ProcVarsList()\DeclareLabel = "Unprotected" : *Result\Unprotected + ProcVarsList()\DisplayName + ", " : EndIf
  Next
  If *Result\Unused           : *Result\Unused           = GetTextFromCatalogPB("UnusedVariables:")      + " " + FastLeft(*Result\Unused,           Len(*Result\Unused) - 2)           : EndIf
  If *Result\Unprotected      : *Result\Unprotected      = GetTextFromCatalogPB("UnprotectedVariables:") + " " + FastLeft(*Result\Unprotected,      Len(*Result\Unprotected) - 2)      : EndIf
  If *Result\CalledProcedures : *Result\CalledProcedures = FastLeft(*Result\CalledProcedures, Len(*Result\CalledProcedures) - 2)                                                       : EndIf
  If *Result\MissingElements  : *Result\MissingElements  = GetTextFromCatalogPB("MissingElements:")      + " " + FastLeft(*Result\MissingElements,  Len(*Result\MissingElements) - 2)  : EndIf
  ProcedureReturn #True
EndProcedure
;

; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 417
; FirstLine = 260
; Folding = +Ow3AYxHKAcBAiDA+
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb
; DisableDebugger