; *******************************************************************************
;
;                      Find/Replace Window for PBBrowser
;
;   NOTE: The gadget structure declaration (Structure _FR_Gadgets and FR_Gadgets)
;         is located in Section 6 of PBBrowserDeclarations.pb.
;
; *******************************************************************************
;
; ===============================================================================
; 1. DIALOG LAYOUT, THEME & LOCALIZATION
; ===============================================================================
;- Dialog layout, theme, and localization
;{
Procedure GUI_FR_SetReplaceState()
  ;
  ; Role: Synchronizes the replacement input field's read-only/color state and the
  ;       replacement action button's enabled state with the 'ReplaceCheck' checkbox
  ;       and the current scope list count.
  ;
  Protected ReplaceEnabled = GetGadgetState(FR_Gadgets\ReplaceCheck)
  ;
  DisableGadget(FR_Gadgets\ReplaceActionButton, Bool(ReplaceEnabled = 0 Or CountGadgetItems(FR_Gadgets\Scopes) = 0))
  ;
  If ReplaceEnabled
    SetGadgetColorEx(FR_Gadgets\ReplaceString, #PB_Gadget_FrontColor, InterfaceColorPresets()\TextColor, 1)
    SendMessage_(GadgetID(FR_Gadgets\ReplaceString), #EM_SETREADONLY, #False, 0)
  Else
    SetGadgetColorEx(FR_Gadgets\ReplaceString, #PB_Gadget_FrontColor, GetPBBNumParameter("PBBUnusedElementColor"), 1)
    SendMessage_(GadgetID(FR_Gadgets\ReplaceString), #EM_SETREADONLY, #True, 0)
  EndIf
EndProcedure
;
Procedure GUI_FR_SetGadgetsTitles()
  ;
  ; Role: Updates the window title, gadget captions, action button symbols, and balloon
  ;       tooltips according to the active language catalog.
  ;
  Protected Loupe$ = Space(2)
  Protected txa$
  ;
  SetWindowTitle(FR_Gadgets\Window, GetTextFromCatalogPB("FindReplaceTitle"))
  SetGadgetText(FR_Gadgets\SearchLabel, GetTextFromCatalogPB("Search"))
  ;
  ; Build the UTF-16 surrogate pair for the magnifying glass symbol ("🔍", U+1F50D),
  ; since Chr() only accepts code points up to $FFFF:
  PokeW(@Loupe$, $D83D)
  PokeW(@Loupe$ + 2, $DD0D)
  SetGadgetText(FR_Gadgets\SearchActionButton, Loupe$)
  ;
  txa$ = GetTextFromCatalogPB("SearchInstructions")
  RemoveBalloonToolTip(FR_Gadgets\ToolTip, FR_Gadgets\SearchString)
  AddBalloonToolTip(FR_Gadgets\ToolTip, FR_Gadgets\SearchString, txa$)
  ;
  txa$ = GetTextFromCatalogPB("ToolTipSearchAction")
  RemoveBalloonToolTip(FR_Gadgets\ToolTip, FR_Gadgets\SearchActionButton)
  AddBalloonToolTip(FR_Gadgets\ToolTip, FR_Gadgets\SearchActionButton, txa$)
  ;
  SetGadgetText(FR_Gadgets\ReplaceCheck, GetTextFromCatalogPB("ReplaceWith"))
  SetGadgetText(FR_Gadgets\ReplaceActionButton, Chr($25B6))
  txa$ = GetTextFromCatalogPB("ToolTipReplaceAction")
  RemoveBalloonToolTip(FR_Gadgets\ToolTip, FR_Gadgets\ReplaceActionButton)
  AddBalloonToolTip(FR_Gadgets\ToolTip, FR_Gadgets\ReplaceActionButton, txa$)
  ;
  SetGadgetText(FR_Gadgets\ReplaceNote,     GetTextFromCatalogPB("ReplaceNote"))
  SetGadgetText(FR_Gadgets\MatchExact,      GetTextFromCatalogPB("MatchExact"))
  SetGadgetText(FR_Gadgets\MatchStartsWith, GetTextFromCatalogPB("MatchStartsWith"))
  SetGadgetText(FR_Gadgets\MatchContains,   GetTextFromCatalogPB("MatchContains"))
  SetGadgetText(FR_Gadgets\IgnoreComments,  GetTextFromCatalogPB("IgnoreComments"))
  SetGadgetText(FR_Gadgets\IgnoreStrings,   GetTextFromCatalogPB("IgnoreStrings"))
EndProcedure
;
Procedure GUI_FR_ColorizeAndSetFontsOfGadgets()
  ;
  ; Role: Applies the active color theme and configured fonts to the Find/Replace window
  ;       and all its child gadgets via contiguous structure pointer traversal.
  ;
  Protected *GadgetAddress, GadgetID
  ;
  If FR_Gadgets\Window < 0 Or IsWindow(FR_Gadgets\Window) = #False
    ProcedureReturn
  EndIf
  ;
  SetWindowColor(FR_Gadgets\Window, InterfaceColorPresets()\BackGroundColor)
  ;
  *GadgetAddress = @FR_Gadgets\SearchLabel
  Repeat
    GadgetID = PeekI(*GadgetAddress)
    If IsGadget(GadgetID)
      If GadgetID = FR_Gadgets\SearchLabel Or GadgetID = FR_Gadgets\ReplaceCheck
        SetGadgetColorEx(GadgetID, #PB_Gadget_FrontColor, GetPBBNumParameter("PBBTitleColor"), 1)
        SetGadgetFont(GadgetID, PBBTitleFont)
      Else
        If GadgetID = FR_Gadgets\ReplaceString And GetGadgetState(FR_Gadgets\ReplaceCheck) = 0
          SetGadgetColorEx(GadgetID, #PB_Gadget_FrontColor, GetPBBNumParameter("PBBUnusedElementColor"), 1)
        Else
          SetGadgetColorEx(GadgetID, #PB_Gadget_FrontColor, InterfaceColorPresets()\TextColor, 1)
        EndIf
        SetGadgetFont(GadgetID, PBBAllGadgetsFont)
      EndIf
      SetGadgetColorEx(GadgetID, #PB_Gadget_BackColor, InterfaceColorPresets()\BackGroundColor, 1)
    EndIf
    *GadgetAddress + SizeOf(Integer)
  Until *GadgetAddress > @FR_Gadgets\IgnoreStrings
  ;
  ApplyDarkModeToWindow(FR_Gadgets\Window)
EndProcedure
;
Procedure GUI_FR_ResizeGadgets()
  ;
  ; Role: Dynamically repositions and resizes all gadgets when the Find/Replace window is resized.
  ;
  Protected EndingRight        = WindowWidth(FR_Gadgets\Window) - FR_Gadgets\ExternalMargin
  Protected SecondColumnPos    = FR_Gadgets\ExternalMargin + FR_Gadgets\ButtonWidth + FR_Gadgets\InternalMargin
  Protected CheckAndRadioWidth = (EndingRight - FR_Gadgets\ExternalMargin - FR_Gadgets\InternalMargin) / 2
  Protected ThirdColumnPos     = FR_Gadgets\ExternalMargin + CheckAndRadioWidth + FR_Gadgets\InternalMargin
  Protected ActionBtnSize      = FR_Gadgets\ButtonHeight
  Protected SearchStrWidth     = EndingRight - SecondColumnPos - ActionBtnSize - FR_Gadgets\InternalMargin
  Protected BottomY            = WindowHeight(FR_Gadgets\Window) - 65
  Protected VPos               = FR_Gadgets\ExternalMargin
  ;
  ResizeGadget(FR_Gadgets\SearchLabel,        FR_Gadgets\ExternalMargin,    VPos, FR_Gadgets\ButtonWidth, FR_Gadgets\ButtonHeight)
  ResizeGadget(FR_Gadgets\SearchString,       SecondColumnPos,              VPos, SearchStrWidth,         FR_Gadgets\ButtonHeight)
  ResizeGadget(FR_Gadgets\SearchActionButton, EndingRight - ActionBtnSize,  VPos, ActionBtnSize,          FR_Gadgets\ButtonHeight)
  ;
  VPos + FR_Gadgets\ButtonHeight + FR_Gadgets\InternalMargin - 4
  ResizeGadget(FR_Gadgets\MatchExact,         FR_Gadgets\ExternalMargin,    VPos, CheckAndRadioWidth,     FR_Gadgets\ButtonHeight)
  ResizeGadget(FR_Gadgets\IgnoreComments,     ThirdColumnPos,               VPos, CheckAndRadioWidth,     FR_Gadgets\ButtonHeight)
  ;
  VPos + FR_Gadgets\ButtonHeight - 4
  ResizeGadget(FR_Gadgets\MatchStartsWith,    FR_Gadgets\ExternalMargin,    VPos, CheckAndRadioWidth,     FR_Gadgets\ButtonHeight)
  ResizeGadget(FR_Gadgets\IgnoreStrings,      ThirdColumnPos,               VPos, CheckAndRadioWidth,     FR_Gadgets\ButtonHeight)
  ;
  VPos + FR_Gadgets\ButtonHeight - 4
  ResizeGadget(FR_Gadgets\MatchContains,      FR_Gadgets\ExternalMargin,    VPos, CheckAndRadioWidth,     FR_Gadgets\ButtonHeight)
  ;
  VPos + FR_Gadgets\ButtonHeight + FR_Gadgets\InternalMargin
  ResizeGadget(FR_Gadgets\Scopes,             FR_Gadgets\ExternalMargin,    VPos, EndingRight - FR_Gadgets\ExternalMargin, BottomY - VPos - FR_Gadgets\InternalMargin)
  ;
  ResizeGadget(FR_Gadgets\ReplaceCheck,        FR_Gadgets\ExternalMargin,   BottomY, FR_Gadgets\ButtonWidth, FR_Gadgets\ButtonHeight)
  ResizeGadget(FR_Gadgets\ReplaceString,       SecondColumnPos,             BottomY, SearchStrWidth,         FR_Gadgets\ButtonHeight)
  ResizeGadget(FR_Gadgets\ReplaceActionButton, EndingRight - ActionBtnSize, BottomY, ActionBtnSize,          FR_Gadgets\ButtonHeight)
  ;
  VPos = BottomY + FR_Gadgets\ButtonHeight + FR_Gadgets\InternalMargin - 4
  ResizeGadget(FR_Gadgets\ReplaceNote,         FR_Gadgets\ExternalMargin,   VPos, EndingRight - FR_Gadgets\ExternalMargin, FR_Gadgets\ButtonHeight + 10)
EndProcedure
;}
;
; ===============================================================================
; 2. SEARCH DISPATCH & TWO-PASS SCOPE REPLACEMENT
; ===============================================================================
;- Search dispatch and two-pass scope replacement
;{
Procedure GUI_FR_RefreshScopes(*LastSearch.SearchDetails = 0, SelectedIndex = -1)
  ;
  ; Role: Populates the scopes ListViewGadget from SearchMenuChoices() (plus the trailing
  ;       'All' option) and highlights 'SelectedIndex'.
  ;
  Shared SearchMenuChoices()
  Protected ItemDetails$
  ;
  ClearGadgetItems(FR_Gadgets\Scopes)
  ;
  If ListSize(SearchMenuChoices()) > 0
    ForEach SearchMenuChoices()
      ItemDetails$ = ComputeExpressionDetails(SearchMenuChoices())
      AddGadgetItem(FR_Gadgets\Scopes, -1, SearchMenuChoices()\ElementName$ + #SpecialSpace$ + ItemDetails$)
    Next
  EndIf
  ;
  AddGadgetItem(FR_Gadgets\Scopes, -1, GetTextFromCatalogPB("All"))
  ;
  If SelectedIndex >= 0 And SelectedIndex < CountGadgetItems(FR_Gadgets\Scopes)
    SetGadgetState(FR_Gadgets\Scopes, SelectedIndex)
  Else
    SetGadgetState(FR_Gadgets\Scopes, 0)
  EndIf
  ;
  GUI_FR_SetReplaceState()
EndProcedure
;
Procedure.s GUI_FR_GetSearchExpression()
  ;
  ; Role: Returns the trimmed search expression from FR_Gadgets\SearchString decorated
  ;       with #PBBJoker$ according to the active match mode radio button (preserving an
  ;       explicit leading #PBBJoker$ in Exact mode when entered manually).
  ;
  Protected SExpression$ = Trim(GetGadgetText(FR_Gadgets\SearchString))
  ;
  If SExpression$ <> ""
    If GetGadgetState(FR_Gadgets\MatchStartsWith)
      SExpression$ = Trim(SExpression$, #PBBJoker$) + #PBBJoker$
    ElseIf GetGadgetState(FR_Gadgets\MatchContains)
      SExpression$ = #PBBJoker$ + Trim(SExpression$, #PBBJoker$) + #PBBJoker$
    ElseIf Left(SExpression$, Len(#PBBJoker$)) = #PBBJoker$ And Right(SExpression$, Len(#PBBJoker$)) <> #PBBJoker$
      ; Preserve a standalone leading wildcard (e.g., "…Init") in Exact mode:
      SExpression$ = RTrim(SExpression$, #PBBJoker$)
    Else
      SExpression$ = Trim(SExpression$, #PBBJoker$)
    EndIf
  EndIf
  ;
  ProcedureReturn SExpression$
EndProcedure
;
Procedure GUI_FR_ExecuteSearch(*LastSearch.SearchDetails = 0)
  ;
  ; Role: Reads the search string and matching options from the Find/Replace dialog,
  ;       updates '*LastSearch' and the main search field, and schedules a PBBrowser search.
  ;
  Shared ExecuteSearch ; Signals PBBrowserMainProc() how the queued search must be executed.
  Protected SExpression$ = GUI_FR_GetSearchExpression()
  ;
  If SExpression$ = "" Or *LastSearch = 0
    ProcedureReturn
  EndIf
  ;
  SetGadgetText(GPBBGadgets\SearchedExpression_gadget, SExpression$)
  PrepareSearchDetailsElement(*LastSearch, SExpression$)
  *LastSearch\SearchInCom       = Bool(GetGadgetState(FR_Gadgets\IgnoreComments) = 0)
  *LastSearch\SearchInQuotes    = Bool(GetGadgetState(FR_Gadgets\IgnoreStrings) = 0)
  *LastSearch\DontCheckPosition = #True
  *LastSearch\HasParentheses    = -1
  ;
  ExecuteSearch = #PBBUserWillChoose | #PBBSearchExact
EndProcedure
;
Structure ReplaceMatchStruct
  FileName.s    ; Full path of the source file containing the validated occurrence.
  PosInCode.i   ; 1-based starting character offset of the occurrence in the filtered buffer.
  MatchLen.i    ; Exact character span to replace in the filtered buffer.
  ReplaceWith.s ; Context-adapted replacement string for this specific occurrence.
EndStructure
;
Procedure GUI_FR_ExecuteReplacement(*LastSearch.SearchDetails)
  ;
  ; Role: Executes a two-pass scope-aware replacement across the project.
  ;       - Pass 1 collects all occurrences validated by FindNextMatchInProject(), adapts replacement
  ;         suffixes (.DotType and parentheses) to each occurrence's local syntax without mutating
  ;         in-memory buffers, and builds a concrete before/after confirmation preview when ambiguity exists.
  ;       - Pass 2 optionally creates a project copy via AskAndPrepareProjectCopy(), splices replacements
  ;         directly into the raw disk content in reverse positional order via FilteredPosToRawPos()
  ;         (strictly preserving original encoding, BOM, line endings, and tabs), and updates project state.
  ;
  Protected NewList MatchesToReplace.ReplaceMatchStruct()
  Protected NewList FilesToChange.FilesToChangeStruct()
  Protected FIPS.FIPSPosStruct
  Protected TestRemainingSearch.SearchDetails
  Protected TargetDir.String, RawDiskCode.String, ModifiedCode.String
  Protected *BufStart, *Char.Character
  Protected Reset = #True, FoundPos, MatchLen, ActualMatchLen
  Protected TokenStart, TokenEnd, SuffixEnd, BufferLen, PrevChar, NextChar, PreviewCount
  Protected ReplaceText$, ReplaceBase$, ReplaceDotType$, ActualReplaceWith$
  Protected TrailParen$, PreviewLine$, PreviewList$, NewMainFile$
  Protected NeedConfirmation, Result, WGadget, FileEncoding, HasBOM, FileModified, AnyFileModified
  ;
  If *LastSearch = 0 Or *LastSearch\ElementPurifiedName$ = "" Or GetGadgetState(FR_Gadgets\ReplaceCheck) = 0
    ProcedureReturn #False
  EndIf
  ;
  ; Ensure full AST exploration is completed before evaluating scope compatibility:
  If ExplorationState < #ExplorationState_Completed
    OnePassSetElements(#FinishCompletionNow)
  EndIf
  ;
  ; -------------------------------------------------------------------------
  ; 1. ANALYZE REPLACEMENT STRING & DETECT POTENTIAL AMBIGUITY
  ; -------------------------------------------------------------------------
  ReplaceText$ = GetGadgetText(FR_Gadgets\ReplaceString)
  ;
  If *LastSearch\NotAnElement = #False
    ; Determine whether a confirmation preview is required due to type/parenthesis suffixes or partial match modes:
    NeedConfirmation = Bool(*LastSearch\VarDotType$ <> "" Or
                            FindString(GetGadgetText(FR_Gadgets\SearchString), ".") Or
                            FindString(GetGadgetText(FR_Gadgets\SearchString), "(") Or
                            FindString(ReplaceText$, ".") Or
                            FindString(ReplaceText$, "(") Or
                            GetGadgetState(FR_Gadgets\MatchStartsWith) Or
                            GetGadgetState(FR_Gadgets\MatchContains))
    ;
    ; Strip trailing empty parentheses from the replacement string if the searched element is parenthesized,
    ; since MatchLen only covers the identifier name preceding '(' in the source code:
    If Right(ReplaceText$, 2) = "()" And
       (*LastSearch\HasParentheses = 1 Or FindString(GetGadgetText(FR_Gadgets\SearchString), "("))
      ReplaceText$ = Left(ReplaceText$, Len(ReplaceText$) - 2)
    EndIf
    ;
    ; Split replacement identifier and explicit '.DotType' if provided by the user:
    If FindString(ReplaceText$, ".") And Left(ReplaceText$, 1) <> "."
      ReplaceBase$    = StringField(ReplaceText$, 1, ".")
      ReplaceDotType$ = StringField(ReplaceText$, 2, ".")
    Else
      ReplaceBase$    = ReplaceText$
      ReplaceDotType$ = ""
    EndIf
  Else
    ReplaceBase$ = ReplaceText$
  EndIf
  ;
  ; -------------------------------------------------------------------------
  ; 2. PASS 1: COLLECT VALIDATED OCCURRENCES & BUILD PREVIEW SAMPLES
  ; -------------------------------------------------------------------------
  Repeat
    FoundPos = FindNextMatchInProject(*LastSearch, @FIPS, Reset, 0, @MatchLen, 0)
    Reset    = #False
    If FoundPos <= 0
      Break
    EndIf
    ;
    ActualMatchLen     = MatchLen
    ActualReplaceWith$ = ReplaceBase$
    ;
    If *LastSearch\NotAnElement = #False
      ; Locate the full identifier boundaries [TokenStart, TokenEnd] around the matched substring via direct pointer scan:
      *BufStart  = @FIPS\SearchString\s
      BufferLen  = Len(FIPS\SearchString\s)
      TokenStart = FoundPos
      *Char      = PosToChar(*BufStart, TokenStart)
      While TokenStart > 1
        *Char - SizeOf(Character)
        PrevChar = *Char\c
        If IsPBSeparator(PrevChar) = #False Or
           (TokenStart = FoundPos And (PrevChar = '*' Or PrevChar = '#'))
          TokenStart - 1
        Else
          Break
        EndIf
      Wend
      ;
      TokenEnd = FoundPos + MatchLen
      *Char    = PosToChar(*BufStart, TokenEnd)
      While TokenEnd <= BufferLen And IsPBSeparator(*Char\c) = #False
        TokenEnd + 1
        GoNextChar(*Char)
      Wend
      ;
      ; Inspect any immediate '.Type' or '\Field' suffix following the identifier:
      SuffixEnd = TokenEnd
      NextChar  = 0
      If TokenEnd <= BufferLen
        NextChar = *Char\c
      EndIf
      If NextChar = '.' Or NextChar = '\'
        SuffixEnd + 1
        GoNextChar(*Char)
        While SuffixEnd <= BufferLen And IsPBSeparator(*Char\c) = #False
          SuffixEnd + 1
          GoNextChar(*Char)
        Wend
        ;
        ; When the user specified a new '.DotType' in ReplaceString and the code occurrence has an explicit '.Type',
        ; extend the replacement span to substitute the old '.Type' with the new one; otherwise replace only the base name:
        If NextChar = '.' And TokenEnd = FoundPos + MatchLen And ReplaceDotType$ <> ""
          ActualMatchLen     = SuffixEnd - FoundPos
          ActualReplaceWith$ = ReplaceBase$ + "." + ReplaceDotType$
        EndIf
      EndIf
      ;
      ; Collect up to 8 distinct concrete before/after transformation examples for the confirmation dialog:
      If NeedConfirmation And PreviewCount < 8
        If SuffixEnd <= BufferLen And *Char\c = '('
          TrailParen$ = "()"
        Else
          TrailParen$ = ""
        EndIf
        ;
        PreviewLine$ = FastMid(FIPS\SearchString\s, TokenStart, SuffixEnd - TokenStart) + TrailParen$ +
                       " --> " +
                       FastMid(FIPS\SearchString\s, TokenStart, FoundPos - TokenStart) +
                       ActualReplaceWith$ +
                       FastMid(FIPS\SearchString\s, FoundPos + ActualMatchLen, SuffixEnd - (FoundPos + ActualMatchLen)) +
                       TrailParen$
        ;
        If FindString(#CR$ + PreviewList$, #CR$ + PreviewLine$ + #CR$) = 0
          PreviewList$ + PreviewLine$ + #CR$
          PreviewCount + 1
        EndIf
      EndIf
    EndIf
    ;
    ; Advance the iterator to the end of the replaced span to prevent overlapping hits:
    If ActualMatchLen > 1
      FIPS\PosInCode = FoundPos + ActualMatchLen - 1
    EndIf
    ;
    AddElement(MatchesToReplace())
    MatchesToReplace()\FileName    = FIPS\FileName$
    MatchesToReplace()\PosInCode   = FoundPos
    MatchesToReplace()\MatchLen    = ActualMatchLen
    MatchesToReplace()\ReplaceWith = ActualReplaceWith$
  ForEver
  ;
  If ListSize(MatchesToReplace()) = 0
    AlertInPBBWindow(GetTextFromCatalogPB("NoElementFound"))
    ProcedureReturn #False
  EndIf
  ;
  ; -------------------------------------------------------------------------
  ; 3. USER CONFIRMATION ON AMBIGUOUS OR QUALIFIED REPLACEMENTS
  ; -------------------------------------------------------------------------
  If NeedConfirmation And PreviewList$ <> ""
    PreviewList$ = GetGadgetText(FR_Gadgets\Scopes) + #CR$ + #CR$ + PreviewList$
    If PreviewCount >= 8
      PreviewList$ + "..." + #CR$
    EndIf
    If AlertInPBBWindow(PreviewList$, GetTextFromCatalogPB("FindReplaceTitle"), 1, 0, 0, 500, "", #AW_OKOrCancel_OKByDefault, FR_Gadgets\Window) <> #PB_MessageRequester_Yes
      ProcedureReturn #False
    EndIf
  EndIf
  ;
  ; Sort all collected occurrences by descending position so replacements within each file
  ; are strictly applied from the end of the buffer toward the beginning:
  SortStructuredList(MatchesToReplace(), #PB_Sort_Descending, OffsetOf(ReplaceMatchStruct\PosInCode), #PB_Integer)
  ;
  ; -------------------------------------------------------------------------
  ; 4. PASS 2: BACKUP PROMPT & REVERSE-ORDER IN-MEMORY REPLACEMENT
  ; -------------------------------------------------------------------------
  Result = AskAndPrepareProjectCopy(FilesToChange(), @TargetDir)
  If Result = #PB_MessageRequester_Cancel
    ProcedureReturn #False
  EndIf
  ;
  DisablePBBWindow()
  WGadget = OpenWaitWindow(GPBBGadgets\PBBWindow, "")
  ;
  ForEach FilesToChange()
    If FilesToChange()\SourceFile = PBBFicPrincipalPB$
      NewMainFile$ = FilesToChange()\TargetFile
    EndIf
    ;
    ; Check whether at least one validated match belongs to this file:
    FileModified = #False
    ForEach MatchesToReplace()
      If MatchesToReplace()\FileName = FilesToChange()\SourceFile
        If FileModified = #False
          If IsGadget(WGadget)
            SetGadgetText(WGadget, GetTextFromCatalogPB("FindReplaceTitle") + #CR$ + GetFilePart(FilesToChange()\TargetFile))
            UpdateWindow_(GadgetID(WGadget))
          EndIf
          RawDiskCode\s = FileToText(FilesToChange()\SourceFile, #PB_UTF8, 1)
          If RawDiskCode\s = ""
            Break
          EndIf
          FileEncoding   = BOM
          HasBOM         = FileHasBOM
          ModifiedCode\s = RawDiskCode\s
          FileModified   = #True
        EndIf
        ;
        ; Splice the context-adapted replacement string directly at the corresponding raw disk coordinates:
        ModifiedCode\s = FastLeft(ModifiedCode\s, FilteredPosToRawPos(@RawDiskCode, MatchesToReplace()\PosInCode) - 1) +
                         MatchesToReplace()\ReplaceWith +
                         FastMid(ModifiedCode\s, FilteredPosToRawPos(@RawDiskCode, MatchesToReplace()\PosInCode + MatchesToReplace()\MatchLen))
      EndIf
    Next
    ;
    If FileModified
      If TexteDansFichier(FilesToChange()\TargetFile, ModifiedCode\s, FileEncoding, 0, HasBOM)
        AnyFileModified = #True
      Else
        AlertInPBBWindow(GetTextFromCatalogPB("ErrorWithFile") + " " + FilesToChange()\TargetFile)
      EndIf
    EndIf
  Next
  ;
  If IsGadget(WGadget)
    DestroyParentWindow(WGadget)
  EndIf
  EnablePBBWindow()
  ;
  If AnyFileModified
    AlertInPBBWindow(GetTextFromCatalogPB("FilesAreUpdated"))
    If Result = #PB_MessageRequester_Yes And NewMainFile$ <> ""
      ; Switch PBBrowser to the newly created project copy:
      PBBUpdateMainFile(NewMainFile$, #True)
    ElseIf Result = #PB_MessageRequester_No
      ; Refresh modified buffers in memory and trigger panel updates:
      UpdateAllFilesInMemory()
    EndIf
    ;
    ; Check whether any occurrences of the old expression still remain in other scopes across the project:
    CopyStructure(*LastSearch, @TestRemainingSearch, SearchDetails)
    TestRemainingSearch\SearchPurifiedOnly   = #True
    TestRemainingSearch\StartingContextLimit = 0
    TestRemainingSearch\EndingContextLimit   = 0
    TestRemainingSearch\ElementFile$         = ""
    ;
    If FindNextMatchInProject(@TestRemainingSearch, @FIPS, #True, 0, 0, 0) = 0 And Trim(ReplaceBase$) <> ""
      ; The old expression has completely disappeared from the project; switch to the new expression:
      SetGadgetText(FR_Gadgets\SearchString, ReplaceText$)
    EndIf
    ; Trigger search to display either the new expression or the remaining scopes of the old one:
    GUI_FR_ExecuteSearch(*LastSearch)
  EndIf
  ;
  ProcedureReturn AnyFileModified
EndProcedure
;}
;
; ===============================================================================
; 3. WINDOW LIFECYCLE & EVENT HANDLING
; ===============================================================================
;- Window lifecycle and event handling
;{
#Menu_FR_Enter  = 63600
#Menu_FR_Escape = 63601
;
Procedure GUI_FR_SyncGadgetsWithSearch(*LastSearch.SearchDetails, RefreshScopes = #False)
  ;
  ; Role: Synchronizes the Find/Replace search field, match mode radio buttons,
  ;       comment/string checkboxes, and optionally rebuilds and selects the active scope
  ;       in the scopes list from '*LastSearch'.
  ;
  Shared SearchMenuChoices()
  Protected RawExpr$, PurifiedExpr$, DisplayExpr$
  Protected HasLeftJoker, HasRightJoker
  Protected TempSearch.SearchDetails
  ;
  If FR_Gadgets\Window < 0 Or IsWindow(FR_Gadgets\Window) = #False Or *LastSearch = 0
    ProcedureReturn
  EndIf
  ;
  RawExpr$      = *LastSearch\ElementName$
  PurifiedExpr$ = *LastSearch\ElementPurifiedName$
  If RawExpr$ <> ""
    HasLeftJoker  = Bool(Left(RawExpr$, Len(#PBBJoker$)) = #PBBJoker$ Or Left(PurifiedExpr$, Len(#PBBJoker$)) = #PBBJoker$)
    HasRightJoker = Bool(Right(RawExpr$, Len(#PBBJoker$)) = #PBBJoker$ Or Right(PurifiedExpr$, Len(#PBBJoker$)) = #PBBJoker$)
    ;
    If HasLeftJoker And HasRightJoker
      SetGadgetState(FR_Gadgets\MatchContains, 1)
      DisplayExpr$ = Trim(RawExpr$, #PBBJoker$)
    ElseIf HasRightJoker
      SetGadgetState(FR_Gadgets\MatchStartsWith, 1)
      DisplayExpr$ = RTrim(RawExpr$, #PBBJoker$)
    Else
      SetGadgetState(FR_Gadgets\MatchExact, 1)
      ; Keep a standalone leading wildcard (e.g., "…Init") intact when displayed in Exact mode:
      DisplayExpr$ = RTrim(RawExpr$, #PBBJoker$)
    EndIf
    ;
    If GetGadgetText(FR_Gadgets\SearchString) <> DisplayExpr$
      SetGadgetText(FR_Gadgets\SearchString, DisplayExpr$)
    EndIf
    SetGadgetState(FR_Gadgets\IgnoreComments, Bool(*LastSearch\SearchInCom = #False))
    SetGadgetState(FR_Gadgets\IgnoreStrings,  Bool(*LastSearch\SearchInQuotes = #False))
    If RefreshScopes
      If *LastSearch\NotAnElement
        ClearList(SearchMenuChoices())
        GUI_FR_RefreshScopes(*LastSearch, 0)
      Else
        ; Rebuild untruncated scope choices and select the active scope (or 'All') via GetInfosFromData():
        CopyStructure(*LastSearch, @TempSearch, SearchDetails)
        GetInfosFromData(@TempSearch, #PBBSearchExact)
        *LastSearch\MultipleResults = TempSearch\MultipleResults
      EndIf
    EndIf
  Else
    SetGadgetText(FR_Gadgets\SearchString, "")
    SetGadgetState(FR_Gadgets\IgnoreComments, 1)
    SetGadgetState(FR_Gadgets\IgnoreStrings,  1)
    ClearGadgetItems(FR_Gadgets\Scopes)
    GUI_FR_SetReplaceState()
  EndIf
EndProcedure
;
Procedure.i GUI_FR_OpenWindow(WSticky = #PB_Default, ParentWindowID = 0, *LastSearch.SearchDetails = 0)
  ;
  ; Role: Creates and displays the modeless Find/Replace dialog (or brings it to the front
  ;       and resynchronizes its state if already open).
  ;
  If FR_Gadgets\Window >= 0 And IsWindow(FR_Gadgets\Window)
    GUI_FR_SyncGadgetsWithSearch(*LastSearch, #True)
    SetActiveWindow(FR_Gadgets\Window)
    ProcedureReturn #True
  EndIf
  ;
  Protected WindowFlags   = #PB_Window_Invisible | #PB_Window_SystemMenu | #PB_Window_SizeGadget
  Protected ExtendedFlags = #WS_EX_TOOLWINDOW
  ;
  FR_Gadgets\Window = OpenWindowFromPrefCoordonnates(PBBrowserPrefile$, GetTextFromCatalogPB("FindReplaceTitle"), 500, 240,
                                                     WindowFlags, "_FR_Window", ParentWindowID, ExtendedFlags)
  ;
  If IsWindow(FR_Gadgets\Window)
    WindowBounds(FR_Gadgets\Window, 380, 220, #PB_Ignore, #PB_Ignore)
    StickyWindow(FR_Gadgets\Window, WSticky)
    BindEvent(#PB_Event_SizeWindow, @GUI_FR_ResizeGadgets(), FR_Gadgets\Window)
    ;
    FR_Gadgets\ButtonWidth    = 100
    FR_Gadgets\ButtonHeight   = 22
    FR_Gadgets\ExternalMargin = 12
    FR_Gadgets\InternalMargin = 5
    ;
    ; Search label, input field, and magnifying glass action button:
    FR_Gadgets\SearchLabel        = TextGadget  (#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\SearchString       = StringGadget(#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\SearchActionButton = ButtonGadget(#PB_Any, 1, 1, 1, 1, "")
    ;
    FR_Gadgets\ToolTip = InitBalloonToolTip(FR_Gadgets\Window)
    ;
    ; Match mode radio group (closed immediately below by IgnoreComments):
    FR_Gadgets\MatchExact         = OptionGadget(#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\MatchStartsWith    = OptionGadget(#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\MatchContains      = OptionGadget(#PB_Any, 1, 1, 1, 1, "")
    SetGadgetState(FR_Gadgets\MatchExact, 1)
    ;
    ; Comments and string literals are ignored by default:
    FR_Gadgets\IgnoreComments     = CheckBoxGadget(#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\IgnoreStrings      = CheckBoxGadget(#PB_Any, 1, 1, 1, 1, "")
    SetGadgetState(FR_Gadgets\IgnoreComments, 1)
    SetGadgetState(FR_Gadgets\IgnoreStrings,  1)
    ;
    FR_Gadgets\Scopes             = ListViewGadget(#PB_Any, 1, 1, 1, 1)
    ;
    ; Replacement activation checkbox, input field, action button, and scope note:
    FR_Gadgets\ReplaceCheck        = CheckBoxGadget(#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\ReplaceString       = StringGadget  (#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\ReplaceActionButton = ButtonGadget  (#PB_Any, 1, 1, 1, 1, "")
    FR_Gadgets\ReplaceNote         = TextGadget    (#PB_Any, 1, 1, 1, 1, "")
    ;
    GUI_FR_SetGadgetsTitles()
    GUI_FR_ColorizeAndSetFontsOfGadgets()
    GUI_FR_ResizeGadgets()
    GUI_FR_SetReplaceState()
    ;
    AddKeyboardShortcut(FR_Gadgets\Window, #PB_Shortcut_Return, #Menu_FR_Enter)
    AddKeyboardShortcut(FR_Gadgets\Window, #PB_Shortcut_Escape, #Menu_FR_Escape)
    ;
    GUI_FR_SyncGadgetsWithSearch(*LastSearch, #True)
    ;
    HideWindow(FR_Gadgets\Window, #False)
    SetActiveGadget(FR_Gadgets\SearchString)
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure.i GUI_FR_HandleEvent(EventID, *LastSearch.SearchDetails = 0)
  ;
  ; Role: Processes Find/Replace dialog events forwarded from the main application event loop.
  ; Output: Returns #True when the event belonged to and was handled by the Find/Replace dialog.
  ;
  Protected Quit
  Static FR_DisableFocusEvents
  ;
  If FR_Gadgets\Window < 0 Or IsWindow(FR_Gadgets\Window) = #False
    ProcedureReturn #False
  EndIf
  ;
  If EventWindow() <> FR_Gadgets\Window
    ProcedureReturn #False
  EndIf
  ;
  Select EventID
    Case #PB_Event_CloseWindow
      Quit = #True
      ;
    Case #PB_Event_Menu
      Select EventMenu()
        Case #Menu_FR_Enter
          GUI_FR_ExecuteSearch(*LastSearch)
          ProcedureReturn #True
        Case #Menu_FR_Escape
          Quit = #True
      EndSelect
      ;
    Case #PB_Event_Gadget
      Select EventGadget()
        Case FR_Gadgets\Scopes
          If EventType() = #PB_EventType_LeftClick Or EventType() = #PB_EventType_Change
            SelectSearchScopeChoice(GetGadgetState(FR_Gadgets\Scopes), *LastSearch, #False)
          EndIf
          ProcedureReturn #True
          ;
        Case FR_Gadgets\SearchActionButton
          GUI_FR_ExecuteSearch(*LastSearch)
          ProcedureReturn #True
          ;
        Case FR_Gadgets\ReplaceActionButton
          GUI_FR_ExecuteReplacement(*LastSearch)
          ProcedureReturn #True
          ;
        Case FR_Gadgets\ReplaceCheck
          If FR_DisableFocusEvents = #False
            GUI_FR_SetReplaceState()
            If GetGadgetState(FR_Gadgets\ReplaceCheck)
              SetActiveGadget(FR_Gadgets\ReplaceString)
            EndIf
          Else
            SetGadgetState(FR_Gadgets\ReplaceCheck, #False)
            FR_DisableFocusEvents = #False
          EndIf
          ProcedureReturn #True
          ;
        Case FR_Gadgets\SearchString
          If EventType() = #PB_EventType_Focus
            PostMessage_(GadgetID(FR_Gadgets\SearchString), #EM_SETSEL, 0, -1)
          ElseIf EventType() = #PB_EventType_Change
            If *LastSearch = 0 Or Trim(GetGadgetText(FR_Gadgets\SearchString), #PBBJoker$) <> Trim(*LastSearch\ElementName$, #PBBJoker$)
              ; Invalidate scope list and disable replacement until the modified expression is searched:
              ClearGadgetItems(FR_Gadgets\Scopes)
              GUI_FR_SetReplaceState()
            EndIf
          EndIf
          ProcedureReturn #True
          ;
        Case FR_Gadgets\ReplaceString
          If EventType() = #PB_EventType_Focus
            If FR_DisableFocusEvents = #False
              If GetGadgetState(FR_Gadgets\ReplaceCheck) = 0
                SetGadgetState(FR_Gadgets\ReplaceCheck, #True)
                GUI_FR_SetReplaceState()
                SetActiveGadget(FR_Gadgets\ReplaceString)
                FR_DisableFocusEvents = #True
              EndIf
            Else
              FR_DisableFocusEvents = #False
            EndIf
            PostMessage_(GadgetID(FR_Gadgets\ReplaceString), #EM_SETSEL, 0, -1)
          ElseIf EventType() = #PB_EventType_LostFocus
            If Trim(GetGadgetText(FR_Gadgets\ReplaceString)) = "" And FR_DisableFocusEvents = #False
              SetGadgetState(FR_Gadgets\ReplaceCheck, #False)
              GUI_FR_SetReplaceState()
              If GetActiveGadget() = FR_Gadgets\ReplaceCheck
                FR_DisableFocusEvents = #True
              EndIf
            EndIf
          EndIf
          ProcedureReturn #True
          ;
        Case FR_Gadgets\MatchExact, FR_Gadgets\MatchStartsWith, FR_Gadgets\MatchContains,
             FR_Gadgets\IgnoreComments, FR_Gadgets\IgnoreStrings
          GUI_FR_ExecuteSearch(*LastSearch)
          ProcedureReturn #True
      EndSelect
  EndSelect
  ;
  If Quit
    UnbindEvent(#PB_Event_SizeWindow, @GUI_FR_ResizeGadgets(), FR_Gadgets\Window)
    RecordWindowDimInPref(PBBrowserPrefile$, FR_Gadgets\Window, "_FR_Window")
    CloseWindow(FR_Gadgets\Window)
    FR_Gadgets\Window = #PB_Ignore
  EndIf
  ;
  ; Every event targeting FR_Gadgets\Window is owned by this dialog and must not leak to PBBWindow:
  ProcedureReturn #True
EndProcedure
;}
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 250
; FirstLine = 230
; Folding = ---
; EnableXP
; DPIAware