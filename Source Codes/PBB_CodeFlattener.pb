; *******************************************************************************
;
;                           PBBrowser_FlattenTool.pb
;
;         Standalone Single-File .pb Exporter & Unused-Code Pruner Tool
;
; *******************************************************************************
;
; ===============================================================================
; 1. FLATTEN EXPORTER CONSTANTS & STRUCTURES
; ===============================================================================
;
EnumerationBinary FT_OptionFlags
  #FT_AddFileBanners  = 1   ; Insert '; === File.pb ===' banners when switching files (unless EraseComments is set).
  #FT_PruneDeadZones  = 2   ; Strip inactive CompilerIf / CompilerSelect intervals recorded in DeadZones().
  #FT_PruneUnused     = 4   ; Strip unused top-level elements (\Used = #False) and their orphan 'Declare' headers.
  #FT_EraseComments   = 8   ; Strip all code comments via DeleteCommentsAndSpacesFromPBCode().
  #FT_CleanSpaces     = 16  ; Strip indentation, extra spaces, and blank lines via DeleteCommentsAndSpacesFromPBCode().
  #FT_RewriteBinPaths = 32  ; Rewrite IncludeBinary and UseIcon paths relative to the exported .pb file directory.
  #FT_OutClipboard    = 64  ; Copy the generated standalone .pb code to the Windows clipboard.
  #FT_OutSaveFile     = 128 ; Write the generated standalone .pb code to disk.
  #FT_OutOpenInIDE    = 256 ; Open the saved .pb file directly in the PureBasic IDE via PureBasicProgAddr$.
EndEnumeration
;
Structure FT_ConfigStruct
  Flags.i
  OutputFile.s
EndStructure
;
Structure FT_GadgetsStruct
  Window.i
  ;
  ; Contiguous gadget IDs ordered for pointer-based font and theme styling:
  LblSecContent.i
  ChkFileBanners.i
  ChkPruneDeadZones.i
  ChkPruneUnused.i
  ChkEraseComments.i
  ChkCleanSpaces.i
  ChkRewriteBinPaths.i
  ;
  LblSecOutput.i
  ChkClipboard.i
  ChkSaveFile.i
  StrOutFile.i
  BtnBrowse.i
  ChkOpenInIDE.i
  ;
  BtnCancel.i
  BtnOK.i
EndStructure
;
Structure FT_SkipRange
  StartPos.i
  EndPos.i
EndStructure
;
; ===============================================================================
; 2. DEFAULT PATH & OPTIONS DIALOG MANAGEMENT
; ===============================================================================
;- Default path and options dialog helpers
;{
Procedure.s FT_BuildDefaultOutputFile()
  ;
  ; Role: Computes the default output path ('<ProjectOrMain>_Standalone.pb') next to the active project or main file.
  ;
  Protected BaseFile$ = PBBCurrentProject$
  ;
  If BaseFile$ = ""
    BaseFile$ = PBBFicPrincipalPB$
  EndIf
  If BaseFile$ = ""
    ProcedureReturn ""
  EndIf
  ProcedureReturn GetPathPart(BaseFile$) + FilePartSsExt(BaseFile$) + "_Standalone.pb"
EndProcedure
;
Procedure FT_ApplyTheme(*GD.FT_GadgetsStruct)
  ;
  ; Role: Applies PBBrowser's active color preset and fonts to the dialog and all its child gadgets.
  ;
  Protected *GadgetAddress.Integer, GadgetID
  Protected GadgetList$ = ""
  ;
  *GadgetAddress = @*GD\LblSecContent
  While *GadgetAddress <= @*GD\BtnOK
    GadgetID = *GadgetAddress\i
    If IsGadget(GadgetID)
      GadgetList$ + Str(GadgetID) + ","
    EndIf
    *GadgetAddress + SizeOf(Integer)
  Wend
  ;
  SetFontAndGadgetsColors(*GD\Window, InterfaceColorPresets(), GadgetList$)
  ;
  *GadgetAddress = @*GD\LblSecContent
  While *GadgetAddress <= @*GD\BtnOK
    GadgetID = *GadgetAddress\i
    If IsGadget(GadgetID)
      If GadgetID = *GD\LblSecContent Or GadgetID = *GD\LblSecOutput
        SetGadgetColorEx(GadgetID, #PB_Gadget_FrontColor, GetPBBNumParameter("PBBTitleColor"))
        SetGadgetFont(GadgetID, PBBTitleFont)
      Else
        SetGadgetFont(GadgetID, PBBAllGadgetsFont)
      EndIf
    EndIf
    *GadgetAddress + SizeOf(Integer)
  Wend
EndProcedure
;
Procedure.i FT_UpdateWindowState(*GD.FT_GadgetsStruct)
  ;
  ; Role: Synchronizes dependent checkboxes (disabling file banners when all comments are erased,
  ;       and toggling file-output controls) and returns #True when validation is allowed.
  ;
  Protected SaveEnabled, EraseCom, CanValidate
  ;
  EraseCom    = Bool(GetGadgetState(*GD\ChkEraseComments) = #PB_Checkbox_Checked)
  SaveEnabled = Bool(GetGadgetState(*GD\ChkSaveFile) = #PB_Checkbox_Checked)
  ;
  DisableGadget(*GD\ChkFileBanners, EraseCom)
  DisableGadget(*GD\StrOutFile, Bool(SaveEnabled = #False))
  DisableGadget(*GD\BtnBrowse, Bool(SaveEnabled = #False))
  DisableGadget(*GD\ChkOpenInIDE, Bool(SaveEnabled = #False Or FileSize(PureBasicProgAddr$) < 2))
  ;
  ; Path rewriting requires a target output directory; keep enabled whenever StrOutFile holds a path:
  DisableGadget(*GD\ChkRewriteBinPaths, Bool(Trim(GetGadgetText(*GD\StrOutFile)) = ""))
  ;
  CanValidate = Bool(GetGadgetState(*GD\ChkClipboard) = #PB_Checkbox_Checked Or
                     (SaveEnabled And Trim(GetGadgetText(*GD\StrOutFile)) <> ""))
  DisableGadget(*GD\BtnOK, Bool(CanValidate = #False))
  ;
  ProcedureReturn CanValidate
EndProcedure
;
Procedure.i FlattenToolOptionsWindow(*Config.FT_ConfigStruct)
  ;
  ; Role: Displays the themed modal configuration dialog for the Standalone Single-File Exporter.
  ; Output: Populates '*Config' and returns #True if validated, or #False if cancelled.
  ;
  Protected GD.FT_GadgetsStruct
  Protected WWidth = 540, WHeight = 410, Margins = 18, LineHeight = 22, ButtonHeight = 28
  Protected OX, OY, ParentWindowID, YPos, InnerX, FullInnerW
  Protected Event, CanValidate, Validated = #False
  Protected DefaultOutFile$, PickedFile$
  ;
  InnerX         = Margins + 8
  FullInnerW     = WWidth - InnerX - Margins
  DefaultOutFile$= FT_BuildDefaultOutputFile()
  ParentWindowID = ComputeWinOrigins(@OX, @OY, WWidth, WHeight, GPBBGadgets\PBBWindow)
  ;
  DisablePBBWindow()
  GD\Window = OpenWindow(#PB_Any, OX, OY, WWidth, WHeight, GetTextFromCatalogPB("FlattenToolTitle"),
                         #PB_Window_SystemMenu | #PB_Window_Invisible, ParentWindowID)
  If GD\Window
    ApplyDarkModeToWindow(GD\Window)
    StickyWindow(GD\Window, 1)
    ;
    ; --- Section 1: Code Content and Cleanup ---
    YPos                   = 14
    GD\LblSecContent       = TextGadget(#PB_Any, Margins, YPos, WWidth - Margins * 2, 20, GetTextFromCatalogPB("FT_SecContent"))
    YPos + 24
    GD\ChkFileBanners      = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("FT_AddFileBanners"))
    SetGadgetState(GD\ChkFileBanners, #PB_Checkbox_Checked)
    YPos + 26
    GD\ChkPruneDeadZones   = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("FT_PruneDeadZones"))
    SetGadgetState(GD\ChkPruneDeadZones, #PB_Checkbox_Checked)
    YPos + 26
    GD\ChkPruneUnused      = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("FT_PruneUnused"))
    YPos + 26
    GD\ChkEraseComments    = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("FT_EraseComments"))
    YPos + 26
    GD\ChkCleanSpaces      = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("FT_CleanSpaces"))
    YPos + 26
    GD\ChkRewriteBinPaths  = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("FT_RewriteBinPaths"))
    SetGadgetState(GD\ChkRewriteBinPaths, #PB_Checkbox_Checked)
    ;
    ; --- Section 2: Output Destination ---
    YPos                   = YPos + 34
    GD\LblSecOutput        = TextGadget(#PB_Any, Margins, YPos, WWidth - Margins * 2, 20, GetTextFromCatalogPB("FT_SecOutput"))
    YPos + 24
    GD\ChkClipboard        = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_CopyClipboard"))
    YPos + 26
    GD\ChkSaveFile         = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_SaveToFile"))
    SetGadgetState(GD\ChkSaveFile, #PB_Checkbox_Checked)
    YPos + 26
    GD\StrOutFile          = StringGadget(#PB_Any, InnerX + 18, YPos, FullInnerW - 54, 24, DefaultOutFile$)
    GD\BtnBrowse           = ButtonGadget(#PB_Any, InnerX + FullInnerW - 30, YPos, 30, 24, "...")
    YPos + 28
    GD\ChkOpenInIDE        = CheckBoxGadget(#PB_Any, InnerX + 18, YPos, FullInnerW - 18, LineHeight, GetTextFromCatalogPB("FT_OpenInIDE"))
    SetGadgetState(GD\ChkOpenInIDE, Bool(FileSize(PureBasicProgAddr$) >= 2))
    ;
    ; --- Bottom Action Buttons ---
    YPos         = WHeight - ButtonHeight - 14
    GD\BtnCancel = ButtonGadget(#PB_Any, WWidth - Margins - 210, YPos, 100, ButtonHeight, GetTextFromCatalogPB("Cancel"))
    GD\BtnOK     = ButtonGadget(#PB_Any, WWidth - Margins - 100, YPos, 100, ButtonHeight, GetTextFromCatalogPB("OK"), #PB_Button_Default)
    ;
    FT_ApplyTheme(@GD)
    CanValidate = FT_UpdateWindowState(@GD)
    AddKeyboardShortcut(GD\Window, #PB_Shortcut_Return, 100)
    AddKeyboardShortcut(GD\Window, #PB_Shortcut_Escape, 101)
    ClearKeyboardBuffer(GD\Window)
    HideWindow(GD\Window, #False)
    ;
    Repeat
      Event = WaitWindowEvent()
      If EventWindow() <> GD\Window
        SetActiveWindow(GD\Window)
      Else
        Select Event
          Case #PB_Event_CloseWindow
            Break
            ;
          Case #PB_Event_Menu
            Select EventMenu()
              Case 100
                If CanValidate : Validated = #True : EndIf
              Case 101
                Break
            EndSelect
            ;
          Case #PB_Event_Gadget
            Select EventGadget()
              Case GD\BtnOK
                If CanValidate : Validated = #True : EndIf
              Case GD\BtnCancel
                Break
              Case GD\BtnBrowse
                PickedFile$ = SaveFileRequester(GetTextFromCatalogPB("FlattenToolTitle"),
                                                GetGadgetText(GD\StrOutFile),
                                                GetTextFromCatalogPB("FT_FileFilter"), 0)
                If PickedFile$ <> ""
                  If GetExtensionPart(PickedFile$) = "" And SelectedFilePattern() = 0
                    PickedFile$ + ".pb"
                  EndIf
                  SetGadgetText(GD\StrOutFile, PickedFile$)
                EndIf
            EndSelect
            CanValidate = FT_UpdateWindowState(@GD)
        EndSelect
        ;
        If Validated
          ClearStructure(*Config, FT_ConfigStruct)
          *Config\OutputFile = Trim(GetGadgetText(GD\StrOutFile))
          If GetGadgetState(GD\ChkFileBanners) = #PB_Checkbox_Checked And
             GetGadgetState(GD\ChkEraseComments) = #PB_Checkbox_Unchecked
            *Config\Flags | #FT_AddFileBanners
          EndIf
          If GetGadgetState(GD\ChkPruneDeadZones) = #PB_Checkbox_Checked  : *Config\Flags | #FT_PruneDeadZones  : EndIf
          If GetGadgetState(GD\ChkPruneUnused) = #PB_Checkbox_Checked     : *Config\Flags | #FT_PruneUnused     : EndIf
          If GetGadgetState(GD\ChkEraseComments) = #PB_Checkbox_Checked   : *Config\Flags | #FT_EraseComments   : EndIf
          If GetGadgetState(GD\ChkCleanSpaces) = #PB_Checkbox_Checked     : *Config\Flags | #FT_CleanSpaces     : EndIf
          If GetGadgetState(GD\ChkRewriteBinPaths) = #PB_Checkbox_Checked And *Config\OutputFile <> ""
            *Config\Flags | #FT_RewriteBinPaths
          EndIf
          If GetGadgetState(GD\ChkClipboard) = #PB_Checkbox_Checked       : *Config\Flags | #FT_OutClipboard    : EndIf
          If GetGadgetState(GD\ChkSaveFile) = #PB_Checkbox_Checked
            *Config\Flags | #FT_OutSaveFile
            If GetGadgetState(GD\ChkOpenInIDE) = #PB_Checkbox_Checked
              *Config\Flags | #FT_OutOpenInIDE
            EndIf
          EndIf
          Break
        EndIf
      EndIf ; EndIf EventWindow() = GD\Window
    ForEver
    ;
    ClearKeyboardBuffer(GD\Window)
    CloseWindow(GD\Window)
  EndIf
  ;
  EnablePBBWindow()
  ProcedureReturn Validated
EndProcedure
;}
;
; ===============================================================================
; 3. SKIP-INTERVAL COLLECTION & SINGLE-PASS SEGMENT REWRITER
; ===============================================================================
;- Skip-interval builder and segment stream rewriter
;{
Procedure FT_AddExpandedSkipRange(*Buf.String, IDELimitPos, StartPos, EndPos, List RawRanges.FT_SkipRange())
  ;
  ; Role: Expands [StartPos, EndPos] to clean statement boundaries while preserving required ':' or #CR
  ;       separators between any surviving statements on the left and right of the skipped span.
  ;
  Protected *Base, *LeftScan.Character, *RightScan.Character, *Limit.Character
  Protected AtLineStart, AtLineEnd
  ;
  If *Buf = 0 Or StartPos < 1 Or EndPos < StartPos Or StartPos > IDELimitPos
    ProcedureReturn
  EndIf
  If EndPos > IDELimitPos : EndPos = IDELimitPos : EndIf
  ;
  *Base  = @*Buf\s
  *Limit = PosToChar(*Base, IDELimitPos) + SizeOf(Character)
  ;
  ; 1. Inspect left context (line start vs ':' after an upstream statement on the same line):
  *LeftScan = PosToChar(*Base, StartPos) - SizeOf(Character)
  SkipSpacesBackward(*LeftScan, *Base)
  If *LeftScan < *Base : AtLineStart = #True : Else : AtLineStart = Bool(*LeftScan\c = #CR) : EndIf
  ;
  ; 2. Inspect right context (EOL vs downstream statement on the same line):
  *RightScan = PosToChar(*Base, EndPos) + SizeOf(Character)
  SkipSpaces(*RightScan)
  If *RightScan < *Limit And *RightScan\c = ';'
    GoToEOL_AndExtractComment(@*RightScan)
  EndIf
  AtLineEnd = Bool(*RightScan >= *Limit Or *RightScan\c = #CR Or *RightScan\c = 0)
  ;
  ; 3. Adjust StartPos and EndPos so surviving neighbors remain separated by ':' or #CR:
  If AtLineStart
    StartPos = ComputeLenFromChars(*Base, *LeftScan) + 2
    If AtLineEnd
      ; Standalone line(s): consume trailing comment and the terminating #CR
      EndPos = ComputeLenFromChars(*Base, *RightScan) + Bool(*RightScan < *Limit And *RightScan\c = #CR)
    ElseIf *RightScan\c = ':' And PeekC(*RightScan + SizeOf(Character)) <> ':'
      ; Span starts a multi-statement line ('Span : StmtB'): consume trailing ':' and spaces
      GoNextChar(*RightScan)
      SkipSpaces(*RightScan)
      EndPos = ComputeLenFromChars(*Base, *RightScan)
    Else
      ; Multi-line DeadZone ending before an indented directive: preserve the directive's leading spaces
      *RightScan = PosToChar(*Base, EndPos)
      SkipSpacesBackward(*RightScan, PosToChar(*Base, StartPos))
      If *RightScan\c = #CR : EndPos = ComputeLenFromChars(*Base, *RightScan) + 1 : EndIf
    EndIf
  ElseIf AtLineEnd
    ; Trailing span after ':' ('StmtA : Span'): strip preceding ':' + spaces, but keep the line's #CR
    If *LeftScan\c = ':' And (*LeftScan = *Base Or PeekC(*LeftScan - SizeOf(Character)) <> ':')
      *LeftScan - SizeOf(Character)
      SkipSpacesBackward(*LeftScan, *Base)
      StartPos = ComputeLenFromChars(*Base, *LeftScan) + 2
    EndIf
    EndPos = ComputeLenFromChars(*Base, *RightScan)
  Else
    ; Middle span ('StmtA : Span : StmtB' or 'CompilerIf 0 : Dead : CompilerEndIf'):
    ; keep the left ':' intact (plus one space if present) and consume any trailing ':' + spaces
    If PeekC(PosToChar(*Base, StartPos)) = #Space : StartPos + 1 : EndIf
    If *RightScan\c = ':' And PeekC(*RightScan + SizeOf(Character)) <> ':'
      GoNextChar(*RightScan)
      SkipSpaces(*RightScan)
    EndIf
    EndPos = ComputeLenFromChars(*Base, *RightScan)
  EndIf
  ;
  If EndPos >= StartPos
    AddElement(RawRanges())
    RawRanges()\StartPos = StartPos
    RawRanges()\EndPos   = EndPos
  EndIf
EndProcedure
;
Procedure.i FT_CanPruneTopLevelVarDecl(Decl$, FileName$, ModuleID)
  ;
  ; Role: Verifies that a top-level variable declaration ('Global', 'Threaded', or 'Define')
  ;       is not a 'ReDim' and that every sibling variable on a comma-separated declaration line
  ;       is also unused before allowing the statement to be stripped.
  ;
  Protected *Char.Character
  Protected Args$, Item$, LHS$, Tok$
  Protected Idx, HasPar, VarID, CanPrune = #True
  ;
  If Left(LCase(Trim(Decl$)), 5) = "redim"
    ProcedureReturn #False
  EndIf
  Args$ = ParseArguments(Decl$)
  If FindString(Args$, #TAB$) = 0
    ProcedureReturn #True
  EndIf
  ;
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
  CompilerEndIf
  PushListPosition(ElementList())
  Repeat
    Idx + 1
    Item$ = Trim(StringField(Args$, Idx, #TAB$))
    If Item$ = "" : Break : EndIf
    LHS$   = Trim(StringField(Item$, 1, "="))
    HasPar = Bool(FindString(LHS$, "(") > 0)
    LHS$   = Trim(StringField(LHS$, 1, "("))
    *Char  = @LHS$
    Repeat
      SkipSpaces(*Char)
      Tok$ = ExtractOrSkipNextToken(@LHS$, @*Char)
      ExtractOptionalDotType(@LHS$, @*Char)
    Until Tok$ = "" Or FindString(",global,threaded,define,newlist,newmap,dim,redim,", "," + LCase(Tok$) + ",") = 0
    If Tok$ = ""
      CanPrune = #False
      Break
    EndIf
    VarID = FindElementInList(LCase(Tok$), #PBBVariable, 0, FileName$, -1, HasPar, "", 0, ModuleID)
    If VarID = 0 Or ElementList()\Used
      CanPrune = #False
      Break
    EndIf
  ForEver
  PopListPosition(ElementList())
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
  CompilerEndIf
  ;
  ProcedureReturn CanPrune
EndProcedure
;
Procedure FT_BuildFileSkipRanges(FileName$, Flags, List MergedRanges.FT_SkipRange())
  ;
  ; Role: Builds a flat, sorted, non-overlapping list of 1-based character intervals to omit
  ;       from 'FileName$' (combining inactive DeadZones() and unused AST elements + orphan Declares).
  ;
  Protected *Buf.String = GetPointedCodeFromFile(FileName$, #False)
  Protected *SavedElem
  Protected IDELimitPos, ElemType, ModID, CanPruneElem
  Protected NewList RawRanges.FT_SkipRange()
  Protected NewList CEDList.CED_Structure()
  ;
  ClearList(MergedRanges())
  If *Buf = 0 Or *Buf\s = ""
    ProcedureReturn
  EndIf
  IDELimitPos = FindIDENotations(*Buf) - 1
  If IDELimitPos < 1
    ProcedureReturn
  EndIf
  ;
  ; 1. Collect inactive CompilerIf / CompilerSelect zones when requested:
  If (Flags & #FT_PruneDeadZones) And IsFileInMemory(FileName$)
    ForEach PBB_FileListing()\DeadZones()
      FT_AddExpandedSkipRange(*Buf, IDELimitPos, PBB_FileListing()\DeadZones()\StartingPos,
                              PBB_FileListing()\DeadZones()\EndingPos, RawRanges())
    Next
  EndIf
  ;
  ; 2. Collect unused top-level elements and any secondary 'Declare' / 'Module' / top-level 'Define' blocks:
  If Flags & #FT_PruneUnused
    *SavedElem = @ElementList()
    ForEach ElementList()
      If ElementList()\Used = #False And ElementList()\StartingPos > 0
        ElemType     = ElementList()\Type
        ModID        = ElementList()\EnclosingModuleID
        CanPruneElem = #False
        Select ElemType
          Case #PBBProcedure, #PBBMacro, #PBBStructure, #PBBInterface, #PBBEnumeration, #PBBModule, #PBBImport
            CanPruneElem = #True
          Case #PBBLabel
            CanPruneElem = Bool(ModID <= 0)
          Case #PBBConstant
            ; Only prune standalone constants outside Enumerations
            CanPruneElem = Bool(ElementList()\LocalVarContainer <= 0 And ElementList()\Continued = "")
          Case #PBBVariable
            CanPruneElem = Bool(ElementList()\VariableSpecie = "Global" Or ElementList()\VariableSpecie = "Threaded" Or
                                ElementList()\VariableSpecie = "Define" Or ElementList()\VariableSpecie = "Shared")
        EndSelect
        ;
        If CanPruneElem
          If ElementList()\FileName = FileName$ And
             (ElemType <> #PBBVariable Or
              (ElementList()\LocalVarContainer <= 0 And FT_CanPruneTopLevelVarDecl(ElementList()\Declaration, FileName$, ModID)))
            FT_AddExpandedSkipRange(*Buf, IDELimitPos, ElementList()\StartingPos, ElementList()\EndingPos, RawRanges())
          EndIf
          If ElementList()\Continued <> "" And
             (ElemType = #PBBProcedure Or ElemType = #PBBEnumeration Or ElemType = #PBBModule Or
              ElemType = #PBBImport Or ElemType = #PBBVariable)
            CED_StringToList(ElementList()\Continued, CEDList())
            ForEach CEDList()
              If CEDList()\Filename = FileName$ And
                 (ElemType <> #PBBVariable Or
                  (CEDList()\DirectContainer <= 0 And FT_CanPruneTopLevelVarDecl(CEDList()\Declaration, FileName$, ModID)))
                FT_AddExpandedSkipRange(*Buf, IDELimitPos, CEDList()\StartingPos, CEDList()\EndingPos, RawRanges())
              EndIf
            Next
          EndIf
        EndIf
      EndIf
    Next
    If *SavedElem
      ChangeCurrentElement(ElementList(), *SavedElem)
    EndIf
  EndIf
  ;
  ; 3. Sort and coalesce overlapping or adjacent intervals for O(1) streaming jumps:
  If ListSize(RawRanges()) > 0
    SortStructuredList(RawRanges(), #PB_Sort_Ascending, OffsetOf(FT_SkipRange\StartPos), #PB_Integer)
    ForEach RawRanges()
      If ListSize(MergedRanges()) = 0 Or RawRanges()\StartPos > MergedRanges()\EndPos + 1
        AddElement(MergedRanges())
        MergedRanges()\StartPos = RawRanges()\StartPos
        MergedRanges()\EndPos   = RawRanges()\EndPos
      ElseIf RawRanges()\EndPos > MergedRanges()\EndPos
        MergedRanges()\EndPos = RawRanges()\EndPos
      EndIf
    Next
  EndIf
EndProcedure
;
Procedure.s FT_ProcessSegment(FileName$, StartPos, EndPos, Flags, OutRefFile$, Map ActiveIncludePaths.s(), List SkipRanges.FT_SkipRange())
  ;
  ; Role: Streams slice [StartPos, EndPos] of 'FileName$' in a single pointer pass:
  ;       - Jumps over merged SkipRanges() (DeadZones and unused elements) in O(1),
  ;       - Strips IncludeFile, XIncludeFile, and IncludePath statements while tracking ActiveIncludePaths(),
  ;       - Rewrites IncludeBinary literal/constant paths relative to 'OutRefFile$' when #FT_RewriteBinPaths is set,
  ;       - Evaluates #PB_Compiler_IsMainFile / #PB_Compiler_IsIncludeFile inside included files so guard semantics hold.
  ;
  Protected *Buf.String = GetPointedCodeFromFile(FileName$, #False)
  Protected *StringStartingAdr, *Char.Character, *Limit.Character, *CopyStart.Character, *StmtStart.Character, *TokStart.Character
  Protected IDELimitPos, CurPos, MacroDepth, HasSkipRange, IsStartOfStmt, IsIncludedFile, FirstLow
  Protected TokLCase$, RawExpr$, EvalVal$, FullBinPath$, RewrittenBinPath$, CurIncPath$, OutChunk$
  ;
  If *Buf = 0 Or *Buf\s = "" Or StartPos < 1 Or EndPos < StartPos
    ProcedureReturn ""
  EndIf
  ;
  IDELimitPos = FindIDENotations(*Buf) - 1
  If EndPos > IDELimitPos
    EndPos = IDELimitPos
  EndIf
  If StartPos > EndPos
    ProcedureReturn ""
  EndIf
  ;
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
  CompilerEndIf
  ;
  IsIncludedFile = Bool(LCase(FileName$) <> LCase(PBBFicPrincipalPB$))
  If FindMapElement(ActiveIncludePaths(), LCase(FileName$))
    CurIncPath$ = ActiveIncludePaths()
  EndIf
  ;
  HasSkipRange = Bool(FirstElement(SkipRanges()) <> 0)
  While HasSkipRange And SkipRanges()\EndPos < StartPos
    HasSkipRange = Bool(NextElement(SkipRanges()) <> 0)
  Wend
  ;
  *StringStartingAdr = @*Buf\s
  *Char              = PosToChar(*StringStartingAdr, StartPos)
  *Limit             = PosToChar(*StringStartingAdr, EndPos) + SizeOf(Character)
  *CopyStart         = *Char
  IsStartOfStmt      = #True
  ;
  While *Char < *Limit And *Char\c <> 0
    CurPos = CharToPos(*Char)
    While HasSkipRange And SkipRanges()\EndPos < CurPos
      HasSkipRange = Bool(NextElement(SkipRanges()) <> 0)
    Wend
    If HasSkipRange And CurPos >= SkipRanges()\StartPos
      If *Char > *CopyStart
        OutChunk$ + PeekS(*CopyStart, ComputeLenFromChars(*CopyStart, *Char))
      EndIf
      *Char = PosToChar(*StringStartingAdr, SkipRanges()\EndPos) + SizeOf(Character)
      If *Char > *Limit : *Char = *Limit : EndIf
      *CopyStart    = *Char
      IsStartOfStmt = #True
      HasSkipRange  = Bool(NextElement(SkipRanges()) <> 0)
      Continue
    EndIf
    ;
    If *Char\c = #Space Or *Char\c = #NonBreakableSpace
      SkipSpaces(*Char)
      Continue
    ElseIf *Char\c = #CR
      GoNextChar(*Char)
      IsStartOfStmt = #True
      Continue
    ElseIf *Char\c = ':'
      If IsModuleUsage(@*Char)
        IsStartOfStmt = #False
      Else
        GoNextChar(*Char)
        IsStartOfStmt = #True
      EndIf
      Continue
    ElseIf *Char\c = ';'
      GoToEOL_AndExtractComment(@*Char)
      Continue
    ElseIf IsStringOrCharLiteralStart(*Char)
      ExtractOrSkipNextString(@*Char, *Char\c, #True)
      IsStartOfStmt = #False
      Continue
    EndIf
    ;
    FirstLow = *Char\c | 32
    If IsStartOfStmt And (FirstLow = 'm' Or FirstLow = 'e' Or FirstLow = 'i' Or FirstLow = 'x')
      *StmtStart = *Char
      If *StmtStart > *CopyStart
        While *StmtStart > *CopyStart And
              (PeekC(*StmtStart - SizeOf(Character)) = #Space Or PeekC(*StmtStart - SizeOf(Character)) = #NonBreakableSpace)
          *StmtStart - SizeOf(Character)
        Wend
        If *StmtStart > *CopyStart And PeekC(*StmtStart - SizeOf(Character)) <> #CR
          *StmtStart = *Char
        EndIf
      EndIf
      ;
      TokLCase$     = LCase(ExtractOrSkipNextToken(*StringStartingAdr, @*Char))
      IsStartOfStmt = #False
      If TokLCase$ = "macro"
        MacroDepth + 1
        Continue
      ElseIf TokLCase$ = "endmacro"
        If MacroDepth > 0 : MacroDepth - 1 : EndIf
        Continue
      ElseIf MacroDepth = 0 And (TokLCase$ = "includepath" Or TokLCase$ = "includefile" Or
                                 TokLCase$ = "xincludefile" Or TokLCase$ = "includebinary")
        RawExpr$ = ExtractCompilerDirectiveExpr(*StringStartingAdr, @*Char, #True)
        EvalVal$ = EvalConstExpr(RawExpr$, FileName$, CurPos, -1, #False, #False)
        If Left(EvalVal$, 1) = #DOUBLEQUOTE$
          EvalVal$ = Mid(EvalVal$, 2, Len(EvalVal$) - 2)
        Else
          EvalVal$ = ""
        EndIf
        ;
        If TokLCase$ = "includepath"
          If EvalVal$ <> "" And Right(EvalVal$, 1) <> "\" And Right(EvalVal$, 1) <> "/"
            EvalVal$ + "\"
          EndIf
          CurIncPath$ = EvalVal$
          ActiveIncludePaths(LCase(FileName$)) = CurIncPath$
        EndIf
        ;
        If TokLCase$ = "includebinary"
          If EvalVal$ <> "" And ((Flags & #FT_RewriteBinPaths) Or CurIncPath$ <> "")
            FullBinPath$ = RelativeAddressToFullAddress(CurIncPath$ + EvalVal$, FileName$)
            If (Flags & #FT_RewriteBinPaths) And OutRefFile$ <> ""
              RewrittenBinPath$ = FullAddressToRelativeAddress(FullBinPath$, OutRefFile$)
            Else
              RewrittenBinPath$ = FullBinPath$
            EndIf
            If *Char > *CopyStart
              OutChunk$ + PeekS(*CopyStart, ComputeLenFromChars(*CopyStart, *StmtStart))
            EndIf
            If *StmtStart < *Char And PeekC(*StmtStart) = #Space
              OutChunk$ + PeekS(*StmtStart, ComputeLenFromChars(*StmtStart, PosToChar(*StringStartingAdr, CurPos)))
            EndIf
            OutChunk$ + "IncludeBinary " + #DOUBLEQUOTE$ + RewrittenBinPath$ + #DOUBLEQUOTE$
            *CopyStart = *Char
          EndIf
        Else
          ; Strip IncludePath, IncludeFile, and XIncludeFile statements cleanly (including trailing EOL when standalone):
          If *StmtStart > *CopyStart
            OutChunk$ + PeekS(*CopyStart, ComputeLenFromChars(*CopyStart, *StmtStart))
          EndIf
          SkipSpaces(*Char)
          If *Char < *Limit And *Char\c = ';'
            GoToEOL_AndExtractComment(@*Char)
          EndIf
          If *Char < *Limit And *Char\c = #CR And
             (*StmtStart = *StringStartingAdr Or PeekC(*StmtStart - SizeOf(Character)) = #CR)
            GoNextChar(*Char)
          ElseIf *Char < *Limit And *Char\c = ':' And PeekC(*Char + SizeOf(Character)) <> ':'
            GoNextChar(*Char)
          EndIf
          *CopyStart    = *Char
          IsStartOfStmt = #True
        EndIf
        Continue
      EndIf
      Continue
    EndIf
    ;
    IsStartOfStmt = #False
    If IsIncludedFile And MacroDepth = 0 And *Char\c = '#'
      *TokStart = *Char
      TokLCase$ = LCase(ExtractOrSkipNextToken(*StringStartingAdr, @*Char))
      If TokLCase$ = "#pb_compiler_ismainfile" Or TokLCase$ = "#pb_compiler_isincludefile"
        If *TokStart > *CopyStart
          OutChunk$ + PeekS(*CopyStart, ComputeLenFromChars(*CopyStart, *TokStart))
        EndIf
        OutChunk$ + Str(Bool(TokLCase$ = "#pb_compiler_isincludefile"))
        *CopyStart = *Char
        Continue
      ElseIf *Char > *TokStart
        Continue
      EndIf
    EndIf
    ;
    If IsPBSeparator(*Char\c)
      GoNextChar(*Char)
    Else
      While *Char < *Limit And *Char\c <> 0 And IsPBSeparator(*Char\c) = #False
        GoNextChar(*Char)
      Wend
    EndIf
  Wend
  ;
  If *Char > *CopyStart
    OutChunk$ + PeekS(*CopyStart, ComputeLenFromChars(*CopyStart, *Char))
  EndIf
  ;
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
  CompilerEndIf
  ;
  ProcedureReturn Trim(OutChunk$, #CR$)
EndProcedure
;}
;
; ===============================================================================
; 4. COMPILATION SEQUENCE MERGER, IDE FOOTER & ENTRY POINT
; ===============================================================================
;- Main flattening pipeline and tool registration
;{
Structure FT_FileSkipCache
  List Ranges.FT_SkipRange()
EndStructure
;
Procedure.s FT_BuildMainFooter(MainFile$, Flags, OutRefFile$)
  ;
  ; Role: Extracts the trailing '; IDE Options' block from 'MainFile$', strips obsolete multi-file
  ;       directives ('UseMainFile', 'Folding', 'CursorPosition', 'FirstLine'), and rewrites
  ;       'UseIcon' / 'AddResource' paths relative to 'OutRefFile$' when #FT_RewriteBinPaths is active.
  ;
  Protected *MainBuf.String = GetPointedCodeFromFile(MainFile$, #False)
  Protected IDEPos, LineIdx, EqPos
  Protected RawFooter$, Line$, KeyLow$, PathVal$, FullPath$, Result$
  ;
  If *MainBuf = 0 Or *MainBuf\s = ""
    ProcedureReturn ""
  EndIf
  IDEPos = FindIDENotations(*MainBuf)
  If IDEPos < 1 Or IDEPos > Len(*MainBuf\s)
    ProcedureReturn ""
  EndIf
  ;
  RawFooter$ = FastMid(*MainBuf\s, IDEPos)
  If FindString(RawFooter$, "; IDE Options = PureBasic") = 0
    ProcedureReturn ""
  EndIf
  ;
  Repeat
    LineIdx + 1
    Line$ = Trim(StringField(RawFooter$, LineIdx, #CR$))
    If Line$ = ""
      If LineIdx > CountString(RawFooter$, #CR$) + 1 : Break : EndIf
      Continue
    EndIf
    KeyLow$ = LCase(Line$)
    If Left(KeyLow$, 16) = "; usemainfile = " Or Left(KeyLow$, 12) = "; folding = " Or
       Left(KeyLow$, 19) = "; cursorposition = " Or Left(KeyLow$, 14) = "; firstline = "
      Continue
    EndIf
    If (Flags & #FT_RewriteBinPaths) And OutRefFile$ <> "" And
       (Left(KeyLow$, 12) = "; useicon = " Or Left(KeyLow$, 16) = "; addresource = ")
      EqPos    = FindString(Line$, "=")
      PathVal$ = Trim(Mid(Line$, EqPos + 1))
      If PathVal$ <> ""
        FullPath$ = RelativeAddressToFullAddress(PathVal$, MainFile$)
        Line$     = Left(Line$, EqPos) + " " + FullAddressToRelativeAddress(FullPath$, OutRefFile$)
      EndIf
    EndIf
    Result$ + Line$ + #CR$
  ForEver
  ;
  ProcedureReturn Trim(Result$, #CR$)
EndProcedure
;
Procedure.s FT_GenerateStandaloneCode(*Config.FT_ConfigStruct, WGadget = 0)
  ;
  ; Role: Traverses PBB_CompilationSequence() in exact compiler order, merges all active code slices,
  ;       applies optional comment/whitespace cleanup via DeleteCommentsAndSpacesFromPBCode(),
  ;       appends the sanitized '; IDE Options' footer, and formats CRLF line endings.
  ;
  Protected *SavedSeq  = @PBB_CompilationSequence()
  Protected *SavedFile = @PBB_FileListing()
  Protected SegFile$, LastEmittedFile$, RelDisplay$, Chunk$, Footer$
  Protected MergedCode.String
  Protected NewMap ActiveIncludePaths.s()
  Protected NewMap CachedSkipRanges.FT_FileSkipCache()
  ;
  If ListSize(PBB_CompilationSequence()) = 0 And PBBFicPrincipalPB$ <> ""
    UpdateVirtualPositionsAndSequence(PBBFicPrincipalPB$, #True)
  EndIf
  ;
  ForEach PBB_CompilationSequence()
    SegFile$ = PBB_CompilationSequence()\FileName$
    If IsFileInMemory(SegFile$) = #False Or PBB_FileListing()\ExcludedByCompilerIf
      Continue
    EndIf
    ;
    If IsGadget(WGadget) And SegFile$ <> LastEmittedFile$
      SetGadgetText(WGadget, GetTextFromCatalogPB("FlattenToolTitle") + #CR$ + GetFilePart(SegFile$))
      UpdateWindow_(GadgetID(WGadget))
    EndIf
    ;
    If FindMapElement(CachedSkipRanges(), LCase(SegFile$)) = 0
      AddMapElement(CachedSkipRanges(), LCase(SegFile$))
      FT_BuildFileSkipRanges(SegFile$, *Config\Flags, CachedSkipRanges()\Ranges())
    EndIf
    ;
    Chunk$ = FT_ProcessSegment(SegFile$, PBB_CompilationSequence()\LocalStartPos,
                               PBB_CompilationSequence()\LocalEndPos, *Config\Flags,
                               *Config\OutputFile, ActiveIncludePaths(), CachedSkipRanges()\Ranges())
    If Chunk$ <> ""
      If MergedCode\s <> ""
        MergedCode\s + #CR$ + ";" + #CR$
      EndIf
      If (*Config\Flags & #FT_AddFileBanners) And SegFile$ <> LastEmittedFile$
        RelDisplay$ = FullAddressToRelativeAddress(SegFile$, PBBFicPrincipalPB$)
        MergedCode\s + "; ===========================================================================" + #CR$ +
                       "; " + RelDisplay$ + #CR$ +
                       "; ===========================================================================" + #CR$
      EndIf
      MergedCode\s + Chunk$
      LastEmittedFile$ = SegFile$
    EndIf
  Next
  ;
  If *SavedSeq  : ChangeCurrentElement(PBB_CompilationSequence(), *SavedSeq) : Else : ResetList(PBB_CompilationSequence()) : EndIf
  If *SavedFile : ChangeCurrentElement(PBB_FileListing(), *SavedFile)        : Else : ResetList(PBB_FileListing())        : EndIf
  ;
  If MergedCode\s = ""
    ProcedureReturn ""
  EndIf
  ;
  ; Apply optional comment stripping and/or whitespace compaction before appending the IDE Options footer:
  If *Config\Flags & (#FT_EraseComments | #FT_CleanSpaces)
    DeleteCommentsAndSpacesFromPBCode(@MergedCode, Bool((*Config\Flags & #FT_CleanSpaces) <> 0),
                                      Bool((*Config\Flags & #FT_EraseComments) <> 0))
    MergedCode\s = Trim(MergedCode\s, #CR$)
  EndIf
  ;
  Footer$ = FT_BuildMainFooter(PBBFicPrincipalPB$, *Config\Flags, *Config\OutputFile)
  If Footer$ <> ""
    MergedCode\s + #CR$ + #CR$ + Footer$
  EndIf
  MergedCode\s + #CR$
  ApplyLineEnding(@MergedCode, #LineEnding_CRLF)
  ;
  ProcedureReturn MergedCode\s
EndProcedure
;
Procedure FlattenProjectTool()
  ;
  ; Role: Entry point for the Standalone Single-File Exporter custom tool.
  ;       Completes any pending AST exploration, prompts for export options, generates the
  ;       flattened .pb source, and dispatches it to the clipboard, disk, and/or PureBasic IDE.
  ;
  Protected Config.FT_ConfigStruct
  Protected WGadget, SavedOK
  Protected Output$, ConfirmMsg$ = ""
  ;
  If PBBFicPrincipalPB$ = "" Or ListSize(PBB_FileListing()) = 0
    AlertInPBBWindow(GetTextFromCatalogPB("NoElementFound"))
    ProcedureReturn #False
  EndIf
  ;
  If ExplorationState < #ExplorationState_Completed
    ManagePBBBackgroundTasks(#BackgroundExplorationHasPriority)
  EndIf
  ;
  If FlattenToolOptionsWindow(@Config) = #False
    ProcedureReturn #False
  EndIf
  ;
  ; Guard against overwriting the active main source file itself:
  If (Config\Flags & #FT_OutSaveFile) And LCase(Config\OutputFile) = LCase(PBBFicPrincipalPB$)
    AlertInPBBWindow(GetTextFromCatalogPB("ErrorWithFile") + " " + Config\OutputFile)
    ProcedureReturn #False
  EndIf
  ;
  DisablePBBWindow()
  WGadget = OpenWaitWindow(GPBBGadgets\PBBWindow, GetTextFromCatalogPB("FlattenToolTitle"))
  Output$ = FT_GenerateStandaloneCode(@Config, WGadget)
  If IsGadget(WGadget)
    DestroyParentWindow(WGadget)
  EndIf
  EnablePBBWindow()
  ;
  If Output$ = ""
    AlertInPBBWindow(GetTextFromCatalogPB("NoElementFound"))
    ProcedureReturn #False
  EndIf
  ;
  If Config\Flags & #FT_OutClipboard
    SetClipboardText(Output$)
  EndIf
  ;
  If Config\Flags & #FT_OutSaveFile
    CreatePathFolders(Config\OutputFile)
    SavedOK = TexteDansFichier(Config\OutputFile, Output$, #PB_UTF8, 0, 1)
    If SavedOK = 0
      AlertInPBBWindow(GetTextFromCatalogPB("ErrorWithFile") + " " + Config\OutputFile)
      ProcedureReturn #False
    ElseIf (Config\Flags & #FT_OutOpenInIDE) And FileSize(PureBasicProgAddr$) >= 2
      RunProgram(#DOUBLEQUOTE$ + PureBasicProgAddr$ + #DOUBLEQUOTE$,
                 #DOUBLEQUOTE$ + Config\OutputFile + #DOUBLEQUOTE$,
                 GetPathPart(Config\OutputFile))
      ProcedureReturn #True
    EndIf
  EndIf
  ;
  If (Config\Flags & #FT_OutClipboard) And SavedOK
    ConfirmMsg$ = ReplaceString(GetTextFromCatalogPB("ST_DoneBoth"), "%1", Config\OutputFile)
  ElseIf SavedOK
    ConfirmMsg$ = ReplaceString(GetTextFromCatalogPB("ST_DoneSaved"), "%1", Config\OutputFile)
  ElseIf Config\Flags & #FT_OutClipboard
    ConfirmMsg$ = GetTextFromCatalogPB("ST_DoneClipboard")
  EndIf
  ;
  If ConfirmMsg$ <> ""
    AlertInPBBWindow(ConfirmMsg$, GetTextFromCatalogPB("FlattenToolTitle"))
  EndIf
  ;
  ProcedureReturn #True
EndProcedure
;
RegisterCustomTool("InCatalog:FlattenToolTitle", "InCatalog:FlattenToolEx", @FlattenProjectTool(), 0)
;}
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 495
; FirstLine = 195
; Folding = j---
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb