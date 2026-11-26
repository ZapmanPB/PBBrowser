; *******************************************************************************
;
;                          PBBrowser_SignatureTool.pb
;
;       Procedure Signature Generator & Local Variable Audit Custom Tools
;
; *******************************************************************************
;
; ===============================================================================
; 1. SIGNATURE GENERATOR CONSTANTS & STRUCTURES
; ===============================================================================
;
EnumerationBinary ST_OptionFlags
  #ST_SortAlpha          = 1    ; Sort procedures alphabetically instead of compiler/file order.
  #ST_ExcludeModules     = 2    ; Exclude procedures declared inside Module / DeclareModule blocks.
  #ST_KeepProcKeyword    = 4    ; Keep 'Procedure...' keyword instead of converting to 'Declare...'.
  #ST_KeepHeaderComments = 8    ; Extract leading ';' header comments below each procedure signature.
  #ST_KeepStandaloneCom  = 16   ; Extract top-level ';' comment blocks outside containers (source order).
  #ST_AnnotateCallGraph  = 32   ; Emit '; Calls:' and '; Called by:' comments.
  #ST_AnnotateGlobals    = 64   ; Emit '; Global/Threaded variables:' and '; Shared variables:' comments.
  #ST_AnnotateVarIssues  = 128  ; Emit '; Unprotected variables:' and '; Unused variables:' comments.
  #ST_OutClipboard       = 256  ; Copy the generated text to the Windows clipboard.
  #ST_OutSaveFile        = 512  ; Write the generated text to the target disk file.
  #ST_OutOpenFile        = 1024 ; Launch the saved file with its default Windows application.
EndEnumeration
;
Structure ST_ConfigStruct
  Flags.i
  ScopeFile.s
  OutputFile.s
EndStructure
;
Structure ST_GadgetsStruct
  Window.i
  ;
  ; Contiguous gadget IDs ordered for pointer-based font and theme styling:
  LblSecScope.i
  LblScope.i
  CmbScope.i
  OptSortSource.i
  OptSortAlpha.i
  ChkExcludeModules.i
  ;
  LblSecFormat.i
  OptKwDeclare.i
  OptKwProcedure.i
  ChkHeaderComments.i
  ChkStandaloneComments.i
  ;
  LblSecMeta.i
  ChkCallGraph.i
  ChkGlobalShared.i
  ChkVarIssues.i
  ;
  LblSecOutput.i
  ChkClipboard.i
  ChkSaveFile.i
  StrOutFile.i
  BtnBrowse.i
  ChkOpenFile.i
  ;
  BtnCancel.i
  BtnOK.i
EndStructure
;
Structure ST_ProcEntry
  FileOrder.i
  FileName.s
  RelFileName.s
  StartingPos.i
  EndingPos.i
  SortKey.s
  EnclosingModuleID.i
  ModuleName.s
  IsDeclareModule.b
  Used.b
  PurifiedName.s
  Declaration.s
  UsageParents.s
  UsageChildren.s
  OOE_UsagePos.s
EndStructure
;
Structure ST_RangeItem
  StartPos.i
  EndPos.i
EndStructure
;
Structure ST_CommentBlock
  StartPos.i
  Text.s
EndStructure
;
; ===============================================================================
; 2. PATH & OPTIONS DIALOG MANAGEMENT
; ===============================================================================
;- Path and options dialog helpers
;{
Procedure.s ST_GetDisplayFilePath(FileName$)
  ;
  ; Role: Formats 'FileName$' as a backslash-normalized path relative to PBBFicPrincipalPB$
  ;       for display in the scope selector, file banners, and alphabetical file annotations.
  ;
  If PBBFicPrincipalPB$ <> ""
    ProcedureReturn FullAddressToRelativeAddress(FileName$, PBBFicPrincipalPB$)
  EndIf
  ProcedureReturn GetFilePart(FileName$)
EndProcedure
;
Procedure.s ST_BuildDefaultOutputFile(ScopeFile$, UseProcedureKw = #False)
  ;
  ; Role: Computes the default output path ('_Declares.pb' in Declare mode, or '_Signatures.txt'
  ;       in Procedure mode) matching either the selected source file or the active project.
  ;
  Protected BaseFile$ = ScopeFile$
  ;
  If BaseFile$ = ""
    If PBBCurrentProject$ <> ""
      BaseFile$ = PBBCurrentProject$
    Else
      BaseFile$ = PBBFicPrincipalPB$
    EndIf
  EndIf
  If BaseFile$ = ""
    ProcedureReturn ""
  EndIf
  ;
  If UseProcedureKw
    ProcedureReturn GetPathPart(BaseFile$) + FilePartSsExt(BaseFile$) + "_Signatures.txt"
  EndIf
  ProcedureReturn GetPathPart(BaseFile$) + FilePartSsExt(BaseFile$) + "_Declares.pb"
EndProcedure
;
Procedure ST_ApplyTheme(*GD.ST_GadgetsStruct)
  ;
  ; Role: Applies PBBrowser's active color preset and fonts to the dialog and all its child
  ;       gadgets via SetFontAndGadgetsColors() followed by section-header font/color overrides.
  ;
  Protected *GadgetAddress.Integer, GadgetID
  Protected GadgetList$ = ""
  ;
  *GadgetAddress = @*GD\LblSecScope
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
  ; Override the default preset font with PBBrowser's configured fonts and highlight section headers:
  *GadgetAddress = @*GD\LblSecScope
  While *GadgetAddress <= @*GD\BtnOK
    GadgetID = *GadgetAddress\i
    If IsGadget(GadgetID)
      If GadgetID = *GD\LblSecScope Or GadgetID = *GD\LblSecFormat Or
         GadgetID = *GD\LblSecMeta  Or GadgetID = *GD\LblSecOutput
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
Procedure.i ST_UpdateWindowState(*GD.ST_GadgetsStruct)
  ;
  ; Role: Enforces cross-option dependencies (disabling standalone file comments when alphabetical
  ;       sorting breaks file linearity, toggling file output controls, and guarding OK validation).
  ; Output: Returns #True when at least one valid output destination is active, #False otherwise.
  ;
  Protected SaveEnabled, CanValidate
  ;
  DisableGadget(*GD\ChkStandaloneComments, Bool(GetGadgetState(*GD\OptSortAlpha) <> 0))
  ;
  SaveEnabled = Bool(GetGadgetState(*GD\ChkSaveFile) = #PB_Checkbox_Checked)
  DisableGadget(*GD\StrOutFile, Bool(SaveEnabled = #False))
  DisableGadget(*GD\BtnBrowse, Bool(SaveEnabled = #False))
  DisableGadget(*GD\ChkOpenFile, Bool(SaveEnabled = #False))
  ;
  CanValidate = Bool(GetGadgetState(*GD\ChkClipboard) = #PB_Checkbox_Checked Or
                     (SaveEnabled And Trim(GetGadgetText(*GD\StrOutFile)) <> ""))
  DisableGadget(*GD\BtnOK, Bool(CanValidate = #False))
  ;
  ProcedureReturn CanValidate
EndProcedure
;
Procedure.i SignatureToolOptionsWindow(*Config.ST_ConfigStruct)
  ;
  ; Role: Displays the themed modal dialog for the procedure signature generator, dynamically
  ;       adapting the file selector, output suffix, and extension to the active options.
  ; Output: Populates '*Config' and returns #True if validated, or #False if cancelled.
  ;
  Protected GD.ST_GadgetsStruct
  Protected WWidth = 540, WHeight = 520, Margins = 18, LineHeight = 22, ButtonHeight = 28
  Protected OX, OY, ParentWindowID, YPos, InnerX, FullInnerW
  Protected FileCount, ScopeIdx, UseProcKw, Event, EventGadget, CanValidate, Validated = #False
  Protected AutoOutFile$, NewAutoOutFile$, CurrentOutFile$, PickedFile$
  ;
  FileCount = ListSize(PBB_FileListing())
  Protected Dim ScopeFiles$(FileCount)
  ;
  InnerX         = Margins + 8
  FullInnerW     = WWidth - InnerX - Margins
  AutoOutFile$   = ST_BuildDefaultOutputFile("", #False)
  ParentWindowID = ComputeWinOrigins(@OX, @OY, WWidth, WHeight, GPBBGadgets\PBBWindow)
  ;
  DisablePBBWindow()
  GD\Window = OpenWindow(#PB_Any, OX, OY, WWidth, WHeight, GetTextFromCatalogPB("SignatureToolTitle"),
                         #PB_Window_SystemMenu | #PB_Window_Invisible, ParentWindowID)
  If GD\Window
    ApplyDarkModeToWindow(GD\Window)
    StickyWindow(GD\Window, 1)
    ;
    ; --- Section 1: Scope and Ordering ---
    YPos           = 14
    GD\LblSecScope = TextGadget(#PB_Any, Margins, YPos, WWidth - Margins * 2, 20, GetTextFromCatalogPB("ST_SecScope"))
    YPos + 24
    GD\LblScope    = TextGadget(#PB_Any, InnerX, YPos + 3, 125, LineHeight, GetTextFromCatalogPB("ST_TargetFiles"))
    GD\CmbScope    = ComboBoxGadget(#PB_Any, InnerX + 128, YPos, FullInnerW - 128, 24)
    AddGadgetItem(GD\CmbScope, -1, GetTextFromCatalogPB("ST_AllProjectFiles"))
    ScopeIdx = 0
    ForEach PBB_FileListing()
      ScopeIdx + 1
      ScopeFiles$(ScopeIdx) = PBB_FileListing()\FileName$
      AddGadgetItem(GD\CmbScope, -1, ST_GetDisplayFilePath(PBB_FileListing()\FileName$))
    Next
    SetGadgetState(GD\CmbScope, 0)
    ;
    ; Radio group #1 (closed immediately below by ChkExcludeModules):
    YPos + 30
    GD\OptSortSource     = OptionGadget(#PB_Any, InnerX, YPos, 270, LineHeight, GetTextFromCatalogPB("ST_SortSource"))
    GD\OptSortAlpha      = OptionGadget(#PB_Any, InnerX + 275, YPos, FullInnerW - 275, LineHeight, GetTextFromCatalogPB("ST_SortAlpha"))
    SetGadgetState(GD\OptSortSource, 1)
    YPos + 26
    GD\ChkExcludeModules = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_ExcludeModules"))
    SetGadgetState(GD\ChkExcludeModules, #PB_Checkbox_Checked)
    ;
    ; --- Section 2: Signature and Comments ---
    YPos + 34
    GD\LblSecFormat = TextGadget(#PB_Any, Margins, YPos, WWidth - Margins * 2, 20, GetTextFromCatalogPB("ST_SecFormat"))
    ;
    ; Radio group #2 (closed immediately below by ChkHeaderComments):
    YPos + 23
    GD\OptKwDeclare          = OptionGadget(#PB_Any, InnerX, YPos, 270, LineHeight, GetTextFromCatalogPB("ST_UseDeclare"))
    GD\OptKwProcedure        = OptionGadget(#PB_Any, InnerX + 275, YPos, FullInnerW - 275, LineHeight, GetTextFromCatalogPB("ST_KeepProcedure"))
    SetGadgetState(GD\OptKwDeclare, 1)
    YPos + 26
    GD\ChkHeaderComments     = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_KeepHeaderCom"))
    SetGadgetState(GD\ChkHeaderComments, #PB_Checkbox_Checked)
    YPos + 26
    GD\ChkStandaloneComments = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_KeepStandaloneCom"))
    ;
    ; --- Section 3: Analysis Annotations ---
    YPos + 34
    GD\LblSecMeta      = TextGadget(#PB_Any, Margins, YPos, WWidth - Margins * 2, 20, GetTextFromCatalogPB("ST_SecMeta"))
    YPos + 23
    GD\ChkCallGraph    = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_AnnotateCalls"))
    YPos + 26
    GD\ChkGlobalShared = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_AnnotateGlobals"))
    YPos + 26
    GD\ChkVarIssues    = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_AnnotateVarIssues"))
    ;
    ; --- Section 4: Output Destination ---
    YPos + 34
    GD\LblSecOutput = TextGadget(#PB_Any, Margins, YPos, WWidth - Margins * 2, 20, GetTextFromCatalogPB("ST_SecOutput"))
    YPos + 23
    GD\ChkClipboard = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_CopyClipboard"))
    SetGadgetState(GD\ChkClipboard, #PB_Checkbox_Checked)
    YPos + 26
    GD\ChkSaveFile  = CheckBoxGadget(#PB_Any, InnerX, YPos, FullInnerW, LineHeight, GetTextFromCatalogPB("ST_SaveToFile"))
    YPos + 26
    GD\StrOutFile   = StringGadget(#PB_Any, InnerX + 18, YPos, FullInnerW - 54, 24, AutoOutFile$)
    GD\BtnBrowse    = ButtonGadget(#PB_Any, InnerX + FullInnerW - 30, YPos, 30, 24, "...")
    YPos + 28
    GD\ChkOpenFile  = CheckBoxGadget(#PB_Any, InnerX + 18, YPos, FullInnerW - 18, LineHeight, GetTextFromCatalogPB("ST_OpenAfterSave"))
    SetGadgetState(GD\ChkOpenFile, #PB_Checkbox_Checked)
    ;
    ; --- Bottom Action Buttons ---
    YPos         = WHeight - ButtonHeight - 14
    GD\BtnCancel = ButtonGadget(#PB_Any, WWidth - Margins - 210, YPos, 100, ButtonHeight, GetTextFromCatalogPB("Cancel"))
    GD\BtnOK     = ButtonGadget(#PB_Any, WWidth - Margins - 100, YPos, 100, ButtonHeight, GetTextFromCatalogPB("OK"), #PB_Button_Default)
    ;
    ST_ApplyTheme(@GD)
    CanValidate = ST_UpdateWindowState(@GD)
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
                If CanValidate
                  Validated = #True
                EndIf
              Case 101
                Break
            EndSelect
            ;
          Case #PB_Event_Gadget
            EventGadget = EventGadget()
            Select EventGadget
              Case GD\BtnOK
                If CanValidate
                  Validated = #True
                EndIf
              Case GD\BtnCancel
                Break
              Case GD\CmbScope, GD\OptKwDeclare, GD\OptKwProcedure
                ; Synchronize default output path (or toggle .pb/.txt extension on a custom user path):
                ScopeIdx       = GetGadgetState(GD\CmbScope)
                UseProcKw      = Bool(GetGadgetState(GD\OptKwProcedure) <> 0)
                CurrentOutFile$ = GetGadgetText(GD\StrOutFile)
                If ScopeIdx >= 0 And ScopeIdx <= FileCount
                  NewAutoOutFile$ = ST_BuildDefaultOutputFile(ScopeFiles$(ScopeIdx), UseProcKw)
                  If CurrentOutFile$ = AutoOutFile$ Or Trim(CurrentOutFile$) = ""
                    If CurrentOutFile$ <> NewAutoOutFile$
                      SetGadgetText(GD\StrOutFile, NewAutoOutFile$)
                    EndIf
                  ElseIf EventGadget = GD\OptKwDeclare
                    SetGadgetText(GD\StrOutFile, ChangeFileExtension(CurrentOutFile$, "pb"))
                  ElseIf EventGadget = GD\OptKwProcedure
                    SetGadgetText(GD\StrOutFile, ChangeFileExtension(CurrentOutFile$, "txt"))
                  EndIf
                  AutoOutFile$ = NewAutoOutFile$
                EndIf
              Case GD\BtnBrowse
                UseProcKw   = Bool(GetGadgetState(GD\OptKwProcedure) <> 0)
                PickedFile$ = SaveFileRequester(GetTextFromCatalogPB("SignatureToolTitle"),
                                                GetGadgetText(GD\StrOutFile),
                                                GetTextFromCatalogPB("ST_FileFilter"), UseProcKw)
                If PickedFile$ <> ""
                  If GetExtensionPart(PickedFile$) = ""
                    If SelectedFilePattern() = 0
                      PickedFile$ + ".pb"
                    ElseIf SelectedFilePattern() = 1
                      PickedFile$ + ".txt"
                    EndIf
                  EndIf
                  SetGadgetText(GD\StrOutFile, PickedFile$)
                EndIf
            EndSelect
            CanValidate = ST_UpdateWindowState(@GD)
        EndSelect
        ;
        If Validated
          ClearStructure(*Config, ST_ConfigStruct)
          ScopeIdx = GetGadgetState(GD\CmbScope)
          If ScopeIdx > 0 And ScopeIdx <= FileCount
            *Config\ScopeFile = ScopeFiles$(ScopeIdx)
          EndIf
          If GetGadgetState(GD\OptSortAlpha) <> 0
            *Config\Flags | #ST_SortAlpha
          ElseIf GetGadgetState(GD\ChkStandaloneComments) = #PB_Checkbox_Checked
            *Config\Flags | #ST_KeepStandaloneCom
          EndIf
          If GetGadgetState(GD\ChkExcludeModules) = #PB_Checkbox_Checked : *Config\Flags | #ST_ExcludeModules     : EndIf
          If GetGadgetState(GD\OptKwProcedure) <> 0                      : *Config\Flags | #ST_KeepProcKeyword    : EndIf
          If GetGadgetState(GD\ChkHeaderComments) = #PB_Checkbox_Checked : *Config\Flags | #ST_KeepHeaderComments : EndIf
          If GetGadgetState(GD\ChkCallGraph) = #PB_Checkbox_Checked      : *Config\Flags | #ST_AnnotateCallGraph  : EndIf
          If GetGadgetState(GD\ChkGlobalShared) = #PB_Checkbox_Checked   : *Config\Flags | #ST_AnnotateGlobals    : EndIf
          If GetGadgetState(GD\ChkVarIssues) = #PB_Checkbox_Checked      : *Config\Flags | #ST_AnnotateVarIssues  : EndIf
          If GetGadgetState(GD\ChkClipboard) = #PB_Checkbox_Checked      : *Config\Flags | #ST_OutClipboard       : EndIf
          If GetGadgetState(GD\ChkSaveFile) = #PB_Checkbox_Checked
            *Config\Flags | #ST_OutSaveFile
            *Config\OutputFile = Trim(GetGadgetText(GD\StrOutFile))
            If GetGadgetState(GD\ChkOpenFile) = #PB_Checkbox_Checked
              *Config\Flags | #ST_OutOpenFile
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
; 3. SIGNATURE FORMATTING, COMMENTS & AST ANNOTATIONS
; ===============================================================================
;- Signature formatting, comments, and AST annotations
;{
Procedure.s ST_FormatSignatureLine(RawDecl$, UseDeclareKeyword)
  ;
  ; Role: Normalizes a single-line procedure header from ElementList()\Declaration by converting
  ;       or standardizing the leading Procedure/Declare keyword, preserving any '.Type' or
  ;       '.Module::Type' return suffix, and enforcing a single space before the procedure name.
  ;
  Protected *Char.Character, *Start.Character
  Protected KwLCase$, OutHeader$, TypePart$
  ;
  If RawDecl$ = ""
    ProcedureReturn ""
  EndIf
  ;
  *Char = @RawDecl$
  SkipSpaces(*Char)
  *Start = *Char
  While (*Char\c >= 'a' And *Char\c <= 'z') Or (*Char\c >= 'A' And *Char\c <= 'Z')
    GoNextChar(*Char)
  Wend
  KwLCase$ = LCase(PeekS(*Start, ComputeLenFromChars(*Start, *Char)))
  ;
  Select KwLCase$
    Case "procedurecdll", "declarecdll"
      If UseDeclareKeyword : OutHeader$ = "DeclareCDLL" : Else : OutHeader$ = "ProcedureCDLL" : EndIf
    Case "proceduredll", "declaredll"
      If UseDeclareKeyword : OutHeader$ = "DeclareDLL"  : Else : OutHeader$ = "ProcedureDLL"  : EndIf
    Case "procedurec", "declarec"
      If UseDeclareKeyword : OutHeader$ = "DeclareC"    : Else : OutHeader$ = "ProcedureC"    : EndIf
    Default
      If UseDeclareKeyword : OutHeader$ = "Declare"     : Else : OutHeader$ = "Procedure"     : EndIf
  EndSelect
  ;
  ; Preserve optional '.ReturnType' or '.Module::ReturnType' suffix attached to the keyword:
  SkipSpaces(*Char)
  If *Char\c = '.'
    GoNextChar(*Char)
    SkipSpaces(*Char)
    *Start = *Char
    While IsPBSeparator(*Char\c) = #False
      GoNextChar(*Char)
    Wend
    TypePart$ = PeekS(*Start, ComputeLenFromChars(*Start, *Char))
    SkipSpaces(*Char)
    If IsModuleUsage(@*Char)
      SkipSpaces(*Char)
      *Start = *Char
      While IsPBSeparator(*Char\c) = #False
        GoNextChar(*Char)
      Wend
      TypePart$ + "::" + PeekS(*Start, ComputeLenFromChars(*Start, *Char))
    EndIf
    If TypePart$ <> ""
      OutHeader$ + "." + TypePart$
    EndIf
  EndIf
  ;
  SkipSpaces(*Char)
  ProcedureReturn OutHeader$ + " " + PeekS(*Char)
EndProcedure
;
Procedure ST_AppendCommentLine(*CommentAfterSemicolon, *LineEnd, Indent$, *Buffer.String, *PendingSpacer.Integer)
  ;
  ; Role: Formats one ';' comment line (collapsing empty ';' spacers until non-empty prose follows
  ;       so outer bare ';' lines are stripped) and appends it to '*Buffer\s' with 'Indent$'.
  ;
  Protected CommentBody$ = RTrim(PeekS(*CommentAfterSemicolon, ComputeLenFromChars(*CommentAfterSemicolon, *LineEnd)))
  ;
  If Trim(CommentBody$) = ""
    If *Buffer\s <> ""
      *PendingSpacer\i = #True
    EndIf
  Else
    If *PendingSpacer\i
      *Buffer\s + Indent$ + ";" + #CR$
      *PendingSpacer\i = #False
    EndIf
    If Left(CommentBody$, 1) <> " " And Left(CommentBody$, 1) <> ";" And Left(CommentBody$, 1) <> "-"
      CommentBody$ = " " + CommentBody$
    EndIf
    *Buffer\s + Indent$ + ";" + CommentBody$ + #CR$
  EndIf
EndProcedure
;
Procedure.s ST_ExtractProcHeaderComments(*FileBuffer.String, StartingPos, EndingPos, Indent$)
  ;
  ; Role: Advances past the balanced parameter parentheses of the procedure starting at 'StartingPos'
  ;       (safely ignoring quotes, escaped strings, and multi-line header comments) and extracts
  ;       the contiguous block of leading ';' description comments at the top of the procedure body.
  ; Output: Returns the formatted comment lines prefixed with 'Indent$' (stripping outer bare ';' spacers).
  ;
  Protected *StringStartingAdr, *Char.Character, *Limit.Character, *LineEnd.Character
  Protected HeaderEndPos, PendingSpacer
  Protected Result.String
  ;
  HeaderEndPos = GetProcedureHeaderEndPos(*FileBuffer, StartingPos, EndingPos)
  If HeaderEndPos = 0
    ProcedureReturn ""
  EndIf
  ;
  *StringStartingAdr = @*FileBuffer\s
  *Char              = PosToChar(*StringStartingAdr, HeaderEndPos) + SizeOf(Character)
  *Limit             = PosToChar(*StringStartingAdr, EndingPos)
  ;
  ; Advance to the first physical line below the signature:
  While *Char < *Limit And *Char\c <> #CR And *Char\c <> 0
    GoNextChar(*Char)
  Wend
  If *Char < *Limit And *Char\c = #CR
    GoNextChar(*Char)
  EndIf
  ;
  ; Collect contiguous leading ';' lines until the first executable or declaration statement:
  While *Char < *Limit And *Char\c <> 0
    SkipSpaces(*Char)
    *LineEnd = *Char
    While *LineEnd < *Limit And *LineEnd\c <> #CR And *LineEnd\c <> 0
      GoNextChar(*LineEnd)
    Wend
    ;
    If *Char = *LineEnd
      If Result\s <> ""
        PendingSpacer = #True
      EndIf
    ElseIf *Char\c = ';'
      GoNextChar(*Char)
      ST_AppendCommentLine(*Char, *LineEnd, Indent$, @Result, @PendingSpacer)
    Else
      Break
    EndIf
    ;
    *Char = *LineEnd
    If *Char < *Limit And *Char\c = #CR
      GoNextChar(*Char)
    EndIf
  Wend
  ;
  ProcedureReturn Result\s
EndProcedure
;
Procedure ST_CollectStandaloneComments(FileName$, List CommentBlocks.ST_CommentBlock())
  ;
  ; Role: Scans 'FileName$' upstream of its PureBasic IDE Options footer and collects contiguous
  ;       top-level ';' comment blocks situated outside any block container (Procedure, Macro,
  ;       Structure, Interface, Enumeration, Module, Import, or excluded IsMainFile block).
  ;
  Protected *FileBuffer.String = GetPointedCodeFromFile(FileName$, #False)
  Protected *SavedElem
  Protected *StringStartingAdr, *Char.Character, *Limit.Character, *LineEnd.Character
  Protected IDELimitPos, LinePos, HasActiveRange, IsAtEnd, IsCommentLine, BlockStartPos, PendingSpacer
  Protected BlockText.String
  Protected NewList RawRanges.ST_RangeItem()
  Protected NewList MergedRanges.ST_RangeItem()
  Protected NewList CEDList.CED_Structure()
  ;
  ClearList(CommentBlocks())
  If *FileBuffer = 0 Or *FileBuffer\s = ""
    ProcedureReturn
  EndIf
  ;
  IDELimitPos = FindIDENotations(*FileBuffer) - 1
  If IDELimitPos < 1
    ProcedureReturn
  EndIf
  ;
  ; 1. Gather all container intervals and inactive CompilerIf / CompilerSelect zones for 'FileName$':
  If IsFileInMemory(FileName$)
    ForEach PBB_FileListing()\DeadZones()
      If PBB_FileListing()\DeadZones()\StartingPos > 0
        AddElement(RawRanges())
        RawRanges()\StartPos = PBB_FileListing()\DeadZones()\StartingPos
        If PBB_FileListing()\DeadZones()\EndingPos >= PBB_FileListing()\DeadZones()\StartingPos
          RawRanges()\EndPos = PBB_FileListing()\DeadZones()\EndingPos
        Else
          RawRanges()\EndPos = IDELimitPos
        EndIf
      EndIf
    Next
  EndIf
  ;
  *SavedElem = @ElementList()
  ForEach ElementList()
    If ElementList()\Type >= #PBBProcedure And ElementList()\Type <= #PBBImport And
       ElementList()\Type <> #PBBLabel And ElementList()\Type <> #PBBConstant And ElementList()\Type <> #PBBVariable
      If ElementList()\FileName = FileName$ And ElementList()\StartingPos > 0 And
         ElementList()\EndingPos >= ElementList()\StartingPos
        AddElement(RawRanges())
        RawRanges()\StartPos = ElementList()\StartingPos
        RawRanges()\EndPos   = ElementList()\EndingPos
      EndIf
      If (ElementList()\Type = #PBBEnumeration Or ElementList()\Type = #PBBModule) And ElementList()\Continued <> ""
        CED_StringToList(ElementList()\Continued, CEDList())
        ForEach CEDList()
          If CEDList()\Filename = FileName$ And CEDList()\StartingPos > 0 And
             CEDList()\EndingPos >= CEDList()\StartingPos
            AddElement(RawRanges())
            RawRanges()\StartPos = CEDList()\StartingPos
            RawRanges()\EndPos   = CEDList()\EndingPos
          EndIf
        Next
      EndIf
    EndIf
  Next
  If *SavedElem
    ChangeCurrentElement(ElementList(), *SavedElem)
  EndIf
  ;
  ; Sort and merge overlapping/nested container intervals into a flat non-overlapping list:
  If ListSize(RawRanges()) > 0
    SortStructuredList(RawRanges(), #PB_Sort_Ascending, OffsetOf(ST_RangeItem\StartPos), #PB_Integer)
    ForEach RawRanges()
      If ListSize(MergedRanges()) = 0 Or RawRanges()\StartPos > MergedRanges()\EndPos
        AddElement(MergedRanges())
        MergedRanges()\StartPos = RawRanges()\StartPos
        MergedRanges()\EndPos   = RawRanges()\EndPos
      ElseIf RawRanges()\EndPos > MergedRanges()\EndPos
        MergedRanges()\EndPos = RawRanges()\EndPos
      EndIf
    Next
  EndIf
  ;
  ; 2. Scan top-level lines while jumping directly over excluded container intervals in O(1):
  HasActiveRange     = Bool(FirstElement(MergedRanges()) <> 0)
  *StringStartingAdr = @*FileBuffer\s ; Required by the CharToPos() macro below.
  *Char              = *StringStartingAdr
  *Limit             = PosToChar(*StringStartingAdr, IDELimitPos) + SizeOf(Character)
  ;
  Repeat
    IsAtEnd       = Bool(*Char >= *Limit Or *Char\c = 0)
    IsCommentLine = #False
    ;
    If IsAtEnd = #False
      LinePos = CharToPos(*Char)
      While HasActiveRange And MergedRanges()\EndPos < LinePos
        HasActiveRange = Bool(NextElement(MergedRanges()) <> 0)
      Wend
      ;
      If HasActiveRange And LinePos >= MergedRanges()\StartPos
        *Char = PosToChar(*StringStartingAdr, MergedRanges()\EndPos)
        While *Char < *Limit And *Char\c <> #CR And *Char\c <> 0
          GoNextChar(*Char)
        Wend
        If *Char < *Limit And *Char\c = #CR
          GoNextChar(*Char)
        EndIf
      Else
        SkipSpaces(*Char)
        *LineEnd = *Char
        While *LineEnd < *Limit And *LineEnd\c <> #CR And *LineEnd\c <> 0
          GoNextChar(*LineEnd)
        Wend
        If *Char < *LineEnd And *Char\c = ';'
          IsCommentLine = #True
          GoNextChar(*Char)
          If BlockText\s = ""
            BlockStartPos = LinePos
          EndIf
          ST_AppendCommentLine(*Char, *LineEnd, "", @BlockText, @PendingSpacer)
        EndIf
        *Char = *LineEnd
        If *Char < *Limit And *Char\c = #CR
          GoNextChar(*Char)
        EndIf
      EndIf
    EndIf
    ;
    ; Flush completed comment block upon hitting code, a blank line, a container, or end-of-scan:
    If IsCommentLine = #False And BlockText\s <> ""
      AddElement(CommentBlocks())
      CommentBlocks()\StartPos = BlockStartPos
      CommentBlocks()\Text     = Trim(BlockText\s, #CR$)
      BlockText\s              = ""
      PendingSpacer            = #False
    EndIf
  Until IsAtEnd
EndProcedure
;
Procedure.s ST_BuildProcAnnotations(*Proc.ST_ProcEntry, Flags, ShowAlphaFile, Indent$)
  ;
  ; Role: Builds English ';' analysis comment lines for '*Proc' covering its source file origin
  ;       (in multi-file alphabetical mode), call graph (called/calling procedures), referenced
  ;       Global/Threaded/Shared variables, and Unprotected/Unused local variables.
  ;
  Protected *SavedElem = @ElementList()
  Protected VList.VariableListings
  Protected ParentIdx, ParentID, OOEPos, LineIdx
  Protected Result$, Calls$, Callers$, GlobalVars$, SharedVars$, UnprotectedVars$, UnusedVars$ = ""
  Protected VLine$, VDeclare$, VName$, VState$, Species$, OOEEntry$, OOEFile$
  Protected ResolvedString$ ; Required by the AddUniqueListItem() macro.
  ;
  If ShowAlphaFile
    Result$ + Indent$ + "; File: " + *Proc\RelFileName + #CR$
  EndIf
  ;
  If Flags & (#ST_AnnotateGlobals | #ST_AnnotateVarIssues)
    ListVariables(*Proc\PurifiedName, @VList, #PBBProcedure, *Proc\ModuleName)
  EndIf
  ;
  ; 1. Call graph annotations (called procedures and calling procedures / main scope):
  If Flags & #ST_AnnotateCallGraph
    Calls$ = FillDetail_GetCalledElementsList(*Proc\UsageChildren, #PBBProcedure)
    If Calls$ <> ""
      Result$ + Indent$ + "; Calls: " + Calls$ + #CR$
    EndIf
    ;
    If Trim(*Proc\UsageParents, "/") <> ""
      ParentIdx = 1
      Repeat
        ParentIdx + 1
        VLine$ = StringField(*Proc\UsageParents, ParentIdx, "/")
        If VLine$ = ""
          Break
        EndIf
        ParentID = Val(VLine$)
        If ParentID = -1
          OOEEntry$ = "Out-of-procedure code"
          If *Proc\OOE_UsagePos <> ""
            VDeclare$ = StringField(*Proc\OOE_UsagePos, 1, #CR$)
            OOEFile$  = StringField(VDeclare$, 1, ",")
            OOEPos    = Val(StringField(VDeclare$, 2, ","))
            If OOEFile$ <> "" And OOEPos > 0
              OOEEntry$ + " (" + GetFilePart(OOEFile$) + ", line " + Str(GetNoLineFromFilenameAndPos(OOEFile$, OOEPos)) + ")"
            EndIf
          EndIf
          AddUniqueListItem(Callers$, OOEEntry$)
        ElseIf ParentID > 0 And SelectPBBElement(ParentID)
          If ElementList()\EnclosingModuleID > 0 And ElementList()\EnclosingModuleID <> *Proc\EnclosingModuleID
            AddUniqueListItem(Callers$, GetEnclosingModuleName(ElementList()\EnclosingModuleID) + "::" + ElementList()\Name)
          Else
            AddUniqueListItem(Callers$, ElementList()\Name)
          EndIf
        EndIf
      ForEver
      If Right(Callers$, 2) = ", "
        Callers$ = Left(Callers$, Len(Callers$) - 2)
      EndIf
    EndIf
    ;
    If Callers$ <> ""
      Result$ + Indent$ + "; Called by: " + Callers$ + #CR$
    ElseIf *Proc\Used = #False
      Result$ + Indent$ + "; Called by: None (Unused)" + #CR$
    EndIf
  EndIf
  ;
  ; 2. Single-pass extraction of Global/Threaded/Shared usages and local variable issues:
  If VList\CompleteList <> ""
    LineIdx = 0
    Repeat
      LineIdx + 1
      VLine$ = StringField(VList\CompleteList, LineIdx, #CR$)
      If VLine$ = ""
        Break
      EndIf
      VDeclare$ = StringField(VLine$, 1, #TAB$)
      VName$    = StringField(VLine$, 2, #TAB$)
      VState$   = StringField(VLine$, 3, #TAB$)
      Species$  = StringField(VDeclare$, 1, " ")
      ;
      If (Flags & #ST_AnnotateGlobals) And VState$ = "Used"
        If Species$ = "Global" Or Species$ = "Threaded"
          AddUniqueListItem(GlobalVars$, VName$)
        ElseIf Species$ = "Shared"
          AddUniqueListItem(SharedVars$, VName$)
        EndIf
      EndIf
      ;
      If Flags & #ST_AnnotateVarIssues
        If Species$ = "Unprotected"
          AddUniqueListItem(UnprotectedVars$, VName$)
        ElseIf VState$ = "Unused" And Species$ <> "Global" And Species$ <> "Threaded"
          AddUniqueListItem(UnusedVars$, VName$)
        EndIf
      EndIf
    ForEver
    ;
    If GlobalVars$ <> ""
      Result$ + Indent$ + "; Global/Threaded variables: " + Left(GlobalVars$, Len(GlobalVars$) - 2) + #CR$
    EndIf
    If SharedVars$ <> ""
      Result$ + Indent$ + "; Shared variables: " + Left(SharedVars$, Len(SharedVars$) - 2) + #CR$
    EndIf
    If UnprotectedVars$ <> ""
      Result$ + Indent$ + "; Unprotected variables: " + Left(UnprotectedVars$, Len(UnprotectedVars$) - 2) + #CR$
    EndIf
    If UnusedVars$ <> ""
      Result$ + Indent$ + "; Unused variables: " + Left(UnusedVars$, Len(UnusedVars$) - 2) + #CR$
    EndIf
  EndIf
  ;
  If *SavedElem
    ChangeCurrentElement(ElementList(), *SavedElem)
  EndIf
  ;
  ProcedureReturn Result$
EndProcedure
;
Procedure.s ST_FormatProcBlock(*Proc.ST_ProcEntry, Flags, ShowAlphaFile, SigIndent$, CommentTag$ = "")
  ;
  ; Role: Assembles one complete procedure block (formatted signature line, optional leading
  ;       header comments, and optional analysis annotations). When 'CommentTag$' is non-empty,
  ;       comments out the signature line and appends 'CommentTag$' (e.g., '[Module: Name]').
  ;
  Protected *FileBuffer.String
  Protected SigLine$, BodyIndent$, Block$ = ""
  ;
  SigLine$ = ST_FormatSignatureLine(*Proc\Declaration, Bool((Flags & #ST_KeepProcKeyword) = 0))
  If CommentTag$ <> ""
    SigLine$ = SigIndent$ + "; " + SigLine$ + " ; " + CommentTag$
  Else
    SigLine$ = SigIndent$ + SigLine$
  EndIf
  ;
  BodyIndent$ = SigIndent$ + "  "
  Block$      = SigLine$ + #CR$
  ;
  If Flags & #ST_KeepHeaderComments
    *FileBuffer = GetPointedCodeFromFile(*Proc\FileName, #False)
    Block$ + ST_ExtractProcHeaderComments(*FileBuffer, *Proc\StartingPos, *Proc\EndingPos, BodyIndent$)
  EndIf
  ;
  If ShowAlphaFile Or (Flags & (#ST_AnnotateCallGraph | #ST_AnnotateGlobals | #ST_AnnotateVarIssues))
    Block$ + ST_BuildProcAnnotations(*Proc, Flags, ShowAlphaFile, BodyIndent$)
  EndIf
  ;
  ProcedureReturn Trim(Block$, #CR$)
EndProcedure
;}
;
; ===============================================================================
; 4. MAIN GENERATION PIPELINE & ENTRY POINT
; ===============================================================================
;- Signature generation pipeline and entry point
;{
Procedure ST_AppendSection(*Dest.String, Block$, *LastWasSpaced.Integer, ForceSpacing = #False)
  ;
  ; Role: Appends 'Block$' to '*Dest\s', inserting a blank line (#CR$ + #CR$) when either
  ;       the previous block or 'Block$' spans multiple lines (or 'ForceSpacing' is #True),
  ;       and a single line break (#CR$) between consecutive single-line signatures.
  ;
  Protected NeedsSpacing
  ;
  If Block$ = ""
    ProcedureReturn
  EndIf
  ;
  NeedsSpacing = Bool(ForceSpacing Or FindString(Block$, #CR$) > 0)
  If *Dest\s = ""
    *Dest\s = Block$
  ElseIf *LastWasSpaced\i Or NeedsSpacing
    *Dest\s + #CR$ + #CR$ + Block$
  Else
    *Dest\s + #CR$ + Block$
  EndIf
  *LastWasSpaced\i = NeedsSpacing
EndProcedure
;
Procedure.s ST_GenerateSignatures(*Config.ST_ConfigStruct, WGadget = 0)
  ;
  ; Role: Collects all matching #PBBProcedure elements from ElementList(), orders them either
  ;       by compiler file/source position or alphabetically, and assembles the output in a
  ;       single linear pass over ProcItems() without relying on PBB_FileListing() cursors.
  ;
  Protected *SavedFile = @PBB_FileListing()
  Protected *SavedElem = @ElementList()
  Protected FileIdx, ModID, HasProc, HasComBlock, ProcStep, KeepProcKw
  Protected TotalSpaced, FileSpaced, PubSpaced, PrivSpaced
  Protected CurrentFile$, EntryBlock$, ModSection$, CommentTag$
  Protected TotalOutput.String, FileOutput.String, PubBlocks.String, PrivBlocks.String
  Protected NewList ProcItems.ST_ProcEntry()
  Protected NewList CommentBlocks.ST_CommentBlock()
  Protected NewMap FileOrderMap.i()
  Protected NewMap EmittedModules.i()
  ;
  KeepProcKw = Bool((*Config\Flags & #ST_KeepProcKeyword) <> 0)
  ;
  ; 1. Index permitted files in their chronological compiler inclusion order:
  ForEach PBB_FileListing()
    If *Config\ScopeFile = "" Or PBB_FileListing()\FileName$ = *Config\ScopeFile
      FileIdx + 1
      FileOrderMap(PBB_FileListing()\FileName$) = FileIdx
    EndIf
  Next
  ;
  ; 2. Collect target #PBBProcedure records from ElementList():
  ForEach ElementList()
    If ElementList()\Type = #PBBProcedure And ElementList()\StartingPos > 0 And
       ElementList()\EndingPos > ElementList()\StartingPos
      If (*Config\Flags & #ST_ExcludeModules) And ElementList()\EnclosingModuleID > 0
        Continue
      EndIf
      If FindMapElement(FileOrderMap(), ElementList()\FileName)
        FileIdx = FileOrderMap()
        AddElement(ProcItems())
        ProcItems()\FileOrder         = FileIdx
        ProcItems()\FileName          = ElementList()\FileName
        ProcItems()\RelFileName       = ST_GetDisplayFilePath(ElementList()\FileName)
        ProcItems()\StartingPos       = ElementList()\StartingPos
        ProcItems()\EndingPos         = ElementList()\EndingPos
        ProcItems()\EnclosingModuleID = ElementList()\EnclosingModuleID
        ProcItems()\ModuleName        = GetEnclosingModuleName(ElementList()\EnclosingModuleID)
        ProcItems()\IsDeclareModule   = ElementList()\IsDeclareModule
        ProcItems()\Used              = ElementList()\Used
        ProcItems()\PurifiedName      = ElementList()\PurifiedName
        ProcItems()\Declaration       = ElementList()\Declaration
        ProcItems()\UsageParents      = ElementList()\UsageParents
        ProcItems()\UsageChildren     = ElementList()\UsageChildren
        ProcItems()\OOE_UsagePos      = ElementList()\OOE_UsagePos
        If *Config\Flags & #ST_SortAlpha
          ProcItems()\SortKey = ElementList()\PurifiedName + #TAB$ + RSet(Str(FileIdx), 6, "0") + #TAB$ + RSet(Str(ElementList()\StartingPos), 10, "0")
        Else
          ProcItems()\SortKey = RSet(Str(FileIdx), 6, "0") + #TAB$ + RSet(Str(ElementList()\StartingPos), 10, "0")
        EndIf
      EndIf
    EndIf
  Next
  If *SavedElem
    ChangeCurrentElement(ElementList(), *SavedElem)
  EndIf
  ;
  If ListSize(ProcItems()) > 0
    SortStructuredList(ProcItems(), #PB_Sort_Ascending, OffsetOf(ST_ProcEntry\SortKey), #PB_String)
    ;
    ; 3. Assemble output in a single linear pass over ProcItems():
    If *Config\Flags & #ST_SortAlpha
      ForEach ProcItems()
        ProcStep + 1
        If IsGadget(WGadget) And (ProcStep = 1 Or ProcStep % 20 = 0)
          SetGadgetText(WGadget, GetTextFromCatalogPB("SignatureToolTitle") + #CR$ + ProcItems()\PurifiedName + "()")
          UpdateWindow_(GadgetID(WGadget))
        EndIf
        If ProcItems()\EnclosingModuleID > 0 And KeepProcKw = #False
          CommentTag$ = "[Module: " + ProcItems()\ModuleName + "]"
        Else
          CommentTag$ = ""
        EndIf
        ST_AppendSection(@TotalOutput,
                         ST_FormatProcBlock(@ProcItems(), *Config\Flags, Bool(ListSize(PBB_FileListing()) > 1), "", CommentTag$),
                         @TotalSpaced)
      Next
    Else
      HasProc = Bool(FirstElement(ProcItems()) <> 0)
      While HasProc
        CurrentFile$ = ProcItems()\FileName
        If IsGadget(WGadget)
          SetGadgetText(WGadget, GetTextFromCatalogPB("SignatureToolTitle") + #CR$ + GetFilePart(CurrentFile$))
          UpdateWindow_(GadgetID(WGadget))
        EndIf
        ;
        FileOutput\s = "; ===========================================================================" + #CR$ +
                       "; " + ProcItems()\RelFileName + #CR$ +
                       "; ==========================================================================="
        FileSpaced   = #True
        ;
        HasComBlock = #False
        If *Config\Flags & #ST_KeepStandaloneCom
          ST_CollectStandaloneComments(CurrentFile$, CommentBlocks())
          HasComBlock = Bool(FirstElement(CommentBlocks()) <> 0)
        EndIf
        ;
        ; Consume all procedures belonging to CurrentFile$ in ascending StartingPos order:
        While HasProc And ProcItems()\FileName = CurrentFile$
          While HasComBlock And CommentBlocks()\StartPos < ProcItems()\StartingPos
            ST_AppendSection(@FileOutput, CommentBlocks()\Text, @FileSpaced, #True)
            HasComBlock = Bool(NextElement(CommentBlocks()) <> 0)
          Wend
          ;
          If ProcItems()\EnclosingModuleID <= 0
            ST_AppendSection(@FileOutput, ST_FormatProcBlock(@ProcItems(), *Config\Flags, #False, ""), @FileSpaced)
          Else
            ; Group module procedures at the first encounter of each module.
            ; In 'Declare' (.pb) mode, PureBasic forbids standalone 'Module ... EndModule' blocks containing
            ; only 'Declare' headers without bodies, so private module declares are emitted as comments.
            ModID = ProcItems()\EnclosingModuleID
            If FindMapElement(EmittedModules(), Str(ModID)) = 0
              EmittedModules(Str(ModID)) = #True
              PubBlocks\s                = ""
              PrivBlocks\s               = ""
              PubSpaced                  = #False
              PrivSpaced                 = #False
              PushListPosition(ProcItems())
              ForEach ProcItems()
                If ProcItems()\EnclosingModuleID = ModID
                  If ProcItems()\IsDeclareModule
                    EntryBlock$ = ST_FormatProcBlock(@ProcItems(), *Config\Flags, #False, "  ")
                    ST_AppendSection(@PubBlocks, EntryBlock$, @PubSpaced)
                  ElseIf KeepProcKw
                    EntryBlock$ = ST_FormatProcBlock(@ProcItems(), *Config\Flags, #False, "  ")
                    ST_AppendSection(@PrivBlocks, EntryBlock$, @PrivSpaced)
                  Else
                    EntryBlock$ = ST_FormatProcBlock(@ProcItems(), *Config\Flags, #False, "", "[Private in Module " + ProcItems()\ModuleName + "]")
                    ST_AppendSection(@PrivBlocks, EntryBlock$, @PrivSpaced)
                  EndIf
                EndIf
              Next
              PopListPosition(ProcItems())
              ;
              ModSection$ = ""
              If PubBlocks\s <> "" Or KeepProcKw
                ModSection$ = "DeclareModule " + ProcItems()\ModuleName + #CR$
                If PubBlocks\s <> ""
                  ModSection$ + PubBlocks\s + #CR$
                EndIf
                ModSection$ + "EndDeclareModule"
              EndIf
              If PrivBlocks\s <> ""
                If ModSection$ <> ""
                  ModSection$ + #CR$ + #CR$
                EndIf
                If KeepProcKw
                  ModSection$ + "Module " + ProcItems()\ModuleName + #CR$ + PrivBlocks\s + #CR$ + "EndModule"
                Else
                  ModSection$ + PrivBlocks\s
                EndIf
              EndIf
              ST_AppendSection(@FileOutput, ModSection$, @FileSpaced, #True)
            EndIf
          EndIf
          ;
          HasProc = Bool(NextElement(ProcItems()) <> 0)
        Wend
        ;
        While HasComBlock
          ST_AppendSection(@FileOutput, CommentBlocks()\Text, @FileSpaced, #True)
          HasComBlock = Bool(NextElement(CommentBlocks()) <> 0)
        Wend
        ;
        ST_AppendSection(@TotalOutput, FileOutput\s, @TotalSpaced, #True)
      Wend
    EndIf
  EndIf
  ;
  If *SavedFile
    ChangeCurrentElement(PBB_FileListing(), *SavedFile)
  Else
    ResetList(PBB_FileListing())
  EndIf
  ;
  If TotalOutput\s = ""
    ProcedureReturn ""
  EndIf
  ProcedureReturn ReplaceString(TotalOutput\s, #CR$, #CRLF$) + #CRLF$
EndProcedure
;
Procedure SignatureTool()
  ;
  ; Role: Entry point for the Procedure Signature & Declaration Generator custom tool.
  ;       Ensures AST indexing is complete, prompts the user for generation options, builds
  ;       the declarations list, and dispatches it to the clipboard and/or target file.
  ;
  Protected Config.ST_ConfigStruct
  Protected WGadget, SavedOK
  Protected Output$, ConfirmMsg$ = ""
  ;
  If PBBFicPrincipalPB$ = "" Or ListSize(PBB_FileListing()) = 0
    AlertInPBBWindow(GetTextFromCatalogPB("NoElementFound"))
    ProcedureReturn #False
  EndIf
  ;
  If ExplorationState < #ExplorationState_Completed
    OnePassSetElements(#FinishCompletionNow)
  EndIf
  ;
  If SignatureToolOptionsWindow(@Config) = #False
    ProcedureReturn #False
  EndIf
  ;
  DisablePBBWindow()
  WGadget = OpenWaitWindow(GPBBGadgets\PBBWindow, GetTextFromCatalogPB("SignatureToolTitle"))
  Output$ = ST_GenerateSignatures(@Config, WGadget)
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
  If Config\Flags & #ST_OutClipboard
    SetClipboardText(Output$)
  EndIf
  ;
  If Config\Flags & #ST_OutSaveFile
    SavedOK = TexteDansFichier(Config\OutputFile, Output$, #PB_UTF8, 0, 1)
    If SavedOK = 0
      AlertInPBBWindow(GetTextFromCatalogPB("ErrorWithFile") + " " + Config\OutputFile)
      ProcedureReturn #False
    ElseIf Config\Flags & #ST_OutOpenFile
      RunProgram(Config\OutputFile)
      ProcedureReturn #True
    EndIf
  EndIf
  ;
  If (Config\Flags & #ST_OutClipboard) And SavedOK
    ConfirmMsg$ = ReplaceString(GetTextFromCatalogPB("ST_DoneBoth"), "%1", Config\OutputFile)
  ElseIf SavedOK
    ConfirmMsg$ = ReplaceString(GetTextFromCatalogPB("ST_DoneSaved"), "%1", Config\OutputFile)
  ElseIf Config\Flags & #ST_OutClipboard
    ConfirmMsg$ = GetTextFromCatalogPB("ST_DoneClipboard")
  EndIf
  ;
  If ConfirmMsg$ <> ""
    AlertInPBBWindow(ConfirmMsg$, GetTextFromCatalogPB("SignatureToolTitle"))
  EndIf
  ;
  ProcedureReturn #True
EndProcedure
;
RegisterCustomTool("InCatalog:SignatureToolTitle", "InCatalog:SignatureToolEx", @SignatureTool(), 0)
;}
;
; ===============================================================================
; 5. CHECK PROCEDURE VARIABLES TOOL (PROTECTED / UNUSED AUDIT)
; ===============================================================================
;- Check Procedure Variables (Protected / Unused)
;{
Structure CheckProcItem
  StartingPos.i
  EndingPos.i
  Name$
  ModuleName$
EndStructure
;
Procedure CheckProcOfOneFile(SourceFile$, TargetFile$, *CPV_ModList.String = 0, WGadget = 0)
  ;
  ; Role: Inspects all procedures inventoried in 'SourceFile$' via ListVariables() (preserving
  ;       module scope isolation), and splices missing 'Protected' declarations and unused-variable
  ;       remarks directly into the raw disk content below each procedure's header comments while
  ;       strictly preserving the original encoding, BOM, line endings, tabs, and indentation.
  ; Output: Returns #True if at least one procedure was updated and saved to 'TargetFile$', #False otherwise.
  ;
  Protected *CodeContent.String = GetPointedCodeFromFile(SourceFile$, #False)
  Protected SourceCode.String, ModifiedCode.String, NDeclare.String
  Protected VList.VariableListings
  Protected NewList FileProcs.CheckProcItem()
  Protected *SavedElem, *StringStartingAdr, *Char.Character, *Limit.Character, *LineStart.Character
  Protected cont = 0, ctr = 0, RawDelta = 0
  Protected HeaderEndPos, UnusedPos, LineStartPos, CodeStartPos, RawInsertPos
  Protected FileEncoding, HasBOM, LineEndingFormat
  Protected NDeclareAdd$, SpaceBefore$, FileIndentUnit$, txg$
  ;
  If *CodeContent = 0 Or *CodeContent\s = ""
    ProcedureReturn #False
  EndIf
  ;
  If IsGadget(WGadget)
    txg$ = GetGadgetText(WGadget)
    If Right(txg$, 1) <> #CR$
      txg$ + #CR$
    EndIf
    txg$ + "."
    SetGadgetText(WGadget, txg$)
    UpdateWindow_(GadgetID(WGadget))
  EndIf
  ;
  ; Collect all procedures (Procedure, ProcedureC, ProcedureDLL, ProcedureCDLL) and their owning module in SourceFile$:
  *SavedElem = @ElementList()
  ForEach ElementList()
    If ElementList()\Type = #PBBProcedure And ElementList()\FileName = SourceFile$ And
       ElementList()\StartingPos > 0 And ElementList()\EndingPos > ElementList()\StartingPos
      AddElement(FileProcs())
      FileProcs()\StartingPos = ElementList()\StartingPos
      FileProcs()\EndingPos   = ElementList()\EndingPos
      FileProcs()\Name$       = ElementList()\Name
      FileProcs()\ModuleName$ = GetEnclosingModuleName(ElementList()\EnclosingModuleID)
    EndIf
  Next
  If *SavedElem
    ChangeCurrentElement(ElementList(), *SavedElem)
  EndIf
  ;
  If ListSize(FileProcs()) = 0
    ProcedureReturn #False
  EndIf
  SortStructuredList(FileProcs(), #PB_Sort_Ascending, OffsetOf(CheckProcItem\StartingPos), TypeOf(CheckProcItem\StartingPos))
  ;
  ForEach FileProcs()
    ctr + 1
    If ctr = 10 And IsGadget(WGadget)
      ctr = 0
      If txg$
        txg$ + "."
        SetGadgetText(WGadget, txg$)
        UpdateWindow_(GadgetID(WGadget))
      EndIf
    EndIf
    ;
    If ListVariables(FileProcs()\Name$, VList, #PBBProcedure, FileProcs()\ModuleName$) And
       (VList\Unprotected <> "" Or VList\Unused <> "")
      HeaderEndPos = GetProcedureHeaderEndPos(*CodeContent, FileProcs()\StartingPos, FileProcs()\EndingPos)
      If HeaderEndPos = 0
        HeaderEndPos = FileProcs()\StartingPos
      EndIf
      ;
      ; Build insertion lines while checking only the current procedure bounds for an existing Unused remark:
      NDeclare\s = ""
      If VList\Unprotected <> ""
        NDeclare\s = "Protected " + Trim(StringField(VList\Unprotected, 2, ":")) + " ; ********* Added by PBBrowser." + #CR$
      EndIf
      If VList\Unused <> ""
        NDeclareAdd$ = "; " + VList\Unused + " ********* (PBBrowser comment)" + #CR$
        UnusedPos    = FindString(*CodeContent\s, NDeclareAdd$, HeaderEndPos)
        If UnusedPos = 0 Or UnusedPos >= FileProcs()\EndingPos
          NDeclare\s + NDeclareAdd$
        EndIf
      EndIf
      ;
      If NDeclare\s <> ""
        *StringStartingAdr = @*CodeContent\s
        *Char              = PosToChar(*StringStartingAdr, HeaderEndPos) + SizeOf(Character)
        *Limit             = PosToChar(*StringStartingAdr, FileProcs()\EndingPos)
        ;
        ; Advance past any trailing header comment to the first physical line below Procedure(...):
        While *Char < *Limit And *Char\c <> #CR And *Char\c <> 0
          GoNextChar(*Char)
        Wend
        If *Char < *Limit And *Char\c = #CR
          GoNextChar(*Char)
          ;
          ; Skip leading blank or ';' header comment lines so '; Role:' stays directly under Procedure(...):
          While *Char < *Limit And *Char\c <> 0
            *LineStart = *Char
            SkipSpaces(*Char)
            If *Char < *Limit And (*Char\c = #CR Or *Char\c = ';')
              While *Char < *Limit And *Char\c <> #CR And *Char\c <> 0
                GoNextChar(*Char)
              Wend
              If *Char < *Limit And *Char\c = #CR
                GoNextChar(*Char)
              EndIf
            Else
              Break
            EndIf
          Wend
          ;
          ; Load raw disk content once on the first modified procedure so BOM/FileHasBOM are captured safely:
          If cont = 0
            SourceCode\s = FileToText(SourceFile$, #PB_UTF8, 1)
            If SourceCode\s = ""
              ProcedureReturn #False
            EndIf
            FileEncoding     = BOM
            HasBOM           = FileHasBOM
            LineEndingFormat = DetectLineEndingAndTab(@SourceCode)
            ModifiedCode\s   = SourceCode\s
          EndIf
          ;
          ; Extract exact raw indentation (#TAB and/or spaces) from SourceCode, adding one unit on empty/flat bodies:
          LineStartPos = CharToPos(*LineStart)
          CodeStartPos = CharToPos(*Char)
          RawInsertPos = FilteredPosToRawPos(@SourceCode, LineStartPos)
          SpaceBefore$ = FastMid(SourceCode\s, RawInsertPos, FilteredPosToRawPos(@SourceCode, CodeStartPos) - RawInsertPos)
          If CodeStartPos = LineStartPos Or CodeStartPos >= FileProcs()\EndingPos - 12
            If FileIndentUnit$ = ""
              If LineEndingFormat & #LineEnding_IncludeTab
                FileIndentUnit$ = #TAB$
              Else
                FileIndentUnit$ = DetectFileIndentUnit(*CodeContent)
              EndIf
            EndIf
            SpaceBefore$ + FileIndentUnit$
          EndIf
          ;
          If *CPV_ModList
            If cont = 0
              *CPV_ModList\s + #CR$ + TargetFile$ + #CR$
            EndIf
            If FileProcs()\ModuleName$ <> ""
              *CPV_ModList\s + "    " + FileProcs()\ModuleName$ + "::"
            Else
              *CPV_ModList\s + "    "
            EndIf
            *CPV_ModList\s + FileProcs()\Name$ + ": " + Trim(ReplaceString(NDeclare\s, #CR$, " ")) + #CR$
          EndIf
          ;
          NDeclare\s     = SpaceBefore$ + ReplaceString(Left(NDeclare\s, Len(NDeclare\s) - 1), #CR$, #CR$ + SpaceBefore$) + #CR$
          ApplyLineEnding(@NDeclare, LineEndingFormat)
          ModifiedCode\s = FastLeft(ModifiedCode\s, RawInsertPos + RawDelta - 1) + NDeclare\s + FastMid(ModifiedCode\s, RawInsertPos + RawDelta)
          RawDelta + Len(NDeclare\s)
          cont = 1
        EndIf ; EndIf multi-line procedure body
      EndIf ; EndIf NDeclare\s <> ""
    EndIf ; EndIf ListVariables(...)
  Next ; Next FileProcs()
  ;
  If cont
    ProcedureReturn Bool(TexteDansFichier(TargetFile$, ModifiedCode\s, FileEncoding, 0, HasBOM) <> 0)
  EndIf
  ProcedureReturn #False
EndProcedure
;
Procedure CheckProcVar()
  ;
  ; Role: Audits local variable declarations across all project files selected by the user,
  ;       displaying a summary report of inserted 'Protected' lines and unused-variable remarks.
  ;
  Protected NewList FilesToChange.FilesToChangeStruct()
  Protected TargetDir.String, CPV_ModList.String
  Protected WGadget, Modified = 0, Result
  ;
  If ExplorationState < #ExplorationState_Completed
    OnePassSetElements(#FinishCompletionNow)
  EndIf
  ;
  Result = AskAndPrepareProjectCopy(FilesToChange(), @TargetDir)
  If Result = #PB_MessageRequester_Cancel
    ProcedureReturn #False
  EndIf
  ;
  DisablePBBWindow()
  WGadget = OpenWaitWindow(GPBBGadgets\PBBWindow, "")
  ;
  ForEach FilesToChange()
    If FileSize(FilesToChange()\TargetFile) > 1
      If IsGadget(WGadget)
        SetGadgetText(WGadget, GetTextFromCatalogPB("ProcVariablesTitle") + #CR$ + GetFilePart(FilesToChange()\TargetFile) + ".")
        UpdateWindow_(GadgetID(WGadget))
      EndIf
      If CheckProcOfOneFile(FilesToChange()\SourceFile, FilesToChange()\TargetFile, @CPV_ModList, WGadget)
        Modified = 1
      EndIf
    EndIf
  Next
  ;
  If IsGadget(WGadget)
    DestroyParentWindow(WGadget)
  EndIf
  ;
  If Modified
    AlertInPBBWindow(GetTextFromCatalogPB("FilesAreUpdated") + #CR$ + #CR$ + CPV_ModList\s)
    If TargetDir\s
      RunProgram(TargetDir\s)
    EndIf
    If Result = #PB_MessageRequester_No
      ExplorationState = #ExplorationState_MustRestart
    EndIf
  Else
    AlertInPBBWindow(GetTextFromCatalogPB("NoElementFound"), GetTextFromCatalogPB("ProcVariablesTitle"))
  EndIf
  ;
  EnablePBBWindow()
EndProcedure
;
RegisterCustomTool("InCatalog:ProcVariablesTitle", "InCatalog:ProcVariablesEx", @CheckProcVar())
;}
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 117
; FirstLine = 101
; Folding = ----
; EnableXP
; DPIAware