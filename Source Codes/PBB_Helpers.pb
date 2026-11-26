; **********************************************************************************
;
;                              PBBrowser_Helpers.pb
;
;        Foundational File, Parsing, Sequence & ElementList Helpers for PBBrowser
;                                  Nov. 20/2026
;
; **********************************************************************************
;
; ROLE & ARCHITECTURE:
; Autonomous foundational helper library for PBBrowser (zero forward Declarations).
; Provides path resolution, binary-search buffer diffing, CED/CIMF list serialization,
; low-level pointer/token parsing primitives, code statement extraction, compilation
; sequence construction, dual-mode encompassing container detection, O(1) ElementList()
; indexing, and project backup/line-ending utilities.
;
; FILE ORGANIZATION:
;   • Path & Project File Helpers            - Relative/absolute paths and .pbp file parsing.
;   • Line Numbers, Intervals & Buffer Diffs - SIMD line counting, bounds, jokers, binary diff.
;   • Helpers for the CED List               - Serializes/sorts segmented '\Continued' records.
;   • Inactive CompilerIf Zone Intervals     - O(1) sorted DeadZones() interval management and queries.
;   • Low-Level Parsing Primitives           - Pointer macros, token/string/parenthesis parsers.
;   • Identifier Purification & Code Cleanup - Name purification, '\Field' chains, comment stripping.
;   • Statement Extraction & Line Tokenizer  - Multi-line statement slicer and LocalToken builder.
;   • Virtual Coordinates & Code Sequences   - Compilation timeline and project-wide search iterator.
;   • Encompassing Element Detection         - Dual-mode container boundary lookup and caching.
;   • ElementList() & Map Index Management   - O(1) FixedID allocation, map indexing, lineage links.
;   • Project Backup & Line-Ending Helpers   - Safe project copy and CRLF/LF/CR/TAB restoration.
;
; **********************************************************************************
;
;-                         Path & project file helpers
;
; **********************************************************************************
;
Procedure.s CorrectAddrSyntaxe(Adress$)
  ;
  ; Role: Normalizes file path syntax to match how the PureBasic compiler interprets path separators and relative dots.
  ;
  Protected mAdress$
  ;
  Adress$ = ReplaceString(Adress$, "/", "\")
  Repeat
    mAdress$ = Adress$
    Adress$  = ReplaceString(ReplaceString(Adress$, "\.\", "\"), "...\", "..\")
    If Len(Adress$) > 2
      ; Preserve any leading UNC '\\' prefix while collapsing duplicate internal backslashes
      Adress$ = Left(Adress$, 1) + ReplaceString(Mid(Adress$, 2), "\\", "\")
    EndIf
  Until mAdress$ = Adress$
  If Left(Adress$, 2) = ".\"
    Adress$ = Mid(Adress$, 3)
  EndIf
  ProcedureReturn Adress$
EndProcedure
;
Procedure   GetRootPrefixLen(Path$)
  ;
  ; Role: Returns the mandatory root prefix length (including trailing '\') of a normalized
  ;       Windows drive path ("C:\" -> 3) or UNC share ("\\Server\Share\" -> N), or 0 if relative.
  ;
  Protected SlashPos
  ;
  If Mid(Path$, 2, 2) = ":\" Or (Len(Path$) = 2 And Right(Path$, 1) = ":")
    ProcedureReturn 3
  ElseIf Left(Path$, 2) = "\\"
    SlashPos = FindString(Path$, "\", 3)
    If SlashPos : ProcedureReturn FindString(Path$ + "\", "\", SlashPos + 1) : EndIf
  EndIf
  ProcedureReturn 0
EndProcedure
;
Procedure.s FullAddressToRelativeAddress(Addr$, RefAddr$)
  ;
  ; Role: Converts an absolute disk or UNC path into a backslash-normalized path relative to 'RefAddr$'
  ;       (where 'RefAddr$' is the reference file or directory path). Returns 'Addr$' unchanged when
  ;       'Addr$' is a web URL or resides on a different drive or UNC share.
  ;
  Protected PosInText, RootLen, HadTrailingSlash
  Protected RefLow$, AddrLow$
  ;
  If Addr$ = "" Or RefAddr$ = "" Or LCase(FastLeft(Addr$, 7)) = "http://" Or
     LCase(FastLeft(Addr$, 8)) = "https://" Or LCase(FastLeft(Addr$, 4)) = "www."
    ProcedureReturn Addr$
  EndIf
  ;
  Addr$            = CorrectAddrSyntaxe(Addr$)
  HadTrailingSlash = Bool(Right(Addr$, 1) = "\")
  RefAddr$         = GetPathPart(CorrectAddrSyntaxe(RefAddr$))
  RootLen          = GetRootPrefixLen(Addr$)
  If RootLen = 0 Or LCase(Left(Addr$ + "\", RootLen)) <> LCase(Left(RefAddr$, RootLen))
    ProcedureReturn Addr$
  EndIf
  ;
  RefLow$   = LCase(RefAddr$)
  AddrLow$  = LCase(Addr$)
  PosInText = 1
  While PosInText <= Len(RefLow$) And PosInText <= Len(AddrLow$) And FastMid(RefLow$, PosInText, 1) = FastMid(AddrLow$, PosInText, 1)
    PosInText + 1
  Wend
  ;
  If PosInText > Len(AddrLow$) And FastMid(RefLow$, PosInText, 1) = "\"
    ; 'Addr$' is an exact directory ancestor of 'RefAddr$' without a trailing backslash
    Addr$    = ""
    RefAddr$ = Mid(RefAddr$, PosInText + 1)
  Else
    ; Rewind to the last shared directory separator so partial folder name matches are not split
    While PosInText > RootLen And FastMid(Addr$, PosInText - 1, 1) <> "\"
      PosInText - 1
    Wend
    Addr$    = Mid(Addr$, PosInText)
    RefAddr$ = Mid(RefAddr$, PosInText)
  EndIf
  ;
  While RefAddr$
    PosInText = FindString(RefAddr$, "\")
    If PosInText
      Addr$    = "..\" + Addr$
      RefAddr$ = Mid(RefAddr$, PosInText + 1)
    Else
      RefAddr$ = ""
    EndIf
  Wend
  ;
  If Addr$ = ""
    If HadTrailingSlash : ProcedureReturn ".\" : Else : ProcedureReturn "." : EndIf
  ElseIf HadTrailingSlash = #False And Right(Addr$, 1) = "\"
    ProcedureReturn Left(Addr$, Len(Addr$) - 1)
  EndIf
  ProcedureReturn Addr$
EndProcedure
;
Procedure.s RelativeAddressToFullAddress(RelAddr$, RefAddr$)
  ;
  ; Role: Decodes IncludePath and IncludeFile values by converting a relative path
  ;       into an absolute disk or UNC path anchored on 'RefAddr$'.
  ;
  Protected NewAddr$, PosInText, pf, RootLen, HadTrailingSlash
  ;
  RefAddr$ = CorrectAddrSyntaxe(RefAddr$)
  RelAddr$ = CorrectAddrSyntaxe(RelAddr$)
  If RelAddr$ = "" : ProcedureReturn "" : EndIf
  HadTrailingSlash = Bool(Right(RelAddr$, 1) = "\")
  ;
  If GetRootPrefixLen(RelAddr$) > 0
    NewAddr$ = RelAddr$
  ElseIf Left(RelAddr$, 1) = "\"
    ; Root-relative path anchored on the drive or UNC share root of RefAddr$
    RootLen = GetRootPrefixLen(RefAddr$)
    If RootLen > 0
      NewAddr$ = Left(RefAddr$ + "\", RootLen - 1) + RelAddr$
    Else
      NewAddr$ = RelAddr$
    EndIf
  Else
    NewAddr$ = CorrectAddrSyntaxe(GetPathPart(RefAddr$) + RelAddr$)
  EndIf
  ;
  ; Append a temporary trailing '\' so terminal '\..' segments are resolved identically to '\..\'
  RootLen  = GetRootPrefixLen(NewAddr$)
  NewAddr$ = Left(NewAddr$, 1) + ReplaceString(Mid(NewAddr$ + "\", 2), "\\", "\")
  Repeat
    PosInText = FindString(NewAddr$, "\..\")
    If PosInText
      pf = PosInText + 4
      If PosInText <= RootLen
        ; Already at the drive or UNC share root: clamp without stripping the root prefix
        PosInText = RootLen + 1
      Else
        While PosInText > RootLen + 1 And FastMid(NewAddr$, PosInText - 1, 1) <> "\"
          PosInText - 1
        Wend
      EndIf
      NewAddr$ = FastLeft(NewAddr$, PosInText - 1) + FastMid(NewAddr$, pf)
    EndIf
  Until PosInText = 0
  ;
  If HadTrailingSlash = #False And Len(NewAddr$) > RootLen And Right(NewAddr$, 1) = "\"
    NewAddr$ = Left(NewAddr$, Len(NewAddr$) - 1)
  EndIf
  ProcedureReturn NewAddr$
EndProcedure
;
Procedure.s ListProjectsFiles(ProjectAddr$)
  ;
  ; Role: Parses the PureBasic project file 'ProjectAddr$' (.pbp) and returns its list of existing files separated by #CR$.
  ;
  Protected ListOfFiles$, FileName$
  Protected ProjectDescr$ = FileToText(ProjectAddr$)
  Protected PosInProject, mPosInProject, pf
  ;
  Repeat
    PosInProject = FindString(ProjectDescr$, "<file name", PosInProject)
    If PosInProject
      PosInProject + Len("<file name")
      mPosInProject = PosInProject
      ; Look for the opening and closing quotes of the 'name' attribute
      PosInProject = FindString(ProjectDescr$, #DOUBLEQUOTE$, PosInProject)
      If PosInProject And (PosInProject - mPosInProject) < 10
        PosInProject + 1
        pf = FindString(ProjectDescr$, #DOUBLEQUOTE$, PosInProject)
        If pf
          FileName$ = RelativeAddressToFullAddress(FastMid(ProjectDescr$, PosInProject, pf - PosInProject), ProjectAddr$)
          If FileSize(FileName$) <> -1
            ListOfFiles$ + FileName$ + #CR$
          EndIf
        EndIf
      EndIf
    EndIf
  Until PosInProject = 0
  ProcedureReturn ListOfFiles$
EndProcedure
;
; **********************************************************************************
;
;-                   Line numbers, intervals & buffer diffs
;
; **********************************************************************************
;
Procedure.s AddIndexToList(NumIndex, IndexList$)
  ;
  ; Role: Appends 'NumIndex' to a slash-delimited index list if not already present.
  ;       Used for updating the \UsageParents and \UsageChildren fields of ElementList().
  ;
  Protected EPIndex$ = Str(NumIndex) + "/"
  ;
  If IndexList$ = "" : IndexList$ = "/" : EndIf
  If FindString(IndexList$, "/" + EPIndex$) = 0
    IndexList$ + EPIndex$
  EndIf
  If IndexList$ = "/" : IndexList$ = "" : EndIf
  ProcedureReturn IndexList$
EndProcedure
;
Procedure.s RemoveIndexFromList(NumIndex, IndexList$)
  ;
  ; Role: Removes 'NumIndex' from a slash-delimited index list and returns "" when no indices remain.
  ;
  If IndexList$ <> ""
    IndexList$ = ReplaceString(IndexList$, "/" + Str(NumIndex) + "/", "/")
    If IndexList$ = "/" : IndexList$ = "" : EndIf
  EndIf
  ProcedureReturn IndexList$
EndProcedure
;
Procedure   GetNoLineFromStrAndPos(*CodeString.String, PosInCode)
  ;
  ; Role: Returns the 1-based line number corresponding to 'PosInCode' within '*CodeString'.
  ;       Uses an in-place null-termination trick to leverage PureBasic's SIMD-optimized
  ;       native 'CountString()' without allocating temporary substrings on the heap.
  ;
  Protected *CharAtPos.Character
  Protected OldChar.c, LineNumber, BufferLen
  ;
  If *CodeString = 0 Or *CodeString\s = ""
    ProcedureReturn 0
  EndIf
  If PosInCode <= 1
    ProcedureReturn 1
  EndIf
  ;
  BufferLen = Len(*CodeString\s)
  If PosInCode >= BufferLen
    ; Position is at or past the end of the buffer: count directly without writing past the null terminator
    ProcedureReturn CountString(*CodeString\s, #CR$) + 1
  EndIf
  ;
  ; Temporarily place a null terminator right after PosInCode so characters 1 to PosInCode are counted
  *CharAtPos   = @*CodeString\s + PosInCode * SizeOf(Character)
  OldChar      = *CharAtPos\c
  *CharAtPos\c = 0
  LineNumber   = CountString(*CodeString\s, #CR$) + 1
  *CharAtPos\c = OldChar
  ;
  ProcedureReturn LineNumber
EndProcedure
;
Procedure   GetNoLineFromFilenameAndPos(Filename$, PosInCode)
  ;
  ; Role: Retrieves the in-memory code buffer for 'Filename$' and returns the 1-based line number for 'PosInCode'.
  ;
  Protected *Temp.String = GetPointedCodeFromFile(Filename$)
  ;
  If *Temp
    ProcedureReturn GetNoLineFromStrAndPos(*Temp, PosInCode)
  EndIf
EndProcedure
;
Procedure   BoundariesVersusLimits(StartBoundary, EndBoundary, StartLimit, EndLimit)
  ;
  ; Role: Evaluates the spatial relationship of an interval [StartBoundary, EndBoundary]
  ;       relative to a target range [StartLimit, EndLimit].
  ; Output: Returns #BeforeLimits, #AfterLimits, or #InsideLimits (when overlapping or enclosed).
  ;
  Enumeration BoundaryResults
    #BeforeLimits
    #AfterLimits
    #InsideLimits
  EndEnumeration
  ;
  If EndBoundary < StartLimit Or (EndBoundary = StartLimit And StartBoundary < StartLimit)
    ProcedureReturn #BeforeLimits
  ElseIf StartBoundary > EndLimit Or (StartBoundary = EndLimit And EndBoundary > EndLimit)
    ProcedureReturn #AfterLimits
  Else
    ProcedureReturn #InsideLimits
  EndIf
EndProcedure
;
Macro       AddStringElement(TargetVar, NewElement, Separator)
  ;
  ; Role: Safely concatenates strings by inserting a separator only if
  ;       the target variable already contains data. Prevents trailing/leading separators.
  ;
  If Len(NewElement)
    If TargetVar
      TargetVar + Separator + NewElement
    Else
      TargetVar = NewElement
    EndIf
  EndIf
EndMacro
;
Procedure   CompareExpressionsWithJoker(Exp1$, Exp2$)
  ;
  ; Role: Compares two expressions where one of them may contain leading and/or trailing
  ;       wildcard jokers (#PBBJoker$). Returns #True if they match, #False otherwise.
  ;
  EnumerationBinary
    #JokerLeft
    #JokerRight
  EndEnumeration
  ;
  Protected CompareMode ; Bitmask indicating whether #PBBJoker$ was stripped from the left and/or right of Exp1$
  ;
  If Exp1$ = Exp2$
    ProcedureReturn #True
  EndIf
  ;
  If Left(Exp1$, Len(#PBBJoker$)) = #PBBJoker$
    CompareMode | #JokerLeft
    Exp1$ = Mid(Exp1$, Len(#PBBJoker$) + 1)
  EndIf
  If Right(Exp1$, Len(#PBBJoker$)) = #PBBJoker$
    CompareMode | #JokerRight
    Exp1$ = Left(Exp1$, Len(Exp1$) - Len(#PBBJoker$))
  EndIf
  ;
  If CompareMode
    If Exp1$ = ""
      ProcedureReturn #True
    ElseIf (CompareMode & #JokerLeft) And (CompareMode & #JokerRight)
      ProcedureReturn Bool(FindString(Exp2$, Exp1$) > 0)
    ElseIf CompareMode & #JokerLeft
      ProcedureReturn Bool(Right(Exp2$, Len(Exp1$)) = Exp1$)
    Else
      ProcedureReturn Bool(Left(Exp2$, Len(Exp1$)) = Exp1$)
    EndIf
  ElseIf FindString(Exp2$, #PBBJoker$)
    ; Only recurse when Exp2$ actually contains a joker to prevent infinite recursion on unequal strings
    ProcedureReturn CompareExpressionsWithJoker(Exp2$, Exp1$)
  EndIf
  ;
  ProcedureReturn #False
EndProcedure
;
Procedure.s FindNoLine(StrRef$, PosInText)
  ;
  ; Role: Extracts the line number following a clicked filename on the same line of 'StrRef$',
  ;       skipping any intermediate '(Container)' mention before scanning for digits.
  ;
  Protected *Char.Character, *Start.Character
  Protected CharSize = SizeOf(Character)
  ;
  If PosInText < 1 Or PosInText > Len(StrRef$)
    ProcedureReturn ""
  EndIf
  ;
  *Char = @StrRef$ + (PosInText - 1) * CharSize
  While *Char\c <> 0 And *Char\c <> #CR
    If *Char\c = '('
      ; Skip parenthesized container names (e.g., "(ParameterType)" or "(Proc2())") that may contain digits
      While *Char\c <> 0 And *Char\c <> #CR And *Char\c <> ')'
        *Char + CharSize
      Wend
    ElseIf *Char\c >= '0' And *Char\c <= '9'
      *Start = *Char
      While *Char\c >= '0' And *Char\c <= '9'
        *Char + CharSize
      Wend
      ProcedureReturn PeekS(*Start, (*Char - *Start) / CharSize)
    EndIf
    If *Char\c <> 0 And *Char\c <> #CR
      *Char + CharSize
    EndIf
  Wend
  ;
  ProcedureReturn ""
EndProcedure
;
Procedure   FindStartOfDifference(*OldContent.String, *NewContent.String, EndExam = 0)
  ;
  ; Role: Searches for the 1-based position where '*OldContent\s' and '*NewContent\s' first differ
  ;       using a binary search with direct 'CompareMemory()' calls (zero heap allocations).
  ; Parameters:
  ;   EndExam - Optional 1-based upper character limit (0 searches the full string length).
  ; Output: Returns the 1-based index of the first differing character, or 0 if identical.
  ;
  Protected Low = 1, High, Mid, DiffPos
  Protected *OldPtr, *NewPtr
  Protected CharSize = SizeOf(Character)
  ;
  If *OldContent = 0 Or *NewContent = 0 : ProcedureReturn 0 : EndIf
  *OldPtr = @*OldContent\s
  *NewPtr = @*NewContent\s
  If *OldPtr = *NewPtr : ProcedureReturn 0 : EndIf
  ;
  High = max(Len(*OldContent\s), Len(*NewContent\s)) + 1
  If EndExam > 0 And EndExam < High
    High = EndExam
  EndIf
  ;
  While Low <= High
    Mid = (Low + High) / 2
    If CompareMemory(*OldPtr + (Low - 1) * CharSize, *NewPtr + (Low - 1) * CharSize, (Mid - Low + 1) * CharSize)
      Low = Mid + 1
    Else
      DiffPos = Mid
      High    = Mid - 1
    EndIf
  Wend
  ;
  ProcedureReturn DiffPos
EndProcedure
;
Procedure   FindEndOfDifference(*OldContent.String, *NewContent.String, StartPos, EndExamOld, EndExamNew)
  ;
  ; Role: Searches backwards from 'EndExamOld' and 'EndExamNew' down to 'StartPos' using binary search
  ;       and 'CompareMemory()' to count matching trailing characters without heap allocations.
  ; Output: Returns the number of identical trailing characters, or 0 if none match.
  ;
  Protected Low = 1, High, Mid, MatchingChars
  Protected *OldPtr, *NewPtr
  Protected CharSize = SizeOf(Character)
  ;
  High = min(EndExamOld - StartPos + 1, EndExamNew - StartPos + 1)
  If High <= 0 : ProcedureReturn 0 : EndIf
  ;
  *OldPtr = @*OldContent\s + (EndExamOld - 1) * CharSize
  *NewPtr = @*NewContent\s + (EndExamNew - 1) * CharSize
  ;
  While Low <= High
    Mid = (Low + High) / 2
    If CompareMemory(*OldPtr - (Mid - 1) * CharSize, *NewPtr - (Mid - 1) * CharSize, Mid * CharSize)
      MatchingChars = Mid
      Low           = Mid + 1
    Else
      High = Mid - 1
    EndIf
  Wend
  ;
  ProcedureReturn MatchingChars
EndProcedure
;
; ----------------------------------------------------------------------------------
;                     ---- IDE options footer boundary ----
; ----------------------------------------------------------------------------------
;
Procedure   FindIDENotations(*Content.String)
  ;
  ; Role: Examines the trailing characters of '*Content' to locate the start position of the PureBasic IDE Options block.
  ;       Starting from the last character, moves backward line by line until a non-comment line (not starting with ";") is encountered.
  ;
  Protected *StartPtr.Character, *Ptr.Character, *PrevPtr.Character
  Protected Length, ScanStartPos, Pos, CutPos
  Protected CharSize = SizeOf(Character)
  ;
  If *Content = 0 Or *Content\s = "" : ProcedureReturn 1 : EndIf
  ;
  Length    = Len(*Content\s)
  *StartPtr = @*Content\s
  ;
  ScanStartPos = Length - 2000
  If ScanStartPos < 1 : ScanStartPos = 1 : EndIf
  ;
  CutPos = Length + 1
  Pos    = Length
  *Ptr   = *StartPtr + (Pos - 1) * CharSize
  ;
  ; Scan backwards line by line directly in memory
  While Pos >= ScanStartPos
    ; Skip trailing CR / LF
    While Pos >= ScanStartPos And (*Ptr\c = #CR Or *Ptr\c = #LF)
      Pos - 1
      *Ptr - CharSize
    Wend
    If Pos < ScanStartPos : Break : EndIf
    ;
    ; Move to the start of the current line
    While Pos > ScanStartPos
      *PrevPtr = *Ptr - CharSize
      If *PrevPtr\c = #CR Or *PrevPtr\c = #LF : Break : EndIf
      Pos - 1
      *Ptr - CharSize
    Wend
    ;
    If *Ptr\c = ';'
      CutPos = Pos
    Else
      Break
    EndIf
    Pos - 1
    *Ptr - CharSize
  Wend
  ;
  Pos = FindString(*Content\s, "; IDE Options = PureBasic", CutPos)
  If Pos : CutPos = Pos : EndIf
  ;
  ProcedureReturn CutPos
EndProcedure
;
; **********************************************************************************
;
;-                          Helpers for the CED List
;
; **********************************************************************************
;
; In some cases (enumerations, procedures, or local variable elements),
; the declaration of the element can be segmented and continued across different
; segments of code or even different files.
; The exploration procedure stores those different segment positions
; in the '\Continued' string field of the element.
; That format is convenient for compact storage, but not for examining the list of positions.
; Therefore, the following structure and procedures are used to populate 'CEDList()'
; from the content of '\Continued' or vice versa.
;
Structure CED_Structure
  Filename.s        ; File in which the declaration is made.
  DirectContainer.l ; FixedID of the procedure inside which the declaration is made,
                    ;   or -1 if the current element was declared outside a procedure.
  StartingPos.l     ; Start position of the declaration in the code.
  EndingPos.l       ; End position of the declaration in the code.
  SharedDotType.s   ; For "Shared" variables only, contains "None" or the dot-type
                    ;   specified by the "Shared" declaration.
  Declaration.s     ; The declaration statement itself.
EndStructure
;
Procedure   CED_RecordToCEDList(Record$, List CEDList.CED_Structure())
  ;
  ; Role: Parses a single tab-delimited continuation record string and appends it to 'CEDList()'.
  ;
  If Record$
    AddElement(CEDList())
    CEDList()\Filename        = StringField(Record$, 1, #TAB$)
    CEDList()\StartingPos     = Val(StringField(Record$, 2, #TAB$))
    CEDList()\EndingPos       = Val(StringField(Record$, 3, #TAB$))
    CEDList()\Declaration     = StringField(Record$, 4, #TAB$)
    CEDList()\DirectContainer = Val(StringField(Record$, 5, #TAB$))
    CEDList()\SharedDotType   = StringField(Record$, 6, #TAB$)
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure   CED_StringToList(Continued$, List CEDList.CED_Structure())
  ;
  ; Role: Populates 'CEDList()' by splitting a multi-line '\Continued' string into individual records.
  ;
  Protected CIndex, ContinuedRecord$
  ;
  ClearList(CEDList())
  Repeat
    CIndex + 1
    ContinuedRecord$ = StringField(Continued$, CIndex, #CR$)
    CED_RecordToCEDList(ContinuedRecord$, CEDList())
  Until ContinuedRecord$ = ""
EndProcedure
;
Procedure.s CED_CreateRecord(Filename$, StartingPos, EndingPos, Declaration$, DirectContainer = -1, SharedDotType$ = "")
  ;
  ; Role: Formats a single continuation record line from its individual properties.
  ;
  ProcedureReturn Filename$ + #TAB$ + Str(StartingPos) + #TAB$ + Str(EndingPos) + #TAB$ + Declaration$ + #TAB$ +
                  Str(DirectContainer) + #TAB$ + SharedDotType$ + #CR$
EndProcedure
;
Procedure.s CED_ListToString(List CEDList.CED_Structure())
  ;
  ; Role: Serializes all elements of 'CEDList()' back into a multi-line string for storage in '\Continued'.
  ;
  Protected ContinuedRecord$
  ;
  ForEach CEDList()
    ContinuedRecord$ + CED_CreateRecord(CEDList()\Filename, CEDList()\StartingPos, CEDList()\EndingPos,
                                        CEDList()\Declaration, CEDList()\DirectContainer,
                                        CEDList()\SharedDotType)
  Next
  ProcedureReturn ContinuedRecord$
EndProcedure
;
Procedure   CED_ToMainDeclaration(List CEDList.CED_Structure())
  ;
  ; Role: Promotes the current 'CEDList()' entry to become the main declaration of the current element
  ;       in 'ElementList()', then removes the promoted entry from 'CEDList()'.
  ;
  Protected SharedDotType$, DeclLCase$, CandType$, RootName$
  Protected DotPos, EndPos
  ;
  If ListIndex(CEDList()) >= 0 And ListIndex(ElementList()) >= 0
    If ElementList()\Type = #PBBVariable And CEDList()\SharedDotType <> ""
      ; Inspect surviving continuations to preserve any explicit dot-type declared on a 'Shared' variable
      PushListPosition(CEDList())
      ForEach CEDList()
        If CEDList()\SharedDotType <> "" And CEDList()\SharedDotType <> "None"
          SharedDotType$ = LCase(CEDList()\SharedDotType)
          Break
        ElseIf CEDList()\SharedDotType = ""
          DeclLCase$ = LCase(CEDList()\Declaration)
          DotPos     = FindString(DeclLCase$, ElementList()\PurifiedName + ".")
          If DotPos > 0
            DotPos + Len(ElementList()\PurifiedName) + 1
          Else
            DotPos = FindString(StringField(DeclLCase$, 1, " "), ".")
            If DotPos > 0 : DotPos + 1 : EndIf
          EndIf
          If DotPos > 0
            While Mid(DeclLCase$, DotPos, 1) = " " : DotPos + 1 : Wend
            EndPos = DotPos
            While EndPos <= Len(DeclLCase$) And
                  (Mid(DeclLCase$, EndPos, 2) = "::" Or IsPBSeparator(Asc(Mid(DeclLCase$, EndPos, 1))) = #False)
              If Mid(DeclLCase$, EndPos, 2) = "::" : EndPos + 2 : Else : EndPos + 1 : EndIf
            Wend
            CandType$ = Mid(DeclLCase$, DotPos, EndPos - DotPos)
            If CandType$ <> ""
              SharedDotType$ = CandType$
              Break
            EndIf
          EndIf
        EndIf
      Next
      PopListPosition(CEDList())
      ;
      RootName$ = StringField(StringField(ElementList()\Name, 1, "("), 1, ".")
      If Right(RootName$, 1) = "$"
        ElementList()\Name             = RootName$ + Left("()", Bool(ElementList()\HasParentheses) * 2)
        ElementList()\VarDotType_LCase = ""
        ElementList()\VarAddedDotType  = #False
      ElseIf SharedDotType$ = ""
        ElementList()\Name             = RootName$ + ".i" + Left("()", Bool(ElementList()\HasParentheses) * 2)
        ElementList()\VarDotType_LCase = "i"
        ElementList()\VarAddedDotType  = #True
      Else
        CandType$ = SharedDotType$
        If FindString(CandType$, "::") : CandType$ = Mid(CandType$, FindString(CandType$, "::") + 2) : EndIf
        ElementList()\Name             = RootName$ + "." + CandType$ + Left("()", Bool(ElementList()\HasParentheses) * 2)
        ElementList()\VarDotType_LCase = SharedDotType$
        ElementList()\VarAddedDotType  = #False
      EndIf
    EndIf
    ;
    ElementList()\FileName          = CEDList()\Filename
    ElementList()\StartingPos       = CEDList()\StartingPos
    ElementList()\EndingPos         = CEDList()\EndingPos
    ElementList()\Declaration       = CEDList()\Declaration
    ElementList()\LocalVarContainer = CEDList()\DirectContainer
    If ElementList()\Type = #PBBProcedure And LCase(Left(CEDList()\Declaration, 7)) = "declare"
      ElementList()\Type = #PBBDeclare
    ElseIf ElementList()\Type = #PBBModule
      ElementList()\IsDeclareModule = Bool(LCase(Left(CEDList()\Declaration, 13)) = "declaremodule")
    EndIf
  EndIf
  DeleteElement(CEDList())
EndProcedure
;
Procedure   CED_SortBasedOnPositions()
  ;
  ; Role: Sorts adjacent records sharing the same file in 'ElementList()\Continued'
  ;       by ascending 'StartingPos' until complete stabilization.
  ;
  Protected NewList CEDList.CED_Structure()
  Protected ActualElement, NextElement, NextStartingPos, Swapped
  Protected NextFileName$
  ;
  CED_StringToList(ElementList()\Continued, CEDList())
  If ListSize(CEDList()) > 1
    Repeat
      Swapped = #False
      ForEach CEDList()
        ActualElement = @CEDList()
        NextElement   = NextElement(CEDList())
        If NextElement
          NextFileName$   = CEDList()\Filename
          NextStartingPos = CEDList()\StartingPos
          PreviousElement(CEDList())
          If NextFileName$ = CEDList()\Filename And NextStartingPos < CEDList()\StartingPos
            SwapElements(CEDList(), ActualElement, NextElement)
            Swapped = #True
          EndIf
        EndIf
      Next
    Until Swapped = #False
    ElementList()\Continued = CED_ListToString(CEDList())
  EndIf
EndProcedure
;
; **********************************************************************************
;
;-               Helpers for inactive CompilerIf / CompilerSelect zones
;
; **********************************************************************************
;
Procedure AddDeadZoneRange(List DeadZones.DeadZoneStruct(), StartPos, EndPos)
  ;
  ; Role: Appends or coalesces a monotonically increasing inactive span [StartPos, EndPos]
  ;       at the tail of 'DeadZones()' in O(1) time.
  ;
  If StartPos >= 1 And EndPos >= StartPos
    If LastElement(DeadZones()) And StartPos <= DeadZones()\EndingPos + 1
      If EndPos > DeadZones()\EndingPos
        DeadZones()\EndingPos = EndPos
      EndIf
    Else
      AddElement(DeadZones())
      DeadZones()\StartingPos = StartPos
      DeadZones()\EndingPos   = EndPos
    EndIf
  EndIf
EndProcedure
;
Procedure.i SkipDeadZoneInCurrentFile(PosInCode)
  ;
  ; Role: If 'PosInCode' falls inside an excluded file or an inactive zone of the currently selected
  ;       PBB_FileListing() entry, returns the first position past that zone; otherwise returns 'PosInCode'.
  ;
  If PosInCode > 0 And ListIndex(PBB_FileListing()) >= 0
    If PBB_FileListing()\ExcludedByCompilerIf
      ProcedureReturn $7FFFFFFF
    EndIf
    ForEach PBB_FileListing()\DeadZones()
      If PosInCode < PBB_FileListing()\DeadZones()\StartingPos
        Break
      ElseIf PosInCode <= PBB_FileListing()\DeadZones()\EndingPos
        ProcedureReturn PBB_FileListing()\DeadZones()\EndingPos + 1
      EndIf
    Next
  EndIf
  ProcedureReturn PosInCode
EndProcedure
;
Procedure.i IsPosInDeadZone(FileName$, PosInCode)
  ;
  ; Role: Returns #True when 'PosInCode' inside 'FileName$' lies within an inactive
  ;       CompilerIf / CompilerSelect zone (or an excluded file), preserving PBB_FileListing() cursor.
  ;
  Protected *SavedFile, Result = #False
  ;
  If FileName$ <> "" And PosInCode > 0
    If ListIndex(PBB_FileListing()) >= 0 : *SavedFile = @PBB_FileListing() : EndIf
    If IsFileInMemory(FileName$)
      Result = Bool(SkipDeadZoneInCurrentFile(PosInCode) > PosInCode)
    EndIf
    If *SavedFile : ChangeCurrentElement(PBB_FileListing(), *SavedFile) : EndIf
  EndIf
  ProcedureReturn Result
EndProcedure
;
; **********************************************************************************
;
;-                        Low-level parsing primitives
;
; **********************************************************************************
;
Enumeration ParsingStateEnum
  #Parsing_Start
  #Parsing_NewToken
  #Parsing_EndOfLine
  #Parsing_EndOfFile
EndEnumeration
;
Enumeration ParsingModeEnum
  #ParsingMode_OnePass
  #ParsingMode_ListVariables
EndEnumeration
;
Macro       GoNextChar(Chr)
  ;
  ; Role: Advances a character pointer to the next character in memory.
  ;
  Chr + SizeOf(Character)
EndMacro
;
Macro       CharToPos(Char)
  ;
  ; Role: Converts a memory character pointer into a 1-based character index relative to *StringStartingAdr.
  ;
  ((Char - *StringStartingAdr) / SizeOf(Character) + 1)
EndMacro
;
Macro       PosToChar(StringStartingAdr, Pos)
  ;
  ; Role: Converts a 1-based character index into a memory character pointer relative to StringStartingAdr.
  ;
  (StringStartingAdr + ((Pos - 1) * SizeOf(Character)))
EndMacro
;
Macro       ComputeLenFromChars(CharStartAdr, CharEndAdr)
  ;
  ; Role: Computes the number of characters between two memory character pointers.
  ;
  (((CharEndAdr) - (CharStartAdr)) / SizeOf(Character))
EndMacro
;
Macro       IsStandaloneLabel(TokenLCase, CharPtr)
  ;
  ; Role: Evaluates to #True if 'CharPtr' points to a single colon ':' (excluding '::')
  ;       and 'TokenLCase' is not a PureBasic control-flow keyword that accepts a trailing colon.
  ;
  Bool(CharPtr\c = ':' And PeekC(CharPtr + SizeOf(Character)) <> ':' And TokenLCase <> "else" And
       TokenLCase <> "default" And TokenLCase <> "break" And TokenLCase <> "compilerelse" And
       TokenLCase <> "compilerdefault")
EndMacro
;
Macro       SkipSpaces(Chr)
  ;
  ; Role: Advances a character pointer past any contiguous standard or non-breaking spaces.
  ;
  While Chr\c = #Space Or Chr\c = #NonBreakableSpace : GoNextChar(Chr) : Wend
EndMacro
;
Macro       SkipSpacesBackward(Chr, MinAdr)
  ;
  ; Role: Moves a character pointer backward past any contiguous standard or non-breaking spaces down to MinAdr.
  ;
  While (Chr) >= (MinAdr) And (Chr\c = #Space Or Chr\c = #NonBreakableSpace) : Chr - SizeOf(Character) : Wend
EndMacro
;
Macro       IsValidIdentifierStart(ChrAsc)
  ;
  ; Role: Evaluates to #True if 'ChrAsc' can start a bare PureBasic ASCII identifier ('A'..'Z', 'a'..'z', '_').
  ; NOTE: A PureBasic element canNOT contain accented characters.
  ;
  Bool((ChrAsc) > '9' And (ChrAsc) <= 'z' And IsPBSeparator(ChrAsc) = 0)
EndMacro
;
Macro       IsChainIdentChar(ChrAsc)
  ;
  ; Role: Evaluates to #True if 'ChrAsc' is a valid ASCII identifier character inside a field chain (including digits and '$').
  ;
  Bool((ChrAsc) >= '$' And (ChrAsc) <= 'z' And IsPBSeparator(ChrAsc) = 0)
EndMacro
;
Macro       IsStringOrCharLiteralStart(Chr)
  ;
  ; Role: Evaluates to #True if 'Chr' points to the start of a standard ("..."), character ('...'), or escaped (~"...") literal.
  ;
  Bool(PeekC(Chr) = #DOUBLEQUOTE Or PeekC(Chr) = #SIMPLEQUOTE Or (PeekC(Chr) = '~' And PeekC((Chr) + SizeOf(Character)) = #DOUBLEQUOTE))
EndMacro
;
Macro       SkipToken
  ;
  ; Role: Records the current character as the last separator/token and advances the character pointer by one position.
  ;
  LastSeparator = *Char\c
  LastToken$    = Chr(*Char\c)
  GoNextChar(*Char)
  ;
  CompilerIf #PBBReport
    If LastToken$ And LastToken$ <> #Space$ And (OPDMod & #OPDFollowsSeparators)
      If Right(OPDTokenLine$, 1) <> "(" And Right(OPDTokenLine$, 1) <> "." And LastToken$ <> ")" And LastToken$ <> "."
        OPDTokenLine$ + " "
      EndIf
      OPDTokenLine$ + LastToken$
    EndIf
  CompilerEndIf
EndMacro
;
Procedure.s ExtractOrSkipNextToken(*StringStartingAdr, *PtrToChar)
  ;
  ; Role: Extracts or skips the next token from the source character stream.
  ; Parameters: *PtrToChar - Pointer to a character pointer inside a code content string.
  ; Prerequisites: PeekI('*PtrToChar') must point inside a valid null-terminated string buffer.
  ; Output: Returns the extracted token and updates the character position via '*PtrToChar'.
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  Protected *TempChar.Character
  Protected *TokenStartChar, TokenLen, Result$
  ;
  SkipSpaces(*Char)
  *TokenStartChar = *Char
  ;
  ; Allow '*' or '#' at the beginning of the token only when they form a valid pointer or constant name
  If *Char\c = '*'
    If RecognizePointer(*StringStartingAdr, CharToPos(*Char))
      GoNextChar(*Char)
    EndIf
  ElseIf *Char\c = '#'
    *TempChar = *Char
    GoNextChar(*TempChar)
    If IsPBSeparator(*TempChar\c) = #False
      GoNextChar(*Char)
    EndIf
  EndIf
  ;
  If *Char = *TokenStartChar And *Char\c = '$'
    ; Hexadecimal literal (e.g., $00FF): consume '$' and all contiguous hex digits as a single token
    GoNextChar(*Char)
    While (*Char\c >= '0' And *Char\c <= '9') Or (*Char\c >= 'a' And *Char\c <= 'f') Or (*Char\c >= 'A' And *Char\c <= 'F')
      GoNextChar(*Char)
    Wend
  ElseIf *Char > *TokenStartChar Or IsPBSeparator(*Char\c) = #False
    ; Standard identifier, keyword, or number: advance until any separator is reached
    While IsPBSeparator(*Char\c) = #False
      GoNextChar(*Char)
      ; In PureBasic identifiers, '$' is only valid as the terminal character of a string variable
      If PeekC(*Char - SizeOf(Character)) = '$'
        Break
      EndIf
    Wend
  EndIf
  ;
  TokenLen = ComputeLenFromChars(*TokenStartChar, *Char)
  If TokenLen
    Result$ = PeekS(*TokenStartChar, TokenLen, #PB_Unicode)
  EndIf
  ;
  CompilerIf #PBBReport
    If Result$ And (OPDMod & #OPDFollowsTokens)
      If Right(OPDTokenLine$, 1) <> "(" And Right(OPDTokenLine$, 1) <> "."
        OPDTokenLine$ + " "
      EndIf
      OPDTokenLine$ + Result$
    EndIf
  CompilerEndIf
  ;
  PokeI(*PtrToChar, *Char)
  ProcedureReturn Result$
EndProcedure
;
; ----------------------------------------------------------------------------------
;           ---- Token, literal & balanced-parentheses extractors ----
; ----------------------------------------------------------------------------------
;
Procedure.s ExtractOrSkipNextString(*PtrToChar, QuoteType = #DOUBLEQUOTE, SkipOnly = #False)
  ;
  ; Role: Advances '*PtrToChar' past a standard ("..."), escaped (~"..."), or character ('...')
  ;       literal, and returns the extracted literal unless 'SkipOnly' is #True.
  ; Parameters:
  ;   *PtrToChar - Address of a Character pointer positioned on the opening quote or '~' prefix.
  ;   QuoteType  - Expected delimiter (#DOUBLEQUOTE, #SIMPLEQUOTE, or '~' for ~"...").
  ;   SkipOnly   - #True to advance '*PtrToChar' without allocating a PureBasic string.
  ; Prerequisites: PeekI('*PtrToChar') must point inside a valid null-terminated string buffer.
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  Protected *TokenStartChar = *Char
  Protected IsEscaped, Result$
  ;
  If *Char\c = '~' And (QuoteType = #DOUBLEQUOTE Or QuoteType = '~') And PeekC(*Char + SizeOf(Character)) = #DOUBLEQUOTE
    ; Escaped string literal (~"..."): consume both '~' and the opening double quote
    QuoteType = #DOUBLEQUOTE
    IsEscaped = #True
    *Char + 2 * SizeOf(Character)
  ElseIf *Char\c = QuoteType And QuoteType <> '~'
    ; Standard string ("...") or character ('...') literal: consume the opening quote
    GoNextChar(*Char)
  Else
    ProcedureReturn ""
  EndIf
  ;
  ; Scan until the closing quote, line break, or end of buffer
  While *Char\c <> 0 And *Char\c <> #CR And *Char\c <> #LF And *Char\c <> QuoteType
    If IsEscaped And *Char\c = '\'
      ; Skip the escaped character (such as \" or \\) without crossing a line break or null terminator
      GoNextChar(*Char)
      If *Char\c = 0 Or *Char\c = #CR Or *Char\c = #LF
        Break
      EndIf
    EndIf
    GoNextChar(*Char)
  Wend
  ;
  If *Char\c = QuoteType
    GoNextChar(*Char)
  EndIf
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowsTokens
      SkipOnly = #False
    EndIf
  CompilerEndIf
  ;
  If SkipOnly = #False
    Result$ = PeekS(*TokenStartChar, ComputeLenFromChars(*TokenStartChar, *Char), #PB_Unicode)
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowsTokens
        If Right(OPDTokenLine$, 1) <> "("
          OPDTokenLine$ + " "
        EndIf
        OPDTokenLine$ + Result$
      EndIf
    CompilerEndIf
  EndIf
  ;
  PokeI(*PtrToChar, *Char)
  ProcedureReturn Result$
EndProcedure
;
Procedure.s GoToEOL_AndExtractComment(*PtrToChar)
  ;
  ; Role: Extracts the comment of the current line, if any, and advances '*PtrToChar' to the line end.
  ; Prerequisites: PeekI('*PtrToChar') must point inside a valid null-terminated string buffer.
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  Protected *TokenStartChar, TokenLen, Result$
  ;
  Repeat
    If *Char\c = 0 Or *Char\c = #CR
      Break
    ElseIf IsStringOrCharLiteralStart(*Char)
      ExtractOrSkipNextString(@*Char, *Char\c, #True)
    ElseIf *Char\c = ';'
      *TokenStartChar = *Char
      While *Char\c <> 0 And *Char\c <> #CR
        GoNextChar(*Char)
      Wend
      TokenLen = ComputeLenFromChars(*TokenStartChar, *Char)
      If TokenLen
        Result$ = PeekS(*TokenStartChar, TokenLen, #PB_Unicode)
      EndIf
      Break
    Else
      GoNextChar(*Char)
    EndIf
  ForEver
  ;
  CompilerIf #PBBReport
    If Result$ And (OPDMod & #OPDFollowsComments)
      OPDTokenLine$ + " " + Result$
    EndIf
  CompilerEndIf
  ;
  PokeI(*PtrToChar, *Char)
  ProcedureReturn Result$
EndProcedure
;
Procedure   CheckForParenthesis(*PtrToChar)
  ;
  ; Role: Checks if an opening parenthesis follows the current PeekI(*PtrToChar) position (read-only).
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  ;
  SkipSpaces(*Char)
  ProcedureReturn Bool(*Char\c = '(')
EndProcedure
;
Procedure   IsModuleUsage(*PtrToChar)
  ;
  ; Role: Checks if the current character pointer is positioned on a module scope operator ('::').
  ; Output: If detected, advances the target pointer past '::' and returns #True; otherwise returns #False.
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  ;
  If *Char\c = ':'
    GoNextChar(*Char)
    If *Char\c = ':'
      GoNextChar(*Char)
      PokeI(*PtrToChar, *Char)
      ProcedureReturn #True
    EndIf
  EndIf
  ProcedureReturn #False
EndProcedure
;
Macro       IsCompilerDirective(TokLCase)
  ;
  ; Role: Evaluates to #True if 'TokLCase' is one of PureBasic's 8 conditional compilation directives.
  ;
  Bool(Left(TokLCase, 8) = "compiler" And
       FindString(",compilerif,compilerelseif,compilerelse,compilerendif," +
                  "compilerselect,compilercase,compilerdefault,compilerendselect,", "," + TokLCase + ","))
EndMacro
;
Procedure   SkipToEndOfStatement(*PtrToChar, *LimitChar.Character = 0, ConsumeDelimiter = #True)
  ;
  ; Role: Advances '*PtrToChar' across literals, '::' scope operators, and trailing comments
  ;       up to the current statement boundary (#CR, standalone ':', '*LimitChar', or 0),
  ;       consuming the terminating #CR or ':' when 'ConsumeDelimiter' is #True.
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  ;
  While (*LimitChar = 0 Or *Char < *LimitChar) And *Char\c <> 0
    If *Char\c = #CR
      If ConsumeDelimiter : GoNextChar(*Char) : EndIf
      Break
    ElseIf *Char\c = ':'
      If IsModuleUsage(@*Char) : Continue : EndIf
      If ConsumeDelimiter : GoNextChar(*Char) : EndIf
      Break
    ElseIf IsStringOrCharLiteralStart(*Char)
      ExtractOrSkipNextString(@*Char, *Char\c, #True)
    ElseIf *Char\c = ';'
      GoToEOL_AndExtractComment(@*Char)
    Else
      GoNextChar(*Char)
    EndIf
  Wend
  If *LimitChar And *Char > *LimitChar : *Char = *LimitChar : EndIf
  PokeI(*PtrToChar, *Char)
EndProcedure
;
Procedure.s ExtractCompilerDirectiveExpr(*StringStartingAdr, *PtrToChar, HasArgument = #True)
  ;
  ; Role: Extracts a compile-time expression up to the statement delimiter (#CR, ':', ';', or 0),
  ;       respecting literals, '::' module prefixes, and multi-line continuations.
  ;       Leaves PeekI(*PtrToChar) positioned directly on the terminating delimiter (#CR, ':', or 0).
  ;
  Protected *Char.Character, *SegStart.Character, *CodeEnd.Character, *TokStart.Character
  Protected Expr$, LastTok$
  ;
  If HasArgument = #False
    SkipToEndOfStatement(*PtrToChar, 0, #False)
    ProcedureReturn ""
  EndIf
  ;
  *Char = PeekI(*PtrToChar)
  Repeat
    SkipSpaces(*Char)
    *SegStart = *Char
    *CodeEnd  = *Char
    LastTok$  = ""
    While *Char\c <> 0 And *Char\c <> #CR
      If IsStringOrCharLiteralStart(*Char)
        ExtractOrSkipNextString(@*Char, *Char\c, #True)
        LastTok$ = "" : *CodeEnd = *Char
      ElseIf *Char\c = ';'
        GoToEOL_AndExtractComment(@*Char)
        Break
      ElseIf *Char\c = ':'
        If IsModuleUsage(@*Char) : *CodeEnd = *Char : Continue : EndIf
        Break
      ElseIf *Char\c = #Space Or *Char\c = #NonBreakableSpace
        SkipSpaces(*Char)
      ElseIf IsPBSeparator(*Char\c)
        LastTok$ = Chr(*Char\c)
        GoNextChar(*Char)
        *CodeEnd = *Char
      Else
        *TokStart = *Char
        While *Char\c <> 0 And IsPBSeparator(*Char\c) = #False : GoNextChar(*Char) : Wend
        *CodeEnd = *Char
        LastTok$ = LCase(PeekS(*TokStart, ComputeLenFromChars(*TokStart, *Char)))
      EndIf
    Wend
    If *CodeEnd > *SegStart
      Expr$ + " " + Trim(PeekS(*SegStart, ComputeLenFromChars(*SegStart, *CodeEnd)))
    EndIf
    If *Char\c = #CR And LastTok$ <> "" And FindString(#PBEndOfSegmentedLine$, #TAB$ + LastTok$ + #TAB$)
      GoNextChar(*Char)
      Continue
    EndIf
    Break
  ForEver
  ;
  PokeI(*PtrToChar, *Char)
  ProcedureReturn Trim(Expr$)
EndProcedure
;
Procedure   SkipBalancedParentheses(*PtrToChar, *LimitChar.Character = 0)
  ;
  ; Role: Advances '*PtrToChar' from an upcoming opening parenthesis '(' past its matching ')',
  ;       supporting multi-line expressions continued with '+', ',', '|', 'And', 'Or', or 'Xor'
  ;       while safely skipping comments and standard or escaped string/character literals.
  ; Parameters:
  ;   *PtrToChar - Address of the caller's Character pointer positioned at or before '('.
  ;   *LimitChar - Optional exclusive upper memory bound (0 to stop only at null terminator).
  ; Output: Returns #True if the matching ')' was consumed, or #False if unclosed before EOL/EOF.
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  Protected *WordStart.Character
  Protected ParenDepth, WordLen, CanContinueLine
  ;
  SkipSpaces(*Char)
  If (*LimitChar = 0 Or *Char < *LimitChar) And *Char\c = '('
    ParenDepth = 1
    GoNextChar(*Char)
    ;
    While ParenDepth > 0 And (*LimitChar = 0 Or *Char < *LimitChar) And *Char\c <> 0
      Select *Char\c
        Case #Space, #NonBreakableSpace
          ; Preserve CanContinueLine across trailing spaces and indented continuation lines
          GoNextChar(*Char)
          ;
        Case ';'
          ; Skip trailing or standalone comments without altering the active continuation state
          While (*LimitChar = 0 Or *Char < *LimitChar) And *Char\c <> 0 And *Char\c <> #CR
            GoNextChar(*Char)
          Wend
          ;
        Case #CR
          ; Cross physical line breaks only when preceded by a valid PureBasic continuation operator
          If CanContinueLine
            GoNextChar(*Char)
          Else
            Break
          EndIf
          ;
        Case #DOUBLEQUOTE, #SIMPLEQUOTE, '~'
          CanContinueLine = #False
          If *Char\c <> '~' Or ((*LimitChar = 0 Or (*Char + SizeOf(Character)) < *LimitChar) And PeekC(*Char + SizeOf(Character)) = #DOUBLEQUOTE)
            ExtractOrSkipNextString(@*Char, *Char\c, #True)
            If *LimitChar And *Char > *LimitChar
              *Char = *LimitChar
            EndIf
          Else
            ; Standalone bitwise NOT operator '~'
            GoNextChar(*Char)
          EndIf
          ;
        Case '('
          ParenDepth + 1
          CanContinueLine = #False
          GoNextChar(*Char)
          ;
        Case ')'
          ParenDepth - 1
          CanContinueLine = #False
          GoNextChar(*Char)
          ;
        Case ',', '+', '|'
          CanContinueLine = #True
          GoNextChar(*Char)
          ;
        Default
          If *Char\c = '#' Or *Char\c = '*' Or IsPBSeparator(*Char\c) = #False
            ; Consume identifier, constant, or pointer token in place and check for 'And', 'Or', 'Xor'
            *WordStart = *Char
            If *Char\c = '#' Or *Char\c = '*'
              GoNextChar(*Char)
            EndIf
            While (*LimitChar = 0 Or *Char < *LimitChar) And IsPBSeparator(*Char\c) = #False
              If *Char\c = '$'
                GoNextChar(*Char)
                Break
              EndIf
              GoNextChar(*Char)
            Wend
            WordLen = ComputeLenFromChars(*WordStart, *Char)
            If WordLen = 2
              CanContinueLine = Bool((*WordStart\c | 32) = 'o' And (PeekC(*WordStart + SizeOf(Character)) | 32) = 'r')
            ElseIf WordLen = 3
              CanContinueLine = Bool(((*WordStart\c | 32) = 'a' And (PeekC(*WordStart + SizeOf(Character)) | 32) = 'n' And
                                      (PeekC(*WordStart + 2 * SizeOf(Character)) | 32) = 'd') Or
                                     ((*WordStart\c | 32) = 'x' And (PeekC(*WordStart + SizeOf(Character)) | 32) = 'o' And
                                      (PeekC(*WordStart + 2 * SizeOf(Character)) | 32) = 'r'))
            Else
              CanContinueLine = #False
            EndIf
          Else
            CanContinueLine = #False
            GoNextChar(*Char)
          EndIf
          ;
      EndSelect
    Wend ; Wend ParenDepth > 0
    ;
    PokeI(*PtrToChar, *Char)
    ProcedureReturn Bool(ParenDepth = 0)
  EndIf
  ;
  ProcedureReturn #False
EndProcedure
;
Procedure.s ParseArguments(ArgList$)
  ;
  ; Role: Replaces top-level argument-separating commas in 'ArgList$' with #TAB$,
  ;       ignoring commas and parentheses inside nested sub-expressions, comments,
  ;       and standard ("..."), escaped (~"..."), or character ('...') literals.
  ;
  Protected *Char.Character
  Protected NbPar
  ;
  If ArgList$ = ""
    ProcedureReturn ""
  EndIf
  ;
  *Char = @ArgList$
  While *Char\c <> 0
    If IsStringOrCharLiteralStart(*Char)
      ExtractOrSkipNextString(@*Char, *Char\c, #True)
    ElseIf *Char\c = ';'
      While *Char\c <> 0 And *Char\c <> #CR
        GoNextChar(*Char)
      Wend
    Else
      If *Char\c = '('
        NbPar + 1
      ElseIf *Char\c = ')' And NbPar > 0
        NbPar - 1
      ElseIf *Char\c = ',' And NbPar = 0
        *Char\c = #TAB
      EndIf
      GoNextChar(*Char)
    EndIf
  Wend
  ;
  ProcedureReturn ArgList$
EndProcedure
;
Procedure.s ExtractBalancedCallParams(*CodeString.String, Pos, *OutEOL.Integer)
  ;
  ; Role: Extracts the parenthesized parameter list following a macro declaration or macro call at 'Pos',
  ;       skipping string literals and multi-line continuations via SkipBalancedParentheses(),
  ;       formats arguments with tab separators via ParseArguments(), and writes the end offset to '*OutEOL'.
  ;
  Protected *StringStartingAdr, *OpenParen.Character, *Scan.Character, *Limit.Character
  Protected InnerLen, BufLen
  ;
  If *OutEOL
    *OutEOL\i = Pos
  EndIf
  If *CodeString = 0 Or *CodeString\s = "" Or Pos < 1
    ProcedureReturn ""
  EndIf
  ;
  BufLen = Len(*CodeString\s)
  If Pos > BufLen
    ProcedureReturn ""
  EndIf
  ;
  *StringStartingAdr = @*CodeString\s
  *Scan              = PosToChar(*StringStartingAdr, Pos)
  *Limit             = PosToChar(*StringStartingAdr, BufLen) + SizeOf(Character)
  SkipSpaces(*Scan)
  ;
  If *Scan < *Limit And *Scan\c = '('
    *OpenParen = *Scan
    If SkipBalancedParentheses(@*Scan, *Limit)
      If *OutEOL
        *OutEOL\i = CharToPos(*Scan)
      EndIf
      InnerLen = ComputeLenFromChars(*OpenParen, *Scan) - 2
      If InnerLen > 0
        ProcedureReturn ParseArguments(PeekS(*OpenParen + SizeOf(Character), InnerLen))
      EndIf
    EndIf
  EndIf
  ProcedureReturn ""
EndProcedure
;
Procedure   GetProcedureHeaderEndPos(*CodeBuffer.String, StartingPos, EndingPos)
  ;
  ; Role: Locates the closing parenthesis ')' of a Procedure, Macro, or Declare header starting
  ;       at 'StartingPos' and bounded by 'EndingPos' (1-based inclusive), handling '.Mod::Type'
  ;       return suffixes and multi-line parameter lists via SkipBalancedParentheses().
  ; Output: Returns the 1-based character position of the closing ')', or 0 if none is found.
  ;
  Protected *StringStartingAdr, *Char.Character, *Limit.Character
  Protected BufferLen
  ;
  If *CodeBuffer = 0 Or *CodeBuffer\s = "" Or StartingPos < 1 Or EndingPos < StartingPos
    ProcedureReturn 0
  EndIf
  ;
  BufferLen = Len(*CodeBuffer\s)
  If StartingPos > BufferLen
    ProcedureReturn 0
  EndIf
  If EndingPos > BufferLen
    EndingPos = BufferLen
  EndIf
  ;
  *StringStartingAdr = @*CodeBuffer\s
  *Char              = PosToChar(*StringStartingAdr, StartingPos)
  *Limit             = PosToChar(*StringStartingAdr, EndingPos) + SizeOf(Character)
  ;
  ; Advance across the keyword, optional '.Mod::Type' suffix, and identifier on the opening header line
  While *Char < *Limit And *Char\c <> 0 And *Char\c <> #CR And *Char\c <> ';' And *Char\c <> '('
    If *Char\c = ':'
      If IsModuleUsage(@*Char)
        Continue
      EndIf
      Break
    EndIf
    GoNextChar(*Char)
  Wend
  ;
  If *Char < *Limit And *Char\c = '('
    If SkipBalancedParentheses(@*Char, *Limit)
      ProcedureReturn CharToPos(*Char) - 1
    EndIf
  EndIf
  ;
  ProcedureReturn 0
EndProcedure
;
Procedure.s ExtractOptionalDotType(*StringStartingAdr, *PtrToChar, DefaultType$ = "")
  ;
  ; Role: Extracts an explicit '.Type' or '.Module::Type' suffix if present at '*PtrToChar',
  ;       or returns 'DefaultType$' unchanged when no dot suffix follows.
  ; Prerequisites: PeekI('*PtrToChar') must point inside a valid null-terminated string buffer.
  ; Output: Returns the extracted type identifier (keeping any "Module::" prefix) or 'DefaultType$',
  ;         and advances '*PtrToChar' past the type token only when a dot is consumed.
  ;
  Protected *Char.Character = PeekI(*PtrToChar)
  ;
  SkipSpaces(*Char)
  If *Char\c = '.'
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowsSeparators
        OPDTokenLine$ + "."
      EndIf
    CompilerEndIf
    GoNextChar(*Char)
    DefaultType$ = ExtractOrSkipNextToken(*StringStartingAdr, @*Char)
    If DefaultType$ <> "" And IsModuleUsage(@*Char)
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowsSeparators
          OPDTokenLine$ + "::"
        EndIf
      CompilerEndIf
      DefaultType$ + "::" + ExtractOrSkipNextToken(*StringStartingAdr, @*Char)
    EndIf
    PokeI(*PtrToChar, *Char)
  EndIf
  ProcedureReturn DefaultType$
EndProcedure
;
Procedure   IsParam(TokenLCase$, MacroProcParams$, IsFieldAccess = #False)
  ;
  ; Role: Checks whether 'TokenLCase$' matches a formal parameter in the tab-delimited 'MacroProcParams$' list
  ;       (ignored when 'IsFieldAccess' is #True).
  ;
  Protected PosInMacroProcParams
  Protected ParamsLCase$, FollowingChar$
  ;
  If IsFieldAccess = #False And Len(MacroProcParams$) > 1
    ParamsLCase$ = LCase(MacroProcParams$)
    Repeat
      PosInMacroProcParams = FindString(ParamsLCase$, #TAB$ + TokenLCase$, PosInMacroProcParams + 1)
      If PosInMacroProcParams
        FollowingChar$ = Mid(ParamsLCase$, PosInMacroProcParams + 1 + Len(TokenLCase$), 1)
        If FollowingChar$ = "" Or FollowingChar$ = #TAB$ Or FollowingChar$ = "." Or FollowingChar$ = "("
          ProcedureReturn #True
        EndIf
        PosInMacroProcParams + Len(TokenLCase$)
      EndIf
    Until PosInMacroProcParams = 0
  EndIf
  ;
  ProcedureReturn #False
EndProcedure
;
Procedure   IsCapturableMacroToken(TokenLCase$, IsQualifiedAccess = #False)
  ;
  ; Role: Determines whether an unindexed identifier encountered inside a Macro body
  ;       can represent a forward-declared project element, filtering out PureBasic
  ;       reserved keywords, native functions, native types, and native constants.
  ; Parameters:
  ;   IsQualifiedAccess - #True when the token is preceded by 'Mod::', allowing module
  ;                       members that share a PureBasic native function, type, or constant name.
  ;
  Shared PBNativeConstantMap()
  Protected FirstAsc = Asc(TokenLCase$)
  ;
  If FirstAsc = '*'
    ; Pointer identifiers never collide with PureBasic keywords, functions, types, or constants
    ProcedureReturn #True
  ElseIf FirstAsc = '#'
    ProcedureReturn Bool(IsQualifiedAccess Or
                         (TokenLCase$ <> "#true" And TokenLCase$ <> "#false" And
                          FindMapElement(PBNativeConstantMap(), UCase(TokenLCase$)) = 0))
  ElseIf IsValidIdentifierStart(FirstAsc)
    If FindMapElement(PBBasicKeywordMap(), TokenLCase$) = 0
      ProcedureReturn Bool(IsQualifiedAccess Or
                           (FindMapElement(PBNativeFunctions(), TokenLCase$) = 0 And
                            FindMapElement(PBBasicNativeTypes(), TokenLCase$) = 0))
    EndIf
  EndIf
  ;
  ProcedureReturn #False
EndProcedure
;
; **********************************************************************************
;
;-                 Identifier purification & code cleanup
;
; **********************************************************************************
;
Procedure.s ExtractWithTarget(*Char.Character)
  ;
  ; Role: Extracts the trimmed target expression following a 'With' keyword without modifying
  ;       the caller's pointer. Stops at end-of-line, comments (';'), or standalone statement
  ;       separators (':') while preserving '::' module prefixes and quoted literals.
  ;
  Protected *Start.Character
  ;
  If *Char = 0 : ProcedureReturn "" : EndIf
  ;
  SkipSpaces(*Char)
  *Start = *Char
  ;
  ; Advance to the end of the 'With' expression, skipping quoted strings and '::' scope operators
  While *Char\c <> 0 And *Char\c <> #CR And *Char\c <> ';'
    If IsStringOrCharLiteralStart(*Char)
      ExtractOrSkipNextString(@*Char, *Char\c, #True)
      Continue
    ElseIf *Char\c = ':'
      If IsModuleUsage(@*Char) : Continue : EndIf
      Break
    EndIf
    GoNextChar(*Char)
  Wend
  ;
  While *Char > *Start And (PeekC(*Char - SizeOf(Character)) = #Space Or PeekC(*Char - SizeOf(Character)) = #NonBreakableSpace)
    *Char - SizeOf(Character)
  Wend
  ;
  If *Char > *Start
    ProcedureReturn PeekS(*Start, ComputeLenFromChars(*Start, *Char))
  EndIf
  ProcedureReturn ""
EndProcedure
;
Procedure.s ExtractPrecedingFieldChain(*BackslashChar.Character, *BufferStart, WithVar$)
  ;
  ; Role: Walks backward from a structure dereference backslash ('\') to extract the
  ;       variable or chained field expression preceding it (e.g., "*SearchParam", "Mod::Var\Sub",
  ;       or "Var.Mod::Type"), stripping indexing expressions inside parentheses/brackets and
  ;       prepending WithVar$ on orphan '\Field' accesses.
  ;
  Protected *Scan.Character = *BackslashChar - SizeOf(Character)
  Protected *Prev.Character
  Protected ParenDepth, BracketDepth, UsedWith
  Protected Chain$
  ;
  SkipSpacesBackward(*Scan, *BufferStart)
  ;
  ; If nothing precedes the backslash on the expression side, switch directly to scanning WithVar$
  If *Scan < *BufferStart Or (IsChainIdentChar(*Scan\c) = #False And *Scan\c <> ')' And *Scan\c <> ']')
    If WithVar$ = "" : ProcedureReturn "" : EndIf
    *BufferStart = @WithVar$
    *Scan        = PosToChar(*BufferStart, Len(WithVar$))
    UsedWith     = #True
  EndIf
  ;
  Repeat
    While *Scan >= *BufferStart
      If *Scan\c = ')'
        If ParenDepth = 0 : Chain$ = "()" + Chain$ : EndIf
        ParenDepth + 1
      ElseIf *Scan\c = '('
        If ParenDepth > 0 : ParenDepth - 1 : Else : Break : EndIf
      ElseIf *Scan\c = ']'
        BracketDepth + 1
      ElseIf *Scan\c = '['
        If BracketDepth > 0 : BracketDepth - 1 : Else : Break : EndIf
      ElseIf ParenDepth = 0 And BracketDepth = 0
        If IsChainIdentChar(*Scan\c) Or *Scan\c = '.' Or *Scan\c = '\'
          Chain$= Chr(*Scan\c) + Chain$
        ElseIf *Scan\c = '*'
          ; Distinguish pointer prefix ('*Ptr', 'Container\*Ptr', 'Mod::*Ptr') from binary multiplication ('a * b\Field')
          *Prev = *Scan - SizeOf(Character)
          SkipSpacesBackward(*Prev, *BufferStart)
          If *Prev >= *BufferStart And (IsChainIdentChar(*Prev\c) Or *Prev\c = ')' Or *Prev\c = ']')
            Break
          EndIf
          Chain$ = "*" + Chain$
          If *Prev >= *BufferStart And
             (*Prev\c = '\' Or (*Prev > *BufferStart And *Prev\c = ':' And PeekC(*Prev - SizeOf(Character)) = ':'))
            *Scan = *Prev
            Continue
          EndIf
          Break
        ElseIf *Scan > *BufferStart And *Scan\c = ':' And PeekC(*Scan - SizeOf(Character)) = ':'
          Chain$ = "::" + Chain$
          *Scan - 2 * SizeOf(Character)
          SkipSpacesBackward(*Scan, *BufferStart)
          Continue
        Else
          Break
        EndIf
      EndIf
      *Scan - SizeOf(Character)
    Wend
    ;
    If UsedWith = #False And Left(Chain$, 1) = "\" And WithVar$ <> ""
      *BufferStart = @WithVar$
      *Scan        = PosToChar(*BufferStart, Len(WithVar$))
      UsedWith     = #True
    Else
      Break
    EndIf
  ForEver
  ;
  ProcedureReturn Chain$
EndProcedure
;
Macro IsQuotedImportTarget(Expr)
  ;
  ; Role: Returns #True only when 'Expr' is a single quoted '.lib' target (e.g., '"kernel32.lib"' or '~"path\foo.lib"'),
  ; preserving standard raw-text searches for all other quoted strings such as '"MyProc()"'.
  ;
  Bool(Len(Expr) >= 7 And LCase(Right(Expr, 5)) = ".lib" + #DOUBLEQUOTE$ And
       ((Left(Expr, 1) = #DOUBLEQUOTE$ And CountString(Expr, #DOUBLEQUOTE$) = 2) Or
        (Left(Expr, 2) = "~" + #DOUBLEQUOTE$ And FindString(Expr, #DOUBLEQUOTE$, 3) = Len(Expr))))
EndMacro
;
Procedure.s PurifyForSearch(Expression$, *NotAnElement.Integer = 0)
  ;
  ; Role: Strips syntax modifiers and jokers to extract the core identifier for search indexing,
  ;       or normalizes spacing around separators when the input is a complex expression.
  ;
  Protected TestNorm$, CoreName$
  Protected IsComplex, HasLeadingJoker, HasTrailingJoker, CharCode
  Protected *Ptr.Character, *NextChar.Character, *PreviousChar.Character
  ;
  Expression$ = Trim(StringField(Expression$, 1, #SpecialSpace$))
  If Expression$ = "" : ProcedureReturn "" : EndIf
  ;
  If Left(LCase(Expression$), 10) = "unnamed - " Or IsQuotedImportTarget(Expression$)
    ProcedureReturn LCase(Expression$)
  EndIf
  ;
  TestNorm$ = Expression$
  While FindString(TestNorm$, "  ")
    TestNorm$ = ReplaceString(TestNorm$, "  ", " ")
  Wend
  ;
  HasLeadingJoker  = Bool(Left(TestNorm$, Len(#PBBJoker$)) = #PBBJoker$)
  HasTrailingJoker = Bool(Right(TestNorm$, Len(#PBBJoker$)) = #PBBJoker$)
  TestNorm$        = Trim(ReplaceString(TestNorm$, #PBBJoker$, ""))
  If TestNorm$ = "" : ProcedureReturn Expression$ : EndIf
  ;
  CharCode = Asc(TestNorm$)
  If CharCode = '*'
    CharCode = Asc(Mid(TestNorm$, 2))
  EndIf
  If CharCode <> '#' And IsValidIdentifierStart(CharCode) = #False
    IsComplex = #True
  EndIf
  ;
  *Ptr = @TestNorm$
  While *Ptr\c
    Select *Ptr\c
      Case #Space, #NonBreakableSpace, #TAB, #CR, #LF ; For copy/paste search expressions
        *NextChar = *Ptr + SizeOf(Character)
        If (*PreviousChar = 0 Or *PreviousChar\c <> '.') And *NextChar\c <> '(' And *NextChar\c <> '.'
          IsComplex = #True
          CoreName$ + " "
        EndIf
      Case '*'
        *NextChar = *Ptr + SizeOf(Character)
        If *Ptr = @TestNorm$ And *NextChar\c <> 0 And IsPBSeparator(*NextChar\c) = #False
          ; Leading pointer prefix stays attached to its identifier
          CoreName$ + "*"
        Else
          IsComplex = #True
          CoreName$ + " * "
        EndIf
      Case '\'
        *NextChar = *Ptr + SizeOf(Character)
        If *NextChar\c <> 0
          IsComplex = #True
          CoreName$ + " \ "
        EndIf
      Case ':'
        *NextChar = *Ptr + SizeOf(Character)
        If *NextChar\c <> 0
          IsComplex = #True
          CoreName$ + " : "
        EndIf
      Case '('
        *NextChar = *Ptr + SizeOf(Character)
        If *NextChar\c <> ')' Or PeekC(*NextChar + SizeOf(Character)) <> 0
          IsComplex = #True
        EndIf
        CoreName$ + " ( "
      Case ')'
        If *PreviousChar = 0 Or *PreviousChar\c <> '('
          IsComplex = #True
        EndIf
        CoreName$ + " ) "
      Case '.'
        CoreName$ + " . "
      Default
        If IsPBSeparator(*Ptr\c)
          IsComplex = #True
          CoreName$ + " " + Chr(*Ptr\c) + " "
        Else
          CoreName$ + Chr(*Ptr\c)
        EndIf
    EndSelect
    *PreviousChar = *Ptr
    *Ptr + SizeOf(Character)
  Wend
  ;
  If IsComplex
    If *NotAnElement
      *NotAnElement\i = #True
    EndIf
    CoreName$ = Trim(CoreName$)
    While FindString(CoreName$, "  ")
      CoreName$ = ReplaceString(CoreName$, "  ", " ")
    Wend
    ProcedureReturn LCase(CoreName$)
  EndIf
  ;
  CoreName$ = StringField(TestNorm$, 1, "(")
  CoreName$ = StringField(CoreName$, 1, ".")
  CoreName$ = StringField(CoreName$, 1, "\")
  CoreName$ = Trim(StringField(CoreName$, 1, ":"))
  ;
  If HasLeadingJoker
    CoreName$ = #PBBJoker$ + CoreName$
  EndIf
  If HasTrailingJoker
    CoreName$ + #PBBJoker$
  EndIf
  ;
  ProcedureReturn LCase(CoreName$)
EndProcedure
;
Procedure.s PurifyElementName(ElementName$, *HasParentheses.Integer = 0, *DotType.string = 0)
  ;
  ; Role: Strips UI suffixes (#SpecialSpace$), structure paths ('\'), argument parentheses ('('),
  ;       type suffixes ('.'), and trailing colons (':') to return the lowercase root identifier
  ;       (preserving 'unnamed - ...' enumeration names).
  ;       Optionally outputs *HasParentheses (set to 1 when '(' is present) and *DotType.
  ;
  ElementName$ = Trim(StringField(ElementName$, 1, #SpecialSpace$))
  If Left(LCase(ElementName$), 10) = "unnamed - " Or IsQuotedImportTarget(ElementName$)
    If *HasParentheses : *HasParentheses\i = 0 : EndIf
    If *DotType : *DotType\s = "" : EndIf
    ProcedureReturn LCase(ElementName$)
  EndIf
  ElementName$ = StringField(ElementName$, 1, "\")
  ;
  If *HasParentheses
    *HasParentheses\i = Bool(FindString(ElementName$, "("))
  EndIf
  ElementName$ = StringField(ElementName$, 1, "(")
  ;
  If *DotType
    *DotType\s = StringField(ElementName$, 2, ".")
  EndIf
  ElementName$= RTrim(Trim(StringField(ElementName$, 1, ".")), ":")
  If FindString(ElementName$, "::") = 0
    ElementName$= Trim(StringField(ElementName$, 1, ":"))
  EndIf
  ;
  ProcedureReturn LCase(ElementName$)
EndProcedure
;
; ----------------------------------------------------------------------------------
;             ---- Comment stripping & whitespace normalization ----
; ----------------------------------------------------------------------------------
;
Procedure   DeleteCommentsAndSpacesFromPBCode(*SearchString.String, CleanSpaces = #True, EraseAllComments = #False)
  ;
  ; Role: Single-pass O(N) pointer stream parser that strips comments and/or normalizes whitespace
  ;       in '*SearchString\s' while protecting string/character literals ("...", ~"...", '...')
  ;       and preserving the trailing PureBasic IDE Options block and EOF line break.
  ; Parameters:
  ;   CleanSpaces      - #True to strip indentation, collapse inline spaces, and remove blank lines;
  ;                      #False to keep indentation and cap consecutive blank lines at one ';'-prefixed line.
  ;   EraseAllComments - #True to strip full-line and inline comments (except IDE Options); #False to retain them.
  ;
  Protected *CharIn.Character, *CharOut.Character, *OutStart, *TokenStart
  Protected InputLen, IDEOptionsPos, IDEOptionsCharPos
  Protected AtLineStart = #True, LineHasCode, ConsecutiveBlankLines
  Protected CurrentChar, SkipLine, PrevC, IsCRLF
  Protected CharSize = SizeOf(Character)
  Protected TrailingEOL$
  ;
  If *SearchString = 0 : ProcedureReturn : EndIf
  InputLen = Len(*SearchString\s)
  If InputLen = 0 : ProcedureReturn : EndIf
  ;
  If Right(*SearchString\s, 2) = #CRLF$
    TrailingEOL$ = #CRLF$
  ElseIf Right(*SearchString\s, 1) = #LF$ Or Right(*SearchString\s, 1) = #CR$
    TrailingEOL$ = Right(*SearchString\s, 1)
  EndIf
  ;
  IDEOptionsPos     = FindIDENotations(*SearchString)
  IDEOptionsCharPos = PosToChar(@*SearchString\s, IDEOptionsPos)
  *OutStart         = AllocateMemory((InputLen + 4096) * CharSize)
  If *OutStart = 0 : ProcedureReturn : EndIf
  ;
  *CharIn  = @*SearchString\s
  *CharOut = *OutStart
  ;
  While *CharIn\c <> 0
    CurrentChar = *CharIn\c
    ;
    ; A. Protect string ("...", ~"...") and character ('...') literals
    If IsStringOrCharLiteralStart(*CharIn)
      *TokenStart = *CharIn
      ExtractOrSkipNextString(@*CharIn, CurrentChar, #True)
      CopyMemory(*TokenStart, *CharOut, *CharIn - *TokenStart)
      *CharOut + (*CharIn - *TokenStart)
      LineHasCode = #True
      AtLineStart = #False
      Continue
    EndIf
    ;
    ; B. Process comments and preserve the trailing IDE Options block
    If CurrentChar = ';'
      If EraseAllComments
        If LineHasCode = #False And *CharIn >= IDEOptionsCharPos
          While *CharIn\c <> 0
            *CharOut\c = *CharIn\c : *CharOut + CharSize
            *CharIn + CharSize
          Wend
          Break
        EndIf
        ; Rewind leading indentation on full-line comments, or trailing horizontal spaces before inline comments:
        While *CharOut > *OutStart
          PrevC = PeekC(*CharOut - CharSize)
          If PrevC = #Space Or PrevC = #NonBreakableSpace Or PrevC = #TAB
            *CharOut - CharSize
          Else
            Break
          EndIf
        Wend
        While *CharIn\c <> 0 And *CharIn\c <> #CR And *CharIn\c <> #LF
          *CharIn + CharSize
        Wend
        If LineHasCode = #False
          If *CharIn\c = #CR And PeekC(*CharIn + CharSize) = #LF
            *CharIn + 2 * CharSize
          ElseIf *CharIn\c = #CR Or *CharIn\c = #LF
            *CharIn + CharSize
          EndIf
          AtLineStart = #True
        EndIf
        Continue
      Else
        ; Copy retained comments verbatim so quotes inside them never trigger literal parsing
        AtLineStart = #False
        LineHasCode = #True
        While *CharIn\c <> 0 And *CharIn\c <> #CR And *CharIn\c <> #LF
          *CharOut\c = *CharIn\c : *CharOut + CharSize
          *CharIn + CharSize
        Wend
        Continue
      EndIf
    EndIf
    ;
    ; C. Handle line breaks and enforce blank-line rules
    If CurrentChar = #CR Or CurrentChar = #LF
      IsCRLF = Bool(CurrentChar = #CR And PeekC(*CharIn + CharSize) = #LF)
      If LineHasCode = #False
        While *CharOut > *OutStart And PeekC(*CharOut - CharSize) <> #CR And PeekC(*CharOut - CharSize) <> #LF
          *CharOut - CharSize
        Wend
        ConsecutiveBlankLines + 1
      Else
        ConsecutiveBlankLines = 0
      EndIf
      ;
      SkipLine = Bool((CleanSpaces And LineHasCode = #False) Or (CleanSpaces = #False And ConsecutiveBlankLines > 1))
      If SkipLine
        *CharIn + (1 + IsCRLF) * CharSize
        AtLineStart = #True
        LineHasCode = #False
        Continue
      EndIf
      ;
      If LineHasCode = #False And CleanSpaces = #False
        *CharOut\c = ';' : *CharOut + CharSize
      EndIf
      ;
      *CharOut\c = CurrentChar : *CharOut + CharSize
      *CharIn + CharSize
      If IsCRLF
        *CharOut\c = #LF : *CharOut + CharSize
        *CharIn + CharSize
      EndIf
      AtLineStart = #True
      LineHasCode = #False
      Continue
    EndIf
    ;
    ; D. Strip leading indentation or collapse consecutive internal whitespace
    If AtLineStart And (CurrentChar = #Space Or CurrentChar = #NonBreakableSpace Or CurrentChar = #TAB)
      If CleanSpaces
        *CharIn + CharSize
        Continue
      EndIf
    ElseIf CurrentChar <> #Space And CurrentChar <> #NonBreakableSpace And CurrentChar <> #TAB
      AtLineStart = #False
      LineHasCode = #True
    EndIf
    ;
    If CleanSpaces And (CurrentChar = #Space Or CurrentChar = #NonBreakableSpace Or CurrentChar = #TAB)
      If *CharOut > *OutStart
        PrevC = PeekC(*CharOut - CharSize)
        If PrevC <> #Space And PrevC <> #NonBreakableSpace And PrevC <> #TAB And PrevC <> #CR And PrevC <> #LF
          *CharOut\c = #Space : *CharOut + CharSize
        EndIf
      EndIf
    Else
      *CharOut\c = CurrentChar : *CharOut + CharSize
    EndIf
    ;
    *CharIn + CharSize
  Wend
  ;
  ; Strip trailing whitespace, empty ';' markers, and trailing line breaks in the output buffer,
  ; then restore the original file's single trailing EOL (if any):
  While *CharOut > *OutStart
    PrevC = PeekC(*CharOut - CharSize)
    If PrevC = #Space Or PrevC = #NonBreakableSpace Or PrevC = #TAB Or PrevC = #CR Or PrevC = #LF
      *CharOut - CharSize
    ElseIf CleanSpaces = #False And PrevC = ';' And
           (*CharOut - CharSize = *OutStart Or
            PeekC(*CharOut - 2 * CharSize) = #CR Or PeekC(*CharOut - 2 * CharSize) = #LF)
      *CharOut - CharSize
    Else
      Break
    EndIf
  Wend
  If CleanSpaces
    *CharIn = *OutStart
    While *CharIn < *CharOut And (*CharIn\c = #Space Or *CharIn\c = #NonBreakableSpace Or *CharIn\c = #TAB Or *CharIn\c = #CR Or *CharIn\c = #LF)
      *CharIn + CharSize
    Wend
  Else
    *CharIn = *OutStart
  EndIf
  If *CharOut > *CharIn
    *SearchString\s = PeekS(*CharIn, ComputeLenFromChars(*CharIn, *CharOut)) + TrailingEOL$
  Else
    *SearchString\s = ""
  EndIf
  FreeMemory(*OutStart)
EndProcedure
;
; **********************************************************************************
;
;-                 Statement extraction & line tokenizer
;
; **********************************************************************************
;
Procedure   IsLineContinuedAtEnd(*CodeString.String, LineEndPos)
  ;
  ; Role: Checks whether the physical line ending at 'LineEndPos' terminates with a line
  ;       continuation operator ('+', '|', ',', 'And', 'Or', 'Xor') outside of comments.
  ; Parameters:
  ;   LineEndPos - 1-based position of the trailing #CR (or buffer end + 1) of the line to inspect.
  ; Output: Returns the 1-based start boundary of the continued line (1 for the first line of the
  ;         buffer, or the position of its leading #CR), or 0 if the line is not continued.
  ;
  Protected pTest = LineEndPos - 1
  Protected pTestEnd, LineStartPos, CommentPos
  Protected LastToken$
  ;
  ; Skip trailing spaces before the line break
  While pTest > 0 And PeekC(@*CodeString\s + (pTest - 1) * SizeOf(Character)) = #Space
    pTest - 1
  Wend
  ;
  If pTest > 0 And PeekC(@*CodeString\s + (pTest - 1) * SizeOf(Character)) <> #CR
    LastToken$ = FastMid(*CodeString\s, pTest, 1)
    If IsPBSeparator(Asc(LastToken$)) = #False
      ; Isolate the trailing word on the line when the last character is not a separator
      pTestEnd = pTest + 1
      While pTest > 0 And IsPBSeparator(Asc(FastMid(*CodeString\s, pTest, 1))) = #False
        pTest - 1
      Wend
      pTest + 1
      LastToken$ = LCase(FastMid(*CodeString\s, pTest, pTestEnd - pTest))
    EndIf
    ;
    If FindString(#PBEndOfSegmentedLine$, #TAB$ + LastToken$ + #TAB$)
      ; Locate the start of this physical line and ensure no ';' comment precedes the operator
      LineStartPos = pTest
      FastFindPrecReturn(*CodeString\s, LineStartPos)
      If LineStartPos = 0 : LineStartPos = 1 : EndIf
      CommentPos = FindInPBPointedCode(*CodeString, ";", LineStartPos, #CheckComAndQuote | #CheckOneKeyWordOnly, #PB_String_CaseSensitive, LineEndPos)
      If CommentPos = 0 Or CommentPos > pTest
        ProcedureReturn LineStartPos
      EndIf
    EndIf
  EndIf
  ;
  ProcedureReturn 0
EndProcedure
;
Procedure.s ExtractSectionAroundExpression(*CodeString.String, Expression$, PosInCode,
                                           *OutComment.String = 0, CheckInCode = #CheckAll | #CheckOneKeyWordOnly,
                                           SStringLength = 0, *PosInLine.Integer = 0, *OriginalStart.Integer = 0,
                                           *OriginalLength.Integer = 0)
  ;
  ; Role: Verifies that 'Expression$' is not within quotes or comments in '*CodeString',
  ;       then extracts the logical statement containing the expression.
  ;       Strips segments on the same line located before or after ":", and concatenates
  ;       adjacent physical lines when a line continuation operator (+, |, ,, And, Or, Xor) is present.
  ; Parameters:
  ;   PosInCode   - If 0, 'Expression$' is searched within '*CodeString'; otherwise, specifies its 1-based position.
  ;   *OutComment - If provided, receives any trailing comment found on the statement line.
  ;   CheckInCode - If #NoCheck, extracts the section assuming it is already confirmed outside quotes or comments.
  ;
  Protected posDepR, posEndR, BreakInLine, CChar, pTest, Cont, LocalPosInLine, SemColPos, Pos
  Protected FullLine.String
  Protected *ScanPtr.Character
  Protected Keyword$
  ;
  If *CodeString = 0
    Alert("Empty pointer in the parameter of ExtractSectionAroundExpression().")
    ProcedureReturn ""
  EndIf
  If SStringLength = 0
    SStringLength = Len(*CodeString\s)
  EndIf
  If SStringLength = 0
    ProcedureReturn ""
  EndIf
  ;
  If PosInCode = 0
    ; Fallback case-insensitive search when caller did not supply a 1-based position
    PosInCode = FindInPBPointedCode(*CodeString, Expression$, 1, CheckInCode, #PB_String_NoCase, SStringLength)
  EndIf
  If PosInCode > SStringLength
    Alert("Wrong PosInCode parameter for ExtractSectionAroundExpression().")
    ProcedureReturn ""
  EndIf
  ;
  If PosInCode And LCase(FastMid(*CodeString\s, PosInCode, Len(Expression$))) <> LCase(Expression$)
    ; Recalculate exact position of the expression within the current line
    posEndR = FindString(*CodeString\s, #CR$, PosInCode)
    If posEndR = 0 : posEndR = SStringLength + 1 : EndIf
    PosInCode = FindInPBPointedCode(*CodeString, Expression$, PosInCode, CheckInCode, #PB_String_NoCase, posEndR)
    If PosInCode = 0
      ProcedureReturn ""
    EndIf
  EndIf
  ;
  If PosInCode > 0
    ; Match the exact casing found in the source code (including when PosInCode = 1)
    Expression$ = FastMid(*CodeString\s, PosInCode, Len(Expression$))
    ;
    If PosInCode > 1
      ; Move back to the start of the current line, then include any preceding continued lines
      posDepR = PosInCode
      FastFindPrecReturn(*CodeString\s, posDepR)
      While posDepR > 1
        pTest = IsLineContinuedAtEnd(*CodeString, posDepR)
        If pTest
          posDepR = pTest
        Else
          Break
        EndIf
      Wend
    EndIf
    ;
    ; Trim leading spaces and carriage returns at the beginning of the extract
    If posDepR = 0 : posDepR = 1 : EndIf
    While posDepR < SStringLength
      CChar = PeekC(PosToChar(@*CodeString\s, posDepR))
      If CChar = #Space Or CChar = #CR
        posDepR + 1
      Else
        Break
      EndIf
    Wend
    If *OriginalStart
      *OriginalStart\i = posDepR
    EndIf
    ;
    ; Find the end of the current line, then include any following continued lines
    posEndR = PosInCode + Len(Expression$)
    Repeat
      posEndR = FindString(*CodeString\s, #CR$, posEndR)
      If posEndR = 0
        posEndR = SStringLength + 1
        Break
      EndIf
      If IsLineContinuedAtEnd(*CodeString, posEndR)
        posEndR + 1
      Else
        Break
      EndIf
    Until posEndR > SStringLength
    ;
    ; Trim trailing spaces and carriage returns at the end of the extract
    While posEndR > 1
      CChar = PeekC(PosToChar(@*CodeString\s, posEndR))
      If CChar = #Space Or CChar = #CR
        posEndR - 1
      Else
        Break
      EndIf
    Wend
    posEndR + 1
    ;
    If posEndR > posDepR
      FullLine\s = FastMid(*CodeString\s, posDepR, posEndR - posDepR)
      If *OriginalLength
        *OriginalLength\i = posEndR - posDepR
      EndIf
      ;
      ; Verify separator boundaries and ensure the expression is outside quotes/comments
      LocalPosInLine = PosInCode - posDepR + 1
      If LocalPosInLine > 0 And LocalPosInLine = FindInPBPointedCode(FullLine, Expression$, LocalPosInLine, CheckInCode)
        ; Remove statement segments preceding ":" on the same line (ignoring "::")
        BreakInLine = 0
        Repeat
          BreakInLine = FindInPBPointedCode(FullLine, ":", BreakInLine + 1, #CheckComAndQuote | #CheckOneKeyWordOnly)
          If BreakInLine
            If Mid(FullLine\s, BreakInLine, 2) = "::"
              BreakInLine + 2
            ElseIf BreakInLine < LocalPosInLine
              FullLine\s = FastMid(FullLine\s, BreakInLine + 1)
              LocalPosInLine - BreakInLine
              BreakInLine = 0
            EndIf
          EndIf
        Until BreakInLine = 0 Or BreakInLine >= LocalPosInLine
        ;
        ; Trim leading spaces while keeping LocalPosInLine synchronized
        While Left(FullLine\s, 1) = " " And LocalPosInLine > 1
          FullLine\s = Mid(FullLine\s, 2)
          LocalPosInLine - 1
        Wend
        ;
        ; Strip any trailing comment and optionally export it via *OutComment
        SemColPos = FindInPBPointedCode(FullLine, ";", LocalPosInLine + Len(Expression$), #CheckComAndQuote | #CheckOneKeyWordOnly)
        If SemColPos
          If *OutComment
            *OutComment\s = FastMid(FullLine\s, SemColPos)
          EndIf
          FullLine\s = FastLeft(FullLine\s, SemColPos - 1)
        ElseIf *OutComment
          *OutComment\s = ""
        EndIf
        ;
        ; Remove statement segments following ":" on the same line (ignoring "::" and label declarations)
        BreakInLine = LocalPosInLine + Len(Expression$)
        Repeat
          BreakInLine = FindInPBPointedCode(FullLine, ":", BreakInLine, #CheckComAndQuote | #CheckOneKeyWordOnly)
          If BreakInLine
            If Mid(FullLine\s, BreakInLine, 2) = "::"
              BreakInLine + 2
            ElseIf LocalPosInLine = 1 And BreakInLine = Len(Expression$) + LocalPosInLine
              BreakInLine + 1
            Else
              FullLine\s  = FastLeft(FullLine\s, BreakInLine - 1)
              BreakInLine = 0
            EndIf
          EndIf
        Until BreakInLine = 0
        ;
        FullLine\s = RTrim(ReplaceString(FullLine\s, #CR$, " "))
      Else
        FullLine\s = ""
      EndIf
    EndIf
  EndIf
  ;
  ; Strip spaces around '.', '\', and '(' outside string/character literals to simplify downstream parsing
  Pos = 1
  While Pos <= Len(FullLine\s)
    *ScanPtr = PosToChar(@FullLine\s, Pos)
    CChar    = *ScanPtr\c
    If IsStringOrCharLiteralStart(*ScanPtr)
      ExtractOrSkipNextString(@*ScanPtr, CChar, #True)
      Pos = ComputeLenFromChars(@FullLine\s, *ScanPtr) + 1
      Continue
    ElseIf CChar = '.' Or CChar = '\' Or CChar = '('
      Cont = #True
      If CChar = '\' And Pos > 1
        ; Preserve space before '\' when preceded by a PureBasic keyword (e.g., "With \Field", "If \Field")
        pTest = Pos - 1
        While pTest > 1 And FastMid(FullLine\s, pTest, 1) = " " : pTest - 1 : Wend
        While pTest > 1 And IsPBSeparator(Asc(FastMid(FullLine\s, pTest, 1))) = #False : pTest - 1 : Wend
        If IsPBSeparator(Asc(FastMid(FullLine\s, pTest, 1))) : pTest + 1 : EndIf
        Keyword$ = Trim(FastMid(FullLine\s, pTest, Pos - pTest))
        If Keyword$ <> "" And FindMapElement(PBBasicKeywordMap(), LCase(Keyword$))
          Cont = #False
        EndIf
      EndIf
      If Cont
        While Pos > 1 And PeekC(PosToChar(@FullLine\s, Pos - 1)) = #Space
          FullLine\s = FastLeft(FullLine\s, Pos - 2) + FastMid(FullLine\s, Pos)
          If Pos - 1 < LocalPosInLine : LocalPosInLine - 1 : EndIf
          Pos - 1
        Wend
      EndIf
      While Pos < Len(FullLine\s) And PeekC(PosToChar(@FullLine\s, Pos + 1)) = #Space
        FullLine\s = FastLeft(FullLine\s, Pos) + FastMid(FullLine\s, Pos + 2)
        If Pos + 1 < LocalPosInLine : LocalPosInLine - 1 : EndIf
      Wend
    EndIf
    Pos + 1
  Wend
  ;
  If *PosInLine
    *PosInLine\i = LocalPosInLine
  EndIf
  ProcedureReturn FullLine\s
EndProcedure
;
; ----------------------------------------------------------------------------------
;                       ---- Line-level AST tokenizer ----
; ----------------------------------------------------------------------------------
;
EnumerationBinary TCL_TokenTypes
  #TokenType_Word
  #TokenType_StringInQuotes
  #TokenType_Symbol
EndEnumeration
;
Structure LocalToken
  Value$
  Value_LCase$
  ModulePrefix$
  Type.i
  Level.i
  PosInLine.i
EndStructure
;
Macro       TCL_AppendToken(TokVal, TokType)
  ;
  ; Role: Inline helper appending a populated LocalToken entry in TokenizeCodeLine() with zero call overhead.
  ;
  AddElement(TokenList())
  TokenList()\Value$    = TokVal
  TokenList()\Type      = TokType
  TokenList()\Level     = CurrentLevel
  TokenList()\PosInLine = CurrentPos
EndMacro
;
Procedure   TokenizeCodeLine(*CodeLineCS, List TokenList.LocalToken(), TargetPosInLine = 0)
  ;
  ; Role: Transforms a line of code into a list of structured tokens to evaluate the local context,
  ;       absorbing 'Mod::Identifier' qualifications into a single token (with 'Mod' stored in ModulePrefix$)
  ;       unless 'TargetPosInLine' specifically points to the module identifier itself.
  ; Parameters:
  ;   *CodeLineCS     - Pointer to a null-terminated string buffer.
  ;   TargetPosInLine - 1-based offset of the inspected expression on the line (0 to merge all 'Mod::' prefixes).
  ;
  Protected *Char.Character = *CodeLineCS
  Protected *TempChar.Character
  Protected CurrentLevel, CurrentPos, QualPos, FirstChar
  Protected TokenStr$, QualName$, ModPrefix$
  ;
  ClearList(TokenList())
  If *CodeLineCS = 0
    ProcedureReturn
  EndIf
  ;
  CompilerIf #PBBReport
    Protected SavedOPDTokenLine$ = OPDTokenLine$
  CompilerEndIf
  ;
  While *Char\c <> 0
    SkipSpaces(*Char)
    If *Char\c = 0 Or *Char\c = #CR Or *Char\c = ';'
      Break
    EndIf
    ;
    CurrentPos = ComputeLenFromChars(*CodeLineCS, *Char) + 1
    ;
    If IsStringOrCharLiteralStart(*Char)
      TCL_AppendToken(ExtractOrSkipNextString(@*Char, *Char\c), #TokenType_StringInQuotes)
      ;
    ElseIf IsPBSeparator(*Char\c) And *Char\c <> '*' And *Char\c <> '#'
      TokenStr$ = Chr(*Char\c)
      If *Char\c = ')'
        CurrentLevel - 1
      EndIf
      TCL_AppendToken(TokenStr$, #TokenType_Symbol)
      If *Char\c = '('
        CurrentLevel + 1
      EndIf
      GoNextChar(*Char)
      ;
      ; Recognize standalone or unmerged module scope operator ('::') as a single symbol token
      If TokenStr$ = ":" And *Char\c = ':'
        TokenList()\Value$ + ":"
        GoNextChar(*Char)
      EndIf
      ;
    Else
      TokenStr$ = ExtractOrSkipNextToken(*CodeLineCS, @*Char)
      If TokenStr$
        ; Absorb 'Mod::Identifier' into a single qualified token unless 'Mod' itself is the targeted token
        FirstChar = Asc(TokenStr$)
        If CurrentPos <> TargetPosInLine And CurrentPos <> TargetPosInLine - 1 And IsValidIdentifierStart(FirstChar)
          *TempChar = *Char
          SkipSpaces(*TempChar)
          If IsModuleUsage(@*TempChar)
            SkipSpaces(*TempChar)
            QualPos   = ComputeLenFromChars(*CodeLineCS, *TempChar) + 1
            QualName$ = ExtractOrSkipNextToken(*CodeLineCS, @*TempChar)
            If QualName$ <> ""
              *Char      = *TempChar
              ModPrefix$ = TokenStr$
              TokenStr$  = QualName$
              CurrentPos = QualPos
            EndIf
          EndIf
        EndIf
        ;
        TCL_AppendToken(TokenStr$, #TokenType_Word)
        TokenList()\Value_LCase$  = LCase(TokenStr$)
        TokenList()\ModulePrefix$ = ModPrefix$
        ModPrefix$ = ""
      Else
        ; Standalone '*' or '#' not beginning an identifier
        TCL_AppendToken(Chr(*Char\c), #TokenType_Symbol)
        GoNextChar(*Char)
      EndIf
    EndIf
  Wend
  ;
  CompilerIf #PBBReport
    OPDTokenLine$ = SavedOPDTokenLine$
  CompilerEndIf
EndProcedure
;
; **********************************************************************************
;
;-                 Virtual coordinates & code sequences
;
; **********************************************************************************
;
Procedure   ComputePositionInVirtualWholeProjectFile(FileName$, PosInFile)
  ;
  ; Role: Converts a local character position within 'FileName$' into an absolute
  ;       character position within the theoretical 'Virtual Whole Project File'.
  ;       A single pass over PBB_FileListing() simultaneously retrieves the parent
  ;       insertion point and sums the TotalVirtualSize of all children included at or before 'PosInFile'.
  ;
  Protected ParentFileName$, PosInParent = -1
  Protected ChildrenVirtualSize
  ;
  PushListPosition(PBB_FileListing())
  ForEach PBB_FileListing()
    If PBB_FileListing()\FileName$ = FileName$
      ParentFileName$ = PBB_FileListing()\ParentFileName
      PosInParent     = PBB_FileListing()\PosInParent
    EndIf
    If PBB_FileListing()\ParentFileName = FileName$
      If PBB_FileListing()\PosInParent <> -1 And PBB_FileListing()\PosInParent <= PosInFile
        ChildrenVirtualSize + PBB_FileListing()\TotalVirtualSize
      EndIf
    EndIf
  Next
  PopListPosition(PBB_FileListing())
  ;
  If ParentFileName$ <> "" And PosInParent <> -1
    ProcedureReturn ComputePositionInVirtualWholeProjectFile(ParentFileName$, PosInParent - 1) + PosInFile + ChildrenVirtualSize
  EndIf
  ProcedureReturn PosInFile + ChildrenVirtualSize
EndProcedure
;
Procedure   IsPositionBefore(ReferenceFile$, ReferencePos, TestedFile$, TestedPos)
  ;
  ; Role: Determines whether a Reference occurrence chronologically precedes a Tested occurrence
  ;       in the final compiler sequence across the entire project (fast-path on identical files,
  ;       O(N) lookup in PBB_CompilationSequence(), or dynamic tree fallback).
  ;
  Protected RefVirtualPos = -1, TestVirtualPos = -1
  ;
  If ReferenceFile$ = TestedFile$
    ProcedureReturn Bool(TestedPos >= ReferencePos)
  EndIf
  If ReferenceFile$ = "" Or TestedFile$ = ""
    ProcedureReturn #False
  EndIf
  ;
  ; Attempt resolution via pre-computed compilation sequence segments
  PushListPosition(PBB_CompilationSequence())
  ForEach PBB_CompilationSequence()
    If RefVirtualPos = -1 And PBB_CompilationSequence()\FileName$ = ReferenceFile$
      If ReferencePos >= PBB_CompilationSequence()\LocalStartPos And ReferencePos <= PBB_CompilationSequence()\LocalEndPos
        RefVirtualPos = PBB_CompilationSequence()\VirtualStartPos + (ReferencePos - PBB_CompilationSequence()\LocalStartPos)
      EndIf
    EndIf
    If TestVirtualPos = -1 And PBB_CompilationSequence()\FileName$ = TestedFile$
      If TestedPos >= PBB_CompilationSequence()\LocalStartPos And TestedPos <= PBB_CompilationSequence()\LocalEndPos
        TestVirtualPos = PBB_CompilationSequence()\VirtualStartPos + (TestedPos - PBB_CompilationSequence()\LocalStartPos)
      EndIf
    EndIf
    If RefVirtualPos <> -1 And TestVirtualPos <> -1
      Break
    EndIf
  Next
  PopListPosition(PBB_CompilationSequence())
  ;
  If RefVirtualPos <> -1 And TestVirtualPos <> -1
    ProcedureReturn Bool(TestVirtualPos >= RefVirtualPos)
  EndIf
  ;
  ; Fallback: compute virtual project positions dynamically
  ProcedureReturn Bool(ComputePositionInVirtualWholeProjectFile(TestedFile$, TestedPos) >= ComputePositionInVirtualWholeProjectFile(ReferenceFile$, ReferencePos))
EndProcedure
;
Procedure   Internal_BuildCompilationSequence(CurrentFile$, StartPos, EndPos)
  ;
  ; Role: Recursively fragments a physical source file into logical compilation segments
  ;       based on its active 'IncludeFile'/'XIncludeFile' directives, injecting included files at their exact positions.
  ;       Note: Call via BuildCompilationSequence() to ensure PBB_CompilationSequence() is cleared upfront.
  ;
  Protected LastSegmentPos = StartPos
  Protected CurrentVirtualPos = ComputePositionInVirtualWholeProjectFile(CurrentFile$, StartPos)
  ;
  PushListPosition(PBB_FileListing())
  ForEach PBB_FileListing()
    If PBB_FileListing()\ExcludedByCompilerIf = #False And
       PBB_FileListing()\ParentFileName = CurrentFile$ And
       PBB_FileListing()\PosInParent >= StartPos And PBB_FileListing()\PosInParent <= EndPos
      ; 1. Create a segment for the code preceding this inclusion
      If PBB_FileListing()\PosInParent > LastSegmentPos
        AddElement(PBB_CompilationSequence())
        PBB_CompilationSequence()\FileName$       = CurrentFile$
        PBB_CompilationSequence()\LocalStartPos   = LastSegmentPos
        PBB_CompilationSequence()\LocalEndPos     = PBB_FileListing()\PosInParent - 1
        PBB_CompilationSequence()\VirtualStartPos = CurrentVirtualPos
        CurrentVirtualPos + (PBB_FileListing()\PosInParent - LastSegmentPos)
      EndIf
      ;
      ; 2. Recursively dive into the included file and advance cursors past it
      Internal_BuildCompilationSequence(PBB_FileListing()\FileName$, 1, Len(PBB_FileListing()\FileContentInList\s))
      CurrentVirtualPos + PBB_FileListing()\TotalVirtualSize
      LastSegmentPos = PBB_FileListing()\PosInParent
    EndIf
  Next
  PopListPosition(PBB_FileListing())
  ;
  ; 3. Create the final segment for the remaining code after the last inclusion
  If LastSegmentPos <= EndPos
    AddElement(PBB_CompilationSequence())
    PBB_CompilationSequence()\FileName$       = CurrentFile$
    PBB_CompilationSequence()\LocalStartPos   = LastSegmentPos
    PBB_CompilationSequence()\LocalEndPos     = EndPos
    PBB_CompilationSequence()\VirtualStartPos = CurrentVirtualPos
  EndIf
EndProcedure
;
Procedure   BuildCompilationSequence(CurrentFile$, StartPos, EndPos)
  ;
  ; Role: Purges PBB_CompilationSequence() and rebuilds the complete project compilation sequence.
  ;
  ClearList(PBB_CompilationSequence())
  Internal_BuildCompilationSequence(CurrentFile$, StartPos, EndPos)
EndProcedure
;
Procedure   UpdateVirtualPositionsAndSequence(RootFileName$, RebuildSequence = #True)
  ;
  ; Role: Computes 'TotalVirtualSize' and 'PosInVirtualFile' for all active files in 'PBB_FileListing()',
  ;       sorts the file list in chronological compilation order, and rebuilds 'PBB_CompilationSequence()'
  ;       when 'RebuildSequence' is #True.
  ;
  Protected OwnSize, PosInParent
  Protected ParentName$
  ;
  ; 1. Initialize each active file's virtual size to its own physical buffer length (0 when excluded by CompilerIf)
  ForEach PBB_FileListing()
    If PBB_FileListing()\ExcludedByCompilerIf
      PBB_FileListing()\TotalVirtualSize = 0
    Else
      PBB_FileListing()\TotalVirtualSize = Len(PBB_FileListing()\FileContentInList\s)
    EndIf
  Next
  ;
  ; 2. Propagate each active included file's size up its entire ancestor chain
  ForEach PBB_FileListing()
    If PBB_FileListing()\ExcludedByCompilerIf
      Continue
    EndIf
    OwnSize     = Len(PBB_FileListing()\FileContentInList\s)
    ParentName$ = PBB_FileListing()\ParentFileName
    PosInParent = PBB_FileListing()\PosInParent
    ;
    PushListPosition(PBB_FileListing())
    While ParentName$ <> "" And PosInParent <> -1
      If IsFileInMemory(ParentName$)
        PBB_FileListing()\TotalVirtualSize + OwnSize
        ParentName$ = PBB_FileListing()\ParentFileName
        PosInParent = PBB_FileListing()\PosInParent
      Else
        Break
      EndIf
    Wend
    PopListPosition(PBB_FileListing())
  Next
  ;
  ; 3. Compute absolute virtual project coordinates and sort files in compiler inclusion order
  ForEach PBB_FileListing()
    PBB_FileListing()\PosInVirtualFile = ComputePositionInVirtualWholeProjectFile(PBB_FileListing()\FileName$, 1)
  Next
  SortStructuredList(PBB_FileListing(), #PB_Sort_Ascending, OffsetOf(FileListingStruct\PosInVirtualFile), #PB_Integer)
  ;
  ; 4. Rebuild the linear compilation sequence segments from the root file when requested
  If RebuildSequence And RootFileName$ <> "" And IsFileInMemory(RootFileName$)
    BuildCompilationSequence(RootFileName$, 1, Len(PBB_FileListing()\FileContentInList\s))
  EndIf
EndProcedure
;
; ----------------------------------------------------------------------------------
;                  ---- Multi-sequence path search engine ----
; ----------------------------------------------------------------------------------
;
Enumeration FIPSModeEnum
  #FIPSMode_LocalOnly
  #FIPSMode_Global
  #FIPSMode_Hybrid
EndEnumeration
;
Structure FIPSPosStruct
  ;
  ; --- Public Context, Configuration & Results ---
  PosInCode.i
  FileName$
  FIPSContinue.i
  ScanMode.i         ; #FIPSMode_LocalOnly, #FIPSMode_Global, or #FIPSMode_Hybrid
  IncludeDeadZones.b ; #True to include DeadZones() and traverse ExcludedByCompilerIf files at the end of Global scans
  SegmentEnd.i
  ;
  ; --- Internal Iterator State (Persisted across sequential calls) ---
  IsLocalScan.i
  InExcludedFiles.b
  EndLimit.i
  InitialLocalStart.i
  TargetFileName$
  *CurrentSeqElement
  *CurrentExclFile
  *SearchString.String
  *SearchString_LCase.String
  SearchState.FIPC_KeywordCache
EndStructure
;
Procedure   FindInWholeProjectSequences(KeywordList$, *FIPSParam.FIPSPosStruct, CheckMode = #CheckAll)
  ;
  ; Role: Universal multi-sequence keyword search engine traversing the compilation timeline.
  ;       Supports #FIPSMode_LocalOnly (bounded within FileName$), #FIPSMode_Global (chronological
  ;       project-wide scan with optional EndLimit cutoff in FileName$), and #FIPSMode_Hybrid
  ;       (local scan followed by global fallback clamped to InitialLocalStart).
  ;       Set *FIPSParam\FIPSContinue = #True to resume from PosInCode + 1 on subsequent calls.
  ; Output: Returns the matched 1-based offset (also stored in *FIPSParam\PosInCode) and populates
  ;         SegmentEnd, *SearchString, and *SearchString_LCase.
  ;
  Shared PBB_CompilationSequence(), PBB_FileListing()
  Protected CurrentPos, CurrentEnd, FoundPos, EndLimit
  ;
  If *FIPSParam = 0 Or KeywordList$ = "" : ProcedureReturn 0 : EndIf
  KeywordList$ = LCase(KeywordList$)
  EndLimit     = *FIPSParam\EndLimit
  ;
  ; 1. Iterator initialization and state recovery
  If *FIPSParam\FIPSContinue
    CurrentPos = *FIPSParam\PosInCode + 1
    If *FIPSParam\InExcludedFiles And *FIPSParam\CurrentExclFile
      ChangeCurrentElement(PBB_FileListing(), *FIPSParam\CurrentExclFile)
    ElseIf *FIPSParam\IsLocalScan = #False And *FIPSParam\CurrentSeqElement
      ChangeCurrentElement(PBB_CompilationSequence(), *FIPSParam\CurrentSeqElement)
    EndIf
  Else
    *FIPSParam\TargetFileName$        = *FIPSParam\FileName$
    *FIPSParam\InExcludedFiles        = #False
    *FIPSParam\CurrentSeqElement      = 0
    *FIPSParam\CurrentExclFile        = 0
    *FIPSParam\SearchState\BufferAddr = 0
    ;
    Select *FIPSParam\ScanMode
      Case #FIPSMode_LocalOnly, #FIPSMode_Hybrid
        If *FIPSParam\FileName$ = "" Or *FIPSParam\EndLimit = 0
          Alert("Wrong Parameters in FindInWholeProjectSequences: Local/Hybrid scan requires FileName$ and EndLimit!")
          ProcedureReturn 0
        EndIf
        *FIPSParam\IsLocalScan       = #True
        *FIPSParam\InitialLocalStart = *FIPSParam\PosInCode
        CurrentPos                   = *FIPSParam\PosInCode
        ;
      Case #FIPSMode_Global
        If *FIPSParam\EndLimit > 0
          If *FIPSParam\FileName$ = ""
            Alert("Wrong Parameters in FindInWholeProjectSequences: Clamped global scan requires FileName$!")
            ProcedureReturn 0
          EndIf
        Else
          *FIPSParam\TargetFileName$ = ""
        EndIf
        *FIPSParam\IsLocalScan = #False
        CurrentPos             = 0
        ResetList(PBB_CompilationSequence())
    EndSelect
  EndIf
  ;
  ; 2. Unified sequential traversal loop
  Repeat
    If *FIPSParam\IsLocalScan
      CurrentEnd = *FIPSParam\EndLimit
      If IsFileInMemory(*FIPSParam\FileName$) = #False
        If *FIPSParam\ScanMode = #FIPSMode_LocalOnly : Break : EndIf
        *FIPSParam\IsLocalScan            = #False
        *FIPSParam\EndLimit               = *FIPSParam\InitialLocalStart
        EndLimit                          = *FIPSParam\InitialLocalStart
        CurrentPos                        = 0
        *FIPSParam\SearchState\BufferAddr = 0
        ResetList(PBB_CompilationSequence())
        Continue
      EndIf
      *FIPSParam\SegmentEnd = CurrentEnd
      ForEach PBB_CompilationSequence()
        If PBB_CompilationSequence()\FileName$ = *FIPSParam\FileName$ And
           PBB_CompilationSequence()\LocalStartPos <= CurrentPos And PBB_CompilationSequence()\LocalEndPos > CurrentPos
          *FIPSParam\SegmentEnd = PBB_CompilationSequence()\LocalEndPos
          Break
        EndIf
      Next
    ElseIf *FIPSParam\InExcludedFiles
      If CurrentPos = 0
        Repeat
          If NextElement(PBB_FileListing()) = 0 : Break 2 : EndIf
        Until PBB_FileListing()\ExcludedByCompilerIf And PBB_FileListing()\FileContentInList\s <> ""
        *FIPSParam\CurrentExclFile = @PBB_FileListing()
        CurrentPos                 = 1
      EndIf
      *FIPSParam\FileName$  = PBB_FileListing()\FileName$
      *FIPSParam\SegmentEnd = Len(PBB_FileListing()\FileContentInList\s)
      CurrentEnd            = *FIPSParam\SegmentEnd
    Else
      If CurrentPos = 0
        If NextElement(PBB_CompilationSequence()) = 0
          If *FIPSParam\IncludeDeadZones And *FIPSParam\TargetFileName$ = ""
            *FIPSParam\InExcludedFiles = #True
            ResetList(PBB_FileListing())
            Continue
          EndIf
          Break
        EndIf
        If *FIPSParam\TargetFileName$ <> "" And PBB_CompilationSequence()\FileName$ = *FIPSParam\TargetFileName$ And
           PBB_CompilationSequence()\LocalStartPos >= EndLimit
          Break
        EndIf
        If IsFileInMemory(PBB_CompilationSequence()\FileName$) = #False : Continue : EndIf
        CurrentPos = PBB_CompilationSequence()\LocalStartPos
      ElseIf IsFileInMemory(PBB_CompilationSequence()\FileName$) = #False
        CurrentPos = 0
        Continue
      EndIf
      *FIPSParam\FileName$  = PBB_CompilationSequence()\FileName$
      *FIPSParam\SegmentEnd = PBB_CompilationSequence()\LocalEndPos
      CurrentEnd            = PBB_CompilationSequence()\LocalEndPos
      If *FIPSParam\TargetFileName$ <> "" And *FIPSParam\FileName$ = *FIPSParam\TargetFileName$ And CurrentEnd >= EndLimit
        CurrentEnd = EndLimit
      EndIf
    EndIf
    ;
    *FIPSParam\SearchString       = PBB_FileListing()\FileContentInList
    *FIPSParam\SearchString_LCase = PBB_FileListing()\FileLCaseContentInList
    ;
    FoundPos = FindInPBPointedCode(*FIPSParam\SearchString_LCase, KeywordList$, CurrentPos, CheckMode, #PB_String_CaseSensitive, CurrentEnd, 0, @*FIPSParam\SearchState)
    If FoundPos > 0
      If *FIPSParam\IncludeDeadZones = #False
        CurrentPos = SkipDeadZoneInCurrentFile(FoundPos)
        If CurrentPos > FoundPos
          If CurrentPos > CurrentEnd : CurrentPos = 0 : Else : Continue : EndIf
          FoundPos = 0
        EndIf
      EndIf
      If FoundPos > 0
        *FIPSParam\PosInCode = FoundPos
        If *FIPSParam\IsLocalScan = #False And *FIPSParam\InExcludedFiles = #False
          *FIPSParam\CurrentSeqElement = @PBB_CompilationSequence()
        EndIf
        ProcedureReturn FoundPos
      EndIf
    EndIf
    ;
    ; Current segment exhausted without a match: handle scope transition
    If *FIPSParam\IsLocalScan
      If *FIPSParam\ScanMode = #FIPSMode_LocalOnly : Break : EndIf
      *FIPSParam\IsLocalScan = #False
      *FIPSParam\EndLimit    = *FIPSParam\InitialLocalStart
      EndLimit               = *FIPSParam\InitialLocalStart
      ResetList(PBB_CompilationSequence())
    ElseIf *FIPSParam\InExcludedFiles = #False And *FIPSParam\TargetFileName$ <> "" And
           *FIPSParam\FileName$ = *FIPSParam\TargetFileName$ And PBB_CompilationSequence()\LocalEndPos >= EndLimit
      Break
    EndIf
    CurrentPos                        = 0
    *FIPSParam\SearchState\BufferAddr = 0
  ForEver
  ;
  *FIPSParam\PosInCode         = 0
  *FIPSParam\SegmentEnd        = 0
  *FIPSParam\CurrentSeqElement = 0
  ProcedureReturn 0
EndProcedure
;
; **********************************************************************************
;
;-                        Encompassing element detection
;
; **********************************************************************************
;
Structure EEDStruct
  ; EncompassingElement info:
  ; Information about the procedure, macro, enumeration, etc.
  ; encompassing a given position.
  ElementName$
  TypeOfElementName$
  PrefixeCache$
  StartingPos.i
  EndingPos.i
  LastSearchPos.i
  FileName$
  SearchStringAdress.i
  LastUnsuccessfulSearchContext$
  LastUnsuccessfulStartingSearchPos.i
EndStructure
;
Macro       EED_MatchOrTrackHorizon(CandStart, CandEnd, CandDecl, CheckParens, BreakDepth)
  ;
  ; Role: Inline zero-overhead helper evaluating whether PosInCode falls inside [CandStart, CandEnd]
  ;       or updating NextUpcomingStartPos for GetEncompassingLimitsAndName_ListMode().
  ;
  If PosInCode >= CandStart And PosInCode <= CandEnd
    *EED\StartingPos  = CandStart
    *EED\EndingPos    = CandEnd
    *EED\ElementName$ = ElementList()\Name
    CompilerIf CheckParens
      If ElementList()\HasParentheses And Right(*EED\ElementName$, 2) <> "()"
        *EED\ElementName$ + "()"
      EndIf
    CompilerEndIf
    If TypeOfElement = #PBBModule And LCase(Left(CandDecl, 13)) = "declaremodule"
      *EED\TypeOfElementName$ = "DeclareModule"
    Else
      *EED\TypeOfElementName$ = PBBTypeNames$(TypeOfElement)
    EndIf
    *EED\FileName$          = FileName$
    *EED\SearchStringAdress = *SearchString
    Break BreakDepth
  ElseIf CandStart > PosInCode And CandStart < NextUpcomingStartPos
    NextUpcomingStartPos = CandStart
  EndIf
EndMacro
;
Procedure   GetEncompassingLimitsAndName_ListMode(TypeOfElement, FileName$, PosInCode, *SearchString, *EED.EEDStruct)
  ;
  ; Role: Resolves the encompassing syntactic container for PosInCode by querying
  ;       the pre-populated ElementList() AST instead of scanning raw source code buffers.
  ;       While background exploration is incomplete, BuildLayout() falls back to text-scanning
  ;       mode; once completed, List Mode eliminates string searches and computes both the active
  ;       container bounds and the file-wide downstream horizon (NextUpcomingStartPos) in a single pass.
  ;
  Protected NextUpcomingStartPos = $7FFFFFFF
  Protected NewList CEDList.CED_Structure()
  ;
  *EED\ElementName$ = ""
  ;
  ForEach ElementList()
    If ElementList()\Type = TypeOfElement
      ; 1. Evaluate active primary declaration bounds (skipping StartingPos = 0 records invalidated during PartialUpdateOfCode)
      If ElementList()\FileName = FileName$ And ElementList()\StartingPos > 0
        EED_MatchOrTrackHorizon(ElementList()\StartingPos, ElementList()\EndingPos, ElementList()\Declaration, 1, 1)
      EndIf
      ;
      ; 2. Evaluate secondary blocks for multi-part Enumerations and DeclareModule/Module pairs
      If (TypeOfElement = #PBBEnumeration Or TypeOfElement = #PBBModule) And ElementList()\Continued <> ""
        CED_StringToList(ElementList()\Continued, CEDList())
        ForEach CEDList()
          If CEDList()\Filename = FileName$
            EED_MatchOrTrackHorizon(CEDList()\StartingPos, CEDList()\EndingPos, CEDList()\Declaration, 0, 2)
          EndIf
        Next
      EndIf
    EndIf
  Next
  ;
  If *EED\ElementName$ = ""
    ; PosInCode is outside any element of this type: record the downstream horizon for EARLY EXIT 4
    *EED\StartingPos   = NextUpcomingStartPos
    *EED\EndingPos     = 0
    *EED\LastSearchPos = PosInCode
  EndIf
EndProcedure
;
Procedure.s GetEncompassingLimitsAndName(TypeOfElement, FileName$, PosInCode,
                                         *SearchString.String, *SearchString_LCase.String, EndLimit,
                                         *EncompassingElementData.EEDStruct = 0, Init = 0)
  ;
  ; Role: Determines, for a given position in the code, whether it resides inside
  ;       an encompassing syntactic element (such as a Procedure, Macro, Structure, etc.).
  ;       Populates *EncompassingElementData if provided and returns the formatted display prefix.
  ;
  CompilerIf #PBBReport
    If OPDMod & #OPDFollowGetEncompassing
      If TypeOfElement And TypeOfElement < #EndEnumPBBTypes
        OPDReport("GetEncompassing for " + PBBTypeNames$(TypeOfElement))
      EndIf
      Protected PreComp$
    EndIf
    If OPDMod & #OPDBuildOutExternalProcedureTiming
      StartTimer(6, "GetEncompassingLimitsAndName")
    EndIf
  CompilerEndIf
  ;
  Protected StartKeyWord$, EndKeyWord$, NoLine$, CatTxt_CodeMismatchAlert$, ReturnValue$
  Protected mStartOfElementPos, pPar, TOE, FoundPos, SearchEndLimit, SearchPos, InDeadZone
  Protected SearchState.FIPC_KeywordCache
  Static Dim EncompassingElementData.EEDStruct(#PBBWith)
  Static mFileName$, *mSearchString
  ;
  If Init Or *mSearchString <> *SearchString Or mFileName$ <> FileName$
    mFileName$     = FileName$
    *mSearchString = *SearchString
    ;
    ; Reset the full cache array across all defined elements to avoid carrying over stale positions
    For TOE = 0 To ArraySize(EncompassingElementData())
      ClearStructure(@EncompassingElementData(TOE), EEDStruct)
    Next
    ;
    ; Pre-load lowercase buffer into reverse search cache
    If *SearchString_LCase
      ReverseFindInPBCode(*SearchString_LCase\s, "", 0)
    EndIf
    ;
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowGetEncompassing
        If PreComp$ : PreComp$ + #CR$ : EndIf
        PreComp$ = "INIT  "
      EndIf
    CompilerEndIf
    ;
  ElseIf EncompassingElementData(TypeOfElement)\EndingPos <> 0 And PosInCode > EncompassingElementData(TypeOfElement)\EndingPos
    ; EARLY EXIT 1: Invalidate cached boundaries once parsing advances past the ending keyword
    EncompassingElementData(TypeOfElement)\StartingPos = 0
    EncompassingElementData(TypeOfElement)\EndingPos   = 0
    ;
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowGetEncompassing
        If PreComp$ : PreComp$ + #CR$ : EndIf
        PreComp$ + PBBTypeNames$(TypeOfElement) + "  PosInCode > EED\EndingPos  "
      EndIf
    CompilerEndIf
    ;
  ElseIf EncompassingElementData(TypeOfElement)\EndingPos <> 0
    ; EARLY EXIT 2: Reuse known element data when PosInCode lies within active boundaries
    If PosInCode < EncompassingElementData(TypeOfElement)\StartingPos
      ; Invalidate cached downstream boundaries when parsing jumps backward to an earlier position
      ClearStructure(@EncompassingElementData(TypeOfElement), EEDStruct)
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowGetEncompassing
          If PreComp$ : PreComp$ + #CR$ : EndIf
          PreComp$ + PBBTypeNames$(TypeOfElement) + "  PosInCode < EED\StartingPos  "
          PreComp$ + #CR$ + "---------------------------------------------"
          OPDReport(PreComp$)
        EndIf
        If OPDMod & #OPDBuildOutExternalProcedureTiming
          ReapTimer(6)
        EndIf
      CompilerEndIf
    Else
      If *EncompassingElementData
        CopyStructure(@EncompassingElementData(TypeOfElement), *EncompassingElementData, EEDStruct)
      EndIf
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowGetEncompassing
          If PreComp$ : PreComp$ + #CR$ : EndIf
          PreComp$ + PBBTypeNames$(TypeOfElement) + "  Pre-existing Data. Use it!  "
          PreComp$ + #CR$ + "---------------------------------------------"
          OPDReport(PreComp$)
        EndIf
        If OPDMod & #OPDBuildOutExternalProcedureTiming
          ReapTimer(6)
        EndIf
      CompilerEndIf
      ProcedureReturn EncompassingElementData(TypeOfElement)\PrefixeCache$
    EndIf
  EndIf
  ;
  If TypeOfElement And *SearchString And *SearchString_LCase
    ; EARLY EXIT 4: Pre-start horizon check when EndingPos = 0 and StartingPos holds the next upcoming start
    If EncompassingElementData(TypeOfElement)\EndingPos = 0 And EncompassingElementData(TypeOfElement)\StartingPos > 0
      If PosInCode < EncompassingElementData(TypeOfElement)\LastSearchPos
        EncompassingElementData(TypeOfElement)\StartingPos   = 0
        EncompassingElementData(TypeOfElement)\LastSearchPos = 0
      ElseIf PosInCode < EncompassingElementData(TypeOfElement)\StartingPos
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowGetEncompassing
            If PreComp$ : PreComp$ + #CR$ : EndIf
            PreComp$ + "  Current position precedes known upcoming element start  "
            OPDReport(PreComp$)
          EndIf
          If OPDMod & #OPDBuildOutExternalProcedureTiming
            ReapTimer(6)
          EndIf
        CompilerEndIf
        ProcedureReturn ""
      EndIf
    EndIf
    ;
    ; Dual-mode dispatch: query ElementList() once exploration is completed (falling back to Text Mode inside DeadZones)
    InDeadZone = IsPosInDeadZone(FileName$, PosInCode)
    If ExplorationState = #ExplorationState_Completed And ListSize(ElementList()) And
       TypeOfElement <> #PBBWith And InDeadZone = #False
      CompilerIf #PBBReport
        If OPDMod & #OPDFollowGetEncompassing
          If PreComp$ : PreComp$ + #CR$ : EndIf
          PreComp$ + "  ListMode lookup for " + PBBTypeNames$(TypeOfElement) + " from " + Str(PosInCode)
        EndIf
      CompilerEndIf
      GetEncompassingLimitsAndName_ListMode(TypeOfElement, FileName$, PosInCode, *SearchString, @EncompassingElementData(TypeOfElement))
    Else
      ; Text-scanning mode: used while background exploration is still running, or for 'With ... EndWith' blocks
      If TypeOfElement = #PBBWith
        StartKeyWord$ = "with"
        EndKeyWord$   = "endwith"
      ElseIf TypeOfElement = #PBBModule
        StartKeyWord$ = "module,declaremodule"
        EndKeyWord$   = "endmodule,enddeclaremodule"
      Else
        StartKeyWord$ = LCase(PBBTypeNames$(TypeOfElement))
        If TypeOfElement = #PBBEnumeration
          StartKeyWord$ + ",enumerationbinary"
        ElseIf TypeOfElement = #PBBProcedure
          StartKeyWord$ + ",procedurec,proceduredll,procedurecdll"
        ElseIf TypeOfElement = #PBBImport
          StartKeyWord$ + ",importc"
        EndIf
        EndKeyWord$ = "end" + LCase(PBBTypeNames$(TypeOfElement))
      EndIf
      ;
      ; EARLY EXIT 3: Restrict forward search limit to enclosing Procedure or Macro boundary (excluding #PBBModule)
      If TypeOfElement > #PBBMacro And TypeOfElement <> #PBBModule
        If PosInCode >= EncompassingElementData(#PBBProcedure)\StartingPos And
           PosInCode <= EncompassingElementData(#PBBProcedure)\EndingPos
          CompilerIf #PBBReport
            If OPDMod & #OPDFollowGetEncompassing
              If PreComp$ : PreComp$ + #CR$ : EndIf
              PreComp$ + "  Advance EndLimit to the end of the encompassing Procedure " +
                         EncompassingElementData(#PBBProcedure)\ElementName$ +
                         " (" + Str(EncompassingElementData(#PBBProcedure)\StartingPos) +
                         " / " + Str(EncompassingElementData(#PBBProcedure)\EndingPos) + ")"
            EndIf
          CompilerEndIf
          EndLimit = EncompassingElementData(#PBBProcedure)\EndingPos
        ElseIf PosInCode >= EncompassingElementData(#PBBMacro)\StartingPos And
               PosInCode <= EncompassingElementData(#PBBMacro)\EndingPos
          CompilerIf #PBBReport
            If OPDMod & #OPDFollowGetEncompassing
              If PreComp$ : PreComp$ + #CR$ : EndIf
              PreComp$ + "  Advance EndLimit to the end of the encompassing Macro  "
            EndIf
          CompilerEndIf
          EndLimit = EncompassingElementData(#PBBMacro)\EndingPos
        EndIf
      EndIf
      ;
      ; EARLY EXIT 5: Skip redundant parsing when this element type was already confirmed absent in the active Procedure
      If TypeOfElement <> #PBBProcedure And
         EncompassingElementData(#PBBProcedure)\ElementName$ And
         EncompassingElementData(TypeOfElement)\LastUnsuccessfulSearchContext$ = EncompassingElementData(#PBBProcedure)\ElementName$ And
         PosInCode > EncompassingElementData(#PBBProcedure)\StartingPos And
         PosInCode < EncompassingElementData(#PBBProcedure)\EndingPos And
         PosInCode >= EncompassingElementData(TypeOfElement)\LastUnsuccessfulStartingSearchPos
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowGetEncompassing
            If PreComp$ : PreComp$ + #CR$ : EndIf
            PreComp$ + "  Skipping previous unsuccessful attempt  " + #CR$
            PreComp$ + "      Procedure scope for " + EncompassingElementData(#PBBProcedure)\ElementName$ + ": " +
                       Str(EncompassingElementData(#PBBProcedure)\StartingPos) + " / " +
                       Str(EncompassingElementData(#PBBProcedure)\EndingPos) + #CR$
            PreComp$ + "      PosInCode: " + Str(PosInCode) + #CR$
            PreComp$ + "      LastUnsuccessfulStartingSearchPos: " + Str(EncompassingElementData(TypeOfElement)\LastUnsuccessfulStartingSearchPos)
            OPDReport(PreComp$)
          EndIf
          If OPDMod & #OPDBuildOutExternalProcedureTiming
            ReapTimer(6)
          EndIf
        CompilerEndIf
        ProcedureReturn ""
      EndIf
      ;
      If EncompassingElementData(TypeOfElement)\EndingPos = 0
        EncompassingElementData(TypeOfElement)\ElementName$ = ""
        ; Modules span across file segments: force SearchEndLimit to full file length for #PBBModule
        SearchEndLimit = EndLimit
        If TypeOfElement = #PBBModule
          SearchEndLimit = Len(*SearchString\s)
        EndIf
        ;
        CompilerIf #PBBReport
          If OPDMod & #OPDFollowGetEncompassing
            If PreComp$ : PreComp$ + #CR$ : EndIf
            PreComp$ + "  Searching " + StartKeyWord$ + "/" + EndKeyWord$ + " from " + Str(PosInCode)
          EndIf
        CompilerEndIf
        ; When PosInCode is in active code, ignore block headers/closers located inside inactive CompilerIf zones
        SearchPos              = PosInCode
        SearchState\BufferAddr = 0
        Repeat
          FoundPos = FindInPBPointedCode(*SearchString_LCase, StartKeyWord$ + "," + EndKeyWord$,
                                         SearchPos, #CheckAll, #PB_String_CaseSensitive, SearchEndLimit, 0, @SearchState)
          If FoundPos > 0 And InDeadZone = #False And IsPosInDeadZone(FileName$, FoundPos)
            SearchPos = FoundPos + 1
            If IsFileInMemory(FileName$)
              SearchPos = SkipDeadZoneInCurrentFile(FoundPos)
            EndIf
            FoundPos = 0
            If SearchPos <= SearchEndLimit : Continue : EndIf
          EndIf
          Break
        ForEver
        ;
        ; Narrow polymorphic module closing keyword to the exact token matched
        If TypeOfElement = #PBBModule And FoundPos > 0
          If FastMid(*SearchString_LCase\s, FoundPos, 16) = "enddeclaremodule"
            EndKeyWord$ = "enddeclaremodule"
          ElseIf FastMid(*SearchString_LCase\s, FoundPos, 9) = "endmodule"
            EndKeyWord$ = "endmodule"
          EndIf
        EndIf
        ;
        If FoundPos = 0
          EncompassingElementData(TypeOfElement)\StartingPos   = SearchEndLimit + 1
          EncompassingElementData(TypeOfElement)\EndingPos     = 0
          EncompassingElementData(TypeOfElement)\LastSearchPos = PosInCode
          CompilerIf #PBBReport
            If OPDMod & #OPDFollowGetEncompassing
              If PreComp$ : PreComp$ + #CR$ : EndIf
              PreComp$ + "  " + StartKeyWord$ + " not found  " + #CR$
              PreComp$ + "  EndLimit is: " + Str(EndLimit) + #CR$
              If TypeOfElement <> #PBBProcedure
                PreComp$ + "  Storing LastUnsuccessfulSearchContext$ : " +
                           EncompassingElementData(#PBBProcedure)\ElementName$ +
                           " / " + Str(PosInCode)
              EndIf
            EndIf
          CompilerEndIf
          If TypeOfElement <> #PBBProcedure
            EncompassingElementData(TypeOfElement)\LastUnsuccessfulSearchContext$    = EncompassingElementData(#PBBProcedure)\ElementName$
            EncompassingElementData(TypeOfElement)\LastUnsuccessfulStartingSearchPos = PosInCode
          EndIf
        ElseIf FoundPos + Len(EndKeyWord$) - 1 <= SearchEndLimit And
               FastMid(*SearchString_LCase\s, FoundPos, Len(EndKeyWord$)) = EndKeyWord$
          ; Closing keyword encountered before opening keyword: PosInCode is inside the element
          CompilerIf #PBBReport
            If OPDMod & #OPDFollowGetEncompassing
              If PreComp$ : PreComp$ + #CR$ : EndIf
              PreComp$ + "  " + EndKeyWord$ + " found first  "
            EndIf
          CompilerEndIf
          ;
          EncompassingElementData(TypeOfElement)\EndingPos = FoundPos + Len(EndKeyWord$)
          SearchPos                                        = PosInCode
          While SearchPos >= 1
            mStartOfElementPos = ReverseFindInPBCode("", StartKeyWord$, SearchPos, #CheckAll, #PB_String_CaseSensitive)
            If mStartOfElementPos > 0 And InDeadZone = #False And IsPosInDeadZone(FileName$, mStartOfElementPos)
              SearchPos          = mStartOfElementPos - 1
              mStartOfElementPos = 0
              Continue
            EndIf
            Break
          Wend
          EncompassingElementData(TypeOfElement)\StartingPos = mStartOfElementPos
          ;
          If EncompassingElementData(TypeOfElement)\StartingPos
            mStartOfElementPos = FirstOccurrenceOfCharPointed(*SearchString_LCase, " :;" + #CR$, EncompassingElementData(TypeOfElement)\StartingPos + 1)
            ;
            ; Capture the exact opening variant for polymorphic blocks
            If (TypeOfElement = #PBBEnumeration Or TypeOfElement = #PBBModule) And mStartOfElementPos > EncompassingElementData(TypeOfElement)\StartingPos
              StartKeyWord$ = FastMid(*SearchString_LCase\s, EncompassingElementData(TypeOfElement)\StartingPos, mStartOfElementPos - EncompassingElementData(TypeOfElement)\StartingPos)
            EndIf
            ;
            If mStartOfElementPos And PeekC(PosToChar(@*SearchString_LCase\s, mStartOfElementPos)) = #Space
              While mStartOfElementPos <= SearchEndLimit And PeekC(PosToChar(@*SearchString_LCase\s, mStartOfElementPos)) = #Space
                mStartOfElementPos + 1
              Wend
              If TypeOfElement = #PBBWith Or TypeOfElement = #PBBImport
                EncompassingElementData(TypeOfElement)\ElementName$ = ExtractWithTarget(PosToChar(@*SearchString\s, mStartOfElementPos))
                If TypeOfElement = #PBBImport And EncompassingElementData(TypeOfElement)\ElementName$ = #DOUBLEQUOTE$ + #DOUBLEQUOTE$
                  EncompassingElementData(TypeOfElement)\ElementName$ = ""
                EndIf
              Else
                pPar = FirstOccurrenceOfCharPointed(*SearchString_LCase, " (;:" + #CR$, mStartOfElementPos)
                If pPar And pPar - mStartOfElementPos
                  EncompassingElementData(TypeOfElement)\ElementName$ = Trim(FastMid(*SearchString\s, mStartOfElementPos, pPar - mStartOfElementPos))
                  While PeekC(PosToChar(@*SearchString_LCase\s, pPar)) = #Space : pPar + 1 : Wend
                  If PeekC(PosToChar(@*SearchString_LCase\s, pPar)) = '('
                    EncompassingElementData(TypeOfElement)\ElementName$ + "()"
                  EndIf
                EndIf
              EndIf
            EndIf
            ;
            If EncompassingElementData(TypeOfElement)\ElementName$ = ""
              NoLine$ = Str(GetNoLineFromStrAndPos(*SearchString_LCase, EncompassingElementData(TypeOfElement)\StartingPos))
              EncompassingElementData(TypeOfElement)\ElementName$ = "UnNamed - " + GetFilePart(FileName$) + "/Line " + NoLine$
            EndIf
            ;
            If TypeOfElement = #PBBModule And LCase(StartKeyWord$) = "declaremodule"
              EncompassingElementData(TypeOfElement)\TypeOfElementName$ = "DeclareModule"
            Else
              EncompassingElementData(TypeOfElement)\TypeOfElementName$ = PBBTypeNames$(TypeOfElement)
            EndIf
            EncompassingElementData(TypeOfElement)\FileName$          = FileName$
            EncompassingElementData(TypeOfElement)\SearchStringAdress = *SearchString
          Else
            ; Closing keyword found without a matching opening keyword upstream
            NoLine$                   = Str(GetNoLineFromStrAndPos(*SearchString_LCase, EncompassingElementData(TypeOfElement)\EndingPos))
            CatTxt_CodeMismatchAlert$ = GetTextFromCatalogPB("CodeMismatchAlert")
            CatTxt_CodeMismatchAlert$ = ReplaceString(CatTxt_CodeMismatchAlert$, "%EndString$%", EndKeyWord$)
            CatTxt_CodeMismatchAlert$ = ReplaceString(CatTxt_CodeMismatchAlert$, "%StartString$%", StartKeyWord$)
            CatTxt_CodeMismatchAlert$ = ReplaceString(CatTxt_CodeMismatchAlert$, "%SearchString\s%", *SearchString_LCase\s)
            CatTxt_CodeMismatchAlert$ = ReplaceString(CatTxt_CodeMismatchAlert$, "%nFile$%", FileName$)
            AlertInPBBWindow(CatTxt_CodeMismatchAlert$ + #CR$ + NoLine$)
          EndIf
        Else
          ; Opening keyword found before closing keyword: PosInCode is outside the element
          CompilerIf #PBBReport
            If OPDMod & #OPDFollowGetEncompassing
              If PreComp$ : PreComp$ + #CR$ : EndIf
              PreComp$ + "  " + StartKeyWord$ + " found first  "
            EndIf
          CompilerEndIf
          EncompassingElementData(TypeOfElement)\StartingPos   = FoundPos
          EncompassingElementData(TypeOfElement)\EndingPos     = 0
          EncompassingElementData(TypeOfElement)\LastSearchPos = PosInCode
        EndIf
      EndIf
    EndIf ; End of Dual-Mode Dispatch
    ;
    If EncompassingElementData(TypeOfElement)\ElementName$
      ReturnValue$ = " In" + #NonBreakableSpace$ + EncompassingElementData(TypeOfElement)\TypeOfElementName$ + #NonBreakableSpace$ + EncompassingElementData(TypeOfElement)\ElementName$
      If TypeOfElement = #PBBProcedure Or TypeOfElement = #PBBMacro Or TypeOfElement = #PBBModule Or
         (EncompassingElementData(#PBBProcedure)\EndingPos = 0 And EncompassingElementData(#PBBMacro)\EndingPos = 0)
        ; Prepend line break when outside procedural contexts
        ReturnValue$ = PBB_LineBreak$ + ReturnValue$
      EndIf
      EncompassingElementData(TypeOfElement)\PrefixeCache$ = ReturnValue$
    EndIf
    ;
    If *EncompassingElementData
      CopyStructure(@EncompassingElementData(TypeOfElement), *EncompassingElementData, EEDStruct)
    EndIf
    ;
    CompilerIf #PBBReport
      If OPDMod & #OPDFollowGetEncompassing
        If *EncompassingElementData And EncompassingElementData(TypeOfElement)\ElementName$
          If PreComp$ : PreComp$ + #CR$ : EndIf
          PreComp$ + "  Found: " + EncompassingElementData(TypeOfElement)\ElementName$
        EndIf
        OPDReport(PreComp$)
      EndIf
      If OPDMod & #OPDBuildOutExternalProcedureTiming
        ReapTimer(6)
      EndIf
    CompilerEndIf
    ;
    ProcedureReturn ReturnValue$
  EndIf
EndProcedure
;
; **********************************************************************************
;
;-                    ElementList() & map index management
;
; **********************************************************************************
;
; 'Elements' correspond to the different parts of the code that can be
; isolated within a codebase. Procedures, macros, and structures,
; for example, fit this 'element' concept.
; The complete list of what is defined as an 'element' is found in
; the 'PBBTypes' enumeration in the declaration file:
;   #PBBProcedure
;   #PBBStructure
;   #PBBMacro
;   #PBBEnumeration
;   #PBBInterface
;   #PBBLabel
;   #PBBConstant
;   #PBBVariable
;   ...
;
; When a code file is opened in PBBrowser, it is immediately
; examined to compile a list of its elements. This list is
; stored in 'ElementList()'.
;
Global Dim ElementFixedIndexes.i(1000)
Global NextElementID = 1
Global PBB_HasModules = #False
;
Procedure   AddIDToMapIndex(ElementPurifiedName$, ElementIndex, Map TargetMap.s())
  ;
  ; Role: Appends 'ElementIndex' to the comma-separated list of IDs associated with
  ;       'ElementPurifiedName$' in 'TargetMap()', avoiding duplicate entries and
  ;       reusing the active map cursor set by the lookup to prevent a second hash computation.
  ;
  Protected Result
  Protected EI_Key$ = Str(ElementIndex) + ","
  ;
  If FindString("," + TargetMap(ElementPurifiedName$), "," + EI_Key$) = 0
    TargetMap() + EI_Key$
    Result = #True
  EndIf
  ProcedureReturn Result
EndProcedure
;
Procedure   RemoveIDFromMapIndex(ElementPurifiedName$, ElementIndex, Map TargetMap.s())
  ;
  ; Role: Removes 'ElementIndex' from the comma-separated list of IDs in 'TargetMap()'
  ;       for 'ElementPurifiedName$', and deletes the map entry entirely if no IDs remain
  ;       so orphan keys never pollute O(1) existence checks via FindMapElement().
  ;
  Protected Result
  Protected Search$, Current$
  ;
  If FindMapElement(TargetMap(), ElementPurifiedName$)
    Current$ = "," + TargetMap()
    Search$  = "," + Str(ElementIndex) + ","
    If FindString(Current$, Search$)
      Current$ = ReplaceString(Current$, Search$, ",")
      If Current$ = ","
        DeleteMapElement(TargetMap())
      Else
        TargetMap() = Mid(Current$, 2)
      EndIf
      Result = #True
    EndIf
  EndIf
  ProcedureReturn Result
EndProcedure
;
Procedure   UpdateParametersInMapIndex(Parameters$, ElementIndex, RemoveMode = #False)
  ;
  ; Role: Extracts each tab-delimited formal parameter of a Procedure or Macro,
  ;       normalizes its identifier via PurifyElementName(), and either registers or
  ;       removes 'ElementIndex' in ParamMapIndex() for O(1) parameter homonym resolution.
  ;
  Protected ParamIdx
  Protected CleanParams$ = Trim(Parameters$, #TAB$)
  Protected ParamToken$, ParamPurified$
  ;
  If CleanParams$ = ""
    ProcedureReturn
  EndIf
  ;
  Repeat
    ParamIdx + 1
    ParamToken$ = StringField(CleanParams$, ParamIdx, #TAB$)
    If ParamToken$ = "" : Break : EndIf
    ParamPurified$ = PurifyElementName(ParamToken$)
    If ParamPurified$ <> ""
      If RemoveMode
        RemoveIDFromMapIndex(ParamPurified$, ElementIndex, ParamMapIndex())
      Else
        AddIDToMapIndex(ParamPurified$, ElementIndex, ParamMapIndex())
      EndIf
    EndIf
  ForEver
EndProcedure
;
Macro       RemoveParametersFromMapIndex(Parameters, ElementIndex)
  ;
  ; Role: Removes all formal parameters of 'ElementIndex' from ParamMapIndex().
  ;
  UpdateParametersInMapIndex(Parameters, ElementIndex, #True)
EndMacro
;
Procedure   AddPBBElement(TypeOfElement, ElementName$, PurifiedName$ = "")
  ;
  ; Role: Creates a new element in ElementList(), assigns it a permanent ID,
  ;       stores its lowercase purified name in ElementMapIndex() with the associated ID,
  ;       and stores its memory pointer in ElementFixedIndexes().
  ;
  Protected EFIArraySize
  ;
  If PurifiedName$ = ""
    PurifiedName$ = PurifyElementName(ElementName$)
  EndIf
  ;
  AddElement(ElementList())
  ElementList()\Type         = TypeOfElement
  ElementList()\FixedID      = NextElementID
  ElementList()\Name         = ElementName$
  ElementList()\PurifiedName = PurifiedName$
  ;
  AddIDToMapIndex(ElementList()\PurifiedName, NextElementID, ElementMapIndex())
  ;
  EFIArraySize = ArraySize(ElementFixedIndexes())
  If NextElementID > EFIArraySize
    ReDim ElementFixedIndexes(EFIArraySize + 1000)
  EndIf
  If TypeOfElement = #PBBModule
    PBB_HasModules = #True
  EndIf
  ElementFixedIndexes(NextElementID) = @ElementList()
  NextElementID + 1
  ;
  ProcedureReturn ElementList()\FixedID
EndProcedure
;
Procedure   DeletePBBElement()
  ;
  ; Role: Deletes the current element from ElementList(), removes its name and index
  ;       from ElementMapIndex() and ParamMapIndex(), and invalidates its pointer in ElementFixedIndexes().
  ;
  If ElementList()\Used And ElementList()\Type < #EndEnumPBBElementTypes
    ListOfUsedElementsNbr(ElementList()\Type) - 1
  EndIf
  If ElementList()\Parameters <> ""
    RemoveParametersFromMapIndex(ElementList()\Parameters, ElementList()\FixedID)
  EndIf
  ElementFixedIndexes(ElementList()\FixedID) = 0
  RemoveIDFromMapIndex(ElementList()\PurifiedName, ElementList()\FixedID, ElementMapIndex())
  DeleteElement(ElementList(), #True)
EndProcedure
;
Procedure   SelectPBBElement(FixedID)
  ;
  ; Role: Instantly positions ElementList() on the element corresponding to 'FixedID' in O(1).
  ; Output: Returns the element's memory pointer if it exists and is selected, or 0 otherwise.
  ;
  If FixedID > 0 And FixedID < NextElementID And ElementFixedIndexes(FixedID) <> 0
    ChangeCurrentElement(ElementList(), ElementFixedIndexes(FixedID))
    ProcedureReturn @ElementList()
  EndIf
  ProcedureReturn 0
EndProcedure
;
Procedure   LinkElementUsage(ParentID, ChildID)
  ;
  ; Role: Registers a bidirectional usage dependency between 'ParentID' (\UsageChildren)
  ;       and 'ChildID' (\UsageParents) in O(1) via ElementFixedIndexes() without moving ElementList().
  ;
  Protected *ParentElem.Elements, *ChildElem.Elements
  ;
  If ParentID > 0 And ParentID < NextElementID And ChildID > 0 And ChildID < NextElementID
    *ParentElem = ElementFixedIndexes(ParentID)
    *ChildElem  = ElementFixedIndexes(ChildID)
    If *ParentElem And *ChildElem
      *ParentElem\UsageChildren = AddIndexToList(ChildID, *ParentElem\UsageChildren)
      *ChildElem\UsageParents   = AddIndexToList(ParentID, *ChildElem\UsageParents)
    EndIf
  EndIf
EndProcedure
;
Procedure.s GetEnclosingModuleName(EnclosingModuleID = 0)
  ;
  ; Role: Returns the name of the module enclosing the current element in ElementList()
  ;       (or corresponding to 'EnclosingModuleID' if explicitly provided), or "" if in global scope.
  ;       Reads directly from ElementFixedIndexes() in O(1) without moving ElementList().
  ;
  Protected *ModElement.Elements
  ;
  If EnclosingModuleID = 0 And @ElementList() And ElementList()\Type <> #PBBModule
    EnclosingModuleID = ElementList()\EnclosingModuleID
  EndIf
  If EnclosingModuleID > 0 And EnclosingModuleID < NextElementID
    *ModElement = ElementFixedIndexes(EnclosingModuleID)
    If *ModElement
      ProcedureReturn *ModElement\Name
    EndIf
  EndIf
  ProcedureReturn ""
EndProcedure
;
; **********************************************************************************
;
;-                    Project backup & line-ending helpers
;
; **********************************************************************************
;
Structure FilesToChangeStruct
  SourceFile.s
  TargetFile.s
EndStructure
;
Procedure.s GetDefaultProjectCopyDir()
  ;
  ; Role: Computes the first non-existing numbered backup directory path next to the main project file.
  ;
  Protected ParentTargetDir$ = GetPathPart(PBBFicPrincipalPB$) + FilePartSsExt(PBBFicPrincipalPB$)
  Protected TargetDir$, CopyIndex = 1
  ;
  Repeat
    TargetDir$ = ParentTargetDir$ + " (" + Str(CopyIndex) + ")\"
    CopyIndex + 1
  Until FileSize(TargetDir$) <> -2
  ;
  ProcedureReturn TargetDir$
EndProcedure
;
Procedure.i PrepareProjectTargetFiles(MakeCopy, List FilesToChange.FilesToChangeStruct(), *TargetDir.String = 0)
  ;
  ; Role: Populates FilesToChange() from PBB_FileListing() in a single pass and either copies
  ;       in-project source files into *TargetDir\s (excluding external files) or maps TargetFile to SourceFile.
  ; Output: Returns #True on success, #False if directory creation or file copying fails.
  ;
  Protected SourceDir$, TargetDir$, SrcFile$
  ;
  ClearList(FilesToChange())
  If MakeCopy
    If *TargetDir = 0 Or Trim(*TargetDir\s) = ""
      ProcedureReturn #False
    EndIf
    TargetDir$ = ReplaceString(Trim(*TargetDir\s), "/", "\")
    If Right(TargetDir$, 1) <> "\"
      TargetDir$ + "\"
    EndIf
    *TargetDir\s = TargetDir$
    If CreatePathFolders(TargetDir$) = #False
      ProcedureReturn #False
    EndIf
    ;
    SourceDir$ = GetPathPart(PBBFicPrincipalPB$)
    ForEach PBB_FileListing()
      SrcFile$ = PBB_FileListing()\FileName$
      ; Only copy files located inside the main project root directory
      If LCase(Left(SrcFile$, Len(SourceDir$))) = LCase(SourceDir$)
        AddElement(FilesToChange())
        FilesToChange()\SourceFile = SrcFile$
        FilesToChange()\TargetFile = TargetDir$ + Mid(SrcFile$, Len(SourceDir$) + 1)
        If FileSize(SrcFile$) >= 0
          CreatePathFolders(FilesToChange()\TargetFile)
          If CopyFile(SrcFile$, FilesToChange()\TargetFile) = 0
            AlertInPBBWindow(GetTextFromCatalogPB("CopyFailed"))
            ProcedureReturn #False
          EndIf
        EndIf
      EndIf
    Next
  Else
    If *TargetDir
      *TargetDir\s = ""
    EndIf
    ForEach PBB_FileListing()
      AddElement(FilesToChange())
      FilesToChange()\SourceFile = PBB_FileListing()\FileName$
      FilesToChange()\TargetFile = PBB_FileListing()\FileName$
    Next
  EndIf
  ;
  ProcedureReturn #True
EndProcedure
;
Procedure AskAndPrepareProjectCopy(List FilesToChange.FilesToChangeStruct(), *TargetDir.String = 0)
  ;
  ; Role: Prompts the user to create a backup copy of the project before applying batch modifications,
  ;       copying only the files referenced in PBB_FileListing() into a new folder, and populates FilesToChange().
  ; Output: Returns #PB_MessageRequester_Yes (copy), #PB_MessageRequester_No (in-place), or #PB_MessageRequester_Cancel.
  ;
  Protected Res
  Protected ChosenDir.String
  ;
  If Left(GetFilePart(PBBFicPrincipalPB$), Len("PureBasic_TempFile")) = "PureBasic_TempFile"
    AlertInPBBWindow(GetTextFromCatalogPB("NotForTempFile"))
    ProcedureReturn #PB_MessageRequester_Cancel
  EndIf
  ;
  Res = AlertInPBBWindow(GetTextFromCatalogPB("MakeACopy"), "", 1, 0, 0, 0, "", #AW_YesNoOrCancel_YesByDefault)
  If Res = #PB_MessageRequester_Cancel
    ProcedureReturn #PB_MessageRequester_Cancel
  EndIf
  ;
  If Res = #PB_MessageRequester_Yes
    ChosenDir\s = InputRequester("", GetTextFromCatalogPB("SelectFolder"), GetDefaultProjectCopyDir())
    If ChosenDir\s = ""
      ProcedureReturn #PB_MessageRequester_Cancel
    EndIf
  EndIf
  ;
  If PrepareProjectTargetFiles(Bool(Res = #PB_MessageRequester_Yes), FilesToChange(), @ChosenDir) = #False
    ProcedureReturn #PB_MessageRequester_Cancel
  EndIf
  ;
  If *TargetDir
    *TargetDir\s = ChosenDir\s
  EndIf
  ProcedureReturn Res
EndProcedure
;
; ----------------------------------------------------------------------------------
;                  ---- Line-ending & indentation restoration ----
; ----------------------------------------------------------------------------------
;
EnumerationBinary LineEndingTypes
  #LineEnding_CRLF       = 1
  #LineEnding_LF         = 2
  #LineEnding_CR         = 4
  #LineEnding_IncludeTab = 8
EndEnumeration
;
Procedure   DetectLineEndingAndTab(*Code.String)
  ;
  ; Role: Detects the line ending format used in '*Code\s' and checks for leading indentation tabs.
  ;
  Protected Result
  ;
  If FindString(*Code\s, #LF$ + #TAB$) Or FindString(*Code\s, #CR$ + #TAB$) Or Left(*Code\s, 1) = #TAB$
    Result | #LineEnding_IncludeTab
  EndIf
  ;
  If FindString(*Code\s, #CRLF$)
    ProcedureReturn #LineEnding_CRLF | Result
  ElseIf FindString(*Code\s, #LF$)
    ProcedureReturn #LineEnding_LF | Result
  ElseIf FindString(*Code\s, #CR$)
    ProcedureReturn #LineEnding_CR | Result
  EndIf
  ProcedureReturn #LineEnding_CRLF | Result
EndProcedure
;
Procedure.s DetectFileIndentUnit(*Code.String)
  ;
  ; Role: Scans '*Code\s' to infer the smallest positive leading-space indentation unit used
  ;       on active code lines (e.g., 2 or 4 spaces), falling back to PBBTabSpaces$ if flat.
  ;
  Protected *Char.Character
  Protected SpaceCount, MinIndent = 0
  ;
  If *Code And *Code\s <> ""
    *Char = @*Code\s
    While *Char\c <> 0
      SpaceCount = 0
      While *Char\c = #Space
        SpaceCount + 1
        GoNextChar(*Char)
      Wend
      If *Char\c <> 0 And *Char\c <> #CR And *Char\c <> #LF And *Char\c <> ';'
        If SpaceCount > 0 And (MinIndent = 0 Or SpaceCount < MinIndent)
          MinIndent = SpaceCount
          If MinIndent = 1
            Break
          EndIf
        EndIf
      EndIf
      While *Char\c <> 0 And *Char\c <> #CR And *Char\c <> #LF
        GoNextChar(*Char)
      Wend
      While *Char\c = #CR Or *Char\c = #LF
        GoNextChar(*Char)
      Wend
    Wend
  EndIf
  ;
  If MinIndent > 0
    ProcedureReturn Space(MinIndent)
  EndIf
  ProcedureReturn PBBTabSpaces$
EndProcedure
;
Procedure   FilteredPosToRawPos(*RawCode.String, FilteredPos)
  ;
  ; Role: Maps a 1-based character offset from a filtered code buffer (where #TAB$ was expanded
  ;       to PBBTabSpaces$ and #CRLF$ collapsed to #CR$) to its exact 1-based offset in '*RawCode\s'.
  ;
  Protected *Char.Character
  Protected CurPos = 1, TabStep = Len(PBBTabSpaces$)
  ;
  If *RawCode = 0 Or FilteredPos <= 1
    ProcedureReturn 1
  EndIf
  If TabStep < 1
    TabStep = 2
  EndIf
  ;
  *Char = @*RawCode\s
  While *Char\c <> 0 And CurPos < FilteredPos
    If *Char\c = #TAB
      If CurPos + TabStep > FilteredPos
        Break
      EndIf
      CurPos + TabStep
    Else
      If *Char\c = #CR And PeekC(*Char + SizeOf(Character)) = #LF
        GoNextChar(*Char)
      EndIf
      CurPos + 1
    EndIf
    GoNextChar(*Char)
  Wend
  ;
  ProcedureReturn ComputeLenFromChars(@*RawCode\s, *Char) + 1
EndProcedure
;
Procedure   ApplyLineEnding(*Code.String, LineEndingFormat)
  ;
  ; Role: Converts #CR$ line breaks in '*Code\s' to the target line-ending format.
  ;
  If *Code = 0 Or *Code\s = ""
    ProcedureReturn #False
  EndIf
  ;
  Select LineEndingFormat & (#LineEnding_CRLF | #LineEnding_LF | #LineEnding_CR)
    Case #LineEnding_LF
      *Code\s = ReplaceString(*Code\s, #CR$, #LF$)
    Case #LineEnding_CR
      ; Line endings are already #CR$, no replacement needed
    Default
      *Code\s = ReplaceString(*Code\s, #CR$, #CRLF$)
  EndSelect
  ;
  ProcedureReturn #True
EndProcedure
;
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 2652
; FirstLine = 426
; Folding = AAAAAEAAAAAAkAgBAA+
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb