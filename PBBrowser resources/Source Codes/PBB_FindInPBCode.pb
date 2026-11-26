;***********************************************************************
;
;                            FindInPBCode.pb
;                        Part of PBBrowser project
;                            Zapman - Nov 2026
;
; Set of functions to find and replace strings in a PureBasic source file.
;
;***********************************************************************
;
CompilerIf #PB_Compiler_IsMainFile
  EnableExplicit
  XIncludeFile "../../Zapman libraries/FastStringFunctions.pbi"
CompilerEndIf
;
CompilerIf Not Defined(DOUBLEQUOTE, #PB_Constant)
  #DOUBLEQUOTE = 34
CompilerEndIf
CompilerIf Not Defined(SIMPLEQUOTE, #PB_Constant)
  #SIMPLEQUOTE = 39
CompilerEndIf
CompilerIf Not Defined(NonBreakableSpace$, #PB_Constant)
  #NonBreakableSpace$ = Chr(160)
CompilerEndIf
CompilerIf Not Defined(SpecialPoint$, #PB_Constant)
  #SpecialPoint$ = Chr(8226)
CompilerEndIf
CompilerIf Not Defined(ThisPlaceChar$, #PB_Constant)
  #ThisPlaceChar$ = Chr(8227)
CompilerEndIf
CompilerIf Not Defined(ItsHereChar$, #PB_Constant)
  #ItsHereChar$ = Chr(8259)
CompilerEndIf
;
; Include ASCII backtick (`) and common typographic quotes/apostrophes (‘ ’ « » “ ”) for comment/string tokenization
Global ListPBSepWithoutEndChars$ = " ,();+-=" + #DOUBLEQUOTE$ + "'" + #TAB$ + #NonBreakableSpace$ +
                                   "/*\><@{}|~[]&.!%^?`" + #SpecialPoint$ + #ThisPlaceChar$ +
                                   #ItsHereChar$ + Chr($2018) + Chr($2019) + Chr($00AB) +
                                   Chr($00BB) + Chr($201C) + Chr($201D)
Global ListPBSep$ = ":" + #CR$ + ListPBSepWithoutEndChars$
;
Global Dim IsPBSeparator.b(65535)
Define ctSep.i, TcharCode.i
IsPBSeparator(0) = #True
For ctSep = 1 To Len(ListPBSep$)
  TcharCode = Asc(Mid(ListPBSep$, ctSep, 1))
  If TcharCode >= 0 And TcharCode <= 65535
    IsPBSeparator(TcharCode) = #True
  EndIf
Next
;
EnumerationBinary CheckModesForFindInPBCode
  #NoCheck
  #CheckLeft
  #CheckRight
  #CheckCom
  #CheckQuote
  ;
  #CheckOneKeyWordOnly
  ;
  #CheckComAndQuote         = #CheckCom | #CheckQuote
  #CheckLeftAndRight        = #CheckLeft | #CheckRight
  #CheckLeftAndComAndQuote  = #CheckComAndQuote | #CheckLeft
  #CheckRightAndComAndQuote = #CheckComAndQuote | #CheckRight
  #CheckAll                 = #CheckComAndQuote | #CheckLeftAndRight
EndEnumeration
;
Enumeration CodeLineContext
  #CodeCtx_Code
  #CodeCtx_Comment
  #CodeCtx_Quote
EndEnumeration
;
Structure FIPC_KeywordCache
  BufferAddr.i
  SafeEndPos.i
  EndLimit.i
  FMode.i
  OneKeyFlag.i
  LastMatchPos.i
  NbKeywords.i
  SimpleSearch.i
  CachedKeywords$
  KeywordBuf.String
  KeyPos.i[32]
  KeyLen.i[32]
  KeyPtr.i[32]
EndStructure
;
Procedure CheckCodeLineContext(*CodeString.String, PosInCode, AllowSimpleQuote = #True, BeforePosOnly = #False, *SpanEndPos.Integer = 0)
  ;
  ; Role: Determines whether the 1-based 'PosInCode' offset inside '*CodeString'
  ;       (or the point immediately preceding it when 'BeforePosOnly' is #True)
  ;       resides in active code, a comment (';'), or a string/character literal
  ;       ("...", ~"..." with backslash escapes, or '...' when 'AllowSimpleQuote' is #True).
  ;       Optionally writes the 1-based end position of the enclosing comment or literal to '*SpanEndPos'.
  ; Output: Returns #CodeCtx_Code, #CodeCtx_Comment, or #CodeCtx_Quote.
  ;
  Protected *BufferStart.Character, *LineStart.Character, *ScanChar.Character
  Protected *TargetChar.Character, *LimitChar.Character
  Protected QuoteChar.c, IsEscaped.i
  ; Character pointers bounding the physical line and tracking active quote or comment spans.
  ;
  If *CodeString = 0 Or PosInCode < 1
    ProcedureReturn #CodeCtx_Code
  EndIf
  ;
  *BufferStart = @*CodeString\s
  If *BufferStart = 0 Or *BufferStart\c = 0
    ProcedureReturn #CodeCtx_Code
  EndIf
  ;
  ; Rewind from the target character to the beginning of its physical line
  *TargetChar = *BufferStart + (PosInCode - 1) * SizeOf(Character)
  *LineStart  = *TargetChar
  While *LineStart > *BufferStart
    *ScanChar = *LineStart - SizeOf(Character)
    If *ScanChar\c = #CR Or *ScanChar\c = #LF
      Break
    EndIf
    *LineStart = *ScanChar
  Wend
  ;
  If BeforePosOnly
    *LimitChar = *TargetChar
  Else
    *LimitChar = *TargetChar + SizeOf(Character)
  EndIf
  ;
  ; Scan forward across the line up to *LimitChar to resolve comments and literals
  *ScanChar = *LineStart
  While *ScanChar < *LimitChar And *ScanChar\c <> 0 And *ScanChar\c <> #CR And *ScanChar\c <> #LF
    If *ScanChar\c = ';'
      If *SpanEndPos
        *ScanChar = *TargetChar
        While *ScanChar\c <> 0 And *ScanChar\c <> #CR And *ScanChar\c <> #LF
          *ScanChar + SizeOf(Character)
        Wend
        *SpanEndPos\i = (*ScanChar - *BufferStart) / SizeOf(Character)
      EndIf
      ProcedureReturn #CodeCtx_Comment
    ElseIf (*ScanChar\c = '~' And (Not BeforePosOnly Or (*ScanChar + SizeOf(Character)) < *LimitChar) And
            PeekC(*ScanChar + SizeOf(Character)) = #DOUBLEQUOTE) Or
           *ScanChar\c = #DOUBLEQUOTE Or
           (AllowSimpleQuote And *ScanChar\c = #SIMPLEQUOTE)
      If *ScanChar\c = '~'
        QuoteChar = #DOUBLEQUOTE
        IsEscaped = #True
        *ScanChar + 2 * SizeOf(Character)
      Else
        QuoteChar = *ScanChar\c
        IsEscaped = #False
        *ScanChar + SizeOf(Character)
      EndIf
      ;
      While *ScanChar\c <> 0 And *ScanChar\c <> #CR And *ScanChar\c <> #LF And *ScanChar\c <> QuoteChar
        If IsEscaped And *ScanChar\c = '\'
          *ScanChar + SizeOf(Character)
          If *ScanChar\c = 0 Or *ScanChar\c = #CR Or *ScanChar\c = #LF
            Break
          EndIf
        EndIf
        *ScanChar + SizeOf(Character)
      Wend
      ;
      If *ScanChar\c = QuoteChar
        *ScanChar + SizeOf(Character)
        If *TargetChar < *ScanChar
          If *SpanEndPos
            *SpanEndPos\i = (*ScanChar - *BufferStart) / SizeOf(Character)
          EndIf
          ProcedureReturn #CodeCtx_Quote
        EndIf
      Else
        ; Unclosed literal extends to the end of the physical line, covering *TargetChar
        If *SpanEndPos
          *SpanEndPos\i = (*ScanChar - *BufferStart) / SizeOf(Character)
        EndIf
        ProcedureReturn #CodeCtx_Quote
      EndIf
    Else
      *ScanChar + SizeOf(Character)
    EndIf
  Wend
  ;
  ProcedureReturn #CodeCtx_Code
EndProcedure
;
Procedure FIPC_SplitKeywords(KeywordList$, *Cache.FIPC_KeywordCache, *KeywordBuf.String, ReverseKeywords = #False)
  ;
  ; Role: Splits up to 31 comma-separated keywords in place inside '*KeywordBuf', unescaping
  ;       '\comma_' sequences without per-item heap allocations and optionally reversing each keyword.
  ; Output: Returns the number of valid non-empty keywords stored at indices 1..NbKeywords in '*Cache'.
  ;
  Protected *Read.Character, *Write.Character, *Start.Character, *Left.Character, *Right.Character
  Protected NbKeywords = 0, KeyLen.i, SwapChar.c
  ; Pointers for single-pass in-place unescaping, null-termination, and optional character reversal.
  ;
  *KeywordBuf\s = KeywordList$
  *Read         = @*KeywordBuf\s
  If *Read = 0
    ProcedureReturn 0
  EndIf
  *Write = *Read
  ;
  While *Read\c <> 0 And NbKeywords < 31
    If *Read\c = ','
      Break
    EndIf
    NbKeywords + 1
    *Start                    = *Write
    *Cache\KeyPtr[NbKeywords] = *Start
    *Cache\KeyPos[NbKeywords] = 0
    ;
    While *Read\c <> 0 And *Read\c <> ','
      If *Read\c = '\' And
         PeekC(*Read + 1 * SizeOf(Character)) = 'c' And
         PeekC(*Read + 2 * SizeOf(Character)) = 'o' And
         PeekC(*Read + 3 * SizeOf(Character)) = 'm' And
         PeekC(*Read + 4 * SizeOf(Character)) = 'm' And
         PeekC(*Read + 5 * SizeOf(Character)) = 'a' And
         PeekC(*Read + 6 * SizeOf(Character)) = '_'
        *Write\c = ','
        *Write + SizeOf(Character)
        *Read  + 7 * SizeOf(Character)
      Else
        *Write\c = *Read\c
        *Write + SizeOf(Character)
        *Read  + SizeOf(Character)
      EndIf
    Wend
    ;
    KeyLen                    = (*Write - *Start) / SizeOf(Character)
    *Cache\KeyLen[NbKeywords] = KeyLen
    If ReverseKeywords And KeyLen > 1
      *Left  = *Start
      *Right = *Write - SizeOf(Character)
      While *Left < *Right
        SwapChar = *Left\c
        *Left\c  = *Right\c
        *Right\c = SwapChar
        *Left  + SizeOf(Character)
        *Right - SizeOf(Character)
      Wend
    EndIf
    ;
    If *Read\c = ','
      *Read + SizeOf(Character)
      *Write\c = 0
      *Write + SizeOf(Character)
    Else
      *Write\c = 0
      Break
    EndIf
  Wend
  ;
  ProcedureReturn NbKeywords
EndProcedure
;
Procedure FindInPBPointedCode(*SearchString.String, KeywordList$, posdep = 0, CheckMode = #CheckAll, FMode = #PB_String_CaseSensitive, EndLimit = 0, *MatchedIndex.Integer = 0, *SearchState = 0)
  ;
  ; Role: Searches forward for one or more comma-separated keywords in a PureBasic code buffer,
  ;       resolving comment and quote context in a single backward pointer pass via quote-parity invariants
  ;       and delegating to CheckCodeLineContext() only when "'" or "\" appears on the candidate line.
  ; Parameters:
  ;   *SearchString - Pointer to the source code String structure.
  ;   KeywordList$  - Single keyword or comma-separated list of keywords to search.
  ;   posdep        - 1-based starting character position.
  ;   CheckMode     - Bitmask of boundary and context checks (#CheckAll, #CheckOneKeyWordOnly, etc.).
  ;   FMode         - #PB_String_CaseSensitive or #PB_String_NoCase.
  ;   EndLimit      - 1-based upper character limit (0 to scan up to the end of the buffer).
  ;   *MatchedIndex - Optional pointer receiving the 1-based index of the matched keyword in KeywordList$.
  ;   *SearchState  - Optional pointer to a caller-owned FIPC_KeywordCache preserving keyword splits,
  ;                   cached match offsets, and validated slice bounds across monotonic forward loops.
  ; Output: Returns the 1-based character position of the earliest valid match, or 0 if not found.
  ;
  Protected LocalCache.FIPC_KeywordCache
  Protected *State.FIPC_KeywordCache
  Protected *KeyView.String, *SliceView.String
  Protected *MatchPtr.Character, *ScanPtr.Character
  Protected *StartAdress, *EndAdress
  Protected KeyAddr.i, SliceAddr.i, SimpleSearch, IsSliced = #False, OneKeyFlag
  Protected FirstKeywordPos, FirstKeywordLength, FirstKeywordIndex, NbKeywords, nl
  Protected p, SpanEndPos, LineCtx, QuotesCount, SemiColonOdd, SemiColonEven, AmbiguousQuote
  Protected IsComment, IsInQuote, OrigChar.c, CharCode.c
  ; Stack or caller-provided keyword cache, zero-copy string views, and backward quote/comment parity state.
  ;
  If *SearchString = 0
    Debug "Empty pointer"
    CallDebugger
    ProcedureReturn 0
  EndIf
  ;
  *StartAdress = @*SearchString\s
  If *StartAdress = 0 Or PeekC(*StartAdress) = 0 Or KeywordList$ = ""
    ProcedureReturn 0
  EndIf
  ;
  If posdep < 1
    posdep = 1
  EndIf
  ;
  *KeyView   = @KeyAddr
  *SliceView = @SliceAddr
  OneKeyFlag = CheckMode & #CheckOneKeyWordOnly
  ;
  If *SearchState
    *State = *SearchState
    If *State\BufferAddr <> *StartAdress Or *State\EndLimit <> EndLimit Or
       *State\FMode <> FMode Or *State\OneKeyFlag <> OneKeyFlag Or *State\CachedKeywords$ <> KeywordList$
      ; Initialize or refresh persistent search state for the new buffer or query parameters
      *State\BufferAddr      = *StartAdress
      *State\SafeEndPos      = 1
      *State\EndLimit        = EndLimit
      *State\FMode           = FMode
      *State\OneKeyFlag      = OneKeyFlag
      *State\LastMatchPos    = 0
      *State\CachedKeywords$ = KeywordList$
      If OneKeyFlag
        *State\KeywordBuf\s = KeywordList$
        *State\KeyPtr[1]    = @*State\KeywordBuf\s
        *State\KeyLen[1]    = Len(KeywordList$)
        *State\NbKeywords   = 1
        *State\SimpleSearch = #True
      ElseIf FindString(KeywordList$, ",") = 0
        If FindString(KeywordList$, "\comma_")
          *State\KeywordBuf\s = ReplaceString(KeywordList$, "\comma_", ",")
        Else
          *State\KeywordBuf\s = KeywordList$
        EndIf
        *State\KeyPtr[1]    = @*State\KeywordBuf\s
        *State\KeyLen[1]    = Len(*State\KeywordBuf\s)
        *State\NbKeywords   = 1
        *State\SimpleSearch = #True
      Else
        *State\SimpleSearch = #False
        *State\NbKeywords   = FIPC_SplitKeywords(KeywordList$, *State, @*State\KeywordBuf, #False)
      EndIf
    ElseIf *State\LastMatchPos < 0 Or posdep <= *State\LastMatchPos
      ; Caller restarted or moved backward on the same buffer: clear cached offsets while keeping parsed keywords
      *State\SafeEndPos   = 1
      *State\LastMatchPos = 0
      For nl = 1 To *State\NbKeywords
        *State\KeyPos[nl] = 0
      Next
    EndIf
    ;
    NbKeywords   = *State\NbKeywords
    SimpleSearch = *State\SimpleSearch
    If NbKeywords = 0
      *State\LastMatchPos = -1
      ProcedureReturn 0
    EndIf
    If SimpleSearch
      KeyAddr            = *State\KeyPtr[1]
      FirstKeywordLength = *State\KeyLen[1]
      FirstKeywordIndex  = 1
    EndIf
  Else
    *State = @LocalCache
    If OneKeyFlag
      KeyAddr            = @KeywordList$
      FirstKeywordLength = Len(KeywordList$)
      FirstKeywordIndex  = 1
      SimpleSearch       = #True
    ElseIf FindString(KeywordList$, ",") = 0
      If FindString(KeywordList$, "\comma_")
        LocalCache\KeywordBuf\s = ReplaceString(KeywordList$, "\comma_", ",")
        KeyAddr                 = @LocalCache\KeywordBuf\s
        FirstKeywordLength      = Len(LocalCache\KeywordBuf\s)
      Else
        KeyAddr            = @KeywordList$
        FirstKeywordLength = Len(KeywordList$)
      EndIf
      FirstKeywordIndex = 1
      SimpleSearch      = #True
    Else
      SimpleSearch = #False
      NbKeywords   = FIPC_SplitKeywords(KeywordList$, @LocalCache, @LocalCache\KeywordBuf, #False)
      If NbKeywords = 0
        ProcedureReturn 0
      EndIf
    EndIf
  EndIf
  ;
  If EndLimit > 0
    *EndAdress = *StartAdress + (EndLimit * SizeOf(Character))
    OrigChar   = PeekC(*EndAdress)
    ; Temporarily null-terminate the buffer at EndLimit so native FindString() never scans past the limit
    If OrigChar
      PokeC(*EndAdress, 0)
    EndIf
  EndIf
  ;
  If *SearchState
    ; Ensure 'posdep' lies within the null-terminated buffer so O(1) pointer slicing is active immediately
    If posdep > *State\SafeEndPos
      *ScanPtr  = *StartAdress + (*State\SafeEndPos - 1) * SizeOf(Character)
      *MatchPtr = *StartAdress + (posdep - 1) * SizeOf(Character)
      While *ScanPtr < *MatchPtr And *ScanPtr\c <> 0
        *ScanPtr + SizeOf(Character)
      Wend
      If *ScanPtr < *MatchPtr
        If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
        *State\LastMatchPos = -1
        ProcedureReturn 0
      EndIf
      *State\SafeEndPos = posdep
    EndIf
    IsSliced = #True
  EndIf
  ;
  SearchFurtherLabel:
  ;
  If EndLimit > 0 And posdep > EndLimit
    If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
    If *SearchState : *State\LastMatchPos = -1 : EndIf
    ProcedureReturn 0
  EndIf
  ;
  If IsSliced
    SliceAddr = *StartAdress + (posdep - 1) * SizeOf(Character)
    If PeekC(SliceAddr) = 0
      If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
      If *SearchState : *State\LastMatchPos = -1 : EndIf
      ProcedureReturn 0
    EndIf
  EndIf
  ;
  If SimpleSearch
    If IsSliced
      p = FindString(*SliceView\s, *KeyView\s, 1, FMode)
      If p > 0 : p + posdep - 1 : EndIf
    Else
      p = FindString(*SearchString\s, *KeyView\s, posdep, FMode)
    EndIf
    If p = 0 Or (EndLimit > 0 And (p + FirstKeywordLength - 1) > EndLimit)
      If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
      If *SearchState : *State\LastMatchPos = -1 : EndIf
      ProcedureReturn 0
    EndIf
    If p + FirstKeywordLength > *State\SafeEndPos
      *State\SafeEndPos = p + FirstKeywordLength
    EndIf
  Else
    FirstKeywordPos    = 0
    FirstKeywordLength = 0
    FirstKeywordIndex  = 0
    For nl = 1 To NbKeywords
      p = *State\KeyPos[nl]
      ; Only re-run FindString() for keywords whose cached position precedes the new posdep
      If p >= 0 And p < posdep
        KeyAddr = *State\KeyPtr[nl]
        If IsSliced
          p = FindString(*SliceView\s, *KeyView\s, 1, FMode)
          If p > 0 : p + posdep - 1 : EndIf
        Else
          p = FindString(*SearchString\s, *KeyView\s, posdep, FMode)
        EndIf
        If p = 0
          *State\KeyPos[nl] = -1
        Else
          *State\KeyPos[nl] = p
          If p + *State\KeyLen[nl] > *State\SafeEndPos
            *State\SafeEndPos = p + *State\KeyLen[nl]
          EndIf
        EndIf
      EndIf
      ; Select the earliest match, or the longest keyword when two keywords start at the exact same position
      If p > 0
        If FirstKeywordPos = 0 Or p < FirstKeywordPos Or (p = FirstKeywordPos And *State\KeyLen[nl] > FirstKeywordLength)
          FirstKeywordPos    = p
          FirstKeywordLength = *State\KeyLen[nl]
          FirstKeywordIndex  = nl
        EndIf
      EndIf
    Next
    ;
    p = FirstKeywordPos
    If p = 0 Or (EndLimit > 0 And (p + FirstKeywordLength - 1) > EndLimit)
      If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
      If *SearchState : *State\LastMatchPos = -1 : EndIf
      ProcedureReturn 0
    EndIf
  EndIf
  ;
  ; 'p + FirstKeywordLength' is now proven inside the buffer, enabling O(1) pointer slicing on all retries
  IsSliced  = #True
  *MatchPtr = *StartAdress + (p - 1) * SizeOf(Character)
  ;
  ; Verify that the matched word is surrounded by valid PureBasic separators (IsPBSeparator(0) = #True covers EOF/EndLimit)
  If (CheckMode & #CheckLeft And p > 1 And
      IsPBSeparator(PeekC(*MatchPtr - SizeOf(Character))) = #False And
      IsPBSeparator(*MatchPtr\c) = #False) Or
     (CheckMode & #CheckRight And
      IsPBSeparator(PeekC(*MatchPtr + FirstKeywordLength * SizeOf(Character))) = #False And
      IsPBSeparator(PeekC(*MatchPtr + (FirstKeywordLength - 1) * SizeOf(Character))) = #False)
    posdep = p + FirstKeywordLength
    Goto SearchFurtherLabel
  EndIf
  ;
  If (CheckMode & (#CheckCom | #CheckQuote)) And p > 1
    ; Single-pass backward parity check: Q_before(i) is even iff Q_after(i) has the same parity as Q_total
    *ScanPtr       = *MatchPtr
    QuotesCount    = 0
    SemiColonOdd   = #False
    SemiColonEven  = #False
    AmbiguousQuote = #False
    ;
    While *ScanPtr > *StartAdress
      *ScanPtr - SizeOf(Character)
      CharCode = *ScanPtr\c
      If CharCode = #CR Or CharCode = #LF
        Break
      ElseIf CharCode = #DOUBLEQUOTE
        QuotesCount + 1
      ElseIf CharCode = ';'
        If QuotesCount & 1
          SemiColonOdd = #True
        Else
          SemiColonEven = #True
        EndIf
      ElseIf CharCode = #SIMPLEQUOTE Or CharCode = '\'
        AmbiguousQuote = #True
      EndIf
    Wend
    ;
    If AmbiguousQuote
      ; "'" or "\" breaks pure backward double-quote parity: resolve exact context via CheckCodeLineContext()
      LineCtx = CheckCodeLineContext(*SearchString, p, #True, #True, @SpanEndPos)
      If LineCtx = #CodeCtx_Comment And (CheckMode & #CheckCom)
        If PeekC(*StartAdress + SpanEndPos * SizeOf(Character)) = 0
          If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
          If *SearchState : *State\LastMatchPos = -1 : EndIf
          ProcedureReturn 0
        EndIf
        posdep = SpanEndPos + 2
        If posdep > *State\SafeEndPos : *State\SafeEndPos = posdep : EndIf
        Goto SearchFurtherLabel
      ElseIf LineCtx = #CodeCtx_Quote And (CheckMode & #CheckQuote)
        posdep = SpanEndPos + 1
        If posdep > *State\SafeEndPos : *State\SafeEndPos = posdep : EndIf
        If posdep < p + FirstKeywordLength
          posdep   = p + FirstKeywordLength
          IsSliced = Bool(posdep <= *State\SafeEndPos)
        EndIf
        Goto SearchFurtherLabel
      EndIf
    ElseIf QuotesCount Or SemiColonEven
      If QuotesCount & 1
        IsComment = SemiColonOdd
        IsInQuote = Bool(Not SemiColonOdd)
      Else
        IsComment = SemiColonEven
        IsInQuote = #False
      EndIf
      ;
      If IsComment And (CheckMode & #CheckCom)
        *ScanPtr = *MatchPtr + FirstKeywordLength * SizeOf(Character)
        While *ScanPtr\c <> 0 And *ScanPtr\c <> #CR And *ScanPtr\c <> #LF
          *ScanPtr + SizeOf(Character)
        Wend
        If *ScanPtr\c = 0
          If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
          If *SearchState : *State\LastMatchPos = -1 : EndIf
          ProcedureReturn 0
        EndIf
        posdep = ((*ScanPtr - *StartAdress) / SizeOf(Character)) + 2
        If posdep > *State\SafeEndPos : *State\SafeEndPos = posdep : EndIf
        Goto SearchFurtherLabel
      EndIf
      ;
      If IsInQuote And (CheckMode & #CheckQuote)
        posdep = p + FirstKeywordLength
        Goto SearchFurtherLabel
      EndIf
    EndIf
  EndIf
  ;
  If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
  If *SearchState : *State\LastMatchPos = p : EndIf
  If *MatchedIndex : *MatchedIndex\i = FirstKeywordIndex : EndIf
  ProcedureReturn p
EndProcedure
;
Procedure ReverseFindInPBCode(txt$, KeywordList$, posdep = 0, CheckMode = #CheckAll, FMode = #PB_String_CaseSensitive)
  ;
  ; Role: Searches backward (upstream of 'posdep') for one or more comma-separated keywords
  ;       in a PureBasic code page, ensuring matches are properly bounded and outside comments or quotes.
  ; Parameters:
  ;   txt$         - Source buffer to cache in reversed form (pass "" on subsequent calls to reuse cache).
  ;   KeywordList$ - Comma-separated list of keywords (use "\comma_" for a literal comma, or #CheckOneKeyWordOnly).
  ;   posdep       - 1-based character position from which to search backward.
  ;   CheckMode    - Bitmask of separator and comment/quote checks (#CheckAll, #CheckLeft, #CheckRight, etc.).
  ; Output: Returns the 1-based position of the closest upstream match, or 0 if not found.
  ;
  Static SearchString$, SStringLength
  Protected Cache.FIPC_KeywordCache
  Protected KeywordBuf.String
  Protected *KeyView.String, *SliceView.String
  Protected *RevStart, *BufEnd, *AfterLeftPtr.Character, *LineStartRev.Character
  Protected *ScanPtr.Character, *OpenQuoteRev.Character
  Protected KeyAddr.i, SliceAddr.i, SimpleSearch
  Protected FirstKeywordPos, FirstKeywordLength, NbKeywords, nl, p
  Protected HasSpecial, InLiteral, QuoteChar.c
  ; Cache the reversed buffer across calls and inspect line prefixes via zero-allocation pointer scans.
  ;
  If txt$
    SearchString$ = ReverseString(txt$)
    SStringLength = Len(SearchString$)
  EndIf
  ;
  If SStringLength = 0 Or KeywordList$ = ""
    ProcedureReturn 0
  EndIf
  ;
  If posdep
    posdep = SStringLength - posdep + 1
  EndIf
  If posdep < 1
    posdep = 1
  ElseIf posdep > SStringLength
    ProcedureReturn 0
  EndIf
  ;
  *RevStart  = @SearchString$
  *BufEnd    = *RevStart + SStringLength * SizeOf(Character)
  *KeyView   = @KeyAddr
  *SliceView = @SliceAddr
  ;
  If CheckMode & #CheckOneKeyWordOnly
    KeywordBuf\s       = ReverseString(KeywordList$)
    KeyAddr            = @KeywordBuf\s
    FirstKeywordLength = Len(KeywordBuf\s)
    SimpleSearch       = #True
  Else
    NbKeywords = FIPC_SplitKeywords(KeywordList$, @Cache, @KeywordBuf, #True)
    If NbKeywords = 0
      ProcedureReturn 0
    ElseIf NbKeywords = 1
      KeyAddr            = Cache\KeyPtr[1]
      FirstKeywordLength = Cache\KeyLen[1]
      SimpleSearch       = #True
    Else
      SimpleSearch = #False
    EndIf
  EndIf
  ;
  ReverseSearchFurtherLabel:
  ;
  If posdep > SStringLength
    ProcedureReturn 0
  EndIf
  SliceAddr = *RevStart + (posdep - 1) * SizeOf(Character)
  ;
  If SimpleSearch
    FirstKeywordPos = FindString(*SliceView\s, *KeyView\s, 1, FMode)
    If FirstKeywordPos > 0
      FirstKeywordPos + posdep - 1
    EndIf
  Else
    FirstKeywordPos    = 0
    FirstKeywordLength = 0
    For nl = 1 To NbKeywords
      p = Cache\KeyPos[nl]
      If p >= 0 And p < posdep
        KeyAddr = Cache\KeyPtr[nl]
        p       = FindString(*SliceView\s, *KeyView\s, 1, FMode)
        If p = 0
          Cache\KeyPos[nl] = -1
        Else
          p + posdep - 1
          Cache\KeyPos[nl] = p
        EndIf
      EndIf
      ; In the reversed buffer, the smallest 'p' is the closest upstream match in normal order;
      ; when two keywords share the same right boundary 'p', the longer keyword wins
      If p > 0
        If FirstKeywordPos = 0 Or p < FirstKeywordPos Or (p = FirstKeywordPos And Cache\KeyLen[nl] > FirstKeywordLength)
          FirstKeywordPos    = p
          FirstKeywordLength = Cache\KeyLen[nl]
        EndIf
      EndIf
    Next
  EndIf
  ;
  If FirstKeywordPos = 0
    ProcedureReturn 0
  EndIf
  ;
  p = FirstKeywordPos + FirstKeywordLength
  ;
  ; In the reversed buffer, 1-based index 'p' is the left neighbor in normal order (#CheckLeft),
  ; while 'FirstKeywordPos - 1' is the right neighbor in normal order (#CheckRight)
  If (CheckMode & #CheckLeft And p <= SStringLength And
      IsPBSeparator(PeekC(*RevStart + (p - 1) * SizeOf(Character))) = #False And
      IsPBSeparator(PeekC(*RevStart + (p - 2) * SizeOf(Character))) = #False) Or
     (CheckMode & #CheckRight And FirstKeywordPos > 1 And
      IsPBSeparator(PeekC(*RevStart + (FirstKeywordPos - 2) * SizeOf(Character))) = #False And
      IsPBSeparator(PeekC(*RevStart + (FirstKeywordPos - 1) * SizeOf(Character))) = #False)
    posdep = FirstKeywordPos + 1
    Goto ReverseSearchFurtherLabel
  EndIf
  ;
  If (CheckMode & (#CheckCom | #CheckQuote)) And p <= SStringLength
    ; Scan upward in reversed memory to locate #CR/#LF (line start in normal order) and detect ';', '"', or "'"
    *AfterLeftPtr = *RevStart + (p - 1) * SizeOf(Character)
    *LineStartRev = *AfterLeftPtr
    HasSpecial    = #False
    While *LineStartRev < *BufEnd And *LineStartRev\c <> #CR And *LineStartRev\c <> #LF
      If *LineStartRev\c = ';' Or *LineStartRev\c = #DOUBLEQUOTE Or *LineStartRev\c = #SIMPLEQUOTE
        HasSpecial = #True
      EndIf
      *LineStartRev + SizeOf(Character)
    Wend
    ;
    If HasSpecial And *LineStartRev > *AfterLeftPtr
      ; Decrementing from (*LineStartRev - 2) down to *AfterLeftPtr reads the line left-to-right in normal order
      *ScanPtr = *LineStartRev - SizeOf(Character)
      While *ScanPtr >= *AfterLeftPtr
        If *ScanPtr\c = ';'
          If CheckMode & #CheckCom
            posdep = ((*ScanPtr - *RevStart) / SizeOf(Character)) + 2
            Goto ReverseSearchFurtherLabel
          EndIf
          Break
        ElseIf *ScanPtr\c = '~' And *ScanPtr > *AfterLeftPtr And PeekC(*ScanPtr - SizeOf(Character)) = #DOUBLEQUOTE
          *OpenQuoteRev = *ScanPtr
          *ScanPtr - 2 * SizeOf(Character)
          InLiteral = #True
          While *ScanPtr >= *AfterLeftPtr
            If *ScanPtr\c = '\'
              *ScanPtr - SizeOf(Character)
              If *ScanPtr < *AfterLeftPtr
                Break
              EndIf
              *ScanPtr - SizeOf(Character)
            ElseIf *ScanPtr\c = #DOUBLEQUOTE
              InLiteral = #False
              *ScanPtr - SizeOf(Character)
              Break
            Else
              *ScanPtr - SizeOf(Character)
            EndIf
          Wend
          If InLiteral
            If CheckMode & #CheckQuote
              posdep = ((*OpenQuoteRev - *RevStart) / SizeOf(Character)) + 2
              Goto ReverseSearchFurtherLabel
            EndIf
            Break
          EndIf
        ElseIf *ScanPtr\c = #DOUBLEQUOTE Or *ScanPtr\c = #SIMPLEQUOTE
          QuoteChar     = *ScanPtr\c
          *OpenQuoteRev = *ScanPtr
          *ScanPtr - SizeOf(Character)
          InLiteral = #True
          While *ScanPtr >= *AfterLeftPtr
            If *ScanPtr\c = QuoteChar
              InLiteral = #False
              *ScanPtr - SizeOf(Character)
              Break
            EndIf
            *ScanPtr - SizeOf(Character)
          Wend
          If InLiteral
            If CheckMode & #CheckQuote
              posdep = ((*OpenQuoteRev - *RevStart) / SizeOf(Character)) + 2
              Goto ReverseSearchFurtherLabel
            EndIf
            Break
          EndIf
        Else
          *ScanPtr - SizeOf(Character)
        EndIf
      Wend
    EndIf
  EndIf
  ;
  ProcedureReturn SStringLength - p + 2
EndProcedure
;
Procedure FindInPBStringCode(SearchString$, KeywordList$, posdep = 0, CheckMode = #CheckAll, FMode = #PB_String_CaseSensitive)
  ;
  ; Role: Wrapper around FindInPBPointedCode() operating on a standard string value.
  ;
  Protected SearchString.String\s = SearchString$
  ; Wraps the string value into a String structure for FindInPBPointedCode().
  ;
  ProcedureReturn FindInPBPointedCode(SearchString, KeywordList$, posdep, CheckMode, FMode)
EndProcedure
;
Procedure FindReplaceInPBPointedCode(*SearchString.String, KeywordList$, ReplaceKeywordList$, posdep = 0, CheckMode = #CheckAll, FMode = #PB_String_CaseSensitive)
  ;
  ; Role: Searches forward for one or more keywords in a PureBasic code buffer and replaces
  ;       valid matches in a single pass, preventing cascading replacements across multiple keywords.
  ; Parameters:
  ;   *SearchString       - Pointer to the source code String structure to modify in place.
  ;   KeywordList$        - Single keyword or comma-separated list of keywords (use "\comma_" for a literal comma).
  ;   ReplaceKeywordList$ - Replacement word(s): a comma-separated list matching KeywordList$ entries,
  ;                         a single word applied to all keywords, or "" to erase matched keywords.
  ;   posdep              - 1-based starting character position.
  ;   CheckMode           - Bitmask of boundary and context checks (#CheckAll, #CheckOneKeyWordOnly, etc.).
  ;   FMode               - #PB_String_CaseSensitive or #PB_String_NoCase.
  ; Output: Returns the total number of replacements performed.
  ;
  Protected SearchState.FIPC_KeywordCache
  Protected Rep = 0, NbKeywords = 0, MatchedIndex = 0, HasMultipleReplacements
  Protected OldMatchEnd, RepLen, Delta, nl
  Protected Keyword$, DefaultRep$
  Protected Dim KeyLen(31)
  Protected Dim RepStr$(31)
  ; Persistent search cache and pre-computed keyword/replacement tables for single-pass substitution.
  ;
  If *SearchString = 0
    Debug "Empty pointer"
    CallDebugger
    ProcedureReturn 0
  EndIf
  ;
  If KeywordList$ = ""
    ProcedureReturn 0
  EndIf
  ;
  ; Pre-compute keyword lengths and replacement strings once before entering the replacement loop
  If CheckMode & #CheckOneKeyWordOnly
    KeyLen(1)  = Len(KeywordList$)
    RepStr$(1) = ReplaceKeywordList$
  ElseIf FindString(KeywordList$, ",") = 0
    KeyLen(1)  = Len(ReplaceString(KeywordList$, "\comma_", ","))
    RepStr$(1) = ReplaceString(ReplaceKeywordList$, "\comma_", ",")
  Else
    HasMultipleReplacements = Bool(FindString(ReplaceKeywordList$, ",") > 0)
    DefaultRep$             = ReplaceString(ReplaceKeywordList$, "\comma_", ",")
    Repeat
      NbKeywords + 1
      Keyword$ = ReplaceString(StringField(KeywordList$, NbKeywords, ","), "\comma_", ",")
      KeyLen(NbKeywords) = Len(Keyword$)
      If HasMultipleReplacements
        RepStr$(NbKeywords) = ReplaceString(StringField(ReplaceKeywordList$, NbKeywords, ","), "\comma_", ",")
      Else
        RepStr$(NbKeywords) = DefaultRep$
      EndIf
    Until Keyword$ = "" Or NbKeywords = 31
  EndIf
  ;
  ; Execute all replacements in a single forward pass while shifting cached match offsets in place
  Repeat
    posdep = FindInPBPointedCode(*SearchString, KeywordList$, posdep, CheckMode, FMode, 0, @MatchedIndex, @SearchState)
    If posdep And MatchedIndex > 0
      OldMatchEnd     = posdep + KeyLen(MatchedIndex)
      RepLen          = Len(RepStr$(MatchedIndex))
      Delta           = RepLen - KeyLen(MatchedIndex)
      *SearchString\s = FastLeft(*SearchString\s, posdep - 1) + RepStr$(MatchedIndex) + FastMid(*SearchString\s, OldMatchEnd)
      Rep + 1
      posdep + RepLen
      ;
      SearchState\BufferAddr   = @*SearchString\s
      SearchState\LastMatchPos = posdep - 1
      If SearchState\BufferAddr And PeekC(SearchState\BufferAddr) <> 0
        SearchState\SafeEndPos = posdep
        For nl = 1 To SearchState\NbKeywords
          If SearchState\KeyPos[nl] >= OldMatchEnd
            SearchState\KeyPos[nl] + Delta
          ElseIf SearchState\KeyPos[nl] > 0
            SearchState\KeyPos[nl] = 0
          EndIf
        Next
      EndIf
    EndIf
  Until posdep = 0
  ;
  ProcedureReturn Rep
EndProcedure
;
Procedure.s FindReplaceInPBStringCode(SearchString$, SearchKeyword$, ReplaceKeywordList$, posdep = 0, CheckMode = #CheckAll, FMode = #PB_String_CaseSensitive)
  ;
  ; Role: Wrapper around FindReplaceInPBPointedCode() operating on a standard string value.
  ;
  Protected SearchString.String\s = SearchString$
  ; Wraps the string value into a String structure for in-place replacement.
  ;
  FindReplaceInPBPointedCode(SearchString, SearchKeyword$, ReplaceKeywordList$, posdep, CheckMode, FMode)
  ;
  ProcedureReturn SearchString\s
EndProcedure
;
; ***************************************************************************************
;                   BENCHMARK & VALIDATION SUITE (STANDALONE EXECUTION)
; ***************************************************************************************
;
CompilerIf #PB_Compiler_IsMainFile
  ;
  Global Dim StageTimer.q(40)
  Global Dim StageTimerRef.q(40)
  Global Dim TimerLegend$(40)
  Global NbrOfUsedTimers.i
  ;
  Procedure FindInPBPointedCodeOld(*SearchString.String, KeywordList$, posdep = 0, CheckMode = #CheckAll, FMode = #PB_String_CaseSensitive, EndLimit = 0, *MatchedIndex.Integer = 0)
    ;
    ; Role: Historical baseline version of FindInPBPointedCode(), kept strictly unmodified
    ;       inside this standalone block for regression verification and speed benchmarking
    ;       whenever new optimizations are evaluated on FindInPBPointedCode().
    ; Known limitations compared to the current FindInPBPointedCode():
    ;   - Does not handle PureBasic 6.4 escaped string literals (~"..." containing \" or \\).
    ;   - Does not handle character literals containing quotes or semicolons ('"' or ';').
    ;   - Does not support persistent '*SearchState' caching across monotonic search loops.
    ;   - Allocates dynamic arrays and temporary substrings on the heap, and under-reads 2 bytes at p = 1.
    ;
    Protected SimpleSearch, FirstKeyword$, FirstKeywordPos, FirstKeywordIndex, nl, NbKeywords, *StartAdress, *EndAdress
    Protected posdepR, p, ppv, nbg, TestS$, FirstKeywordLength, OrigChar.l
    Protected Dim KeyPos.i(31)
    Protected Dim KeyStr.s(31)
    Protected Dim KeyLen.i(31)
    ; Hold the parsed keywords, their cached match positions, and their respective character lengths.
    ;
    If *SearchString = 0
      Debug "Empty pointer"
      CallDebugger
      ProcedureReturn 0
    EndIf
    ;
    If KeywordList$ = ""
      ProcedureReturn 0
    EndIf
    ;
    If posdep < 1 : posdep = 1 : EndIf
    ;
    If CheckMode & #CheckOneKeyWordOnly
      FirstKeyword$      = KeywordList$
      FirstKeywordLength = Len(FirstKeyword$)
      FirstKeywordIndex  = 1
      SimpleSearch       = 1
    ElseIf FindString(KeywordList$, ",") = 0
      FirstKeyword$      = ReplaceString(KeywordList$, "\comma_", ",")
      FirstKeywordLength = Len(FirstKeyword$)
      FirstKeywordIndex  = 1
      SimpleSearch       = 1
    Else
      SimpleSearch = 0
      ; Pre-extract and unescape all keywords once before entering the search loop
      Repeat
        NbKeywords + 1
        KeyStr(NbKeywords) = ReplaceString(StringField(KeywordList$, NbKeywords, ","), "\comma_", ",")
        KeyLen(NbKeywords) = Len(KeyStr(NbKeywords))
      Until KeyStr(NbKeywords) = "" Or NbKeywords = 31
      If KeyStr(NbKeywords) = "" : NbKeywords - 1 : EndIf
    EndIf
    ;
    *StartAdress = @*SearchString\s
    *EndAdress   = *StartAdress + (EndLimit * SizeOf(CHARACTER))
    ;
    If EndLimit = 0
      EndLimit = Len(*SearchString\s)
    Else
      OrigChar = PeekC(*EndAdress)
      ; Temporarily null-terminate the buffer at EndLimit so native FindString() never scans past the limit
      PokeC(*EndAdress, 0)
    EndIf
    ;
    SearchFurtherLabelOld:
    ;
    If SimpleSearch
      If posdep > EndLimit
        If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
        ProcedureReturn 0
      Else
        p = FindString(*SearchString\s, FirstKeyword$, posdep, FMode)
        FirstKeywordPos = p
        If p = 0 Or (FirstKeywordPos + FirstKeywordLength - 1) > EndLimit
          If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
          ProcedureReturn 0
        EndIf
      EndIf
    Else
      FirstKeyword$      = ""
      FirstKeywordPos    = 0
      FirstKeywordLength = 0
      FirstKeywordIndex  = 0
      For nl = 1 To NbKeywords
        p = KeyPos(nl)
        ; Only re-run FindString() for keywords whose cached position precedes the new posdep
        If p >= 0 And p < posdep
          p = FindString(*SearchString\s, KeyStr(nl), posdep, FMode)
          If p = 0
            KeyPos(nl) = -1
          Else
            KeyPos(nl) = p
          EndIf
        EndIf
        ; Select the earliest match, or the longest keyword when two keywords start at the exact same position
        If p > 0
          If FirstKeywordPos = 0 Or p < FirstKeywordPos Or (p = FirstKeywordPos And KeyLen(nl) > FirstKeywordLength)
            FirstKeywordPos    = p
            FirstKeyword$      = KeyStr(nl)
            FirstKeywordLength = KeyLen(nl)
            FirstKeywordIndex  = nl
          EndIf
        EndIf
      Next
      ;
      p = FirstKeywordPos
      ;
      If p = 0 Or (FirstKeywordPos + FirstKeywordLength - 1) > EndLimit
        If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
        ProcedureReturn 0
      EndIf
    EndIf
    ;
    ; Verify that the matched word is surrounded by valid PureBasic separators
    If (CheckMode & #CheckLeft And p > 1 And
        IsPBSeparator(Asc(FastMid(*SearchString\s, p - 1, 1))) = #False And
        IsPBSeparator(Asc(FastMid(*SearchString\s, p, 1))) = #False) Or
       (CheckMode & #CheckRight And (p + FirstKeywordLength - 1) < EndLimit And
        IsPBSeparator(Asc(FastMid(*SearchString\s, p + FirstKeywordLength, 1))) = #False And
        IsPBSeparator(Asc(FastMid(*SearchString\s, p + FirstKeywordLength - 1, 1))) = #False)
      posdep = p + FirstKeywordLength
      Goto SearchFurtherLabelOld
    EndIf
    ;
    If CheckMode & #CheckCom Or CheckMode & #CheckQuote
      ; Inspect the line prefix preceding the match to check for comments or open string quotes
      posdepR = p
      While posdepR > 0 And PeekC(@*SearchString\s + (posdepR - 1) * SizeOf(CHARACTER)) <> #CR
        posdepR - 1
      Wend
      posdepR + 1
      ;
      If p - posdepR > 0
        TestS$ = Trim(FastMid(*SearchString\s, posdepR, p - posdepR))
        If TestS$
          ;
          If CheckMode & #CheckCom
            ppv = 0
            Repeat
              ppv = FindString(TestS$, ";", ppv + 1)
              If ppv
                nbg = CountString(Mid(TestS$, 1, ppv), #DOUBLEQUOTE$)
                If Not(nbg & 1)
                  posdep = FindString(*SearchString\s, #CR$, p)
                  If posdep = 0
                    If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
                    ProcedureReturn 0
                  EndIf
                  Goto SearchFurtherLabelOld
                EndIf
              EndIf
            Until ppv = 0
          EndIf
          ;
          If CheckMode & #CheckQuote
            nbg = CountString(TestS$, #DOUBLEQUOTE$)
            If nbg & 1
              posdep = p + FirstKeywordLength
              Goto SearchFurtherLabelOld
            EndIf
          EndIf
        EndIf
      EndIf
    EndIf
    ;
    If OrigChar : PokeC(*EndAdress, OrigChar) : EndIf
    If *MatchedIndex : *MatchedIndex\i = FirstKeywordIndex : EndIf
    ProcedureReturn p
  EndProcedure
  ;
  Procedure.q TicksHQ()
    ;
    ; Role: High-resolution microsecond clock derived from QueryPerformanceCounter_(),
    ;       offset by 1 second from startup so the very first timestamp is strictly non-zero
    ;       while preventing 64-bit multiplication overflow on long OS uptimes.
    ;
    Static STFreq.q, STBaseTick.q
    Protected t.q
    ; Hardware counter frequency, back-dated initial tick anchor, and current counter sample.
    ;
    If STFreq = 0
      QueryPerformanceFrequency_(@STFreq)
      QueryPerformanceCounter_(@STBaseTick)
      STBaseTick - STFreq
    EndIf
    QueryPerformanceCounter_(@t)
    ProcedureReturn ((t - STBaseTick) * 1000000) / STFreq
  EndProcedure
  ;
  Procedure StartTimer(Nbr, Legend$ = "")
    ;
    ; Role: Starts or resumes timing slot 'Nbr' and optionally assigns its display legend.
    ;
    If Nbr > NbrOfUsedTimers : NbrOfUsedTimers = Nbr : EndIf
    If Legend$ : TimerLegend$(Nbr) = Legend$ : EndIf
    StageTimerRef(Nbr) = TicksHQ()
  EndProcedure
  ;
  Procedure ReapTimer(Nbr)
    ;
    ; Role: Accumulates elapsed microseconds into slot 'Nbr' since the last StartTimer/ReapTimer call.
    ;
    Protected Now.q = TicksHQ()
    ; Capture timestamp immediately upon entry so timer bookkeeping is excluded from measured time.
    ;
    If StageTimerRef(Nbr)
      StageTimer(Nbr) + (Now - StageTimerRef(Nbr))
    Else
      StageTimer(Nbr) = 0
    EndIf
    StageTimerRef(Nbr) = TicksHQ()
    If Nbr > NbrOfUsedTimers : NbrOfUsedTimers = Nbr : EndIf
  EndProcedure
  ;
  Procedure InitAllTimers()
    ;
    ; Role: Resets all recorded durations, reference timestamps, and legends, and warms up TicksHQ().
    ;
    Protected ct.i
    ; Loop index across the global timer slots.
    ;
    TicksHQ()
    For ct = 0 To ArraySize(StageTimer())
      StageTimer(ct)    = 0
      StageTimerRef(ct) = 0
      TimerLegend$(ct)  = ""
    Next
    NbrOfUsedTimers = 0
  EndProcedure
  ;
  Procedure.s FormatMicroseconds(Microseconds.q)
    ;
    ; Role: Formats a microsecond duration into "X sec. Y ms Z µs".
    ;
    Protected Seconds.q, Milliseconds.q, RemainingUs.q
    ; Decomposed time units extracted from the raw microsecond count.
    ;
    If Microseconds < 0 : Microseconds = 0 : EndIf
    Seconds      = Microseconds / 1000000
    Microseconds = Microseconds % 1000000
    Milliseconds = Microseconds / 1000
    RemainingUs  = Microseconds % 1000
    ProcedureReturn Str(Seconds) + " sec. " + RSet(Str(Milliseconds), 3, " ") + " ms " + RSet(Str(RemainingUs), 3, " ") + " µs"
  EndProcedure
  ;
  Procedure.s DebugAllTimers()
    ;
    ; Role: Outputs all active timers to the debugger and returns the formatted multi-line report.
    ;
    Protected ct.i, Result$, ConcResult$, TotTime.q
    ; Iteration counter, per-line text, full concatenated report, and cumulative duration.
    ;
    Debug "______________________________________________________________________"
    For ct = 1 To NbrOfUsedTimers
      If StageTimer(ct) > 0 Or TimerLegend$(ct) <> ""
        If TimerLegend$(ct) And Left(TimerLegend$(ct), 2) <> " -"
          TimerLegend$(ct) = " - " + LSet(TimerLegend$(ct), 38, " ") + " - "
        EndIf
        Result$ = "Timer " + RSet(Str(ct), 2, " ") + TimerLegend$(ct) + " : " + FormatMicroseconds(StageTimer(ct))
        Debug Result$
        ConcResult$ + Result$ + #CRLF$
        TotTime + StageTimer(ct)
      EndIf
    Next
    Result$ = "Total benchmark time : " + FormatMicroseconds(TotTime)
    Debug "______________________________________________________________________"
    Debug Result$
    ConcResult$ + "______________________________________________________________________" + #CRLF$ + Result$ + #CRLF$
    ProcedureReturn ConcResult$
  EndProcedure
  ;
  Procedure.s FIPC_BuildSample(BlockCount, IncludeEscapedLiterals)
    ;
    ; Role: Synthesizes a realistic PureBasic source buffer of tunable size containing indented code,
    ;       comments, quoted strings, false-positive substrings, and optional PB 6.4 escaped literals.
    ;
    Protected NewList Lines.s()
    Protected Result$, QuoteLiteral$, TotalLen.i, idx.i, *Write.Character
    ; Line accumulator and pre-sized output buffer avoiding quadratic string concatenation.
    ;
    QuoteLiteral$ = Chr(#SIMPLEQUOTE) + #DOUBLEQUOTE$ + Chr(#SIMPLEQUOTE)
    For idx = 1 To BlockCount
      AddElement(Lines()) : Lines() = "Procedure ComputeBlock_" + Str(idx) + "(*Node.MyStruct, Mode = #CheckAll)"
      AddElement(Lines()) : Lines() = "  ; Full-line comment mentioning TargetToken and Procedure to test #CheckCom rejection"
      AddElement(Lines()) : Lines() = "  Protected LocalCounter.i, PrefixTargetTokenSuffix.i, Msg$"
      AddElement(Lines()) : Lines() = "  LocalCounter = SanitizeValue(Mode) + TargetToken(LocalCounter, 10)"
      AddElement(Lines()) : Lines() = "  Msg$ = " + #DOUBLEQUOTE$ + "Literal ; with fake TargetToken and EndProcedure inside quotes" + #DOUBLEQUOTE$ + " ; trailing TargetToken comment"
      If IncludeEscapedLiterals
        AddElement(Lines()) : Lines() = "  If PeekC(*Node) <> ';' And PeekC(*Node) <> " + QuoteLiteral$ + " : LocalCounter + TargetToken(1, 2) : EndIf"
        AddElement(Lines()) : Lines() = "  Msg$ = ~" + #DOUBLEQUOTE$ + "Escaped \" + #DOUBLEQUOTE$ + " TargetToken \" + #DOUBLEQUOTE$ + " ; still string" + #DOUBLEQUOTE$ + " + Str(TargetToken(3, 4))"
      Else
        AddElement(Lines()) : Lines() = "  If LocalCounter > 0 : LocalCounter + TargetToken(1, 2) : EndIf"
        AddElement(Lines()) : Lines() = "  Msg$ = " + #DOUBLEQUOTE$ + "Plain string without escaped quotes" + #DOUBLEQUOTE$ + " + Str(TargetToken(3, 4))"
      EndIf
      AddElement(Lines()) : Lines() = "  ProcedureReturn LocalCounter"
      AddElement(Lines()) : Lines() = "EndProcedure"
    Next
    ;
    ForEach Lines()
      TotalLen + Len(Lines()) + 2
    Next
    Result$ = Space(TotalLen)
    *Write  = @Result$
    ForEach Lines()
      CopyMemoryString(Lines(), @*Write)
      *Write\c = #CR : *Write + SizeOf(Character)
      *Write\c = #LF : *Write + SizeOf(Character)
    Next
    ProcedureReturn Result$
  EndProcedure
  ;
  Procedure.i FIPC_VerifyEngines()
    ;
    ; Role: Validates strict result parity between FindInPBPointedCodeOld() and FindInPBPointedCode()
    ;       on standard syntax, and verifies PB 6.4 escaped strings, character literals, backward search,
    ;       and single-pass replacement on modern syntax.
    ; Output: Returns #True if all checks succeed, #False if any discrepancy is detected.
    ;
    Protected StdCode.String, ModernCode.String, ReplaceBuf.String
    Protected StateA.FIPC_KeywordCache, StateB.FIPC_KeywordCache
    Protected PosOld, PosNew, PosState, PosRev, IdxOld, IdxNew, IdxState, Hits, RepCount
    Protected MultiKeys$ = "Procedure,ProcedureReturn,EndProcedure,TargetToken,Protected"
    Protected Dim ExpectedPos.i(130)
    ; Test buffers, persistent search states, and match array cross-checking forward and backward scans.
    ;
    StdCode\s    = FIPC_BuildSample(40, #False)
    ModernCode\s = "TargetToken(0)" + #CRLF$ + FIPC_BuildSample(40, #True)
    ;
    ; 1. Verify strict equivalence (Old = stateless New = stateful New) on standard PureBasic syntax
    PosOld = 1 : PosNew = 1 : PosState = 1
    Repeat
      PosOld   = FindInPBPointedCodeOld(StdCode, MultiKeys$, PosOld, #CheckAll, #PB_String_NoCase, 0, @IdxOld)
      PosNew   = FindInPBPointedCode(StdCode, MultiKeys$, PosNew, #CheckAll, #PB_String_NoCase, 0, @IdxNew)
      PosState = FindInPBPointedCode(StdCode, MultiKeys$, PosState, #CheckAll, #PB_String_NoCase, 0, @IdxState, @StateA)
      If PosOld <> PosNew Or PosNew <> PosState Or IdxOld <> IdxNew Or IdxNew <> IdxState
        Debug "PARITY ERROR on StdCode: Old=" + Str(PosOld) + " New=" + Str(PosNew) + " State=" + Str(PosState)
        ProcedureReturn #False
      EndIf
      If PosNew > 0
        PosOld + 1 : PosNew + 1 : PosState + 1
      EndIf
    Until PosNew = 0
    ;
    ; Verify automatic state reset when restarting from posdep = 1 with the same StateA structure
    PosOld   = FindInPBPointedCodeOld(StdCode, MultiKeys$, 1, #CheckAll, #PB_String_NoCase, 0, @IdxOld)
    PosState = FindInPBPointedCode(StdCode, MultiKeys$, 1, #CheckAll, #PB_String_NoCase, 0, @IdxState, @StateA)
    If PosOld <> PosState Or IdxOld <> IdxState
      Debug "STATE RESET ERROR on StdCode"
      ProcedureReturn #False
    EndIf
    ;
    ; 2. Verify stateless New, stateful New, and ReverseFindInPBCode() on PB 6.4 escaped syntax (121 hits)
    Hits   = 0
    PosNew = 1 : PosState = 1
    Repeat
      PosNew   = FindInPBPointedCode(ModernCode, "TargetToken", PosNew, #CheckAll, #PB_String_CaseSensitive)
      PosState = FindInPBPointedCode(ModernCode, "TargetToken", PosState, #CheckAll, #PB_String_CaseSensitive, 0, 0, @StateB)
      If PosNew <> PosState
        Debug "MODERN STATE MISMATCH: New=" + Str(PosNew) + " State=" + Str(PosState)
        ProcedureReturn #False
      EndIf
      If PosNew > 0
        Hits + 1
        ExpectedPos(Hits) = PosNew
        PosNew + 11 : PosState + 11
      EndIf
    Until PosNew = 0
    ;
    If Hits <> 121
      Debug "UNEXPECTED HIT COUNT on ModernCode: got " + Str(Hits) + " (expected 121)"
      ProcedureReturn #False
    EndIf
    ;
    PosRev = ReverseFindInPBCode(ModernCode\s, "TargetToken", 0, #CheckAll, #PB_String_CaseSensitive)
    While PosRev > 0
      If PosRev <> ExpectedPos(Hits)
        Debug "REVERSE MISMATCH at hit #" + Str(Hits) + ": got " + Str(PosRev) + " expected " + Str(ExpectedPos(Hits))
        ProcedureReturn #False
      EndIf
      Hits - 1
      If PosRev <= 1
        Break
      EndIf
      PosRev = ReverseFindInPBCode("", "TargetToken", PosRev - 1, #CheckAll, #PB_String_CaseSensitive)
    Wend
    If Hits <> 0
      Debug "REVERSE COUNT MISMATCH: remaining=" + Str(Hits)
      ProcedureReturn #False
    EndIf
    ;
    ; 3. Verify non-cascading replacement and SearchState offset shifting in FindReplaceInPBPointedCode()
    ReplaceBuf\s = "Alpha = Beta + '; Beta' + ~" + #DOUBLEQUOTE$ + "Esc \" + #DOUBLEQUOTE$ + " Alpha" + #DOUBLEQUOTE$ + " ; Alpha"
    RepCount     = FindReplaceInPBPointedCode(ReplaceBuf, "Alpha,Beta", "Beta,Alpha", 1, #CheckAll, #PB_String_CaseSensitive)
    If RepCount <> 2 Or Left(ReplaceBuf\s, 14) <> "Beta = Alpha +"
      Debug "REPLACE ERROR: count=" + Str(RepCount) + " text=" + ReplaceBuf\s
      ProcedureReturn #False
    EndIf
    ;
    ProcedureReturn #True
  EndProcedure
  ;
  Procedure FIPC_RunBenchPair(PairIdx, BaseLegend$, *Buf.String, Keys$, Iterations, StepLen, CheckMode, FMode, EndLimit, UseState, CheckParity, *GuardSum.Quad)
    ;
    ; Role: Benchmarks FindInPBPointedCodeOld() against FindInPBPointedCode() on the exact same
    ;       workload, recording results into slots (2*PairIdx - 1) and (2*PairIdx) while optionally
    ;       verifying identical match checksums.
    ;
    Protected SearchState.FIPC_KeywordCache
    Protected TimerOld.i, TimerNew.i, iter.i, pos.i, mIdx.i
    Protected SumOld.q = 0, SumNew.q = 0
    Protected OldTag$, NewTag$
    ; Per-engine checksums, timer slot indices, and persistent cache reset once per outer file scan.
    ;
    TimerOld = PairIdx * 2 - 1
    TimerNew = PairIdx * 2
    If CheckParity
      OldTag$ = " [Old]"
    Else
      OldTag$ = " [Old-Buggy]"
    EndIf
    If UseState
      NewTag$ = " [New+State]"
    Else
      NewTag$ = " [New]"
    EndIf
    ;
    StartTimer(TimerOld, BaseLegend$ + OldTag$)
    For iter = 1 To Iterations
      pos  = 1
      mIdx = 0
      Repeat
        pos = FindInPBPointedCodeOld(*Buf, Keys$, pos, CheckMode, FMode, EndLimit, @mIdx)
        If pos
          SumOld + pos + mIdx
          pos + StepLen
        EndIf
      Until pos = 0
    Next
    ReapTimer(TimerOld)
    ;
    StartTimer(TimerNew, BaseLegend$ + NewTag$)
    If UseState
      For iter = 1 To Iterations
        SearchState\BufferAddr = 0
        pos  = 1
        mIdx = 0
        Repeat
          pos = FindInPBPointedCode(*Buf, Keys$, pos, CheckMode, FMode, EndLimit, @mIdx, @SearchState)
          If pos
            SumNew + pos + mIdx
            pos + StepLen
          EndIf
        Until pos = 0
      Next
    Else
      For iter = 1 To Iterations
        pos  = 1
        mIdx = 0
        Repeat
          pos = FindInPBPointedCode(*Buf, Keys$, pos, CheckMode, FMode, EndLimit, @mIdx)
          If pos
            SumNew + pos + mIdx
            pos + StepLen
          EndIf
        Until pos = 0
      Next
    EndIf
    ReapTimer(TimerNew)
    ;
    If CheckParity And SumOld <> SumNew
      Debug "CHECKSUM MISMATCH in " + BaseLegend$ + ": Old=" + Str(SumOld) + " vs New=" + Str(SumNew)
    EndIf
    *GuardSum\q + SumOld + SumNew
  EndProcedure
  ;
  Procedure FIPC_RunBenchmarkSuite()
    ;
    ; Role: Runs 8 paired head-to-head workloads comparing FindInPBPointedCodeOld() and
    ;       FindInPBPointedCode() across short/long buffers, hits/no-hits, multi-keyword lists,
    ;       EndLimit bounds, and PB 6.4 escaped literals, copying the report to the clipboard.
    ;
    Protected ShortCode.String, LongStdCode.String, LongModernCode.String
    Protected MultiKeys$ = "Procedure,ProcedureReturn,EndProcedure,Macro,EndMacro,Structure,EndStructure,TargetToken"
    Protected Scale.i, ct.i, CheckSum.q
    Protected TotalOld.q, TotalNew.q
    Protected Header$, Summary$, FullReport$
    ; Benchmark code samples, scaling factor, cumulative durations across Tests 1..7, and text report.
    ;
    If Not FIPC_VerifyEngines()
      MessageRequester("FindInPBCode Validation", "Validation failed! Check Debug output.", #PB_MessageRequester_Error)
      ProcedureReturn
    EndIf
    ;
    ShortCode\s      = FIPC_BuildSample(3, #False)    ; ~27 lines (~1.5 KB)
    LongStdCode\s    = FIPC_BuildSample(200, #False)  ; ~1,800 lines (~95 KB, compatible with Old)
    LongModernCode\s = FIPC_BuildSample(200, #True)   ; ~1,800 lines (~102 KB, PB 6.4 ~"..." & ';')
    ;
    CompilerIf #PB_Compiler_Debugger
      Scale   = 1
      Header$ = "[WARNING: Debugger is ON - disable Debugger for true native CPU ratios]" + #CRLF$
    CompilerElse
      Scale   = 3
      Header$ = "[Native x64/x86 execution - Debugger OFF]" + #CRLF$
    CompilerEndIf
    Debug Header$
    InitAllTimers()
    ;
    ; Warm up instruction and data caches equally before starting timed runs
    FindInPBPointedCodeOld(LongStdCode, "AbsentSymbolXYZ", 1, #CheckAll, #PB_String_NoCase)
    FindInPBPointedCode(LongStdCode, "AbsentSymbolXYZ", 1, #CheckAll, #PB_String_NoCase)
    ;
    ; Pairs 1..7: Strictly identical results between Old and New (verified via CheckParity = #True)
    FIPC_RunBenchPair(1, "T1 Short / Multi-Hits", @ShortCode,   "TargetToken",          1200 * Scale, 11, #CheckAll, #PB_String_CaseSensitive, 0,    #False, #True,  @CheckSum)
    FIPC_RunBenchPair(2, "T2 Short / NoHit",      @ShortCode,   "MissingIdentifierXYZ", 2000 * Scale, 20, #CheckAll, #PB_String_NoCase,        0,    #False, #True,  @CheckSum)
    FIPC_RunBenchPair(3, "T3 Long  / 600 Hits",   @LongStdCode, "TargetToken",            12 * Scale, 11, #CheckAll, #PB_String_CaseSensitive, 0,    #False, #True,  @CheckSum)
    FIPC_RunBenchPair(4, "T4 Long  / 600 Hits",   @LongStdCode, "TargetToken",            12 * Scale, 11, #CheckAll, #PB_String_CaseSensitive, 0,    #True,  #True,  @CheckSum)
    FIPC_RunBenchPair(5, "T5 Long  / NoHit",      @LongStdCode, "MissingIdentifierXYZ",   25 * Scale, 20, #CheckAll, #PB_String_NoCase,        0,    #False, #True,  @CheckSum)
    FIPC_RunBenchPair(6, "T6 Long  / 8-Keys",     @LongStdCode, MultiKeys$,                1 * Scale,  5, #CheckAll, #PB_String_NoCase,        0,    #True,  #True,  @CheckSum)
    FIPC_RunBenchPair(7, "T7 Long  / EndLimit",   @LongStdCode, "TargetToken",          1200 * Scale, 11, #CheckAll, #PB_String_CaseSensitive, 2500, #False, #True,  @CheckSum)
    ;
    ; Pair 8: PB 6.4 escaped strings (~"...") & char literals (';') where Old fails on context
    FIPC_RunBenchPair(8, "T8 PB6.4 / ~Quote",     @LongModernCode, "TargetToken",         12 * Scale, 11, #CheckAll, #PB_String_CaseSensitive, 0,    #True,  #False, @CheckSum)
    ;
    ; Sum per-engine totals across Pairs 1..7 (strictly equivalent workloads) outside StageTimer()
    For ct = 1 To 7
      TotalOld + StageTimer(ct * 2 - 1)
      TotalNew + StageTimer(ct * 2)
    Next
    ;
    FullReport$ = Header$ + DebugAllTimers()
    Summary$    = "Cumulative (Tests 1..7 - strictly identical results):" + #CRLF$ +
                  "  1. FindInPBPointedCodeOld() : " + FormatMicroseconds(TotalOld) + #CRLF$ +
                  "  2. FindInPBPointedCode()    : " + FormatMicroseconds(TotalNew) + #CRLF$ +
                  "  (Checksum guard: " + Str(CheckSum) + ")"
    Debug Summary$
    FullReport$ + Summary$
    SetClipboardText(FullReport$)
    ;
    CompilerIf #PB_Compiler_Debugger = 0
      MessageRequester("FindInPBCode Benchmark (copied to clipboard)", FullReport$)
    CompilerEndIf
  EndProcedure
  ;
  FIPC_RunBenchmarkSuite()
  ;
CompilerEndIf
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 782
; FirstLine = 782
; Folding = -----
; EnableXP
; DPIAware