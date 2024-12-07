;
; ******************************************************************
;
;                       Fast string functions
;                       by Zapman - oct. 2024
;
; ******************************************************************
;
Macro FastMid(FMString, FMStartPos, FMLength = -1)
  ; Does the same thing as Mid() without the associated checks.
  ; Therefore, it is imperative that:
  ; 1- FMLength is greater than -2,
  ; 2- FMStartPos is greater than zero and less than the length of FMString,
  ; 3- FMStartPos + FMLength is less than or equal to the length of FMString.
  ;
  ; For strings of random length to extract, FastMid() is about
  ; one third faster than Mid().
  ; When a loop using only Mid() takes one second to execute,
  ; it takes 663 milliseconds to execute with FastMid().
  ;
  ; For strings of small length to extract (FMLength < 10), FastMid()
  ; is almost twice as fast as Mid().
  ; When a loop using only Mid() takes one second to execute,
  ; it takes 540 milliseconds to execute with FastMid().
  PeekS(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER), FMLength)
EndMacro
;
Macro FastLeft(FMString, FMLength)
  ; Does the same thing as Left() without the associated checks.
  ; Therefore, it is imperative that:
  ; 1- FMLength is provided and greater than or equal to zero,
  ; 2- FMLength is less than or equal to the length of FMString.
  ;
  ; FastLeft is between 2 and 400 times faster than Left, depending on the size of FMLength.
  ; The larger FMLength is, the less significant the time gain.
  PeekS(@FMString, FMLength)
EndMacro
;
Macro FastFindPrecReturn(FMString, FMStartPos)
  ; Searches for the carriage return that precedes the position FMStartPos
  While FMStartPos And PeekC(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER)) <> 13
    FMStartPos - 1
  Wend
EndMacro
;
Macro FastFindPrecSpaces(FMString, FMStartPos)
  ; Searches for the first space preceding the position FMStartPos
  While FMStartPos And PeekC(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER)) <> 32
    FMStartPos - 1
  Wend
EndMacro
;
Macro FastSkipPrecSpaces(FMString, FMStartPos)
  ; Searches for the first non-space character upstream of FMStartPos
  While FMStartPos And PeekC(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER)) = 32
    FMStartPos - 1
  Wend
EndMacro
;
;
; Text$ = FichierDanstexte("MonFichier")
; 
; ;
; MaxtestLoops = 1000
; 
; InitTimer(1)
; StartTimer(1)
; For ct= 1 To MaxtestLoops
;     
;   p = Random(100,2)
;   
;   a$ = Left(Text$,p)
;   
; Next
; reapTimer(1)
; Debug "Timer "+Str(1)+" : "+Str(StageTimer(1))
; mt = StageTimer(1)
; 
; InitTimer(1)
; StartTimer(1)
; For ct= 1 To MaxtestLoops
;     
;   p = Random(100,2)
;   
;   a$ = FastLeft(Text$,p)
;   
; Next
; reapTimer(1)
; Debug "Timer "+Str(1)+" : "+Str(StageTimer(1))
; Debug StrF(mt/StageTimer(1))

; IDE Options = PureBasic 6.11 LTS (Windows - x64)
; CursorPosition = 36
; FirstLine = 24
; Folding = -
; EnableXP
; DPIAware