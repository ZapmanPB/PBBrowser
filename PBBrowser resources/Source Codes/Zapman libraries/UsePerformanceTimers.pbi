Global Dim StageTimer.q(40)
Global Dim StageTimerRef.q(40)
Global Dim TimerLegend$(40)
Global NbrOfUsedTimers
;
Procedure TicksHQ() 
  ; Équivalent d'un ElapsedMicroseconds()
  Static STAdjust.q
  Protected t.q, OldTick.q
  If STAdjust = 0
    ; Measure how much PerformanceCounter there is during 100 ms.
    ; This is done only one time during the program startup.
    QueryPerformanceCounter_(@OldTick.q)
    Delay(100)
    QueryPerformanceCounter_(@t.q)
    STAdjust = (t - OldTick) / 100000
  EndIf
  QueryPerformanceCounter_(@t.q)
  ProcedureReturn t/STAdjust
EndProcedure 
;
Procedure StartTimer(Nbr, Legend$ = "")
  StageTimerRef(Nbr) = TicksHQ()
  If Nbr > NbrOfUsedTimers : NbrOfUsedTimers = Nbr : EndIf
  If Legend$ : TimerLegend$(Nbr) = Legend$ : EndIf
EndProcedure
;
Procedure ReapTimer(Nbr)
  If StageTimerRef(Nbr)
    StageTimer(Nbr) + TicksHQ() - StageTimerRef(Nbr)
  Else
    StageTimer(Nbr) = 0
  EndIf
  StageTimerRef(Nbr) = TicksHQ()
  If Nbr > NbrOfUsedTimers : NbrOfUsedTimers = Nbr : EndIf
EndProcedure
;
Procedure InitTimer(Nbr)
  StageTimerRef(Nbr) = TicksHQ()
  StageTimer(Nbr) = 0
  If Nbr > NbrOfUsedTimers : NbrOfUsedTimers = Nbr : EndIf
EndProcedure
;
Procedure InitAllTimers()
  Protected ct
  For ct = 0 To ArraySize(StageTimer())
    StageTimer(ct) = 0
    StageTimerRef(ct) = 0
  Next
EndProcedure
;
Procedure.s FormatMicroseconds(Microseconds.q)
  ;
  ; Formats a duration in microseconds into a readable string "X sec. Y ms Z µs".
  ;
  Protected Seconds.q, Milliseconds.q, RemainingUs.q
  Protected Result$
  ;
  If Microseconds < 0
    Microseconds = 0
  EndIf
  ;
  Seconds = Microseconds / 1000000
  Microseconds = Microseconds % 1000000
  ;
  Milliseconds = Microseconds / 1000
  RemainingUs = Microseconds % 1000
  ;
  Result$ = Str(Seconds) + " sec. " + Str(Milliseconds) + " ms " + Str(RemainingUs) + " µs"
  ;
  ProcedureReturn Result$
EndProcedure

Procedure.s DebugAllTimers()
  Protected ct, nb$, Result$, ConcResult$, TotTime.q
  Debug "_____________________________"
  For ct = 1 To NbrOfUsedTimers
    If StageTimer(ct)
      If TimerLegend$(ct) And Left(TimerLegend$(ct), 2) <> " -"
        TimerLegend$(ct) = " - " + TimerLegend$(ct) + " - "
      EndIf
      Result$ = "Timer "+Str(ct) + TimerLegend$(ct) + " : " + FormatMicroseconds(StageTimer(ct))
      Debug Result$
      ConcResult$ + Result$ + #CR$
      TotTime + StageTimer(ct)
    EndIf
  Next
  Result$ = "Total time  : " + FormatMicroseconds(TotTime)
  Debug Result$
  ConcResult$ + Result$ + #CR$
  ;
  ProcedureReturn ConcResult$
  ;
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  ; With the following, result should be Timer 1 : 1 sec. 0 ms 0 µs
  ; but the result may vary regarding what the computer has to do
  ; outside of the program
  StartTimer(1)
  Delay(1000)
  ReapTimer(1)
  DebugAllTimers()
CompilerEndIf

; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 51
; FirstLine = 23
; Folding = --
; EnableXP
; DPIAware