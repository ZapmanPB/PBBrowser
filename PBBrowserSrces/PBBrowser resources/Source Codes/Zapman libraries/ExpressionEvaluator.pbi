;*********************************************************************
;
;                        ExpressionEvaluator
;                       By Zapman - March 2025
;
; This file should be saved under the name "ExpressionEvaluator.pbi".
;
;*********************************************************************
; This set of procedures evaluates simple functions
; such as "3 + 2 * (5 + 7)".
;
; The main function is EvaluateExpression(expression$)
;
; The set only works for integer numbers,
; but allows the use of hexadecimal or binary values.
; The allowed operators are: */+-~|&!
;
Global EvError
;
Procedure.i ApplyOperator(VLeft.i, op.s, VRight.i)
  Select op
    Case "+"
      ProcedureReturn VLeft + VRight
    Case "-"
      ProcedureReturn VLeft - VRight
    Case "*"
      ProcedureReturn VLeft * VRight
    Case "/"
      If VRight
        ProcedureReturn VLeft / VRight
      Else
        EvError = 1
        ProcedureReturn 0
      EndIf
    Case "|"
      ProcedureReturn VLeft | VRight
    Case "&"
      ProcedureReturn VLeft & VRight
    Case "~"
      ProcedureReturn VLeft | ~ VRight
    Case "!"
      ProcedureReturn VLeft ! VRight
    Case "|~"
      ProcedureReturn VLeft | ~ VRight
    Case "&~"
      ProcedureReturn VLeft & ~ VRight
    Case "!~"
      ProcedureReturn VLeft !~ VRight  
  EndSelect
  ProcedureReturn 0
EndProcedure
;
Procedure.s EvaluateSimpleExpression(expression$)
  ;
  ; Evaluates (calculates) an expression without parentheses.
  ;
  Protected Ops$ = "*/+-~|&!", occ$, op$
  Protected p, l, ps, pe, mpe, sfind
  Protected VLeft, VRight
  ;
  Repeat
    p = 0
    For l = 1 To Len(Ops$)
      occ$ = Mid(Ops$, l, 1)
      sfind = FindString(expression$, occ$)
      If sfind
        p = sfind
        ps = p
        pe = p + 1
        ; Handles operator pairs such as '&~'
        While FindString(Ops$, Mid(expression$, pe, 1)) : pe + 1 : Wend
        mpe = pe
        Break
      EndIf
    Next
    If p
      While ps > 1 And FindString(Ops$ + "()", Mid(expression$, ps - 1, 1)) = 0 : ps - 1 : Wend
      ;
      While pe <= Len(expression$) And FindString(Ops$ + "()", Mid(expression$, pe, 1)) = 0 : pe + 1 : Wend
      ;
      VLeft = Val(Mid(expression$, ps, p - ps))
      VRight = Val(Mid(expression$, mpe, pe - mpe))
      op$ = Mid(expression$, p, mpe - p)
      expression$ = Mid(expression$, 1, ps - 1) + Str(ApplyOperator(VLeft, op$, VRight)) + Mid(expression$, pe)
    EndIf
  Until p = 0 Or ps = 1 Or pe >= Len(expression$)
  If EvError
    ProcedureReturn ""
  Else
    ProcedureReturn expression$
  EndIf
EndProcedure
;
Procedure EvaluateExpression(expression$)
  ; As its name indicates, this procedure attempts
  ; to evaluate the expression passed as a parameter,
  ; i.e., to calculate its value.
  ;
  Protected pe, ps, r$
  EvError = 0
  ;
  expression$ = UCase(ReplaceString(expression$, " ", ""))
  Repeat
    pe = FindString(expression$, ")")
    If pe
      ps = pe - 1
      While ps And Mid(expression$, ps, 1) <> "(" : ps - 1 : Wend
      r$ = EvaluateSimpleExpression(Mid(expression$, ps + 1, pe - ps - 1))
      expression$ = Left(expression$, ps - 1) + r$ + Mid(expression$, pe + 1)
    Else
      expression$ = EvaluateSimpleExpression(ReplaceString(expression$, "(", ""))
    EndIf
  Until pe = 0
  If EvError
    ProcedureReturn 0
  Else
    ProcedureReturn Val(EvaluateSimpleExpression(expression$))
  EndIf
EndProcedure
;
; Example of usage:
CompilerIf #PB_Compiler_IsMainFile
  ; The following won't run when this file is used as 'Included'.
  ;
   expression$ = "$FFF8&%1001"
   result  = EvaluateExpression(expression$)
   Debug "The expression " + expression$ + " is equal to " + Str(result)
   Debug "Verification: " + Str($FFF8&%1001)
   Debug "_____________________________________________"
   expression$ = "(44/2)+3*4"
   result  = EvaluateExpression(expression$)
   Debug "The expression " + expression$ + " is equal to " + Str(result)
   Debug "Verification: " + Str((44/2)+3*4)
   expression$ = "3 + 2 * (5 + 7)"
   result  = EvaluateExpression(expression$)
   Debug "The expression " + expression$ + " is equal to " + Str(result)
   Debug "Verification: " + Str(3 + 2 * (5 + 7))
   
CompilerEndIf

; IDE Options = PureBasic 6.20 Beta 4 (Windows - x64)
; CursorPosition = 103
; Folding = -
; EnableXP
; DPIAware