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
; The main function is EvaluateExpression(Expression$)
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
Procedure.s EvaluateSimpleExpression(Expression$)
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
      sfind = FindString(Expression$, occ$)
      If sfind
        p = sfind
        ps = p
        pe = p + 1
        ; Handles operator pairs such as '&~'
        While FindString(Ops$, Mid(Expression$, pe, 1)) : pe + 1 : Wend
        mpe = pe
        Break
      EndIf
    Next
    If p
      While ps > 1 And FindString(Ops$ + "()", Mid(Expression$, ps - 1, 1)) = 0 : ps - 1 : Wend
      ;
      While pe <= Len(Expression$) And FindString(Ops$ + "()", Mid(Expression$, pe, 1)) = 0 : pe + 1 : Wend
      ;
      VLeft = Val(Mid(Expression$, ps, p - ps))
      VRight = Val(Mid(Expression$, mpe, pe - mpe))
      op$ = Mid(Expression$, p, mpe - p)
      Expression$ = Mid(Expression$, 1, ps - 1) + Str(ApplyOperator(VLeft, op$, VRight)) + Mid(Expression$, pe)
    EndIf
  Until p = 0 Or ps = 1 Or pe >= Len(Expression$)
  If EvError
    ProcedureReturn ""
  Else
    ProcedureReturn Expression$
  EndIf
EndProcedure
;
Procedure EvaluateExpression(Expression$)
  ; As its name indicates, this procedure attempts
  ; to evaluate the expression passed as a parameter,
  ; i.e., to calculate its value.
  ;
  Protected pe, ps, r$
  EvError = 0
  ;
  Expression$ = UCase(ReplaceString(Expression$, " ", ""))
  Repeat
    pe = FindString(Expression$, ")")
    If pe
      ps = pe - 1
      While ps And Mid(Expression$, ps, 1) <> "(" : ps - 1 : Wend
      r$ = EvaluateSimpleExpression(Mid(Expression$, ps + 1, pe - ps - 1))
      Expression$ = Left(Expression$, ps - 1) + r$ + Mid(Expression$, pe + 1)
    Else
      Expression$ = EvaluateSimpleExpression(ReplaceString(Expression$, "(", ""))
    EndIf
  Until pe = 0
  If EvError
    ProcedureReturn 0
  Else
    ProcedureReturn Val(EvaluateSimpleExpression(Expression$))
  EndIf
EndProcedure
;
; Example of usage:
CompilerIf #PB_Compiler_IsMainFile
  ; The following won't run when this file is used as 'Included'.
  ;
   Expression$ = "$FFF8&%1001"
   result  = EvaluateExpression(Expression$)
   Debug "The expression " + Expression$ + " is equal to " + Str(result)
   Debug "Verification: " + Str($FFF8&%1001)
   Debug "_____________________________________________"
   Expression$ = "(44/2)+3*4"
   result  = EvaluateExpression(Expression$)
   Debug "The expression " + Expression$ + " is equal to " + Str(result)
   Debug "Verification: " + Str((44/2)+3*4)
   Expression$ = "3 + 2 * (5 + 7)"
   result  = EvaluateExpression(Expression$)
   Debug "The expression " + Expression$ + " is equal to " + Str(result)
   Debug "Verification: " + Str(3 + 2 * (5 + 7))
   
CompilerEndIf

; IDE Options = PureBasic 6.21 Beta 2 (Windows - x86)
; CursorPosition = 122
; FirstLine = 68
; Folding = -
; EnableXP
; DPIAware