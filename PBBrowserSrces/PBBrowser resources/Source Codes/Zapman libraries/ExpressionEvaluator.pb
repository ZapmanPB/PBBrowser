; **************************************************
;
;               ExpressionEvaluator
;              By Zapman - July 2024
;
;***************************************************
; Ce jeu de procédures permet d'évaluer des fonctions
; simples telles que "3 + 2 * (5 + 7)".
;
; La fonction principale est EvaluateExpression(expression$)
;
; L'ensemble ne fonctionne que pour des nombres entiers (integer)
; mais tolère l'emploi de valeur hexadécimales ou binaires.
; Les opérateurs admis sont : */+-~|&!
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
      ProcedureReturn VLeft / VRight
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
  ; Évalue (calcule) une expression dépourvue de parenthèses.
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
        ; On gère les doublets d'opérateurs tels que '&~'
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
  ProcedureReturn expression$
EndProcedure
;
Procedure EvaluateExpression(expression$)
  ; Comme son nom l'indique, cette procédure tente
  ; d'évaluer l'expression passée en paramètre,
  ; c'est-à-dire de calculer sa valeur.
  Protected pe, ps, r$
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
  ProcedureReturn Val(EvaluateSimpleExpression(expression$))
EndProcedure
;
; Exemple d'utilisation :
;
;  expression$ = "$FFF8&%1001"
;  result  = EvaluateExpression(expression$)
;  Debug "L'expression " + expression$ + " est égale à " + Str(result)
;  Debug "Vérification : " + Str($FFF8&%1001)
;  Debug "_____________________________________________"
;  expression$ = "(45/2)+3*4"
;  result  = EvaluateExpression(expression$)
;  Debug "L'expression " + expression$ + " est égale à " + Str(result)
;  Debug "Vérification : " + Str((45/2)+3*4)

; IDE Options = PureBasic 6.11 LTS (Windows - x64)
; CursorPosition = 102
; FirstLine = 9
; Folding = 5
; EnableXP
; DPIAware