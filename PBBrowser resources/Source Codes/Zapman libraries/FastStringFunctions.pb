;
Macro FastMid(FMString, FMStartPos, FMLength = -1)
  ; Fait la même chose que Mid() sans les contrôles associés à cette fonction.
  ; Il est par conséquent impératif que :
  ; 1- FMLength soit renseignée et supérieure à zéro,
  ; 2- FMStartPos soit supérieure à zéro et inférieure à la longueur de FMString,
  ; 2- FMStartPos+FMLength soit inférieur ou égal à la longueur de FMString.
  ;
  ; Pour des chaînes à extraire de longueurs aléatoires FastMid() est environ
  ; un tiers plus rapide que Mid().
  ; Quand une boucle utilisant exclusivement Mid() met une seconde à s'exécuter,
  ; elle met 663 millisecondes à s'exécuter avec FastMid()
  ;
  ; Pour des chaînes à extraire de petites tailles (FMLength < 10), FastMid()
  ; est presque deux fois plus rapide que Mid().
  ; Quand une boucle utilisant exclusivement Mid() met une seconde à s'exécuter,
  ; elle met 540 millisecondes à s'exécuter avec FastMid()
  PeekS(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER), FMLength)
EndMacro
;
Macro FastLeft(FMString, FMLength)
  ; Fait la même chose que Left() sans les contrôles associés à cette fonction.
  ; Il est par conséquent impératif que :
  ; 1- FMLength soit renseignée et supérieure ou égale à zéro,
  ; 2- FMLength soit inférieur ou égale à la longueur de FMString.
  ;
  ; FastLeft est entre 2 et 400 fois plus rapide que Left, selon la taille de FMLength.
  ; Plus FMLength est grand, moins le gain de temps est significatif.
  PeekS(@FMString, FMLength)
EndMacro
;
Macro FastSubstituteString(FMString, FMToReplace, FMStartPos)
  ; Remplace une portion de chaîne par une autre de même longueur.
  ; Il est impératif que FMStartPos+len(FMToReplace) soit inférieur ou égal
  ; à len(FMString).
  PokeS(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER), FMToReplace, -1, #PB_String_NoZero)
EndMacro
;
Macro FastFindPrecReturn(FMString, FMStartPos)
  ; Cherche le retour chariot qui précède la position FMStartPos
  While FMStartPos And PeekC(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER)) <> 13
    FMStartPos - 1
  Wend
EndMacro
;
Macro FastFindPrecSpaces(FMString, FMStartPos)
  ; Cherche le premier espace qui précède la position FMStartPos
  While FMStartPos And PeekC(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER)) <> 32
    FMStartPos - 1
  Wend
EndMacro
;
Macro FastSkipPrecSpaces(FMString, FMStartPos)
  ; Cherche le premier caractère qui n'est pas un espace en amont de FMStartPos
  While FMStartPos And PeekC(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER)) = 32
    FMStartPos - 1
  Wend
EndMacro
;
Macro FastSkipFollowingSpaces(FMString, FMStartPos)
  ; Cherche le premier caractère qui n'est pas un espace en aval de FMStartPos
  While FMStartPos And PeekC(@FMString + (FMStartPos - 1) * SizeOf(CHARACTER)) = 32
    FMStartPos + 1
  Wend
EndMacro
;
Macro FastStringRemove(FMString, FMPos1, FMPos2)
  ; Déplace les caractères situés après la position FMPos2 à la position FMPos1.
  ; Ainsi, tout ce qui existait entre FMPos1 et FMPos2 est supprimé de la chaîne.
  ; Cette macro est équivalente à FMString = Left(FMString,FMPos1-1) + Mid(FMString,FMPos2)
  ; Il est impératif que FMPos2 soit supérieure à FMPos1.
  ; Il est ABSOLUMENT IMPERATIF que FMPos1 et FMPos2 soient tous les deux supérieures à zéro
  ; et inférieures à la taille de la chaîne.
  ; Le caractère '#Null' qui est situé à la fin de la chaine est déplacé avec le reste.
  ;
  ; FastStringRemove est environ un tiers plus rapide que FMString = Left(FMString,FMPos1-1) + Mid(FMString,FMPos2)
  ;
  CopyMemory(@FMString + (FMPos2 - 1) * SizeOf(CHARACTER), @FMString + (FMPos1 - 1) * SizeOf(CHARACTER), (Len(FMString) - FMPos2 + 2) * SizeOf(CHARACTER))
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

; IDE Options = PureBasic 6.10 LTS (Windows - x64)
; CursorPosition = 20
; FirstLine = 15
; Folding = --
; EnableXP
; DPIAware