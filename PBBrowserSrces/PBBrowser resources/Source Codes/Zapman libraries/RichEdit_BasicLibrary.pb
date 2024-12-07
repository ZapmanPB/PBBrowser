;
;*********************************************
;            RichEdit_BasicLibrary
;
;      Basic functions for EditorGadgets
;     mainly collected on PureBasic forums
;      and adapted by Zapman - nov. 2024
;*********************************************
;
;
Procedure.l RE_StreamStringOutCallback(*dwCookiePtr, pbBuff, cb, pcb)
  Protected result, StrPtr, ms
  result = 0
  If *dwCookiePtr ; Here, *dwCookiePtr is a pointer to a pointer
    ;               along time, *dwCookiePtr will allways keep the same value (StrPtr pointer adress)
    ;               while the value of StrPtr can change at each iteration.
    StrPtr.i = PeekI(*dwCookiePtr)
    If StrPtr = 0
      StrPtr = AllocateMemory(cb)
      CopyMemory(pbBuff, StrPtr, cb)
    Else
      ms = MemorySize(StrPtr)
      StrPtr = ReAllocateMemory(StrPtr, ms + cb)
      CopyMemory(pbBuff, StrPtr + ms, cb)
    EndIf
    PokeI(*dwCookiePtr, StrPtr)
    ;
  EndIf
  PokeL(pcb, cb)
  If cb = 0
    result = 1
  EndIf
  ProcedureReturn result
EndProcedure
;
Procedure.s RE_GetContent_RTF(Gadget, format = #SF_RTF, UseDirectID = 0)
  ; Get content of a RichEdit Gadget
  ;
  ; Resulting text will be RTF formated if 'format' parameter is undefined.
  ; Resulting text will be simple text formated if 'format' parameter is set to zero,
  ;     (in that case, resulting text is UTF8 encoded).
  ; Resulting text will get whole content of the RichEdit Gadget
  ;     or only selected content if 'format' is set To #SF_TEXT|#SFF_SELECTION Or #SF_RTF|#SFF_SELECTION
  ;
  ;'format' parameter can also be a combination of #SF_RTF and #SF_TEXT with some of the following values:
  ; #SFF_PLAINRTF -- If specified, the rich edit control streams out only the keywords common To all languages, ignoring language-specific keywords. If Not specified, the rich edit control streams out all keywords. You can combine this flag With the SF_RTF Or SF_RTFNOOBJS flag.
  ; #SFF_SELECTION -- If specified, the rich edit control streams out only the contents of the current selection. If Not specified, the control streams out the entire contents. You can combine this flag With any of Data format values.
  ; #SF_UNICODE -- Microsoft Rich Edit 2.0 And later: Indicates Unicode text. You can combine this flag With the SF_TEXT flag.
  ; #SF_USECODEPAGE -- Rich Edit 3.0 And later: Generates UTF-8 RTF As well As text using other code pages. The code page is set in the high word of wParam. For example, For UTF-8 RTF, set wParam To (CP_UTF8 << 16) | SF_USECODEPAGE | SF_RTF.
  ;
  Protected edstr.EDITSTREAM, StrPtr, Ghdl
  Protected Str$ ; Return value
  ;
  #SF_USECODEPAGE   = $20
  #CP_UTF8      = 65001
  ;
  If UseDirectID = 0
    Ghdl = GadgetID(Gadget)
  Else
    Ghdl = Gadget
  EndIf
  ;
  If format = 0 Or format & #SF_TEXT : format | (#CP_UTF8 << 16) | #SF_USECODEPAGE | #SF_TEXT : EndIf
  ;
  StrPtr = 0
  edstr\dwCookie.i = @StrPtr ; Here, *dwCookie is a pointer to a pointer.
  ;                            Along time, *dwCookie will allways keep the same value (StrPtr pointer adress)
  ;                            while the value of StrPtr can change along iterations of RE_StreamStringOutCallback
  edstr\pfnCallback = @RE_StreamStringOutCallback()
  edstr\dwError = 0
  SendMessage_(Ghdl, #EM_STREAMOUT, format, edstr)
  If edstr\dwError
    Str$ = ""
    If StrPtr : FreeMemory(StrPtr) : EndIf
  ElseIf StrPtr
    Str$ = PeekS(StrPtr, MemorySize(StrPtr), #PB_UTF8 | #PB_ByteLength)
    FreeMemory(StrPtr)
  EndIf
  ProcedureReturn Str$
EndProcedure
;
Procedure.s RE_GetGadgetSelectedText(Gadget, format = #SFF_SELECTION | #SF_TEXT)
  ; Get text content of a RichEdit Gadget formated in simple Text$ / UTF8
  ProcedureReturn RE_GetContent_RTF(Gadget, format)
EndProcedure
;
Procedure RE_GetSelection(Gadget, *txtrange.CHARRANGE)
  ProcedureReturn SendMessage_(GadgetID(Gadget), #EM_EXGETSEL, 0, *txtrange)
EndProcedure
;
Procedure RE_SetSelection(Gadget, PosStart, PosEnd)
  ;
  Protected txtrange.CHARRANGE\cpMin = PosStart ; Start of selection. 
  txtrange\cpMax = PosEnd                       ; End of selection
  ProcedureReturn SendMessage_(GadgetID(Gadget), #EM_EXSETSEL, 0, @txtrange)
  ;
EndProcedure
;
Procedure RE_ReplaceSelection(Gadget, ReplaceString$)
  ProcedureReturn SendMessage_(GadgetID(Gadget), #EM_REPLACESEL, 1, @ReplaceString$)
EndProcedure
;
Procedure.s RE_GetGadgetWholeText(Gadget, format = #SF_TEXT)
  ; Get text content of a RichEdit Gadget formated in simple Text$ / UTF8
  ProcedureReturn RE_GetContent_RTF(Gadget, format)
EndProcedure
;
Procedure RE_SaveContent(Gadget, FileName$, format = #SF_RTF, DelPict = 0)
  ; Save content of a RichEdit Gadget formated in UTF8
  ; If format is unsetted, the resulting file will content RTF data
  ; If 'format' is set To #SF_TEXT|SFF_SELECTION Or #SF_RTF|SFF_SELECTION, only selected text will be saved
  ;
  Protected ct, p, pr, po, pf, hFile, Result
  Protected Content$ = RE_GetContent_RTF(Gadget, format)
  ;
  If Content$
    ; La couleur par défaut, qui apparaît bien en noir dans le REGadget, devient
    ; transparente dans le fichier enregistré.
    ; Il faut corriger.
    Content$ = ReplaceString(Content$, "\red255\green255\blue255", "\red0\green0\blue0")
    ;
    If DelPict
      ; On supprime les images
      ct = 0
      Repeat
        p = FindString(Content$, "{\pict")
        If p
          ct = 1
          pr = p + 5
          Repeat
            po = FindString(Content$, "{", pr + 1)
            pf = FindString(Content$, "}", pr + 1)
            If po And (po < pf Or pf = 0)
              pr = po
              ct + 1
            EndIf
            If pf And (pf < po Or po = 0)
              pr = pf
              ct - 1
            EndIf
          Until ct = 0 Or (pr = 0 And pf = 0)
          If pr
            pr + 1
            While Mid(Content$, pr, 1) = " " : pr + 1 : Wend
            While Mid(Content$, p - 1, 1) = " " : p - 1 : Wend
            Content$ = Left(Content$, p - 1) + Mid(Content$, pr)
          EndIf
        EndIf
      Until p = 0
    EndIf
    ;
    ; Open file and write buffer content.
    hFile = CreateFile(#PB_Any, Filename$)
    ;
    If hFile
      ; Write content into file.
      Result = WriteString(hFile, Trim(Content$))
      CloseFile(hFile)
    EndIf
    ProcedureReturn Result
  EndIf
EndProcedure
;
Procedure RE_AdjustZoom(NoGadget, AdjFactor)
  ;
  Protected num.l, denom.l, Factor.f
  ;
  ; Obtenir le niveau de zoom actuel
  SendMessage_(GadgetID(NoGadget), #EM_GETZOOM, @num, @denom)

  ; Calculer le nouveau niveau de zoom
  If num And denom
    Factor = num / denom
  Else
    Factor = 1
  EndIf
  ;
  If AdjFactor
    Factor * (AdjFactor + 100) / 100
  Else
    Factor = 1
  EndIf
  ;
  denom = 10
  num = denom * Factor
  ;
  ; Mettre à jour le niveau de zoom
  SendMessage_(GadgetID(NoGadget), #EM_SETZOOM, num, denom)
EndProcedure
;
; IDE Options = PureBasic 6.11 LTS (Windows - x64)
; CursorPosition = 88
; Folding = --
; EnableXP
; DPIAware