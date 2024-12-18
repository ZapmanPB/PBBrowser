Procedure.s GetFontName(FontID)
  ;
  Protected FontName$
  ;
  If IsFont(FontID)
    FontID = FontID(FontID)
  EndIf
  ;
  CompilerIf #PB_Compiler_OS = #PB_OS_MacOS
    ; returns the font name of FontID
    ;
    Protected string
    If FontID
      string = CocoaMessage(0, FontID, "displayName") ; "familyName" and "fontName" for internal use
                                                      ; use "displayName" for the real name
      If string
        FontName$ = PeekS(CocoaMessage(0, string, "UTF8String"), -1, #PB_UTF8)
      EndIf
    EndIf
  CompilerElseIf #PB_Compiler_OS = #PB_OS_Windows
    Protected finfo.LOGFONT
    GetObject_(fontid, SizeOf(LOGFONT), @finfo)
    FontName$ = PeekS(@finfo\lfFaceName[0])
  CompilerEndIf
  ;
  ProcedureReturn FontName$
  ;
EndProcedure
;
Procedure.f GetFontSize(FontID)
  ;
  Protected FpointSize
  ;
  If IsFont(FontID)
    FontID = FontID(FontID)
  EndIf
  ;
  CompilerIf #PB_Compiler_OS = #PB_OS_MacOS
    ; returns the font size of FontID
    ;
    Protected pointSize.CGFloat = 0.0
    ;
    If FontID
       CocoaMessage(@pointSize, FontID, "pointSize")
    EndIf
    FpointSize = pointSize
    ProcedureReturn FpointSize
  CompilerElseIf #PB_Compiler_OS = #PB_OS_Windows
    ; Get gadget font height in points (to the nearest integer)
    Protected finfo.LOGFONT, height
    Protected hdc = GetDC_(#Null) ; Obtenir le contexte de périphérique
    If GetObject_(FontID, SizeOf(LOGFONT), @finfo.LOGFONT)
      FpointSize = -finfo\lfHeight * 72 / GetDeviceCaps_(hdc, #LOGPIXELSY)
    EndIf
    ProcedureReturn FpointSize
  CompilerEndIf
  ;
EndProcedure
;
Procedure.s GetFontAttributes(FontID)
  ;
  Protected FontAttributes$
  ;
  If IsFont(FontID)
    FontID = FontID(FontID)
  EndIf
  ;
  CompilerIf #PB_Compiler_OS = #PB_OS_MacOS
    ; Pour macOS : vérification des attributs de la police
    Protected isItalic.B = #False, isBold.B = #False, isUnderlined.B = #False
    ;
    If FontID
      CocoaMessage(@isItalic, FontID, "fontDescriptor traits") ; Récupère les traits de la police
      CocoaMessage(@isBold, FontID, "fontDescriptor", "symbolicTraits containsObject:", @"NSFontBoldTrait")
      CocoaMessage(@isItalic, FontID, "fontDescriptor", "symbolicTraits containsObject:", @"NSFontItalicTrait")
      CocoaMessage(@isUnderlined, FontID, "fontDescriptor", "symbolicTraits containsObject:", @"NSFontUnderlineTrait")
    EndIf
    ;
    If isItalic
      FontAttributes$ = "Italic"
    EndIf
    If isBold
      FontAttributes$ + ", Bold, "
    EndIf
    If isUnderlined
      FontAttributes$ + ", Underlined"
    EndIf
    ;
  CompilerElseIf #PB_Compiler_OS = #PB_OS_Windows
    ; Get gadget font attributes
    Protected finfo.LOGFONT, height
    ;
    If GetObject_(FontID, SizeOf(LOGFONT), @finfo.LOGFONT)
      If finfo\lfItalic
        FontAttributes$ = "Italic"
      EndIf
      Select finfo\lfWeight
        Case #FW_THIN
          FontAttributes$ + ", Thin"
        Case #FW_EXTRALIGHT
          FontAttributes$ + ", ExtraLight"
        Case #FW_ULTRALIGHT
          FontAttributes$ + ", UltraLight"
        Case #FW_LIGHT
          FontAttributes$ + ", Light"
        Case #FW_MEDIUM
          FontAttributes$ + ", Medium"
        Case #FW_SEMIBOLD
          FontAttributes$ + ", DemiBold"
        Case #FW_DEMIBOLD
          FontAttributes$ + ", DemiBold"
        Case #FW_BOLD
          FontAttributes$ + ", Bold"
        Case #FW_EXTRABOLD
          FontAttributes$ + ", ExtraBold"
        Case #FW_ULTRABOLD
          FontAttributes$ + ", ExtraLight"
        Case #FW_HEAVY
          FontAttributes$ + ", Heavy"
        Case #FW_BLACK
          FontAttributes$ + ", Black"
      EndSelect
      If finfo\lfUnderline
        FontAttributes$ + ", Underline"
      EndIf
      If finfo\lfStrikeOut
        FontAttributes$ + ", StrikeOut"
      EndIf
    EndIf
  CompilerEndIf
  ;
  If Left(FontAttributes$, 2) = ", "
    FontAttributes$ = Mid(FontAttributes$, 3)
  EndIf
  ProcedureReturn FontAttributes$
  ;
EndProcedure
;
Procedure AttributesTextDescriptionToNumAttributes(FontStyle$)
  ;
  Protected Style = 0
  ;
  If FindString(FontStyle$, "Italic")
    Style = #PB_Font_Italic
  EndIf
  If FindString(FontStyle$, "Bold")
    Style | #PB_Font_Bold
  EndIf
  If FindString(FontStyle$, "StrikeOut")
    Style | #PB_Font_StrikeOut
  EndIf
  If FindString(FontStyle$, "Underline")
    Style | #PB_Font_Underline
  EndIf
  ProcedureReturn Style
EndProcedure
;
Procedure.s NumAttributesToAttributesTextDescription(FontStyle)
  ;
  Protected Style$ = ""
  ;
  If FontStyle & #PB_Font_Italic
    Style$ = ", Italic"
  EndIf
  If FontStyle & #PB_Font_Bold
    Style$ + ", Bold"
  EndIf
  If FontStyle & #PB_Font_StrikeOut
    Style$ + ", StrikeOut"
  EndIf
  If FontStyle & #PB_Font_Underline
    Style$ + ", Underline"
  EndIf
  ProcedureReturn Style$
EndProcedure
;
Procedure GetFontFromDescription(FontDescription$)
  ProcedureReturn LoadFont(#PB_Any, Trim(StringField(FontDescription$, 1, ",")), Val(Trim(StringField(FontDescription$, 2, ","))), AttributesTextDescriptionToNumAttributes(FontDescription$))
EndProcedure
;
Procedure.s GetDescriptionFromFont(Font)
  Protected FontDescription$ = GetFontName(Font) + ", " + GetFontSize(Font)
  Protected Attributes$ = GetFontAttributes(Font)
  If Attributes$
    FontDescription$ + ", " + Attributes$
  EndIf
  ProcedureReturn FontDescription$
EndProcedure
; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; CursorPosition = 2
; Folding = --
; EnableXP
; DPIAware