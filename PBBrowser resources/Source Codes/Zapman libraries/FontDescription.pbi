;
;****************************************************************
;
;                   Font Description Library
;                  v2 - By Zapman - March 2025
;
; This file should be saved under the name "FontDescription.pbi".
;
; It is used by "ZapmanFontRequester.pbi" and "TOM_Functions.pbi".
;
;  This library offers a set of functions to manage a font from
;         a text-based description of it, or vice-versa.
;
;   Textual descriptions of coding objects offer significant
;   advantages for coders: easy memorization and manipulation
;     of parameters and commands, immediate understanding of
;      returned results, the absence of type issues, etc.
;   But they also pose problems and risks: because they do not
;   require the coder to be strictely rigorous in presenting
;     parameters, they can lead to errors that are sometimes
;  difficult to identify. Furthermore, a very detailed analysis
;    of textual commands must be performed to authorize and
;    correctly interpret the possible syntaxes for a command
;                     or its parameters.
;
; The main functions of this library are:
; • LoadFontFromDescription(), which loads a font from a description
;   such as "Segoe UI, 9, Bold" or "Name(Segoe UI), Size(9), Bold".
; • GetFontDescription(), which returns a text description of an existing font.
; • GetFontName(), which returns the name of an existing font.
; • GetFontSize(), which returns the size of an existing font.
; • GetFontStyle(), which returns a style text description of an existing font.
; • NormalizeFontDescription(), which normalizes a simple description
;   and returns a tagged description to avoid any ambiguity in interpretation.
;   For example: "Arial, 12, RGB(255, 0, 0)" will be transformed into:
;   "Name(Arial), Size(12), ForeColor($0000FF)".
; • SimplifyFontDescription(), which generates a simplified description from
;   a normalized description.
; • GetAttributesDescriptionFromNumStyle(), which returns a text description of
;   the styles from a numeric combination of them.
; • GetNumStyleFromAttributesDescription(), which returns a numeric combination
;   of the styles from a text description of them.
;
; Most of the following functions are designed to work with the Microsoft TOM (Text Object Model)
; wich allows many more possibilities to describe a font and its variations than it
; can be done in PureBasic with the LoadFont() and FontRequester native functions.
; i.e.: • The font size is an integer for LoadFont(), while it is decimal for TOM;
;       • They are only four possible styles for LoadFont() (Bold, Italic, Underline, StrikeOut)
;         while they are many more with TOM (Emboss, Engrave and Outline, for example);
;       • TOM offers the possibility to define a back color for the text;
;       • etc.
;
;****************************************************************
;
Global FD_KnownFontKeywords$ = ",Bold,Italic,Emboss,AllCaps,SmallCaps,Engrave,Shadow,OutLine,Underline"
FD_KnownFontKeywords$ + ",StrikeThrough,StrikeOut,Hidden,Protected,Size,Spacing,Position,Kerning,BackColor"
FD_KnownFontKeywords$ + ",ForeColor,Weight,Style,Name,FontName,Font,"
FD_KnownFontKeywords$ = LCase(FD_KnownFontKeywords$)
;
Procedure.s FD_NormalizeStringSyntaxe(FontDescription$)
  ;
  Protected mStyle$
  ;
  Repeat
    mStyle$ = FontDescription$
    FontDescription$ = ReplaceString(FontDescription$, " ,", ",")
    FontDescription$ = ReplaceString(FontDescription$, ", ", ",")
    FontDescription$ = ReplaceString(FontDescription$, ",,", ",")
    FontDescription$ = ReplaceString(FontDescription$, ";", ",")
    FontDescription$ = ReplaceString(FontDescription$, " (", "(")
    FontDescription$ = ReplaceString(FontDescription$, "( ", "(")
    FontDescription$ = ReplaceString(FontDescription$, ") ", ")")
    FontDescription$ = ReplaceString(FontDescription$, " )", ")")
    FontDescription$ = ReplaceString(FontDescription$, ": ", "(")
    FontDescription$ = ReplaceString(FontDescription$, " :", "(")
    FontDescription$ = ReplaceString(FontDescription$, ":", "(")
  Until mStyle$ = FontDescription$
  ProcedureReturn LCase(Trim(FontDescription$))
EndProcedure
;
Procedure.s NormalizeFontDescription(FontDescription$)
  ;
  ; Try to interpretate each possible manners to describe a font,
  ; and normalize its description.
  ;
  ; The classical description order is: FontName, Size, ForeColor, BackColor, Styles
  ; 'Style' can be a succession of comma separated values as Bold, Italic, Underline, etc.
  ;
  ; • The first value of 'FontDescription$' is supposed to be the font name, the second is supposed  to be the size,
  ;   the optional third is supposed to be the forecolor, and the optional fourth is supposed to be the backcolor.
  ;
  ; • The normalized description form is: FontName(Value), Size(Value), ForeColor(Value), BackColor(Value), Bold, Italic, Underline, etc."
  ;   but many less formal declarations are accepted in entry, as "Arial, 12, Bold" (colors are optionnal),
  ;   or "Arial, size:12, BackColor:RGB(255,0,0), Bold"
  ;
  ;   For example, "Segoe UI, 9, Italic" will be transformed in "name(Segoe UI),size(9), italic".
  ;
  ; • Instead of commas, the different values can be separated by ";".
  ;
  ; • Colors given as RGB will be interpreted.
  ; • A color given without 'background' type is considered by default as 'forecolor'.
  ; • Hexadecimal values are accepted if preceeded by "$" or "0x"
  ;
  ; The operation is designed to be reversible by calling SymplifyFontDescription() with
  ; a normalized font description as parameter.
  ;
  Protected ct, StyleUnit$, mStyleUnit$, pp, mpp, pp2, RGBColor$, Color, ColorType$
  Protected NameSet, ForeColorSet, BackColorSet, SizeSet, posDep, Color$, cont
  ;
  FontDescription$ = FD_NormalizeStringSyntaxe(FontDescription$)
  ;
  FontDescription$ = ReplaceString(FontDescription$, "font(", "name(")
  FontDescription$ = ReplaceString(FontDescription$, "fontname(", "name(")
  FontDescription$ = ReplaceString(FontDescription$, "0x", "$")
  ;
  ; Evaluate RGB colors:
  ;
  posDep = 0
  Repeat
    posDep + 1
    pp = FindString(FontDescription$, "rgb", posDep)
    If pp
      mpp = pp
      pp + 3
      While Mid(FontDescription$, pp, 1) = " "
        pp + 1
      Wend
      If Mid(FontDescription$, pp, 1) = "("
        pp + 1
        pp2 = pp
        For ct = 1 To 3
          pp2 = FindString(FontDescription$, ",", pp2 + 1)
          If pp2 = 0
            pp2 = FindString(FontDescription$, ")", pp2 + 1)
            If pp2 = 0
              pp2 = Len(FontDescription$) + 1
            EndIf
            Break
          EndIf
        Next
        mStyleUnit$ = Mid(FontDescription$, mpp, pp2 - mpp)
        RGBColor$ = Mid(FontDescription$, pp, pp2 - pp)
        Color = RGB(Val(StringField(RGBColor$, 1, ",")), Val(StringField(RGBColor$, 2, ",")), Val(StringField(RGBColor$, 3, ",")))
        FontDescription$ = ReplaceString(FontDescription$, mStyleUnit$, "$" + Hex(Color))
      EndIf
    EndIf
  Until pp = 0
  ;
  ; Look for explicite values declarations:
  ct = 1
  Repeat
    StyleUnit$ = StringField(FontDescription$, ct, ",")
    If StyleUnit$
      If FindString(StyleUnit$, "size")
        SizeSet = 1
      ElseIf FindString(StyleUnit$, "name")
        NameSet = 1
      ElseIf FindString(StyleUnit$, "forecolor")
        ForeColorSet = 1
      ElseIf FindString(StyleUnit$, "backcolor")
        BackColorSet = 1
      EndIf
    EndIf
    ct + 1
  Until StyleUnit$ = ""
  ;
  ; Normalize colors values:
  ct = 1
  Repeat
    StyleUnit$ = StringField(FontDescription$, ct, ",")
    If StyleUnit$
      pp = FindString(StyleUnit$, "color")
      If pp
        mStyleUnit$ = StyleUnit$
        mpp = pp
        pp + 5
        While Mid(StyleUnit$, pp, 1) = " "
          pp + 1
        Wend
        If Mid(StyleUnit$, pp, 1) = "("
          pp + 1
          pp2 = FindString(StyleUnit$, ")", pp)
          If pp2 = 0
            pp2 = Len(StyleUnit$) + 1
          EndIf
          Color$ = Trim(Mid(StyleUnit$, pp, pp2 - pp))
          If FindString(StyleUnit$, "backcolor")
            ColorType$ = "backcolor"
            BackColorSet = 1
          Else
            ColorType$ = "forecolor"
            ForeColorSet = 1
          EndIf
          If Val(Color$) = 0 And Val("$" + Color$)
            StyleUnit$ = ColorType$ + "($" + Color$ + ")"
          ElseIf Str(Val(Color$)) = Color$ Or Left(Color$, 1) = "$"
            StyleUnit$ = ColorType$ + "($" + Hex(Val(Color$)) + ")"
          EndIf
          FontDescription$ = ReplaceString(FontDescription$, mStyleUnit$, StyleUnit$)
        EndIf
      EndIf
      ; Add the last parenthesys if it is missing:
      If FindString(StyleUnit$, "(") And FindString(StyleUnit$, ")") = 0
        FontDescription$ = ReplaceString(FontDescription$, StyleUnit$, StyleUnit$ + ")")
        FontDescription$ = ReplaceString(FontDescription$, "))", ")")
      EndIf
    EndIf
    ct + 1
  Until StyleUnit$ = ""
  ;
  cont = 1
  pp = 1
  For ct = 1 To 4
    pp2 = FindString(FontDescription$, ",", pp)
    If pp2 = 0
      pp2 = Len(FontDescription$) + 1
      cont = 0
    EndIf
    StyleUnit$ = Mid(FontDescription$, pp, pp2 - pp)
    ; Look at the first parameter to see if it can be the font name:
    If ct = 1 And NameSet = 0
      If FindString(StyleUnit$, "(") = 0
        If FindString(LCase(FD_KnownFontKeywords$), StyleUnit$) = 0
          FontDescription$ = "name(" + StyleUnit$ + ")" + Mid(FontDescription$, pp2)
          pp2 + 6 ; 6 is the lenght of added caracteres (name()).
        EndIf
      EndIf
    EndIf
    ; Look at the second parameter to see if it can be the font size:
    If ct = 2 And SizeSet = 0
      If FindString(StyleUnit$, "(") = 0
        If StrF(ValF(StyleUnit$)) = StyleUnit$ Or Str(Val(StyleUnit$)) = StyleUnit$ 
          FontDescription$ = Left(FontDescription$, pp - 1) + "size(" + StyleUnit$ + ")" + Mid(FontDescription$, pp2)
          pp2 + 6 ; 6 is the lenght of added caracteres (size()).
        EndIf
      EndIf
    EndIf
    ; Look at the third parameter to see if it can be the font forecolor:
    If ct = 3 And ForeColorSet = 0
      If FindString(StyleUnit$, "(") = 0
        If Val(StyleUnit$) Or StyleUnit$ = "0" Or Left(StyleUnit$, 1) = "$"
          FontDescription$ = Left(FontDescription$, pp - 1) + "forecolor(" + StyleUnit$ + ")" + Mid(FontDescription$, pp2)
          pp2 + 10 ; 10 is the lenght of added caracteres (forecolor()).
        EndIf
      EndIf
    EndIf
    ; Look at the fourth parameter to see if it can be the font backcolor:
    If ct = 4 And BackColorSet = 0
      If FindString(StyleUnit$, "(") = 0
        If Val(StyleUnit$) Or StyleUnit$ = "0" Or Left(StyleUnit$, 1) = "$"
          FontDescription$ = Left(FontDescription$, pp - 1) + "backcolor(" + StyleUnit$ + ")" + Mid(FontDescription$, pp2)
          pp2 + 10 ; 10 is the lenght of added caracteres (backcolor()).
        EndIf
      EndIf
    EndIf
    If cont
      pp = pp2 + 1
    Else
      Break
    EndIf
  Next
  ;
  ProcedureReturn FontDescription$
EndProcedure
;
Procedure.s FD_ExtractParameter(FontDescription$, ParameterName$, Silence = 0)
  ;
  ; This procedure, retrieves the parameter in parenthesis
  ; that follows 'ParameterName$' in the 'FontDescription$' string.
  ;
  Protected pa, pas, limp
  ;
  pa = FindString(FontDescription$, LCase(ParameterName$), 1)
  If pa
    pa + Len(ParameterName$)
    If PeekC(@FontDescription$ + (pa - 1) * SizeOf(CHARACTER)) = Asc("(")
      pa + 1
    EndIf
    pas = pa
    If PeekC(@FontDescription$ + (pa - 1) * SizeOf(CHARACTER)) <> Asc(",")
      limp = Len(FontDescription$)
      While pa <= limp And PeekC(@FontDescription$ + (pa - 1) * SizeOf(CHARACTER)) <> Asc(")")
        pa + 1
      Wend
    EndIf
    ProcedureReturn PeekS(@FontDescription$ + (pas - 1) * SizeOf(CHARACTER), pa - pas)
  ElseIf Silence = 0
    MessageRequester("Error", "Error while extracting parameter:  Wrong parameter name -> " + ParameterName$ + #CR$ + FontDescription$)
  EndIf
EndProcedure
;
Macro FD_CompleteSimpleDescription(StyleName, Description)
  If FindString(FontDescription$, StyleName)
    parameter$ = FD_ExtractParameter(FontDescription$, StyleName)
    If parameter$ <> "none" And parameter$ <> "0" And parameter$ <> "default"
      Description + ", " + UCase(Left(StyleName, 1)) + Mid(StyleName, 2)
    EndIf
  EndIf
EndMacro
;
Procedure.s SymplifyFontDescription(FontDescription$)
  ;
  ; From a normalized type description (as "Name(Segoe UI),Size(9),ForeColor($A05080),Bold")
  ; This function will return a simplified font description as
  ; "Segoe UI, 9, ForeColor:$A05080, Bold".
  ;
  ; The operation is designed to be reversible by calling NormalizeFontDescription() with
  ; a simplified font description as parameter.
  ;
  Protected parameter$, NFontDescription$
  ;
  ; First normalize the string to eventually interpret all its parameters:
  FontDescription$ = NormalizeFontDescription(FontDescription$)
  ;
  ; Then simplify the string:
  If FindString(FontDescription$, "name(")
    parameter$ = FD_ExtractParameter(FontDescription$, "name")
    If parameter$ And parameter$ <> "default"
      NFontDescription$ = UCase(Left(parameter$, 1)) + Mid(parameter$, 2)
    EndIf
  Else
    MessageRequester("Error", "Error while examining the font description:  font name is missing -> " + #CR$ + FontDescription$)
    ProcedureReturn
  EndIf
  If FindString(FontDescription$, "size(")
    parameter$ = FD_ExtractParameter(FontDescription$, "size")
    If parameter$ And parameter$ <> "default"
      NFontDescription$ + ", " + parameter$
    EndIf
  EndIf
  If FindString(FontDescription$, "forecolor(")
    parameter$ = FD_ExtractParameter(FontDescription$, "forecolor")
    If parameter$ And parameter$ <> "default"
      NFontDescription$ + ", ForeColor:" + UCase(parameter$)
    EndIf
  EndIf
  If FindString(FontDescription$, "backcolor(")
    parameter$ = FD_ExtractParameter(FontDescription$, "backcolor")
    If parameter$ And parameter$ <> "default"
      NFontDescription$ + ", BackColor:" + UCase(parameter$)
    EndIf
  EndIf
  ;
  ; The following is designed to be compatible with the font descriptions
  ; used for TOM (Text Object Model).
  FD_CompleteSimpleDescription("bold", NFontDescription$)
  FD_CompleteSimpleDescription("italic", NFontDescription$)
  FD_CompleteSimpleDescription("emboss", NFontDescription$)
  FD_CompleteSimpleDescription("allcaps", NFontDescription$)
  FD_CompleteSimpleDescription("smallcaps", NFontDescription$)
  FD_CompleteSimpleDescription("engrave", NFontDescription$)
  FD_CompleteSimpleDescription("shadow", NFontDescription$)
  FD_CompleteSimpleDescription("outline", NFontDescription$)
  FD_CompleteSimpleDescription("strikethrough", NFontDescription$)
  FD_CompleteSimpleDescription("strikeout", NFontDescription$)
  FD_CompleteSimpleDescription("subscript", NFontDescription$)
  FD_CompleteSimpleDescription("superscript", NFontDescription$)
  FD_CompleteSimpleDescription("hidden", NFontDescription$)
  FD_CompleteSimpleDescription("protected", NFontDescription$)
  ;
  FD_CompleteSimpleDescription("underline", NFontDescription$)
  ;
  If Left(NFontDescription$, 2) = ", "
    NFontDescription$ = Mid(NFontDescription$, 3)
  EndIf
  ;
  ProcedureReturn NFontDescription$
  ;
EndProcedure
;
Procedure GetNumStyleFromAttributesDescription(StyleDescription$)
  ;
  ; Return the style (num value) of the style/font described by 'StyleDescription$'.
  ;
  Protected FontStyle = 0
  StyleDescription$ = LCase(StyleDescription$)
  ;
  If FindString(StyleDescription$, "italic")
    FontStyle = #PB_Font_Italic
  EndIf
  If FindString(StyleDescription$, "bold")
    FontStyle | #PB_Font_Bold
  EndIf
  If FindString(StyleDescription$, "strike")
    FontStyle | #PB_Font_StrikeOut
  EndIf
  If FindString(StyleDescription$, "underline")
    FontStyle | #PB_Font_Underline
  EndIf
  ProcedureReturn FontStyle
EndProcedure
;
Procedure.s GetAttributesDescriptionFromNumStyle(FontStyle)
  ;
  ; Return the style (num value) of the style/font described by 'StyleDescription$'.
  ;
  Protected StyleDescription$ = ""
  ;
  If FontStyle & #PB_Font_Italic
    StyleDescription$ + "Italic, "
  EndIf
  If FontStyle & #PB_Font_Bold
    StyleDescription$ + "Bold, "
  EndIf
  If FontStyle & #PB_Font_StrikeOut
    StyleDescription$ + "StrikeOut, "
  EndIf
  If FontStyle & #PB_Font_Underline
    StyleDescription$ + "Underline, "
  EndIf
  ProcedureReturn Left(StyleDescription$, Len(StyleDescription$) - 2)
EndProcedure
;
Procedure LoadFontFromDescription(FontDescription$)
  ;
  ; Return a font handle from a textual font description.
  ;
  Protected parameter$
  ;
  FontDescription$ = NormalizeFontDescription(FontDescription$)
  ;
  If FindString(FontDescription$, "name(")
    parameter$ = FD_ExtractParameter(FontDescription$, "name")
    If parameter$ And parameter$ <> "default"
      Protected FontName$ = parameter$
    EndIf
  EndIf
  If FindString(FontDescription$, "size(")
    parameter$ = FD_ExtractParameter(FontDescription$, "size")
    If parameter$ And parameter$ <> "default"
      Protected FontSize = Val(parameter$)
    EndIf
  EndIf
  ;
  Protected FontStyle = GetNumStyleFromAttributesDescription(FontDescription$)
  ;
  ProcedureReturn LoadFont(#PB_Any, FontName$, FontSize, FontStyle)
  ;
EndProcedure
;
Procedure.s GetFontName(FontID)
  ;
  ; Return the system name of the font defined by 'FontID'.
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
    ;    
    ; Create a HDC to get the font info:
    Protected hdc = GetDC_(#Null)
    Protected oldFont = SelectObject_(hdc, FontID)
    ;
    ; Get the real font name:
    FontName$ = Space(#LF_FACESIZE) ; Maximum size for font name.
    GetTextFace_(hdc, #LF_FACESIZE, @FontName$)
    If FontName$ = "" : FontName$ = "Arial" : EndIf
    ;
    ; Clean up:
    SelectObject_(hdc, oldFont)
    ReleaseDC_(#Null, hdc)
  CompilerEndIf
  ;
  ProcedureReturn FontName$
  ;
EndProcedure
;
Procedure.f GetFontSize(FontID)
  ;
  ; Return the size of the font defined by 'FontID'.
  ;
  Protected FpointSize.f
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
    Protected finfo.LOGFONT
    Protected hdc = GetDC_(#Null) ; Obtenir le contexte de périphérique
    If GetObject_(FontID, SizeOf(LOGFONT), @finfo.LOGFONT)
      FpointSize = -finfo\lfHeight * 72 / GetDeviceCaps_(hdc, #LOGPIXELSY)
    EndIf
    ProcedureReturn FpointSize
  CompilerEndIf
  ;
EndProcedure
;
Procedure.s GetFontStyles(FontID)
  ;
  Protected lf.LOGFONT, Purelf.LOGFONT
  Protected StyleDescription$, FontName$, PureFont
  ;
  If IsFont(FontID)
    FontID = FontID(FontID)
  EndIf
  ;
  ; Récupérer les infos de la police
  If GetObject_(FontID, SizeOf(LOGFONT), @lf)
    FontName$ = PeekS(@lf\lfFaceName, -1, #PB_Unicode)
    PureFont = LoadFont(#PB_Any, FontName$, 9)
    If GetObject_(FontID(PureFont), SizeOf(LOGFONT), @Purelf)
      ; Compare the weight of the font to the normal weight:
      If lf\lfWeight > Purelf\lfWeight
        StyleDescription$ = "Bold, "
      ElseIf lf\lfWeight < Purelf\lfWeight
        StyleDescription$ = "Light, "
      EndIf
    EndIf
    FreeFont(PureFont)
    ;
    If lf\lfItalic
      StyleDescription$ + "Italic, "
    EndIf
    ;
    If lf\lfUnderline
      StyleDescription$ + "Underline, "
    EndIf
    ;
    If lf\lfStrikeOut
      StyleDescription$ + "StrikeOut, "
    EndIf
    ;
    ProcedureReturn Left(StyleDescription$, Len(StyleDescription$) - 2)
  EndIf
EndProcedure
;
Procedure.s GetFontDescription(FontID)
  Protected FontDescription$ = GetFontName(FontID) + ", " + GetFontSize(FontID)
  Protected StyleDescription$ = GetFontStyles(FontID)
  If StyleDescription$
    FontDescription$ + ", " + StyleDescription$
  EndIf
  ProcedureReturn FontDescription$
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  ; The following won't run when this file is used as 'Included'.
  ;
  Font = LoadFontFromDescription("arial,12,bold")
  Description$ = GetFontDescription(Font)
  FreeFont(Font)
  Debug "Simple description:            " + Description$
  Debug "Normalized description:    " + NormalizeFontDescription(Description$)
  ;
CompilerEndIf

; IDE Options = PureBasic 6.21 Beta 2 (Windows - x86)
; CursorPosition = 285
; FirstLine = 265
; Folding = -r-
; EnableXP
; DPIAware