;
;****************************************************************
;
;                   Font Description Library
;                  v3 - By Zapman - Nov. 2026
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
;   require the coder to be strictly rigorous in presenting
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
; • GetFontStyles(), which returns a style text description of an existing font.
; • NormalizeFontDescription(), which normalizes a simple or mixed-case description
;   and returns a tagged description to avoid any ambiguity in interpretation.
;   For example: "Arial, 12, RGB(255, 0, 0)" or "Name(Arial), Size(12), ForeColor($0000FF)"
;   will be transformed into: "name(Arial),size(12),forecolor($ff)".
; • SimplifyFontDescription(), which generates a simplified description from
;   a normalized description.
; • GetAttributesDescriptionFromNumStyle(), which returns a text description of
;   the styles from a numeric combination of them.
; • GetNumStyleFromAttributesDescription(), which returns a numeric combination
;   of the styles from a text description of them.
;
; Most of the following functions are designed to work with the Microsoft TOM (Text Object Model),
; which allows many more possibilities to describe a font and its variations than
; PureBasic's native LoadFont() and FontRequester() functions.
; e.g.: • The font size is an integer for LoadFont(), while it can be decimal for TOM;
;       • There are only four possible styles for LoadFont() (Bold, Italic, Underline, StrikeOut),
;         while there are many more with TOM (Emboss, Engrave, and Outline, for example);
;       • TOM offers the possibility to define a background color for the text;
;       • etc.
;
;****************************************************************
;
; Store all recognized font attributes and tags in lowercase, framed by leading and trailing
; commas so callers can test exact token membership via "," + Token$ + "," without partial matches:
Global FD_KnownFontKeywords$ = ",bold,italic,light,emboss,allcaps,smallcaps,engrave,shadow,outline,underline" +
                               ",strikethrough,strikeout,subscript,superscript,hidden,protected,size,spacing" +
                               ",position,kerning,backcolor,forecolor,weight,style,name,fontname,font,scaling,"
;
Procedure.s FD_NormalizeStringSyntax(FontDescription$, PreserveCase = #False)
  ;
  ; Role: Standardizes delimiters and strips redundant spaces around separators and parentheses.
  ; Parameters: PreserveCase - Keeps original capitalization when #True; converts to lowercase when #False.
  ; Output: Normalized compact command string.
  ;
  Protected mStyle$
  ;
  ; Convert alternative delimiters once before the space-stripping loop
  FontDescription$ = ReplaceString(FontDescription$, ";", ",")
  FontDescription$ = ReplaceString(FontDescription$, ":", "(")
  ;
  Repeat
    mStyle$ = FontDescription$
    FontDescription$ = ReplaceString(FontDescription$, " ,", ",")
    FontDescription$ = ReplaceString(FontDescription$, ", ", ",")
    FontDescription$ = ReplaceString(FontDescription$, ",,", ",")
    FontDescription$ = ReplaceString(FontDescription$, " (", "(")
    FontDescription$ = ReplaceString(FontDescription$, "( ", "(")
    FontDescription$ = ReplaceString(FontDescription$, ") ", ")")
    FontDescription$ = ReplaceString(FontDescription$, " )", ")")
  Until mStyle$ = FontDescription$
  ;
  FontDescription$ = Trim(Trim(FontDescription$), ",")
  If PreserveCase = #False
    FontDescription$ = LCase(FontDescription$)
  EndIf
  ProcedureReturn FontDescription$
EndProcedure
;
CompilerIf Not Defined(FD_TokenEntry, #PB_Structure)
  Structure FD_TokenEntry
    Tag.s
    Val.s
  EndStructure
CompilerEndIf
;
Procedure.s NormalizeFontDescription(FontDescription$)
  Protected NewList Tokens.FD_TokenEntry()
  Protected *Char.Character
  Protected StyleUnit$, Tag$, Val$, RGBColor$, Color$
  Protected pp, pp2, pp3, mpp, ct
  Protected NameSet, SizeSet, ForeColorSet, BackColorSet, DigitCount, DotCount, HexFlag
  ;
  ; Role: Interprets all possible ways to describe a font and normalizes its description into a tagged format.
  ; Output: Canonical string formatted as "name(Value),size(Value),forecolor($hex),backcolor($hex),styles..."
  ;         The operation is designed to be reversible by calling SimplifyFontDescription() with a normalized
  ;         font description as parameter.
  ;
  ; Supported syntax and interpretation rules:
  ; • The classical description order is: FontName, Size, ForeColor, BackColor, Styles.
  ;   'Styles' can be a succession of comma-separated values such as Bold, Italic, Underline, etc.
  ; • By default, the first untagged value of 'FontDescription$' is interpreted as the font name (unless it is
  ;   a number, a color, or a style keyword), the next numeric value as the size, the optional third untagged
  ;   value as the forecolor, and the optional fourth untagged value as the backcolor.
  ; • The normalized description form is: "name(Value),size(Value),forecolor($hex),backcolor($hex),bold,italic,..."
  ;   Many less formal declarations are accepted as input, such as "Arial, 12, Bold" (colors are optional),
  ;   "Name(Arial), Size(12), ForeColor($0000FF)", "SetForeColor(23756)", or "Arial, size:12, BackColor:RGB(255,0,0), Bold".
  ;   Leading "Set..." and "Font..." prefixes on tags or known style keywords are normalized automatically.
  ; • Instead of commas, the different values can be separated by ";".
  ; • Colors given as RGB(r,g,b) are evaluated and converted to hexadecimal.
  ; • A color given without a 'backcolor' tag is considered by default as 'forecolor' (and a second untagged
  ;   color as 'backcolor').
  ; • Hexadecimal color values are accepted when preceded by "$" or "0x" (or given as hex digits inside a color tag).
  ;
  FontDescription$ = FD_NormalizeStringSyntax(FontDescription$, #True)
  If FontDescription$ = ""
    ProcedureReturn ""
  EndIf
  ;
  ; Evaluate RGB(r,g,b) colors into "$hex" before splitting on commas:
  pp = 1
  Repeat
    pp = FindString(FontDescription$, "rgb(", pp, #PB_String_NoCase)
    If pp
      ; Locate up to the 3rd comma to bound unclosed "RGB:r,g,b" declarations:
      pp2 = pp + 3
      For ct = 1 To 3
        pp2 = FindString(FontDescription$, ",", pp2 + 1)
        If pp2 = 0
          pp2 = Len(FontDescription$) + 1
          Break
        EndIf
      Next
      mpp = FindString(FontDescription$, ")", pp + 4)
      If mpp > 0 And mpp < pp2
        ; Closed parenthesis belongs to RGB(...): extract inner components and consume ")":
        RGBColor$ = Mid(FontDescription$, pp + 4, mpp - pp - 4)
        pp2 = mpp + 1
      Else
        ; Unclosed colon syntax: stop right before the 3rd comma or at the end of the string:
        RGBColor$ = Mid(FontDescription$, pp + 4, pp2 - pp - 4)
      EndIf
      Color$ = "$" + LCase(Hex(RGB(Val(StringField(RGBColor$, 1, ",")) & $FF, Val(StringField(RGBColor$, 2, ",")) & $FF, Val(StringField(RGBColor$, 3, ",")) & $FF)))
      FontDescription$ = Left(FontDescription$, pp - 1) + Color$ + Mid(FontDescription$, pp2)
      pp + Len(Color$)
    EndIf
  Until pp = 0
  ;
  ; Split tokens on top-level commas and normalize all tag prefixes, aliases, and style keywords in a single pass:
  pp = 1
  While pp <= Len(FontDescription$)
    pp2 = FindString(FontDescription$, ",", pp)
    mpp = FindString(FontDescription$, "(", pp)
    If pp2 > 0 And mpp > 0 And mpp < pp2
      ; Preserve internal commas when "(" is closed before any subsequent "(" (multi-arg tag vs unclosed "tag:val"):
      ct  = FindString(FontDescription$, ")", mpp + 1)
      pp3 = FindString(FontDescription$, "(", mpp + 1)
      If ct > 0 And (pp3 = 0 Or ct < pp3)
        pp2 = FindString(FontDescription$, ",", ct + 1)
      EndIf
    EndIf
    If pp2 = 0
      pp2 = Len(FontDescription$) + 1
    EndIf
    If pp2 > pp
      AddElement(Tokens())
      If mpp > 0 And mpp < pp2
        Tag$ = LCase(Mid(FontDescription$, pp, mpp - pp))
        If Mid(FontDescription$, pp2 - 1, 1) = ")"
          Val$ = Mid(FontDescription$, mpp + 1, pp2 - mpp - 2)
        Else
          Val$ = Mid(FontDescription$, mpp + 1, pp2 - mpp - 1)
        EndIf
        ; Canonicalize "set..." / "font..." tag prefixes and color aliases once so downstream logic sees only standard tags:
        If Len(Tag$) > 3 And Left(Tag$, 3) = "set"
          Tag$ = Mid(Tag$, 4)
        EndIf
        If Tag$ = "font"
          Tag$ = "name"
        ElseIf Len(Tag$) > 4 And Left(Tag$, 4) = "font"
          Tag$ = Mid(Tag$, 5)
        EndIf
        If Tag$ = "color" Or Tag$ = "frontcolor" Or Tag$ = "textcolor"
          Tag$ = "forecolor"
        ElseIf Tag$ = "backgroundcolor"
          Tag$ = "backcolor"
        EndIf
        If Tag$ = "name"
          NameSet = #True
        ElseIf Tag$ = "size"
          SizeSet = #True
        ElseIf Tag$ = "forecolor"
          ForeColorSet = #True
        ElseIf Tag$ = "backcolor"
          BackColorSet = #True
        EndIf
        Tokens()\Tag = Tag$
        Tokens()\Val = Val$
      Else
        Val$ = Mid(FontDescription$, pp, pp2 - pp)
        Tag$ = LCase(Val$)
        ; Identify standalone style keywords (including "Set..." variants such as "SetBold") using "*" as a marker:
        If FindString(FD_KnownFontKeywords$, "," + Tag$ + ",")
          Tokens()\Tag = "*"
          Tokens()\Val = Tag$
        ElseIf Len(Tag$) > 3 And Left(Tag$, 3) = "set" And FindString(FD_KnownFontKeywords$, "," + Mid(Tag$, 4) + ",")
          Tokens()\Tag = "*"
          Tokens()\Val = Mid(Tag$, 4)
        Else
          Tokens()\Val = Val$
        EndIf
      EndIf
    EndIf
    pp = pp2 + 1
  Wend
  ;
  ; Classify remaining untagged values (colors, numeric sizes, or font name) and assemble the canonical string:
  FontDescription$ = ""
  ForEach Tokens()
    Tag$ = Tokens()\Tag
    Val$ = Tokens()\Val
    If Tag$ = "*"
      Tag$ = ""
    ElseIf Tag$ = ""
      If LCase(Left(Val$, 2)) = "0x"
        Val$ = "$" + Mid(Val$, 3)
      EndIf
      If Left(Val$, 1) = "$"
        ; Untagged hexadecimal or RGB-converted color defaults to forecolor, then backcolor:
        If ForeColorSet = #False
          Tag$ = "forecolor"
          ForeColorSet = #True
        ElseIf BackColorSet = #False
          Tag$ = "backcolor"
          BackColorSet = #True
        EndIf
      Else
        ; Inspect whether the untagged value is numeric (integer/decimal size, or fallback decimal color):
        *Char = @Val$
        DigitCount = 0
        DotCount = 0
        While *Char\c
          If *Char\c >= '0' And *Char\c <= '9'
            DigitCount + 1
          ElseIf *Char\c = '.'
            DotCount + 1
          Else
            DigitCount = 0
            Break
          EndIf
          *Char + SizeOf(Character)
        Wend
        If DigitCount > 0 And DotCount <= 1
          If SizeSet = #False
            Tag$ = "size"
            SizeSet = #True
          ElseIf DotCount = 0
            If ForeColorSet = #False
              Tag$ = "forecolor"
              ForeColorSet = #True
            ElseIf BackColorSet = #False
              Tag$ = "backcolor"
              BackColorSet = #True
            EndIf
          EndIf
        ElseIf NameSet = #False
          ; Treat the first non-keyword, non-numeric, non-color untagged value as the font name:
          Tag$ = "name"
          NameSet = #True
        Else
          Val$ = LCase(Val$)
        EndIf
      EndIf
    EndIf
    ;
    ; Format the normalized token and add the closing parenthesis if it was missing:
    If Tag$ = ""
      StyleUnit$ = Val$
    ElseIf Tag$ = "name"
      If LCase(Val$) = "default" Or LCase(Val$) = "none"
        Val$ = LCase(Val$)
      EndIf
      StyleUnit$ = "name(" + Val$ + ")"
    ElseIf Tag$ = "forecolor" Or Tag$ = "backcolor"
      ; Normalize color values ("$hex", "0xHex", decimal, or hex digits) while preserving keywords such as "default"
      If LCase(Left(Val$, 2)) = "0x"
        Val$ = "$" + Mid(Val$, 3)
      EndIf
      HexFlag = #False
      DigitCount = 0
      *Char = @Val$
      If *Char\c = '$'
        HexFlag = #True
        *Char + SizeOf(Character)
      EndIf
      While *Char\c
        If *Char\c >= '0' And *Char\c <= '9'
          DigitCount + 1
        ElseIf (*Char\c >= 'a' And *Char\c <= 'f') Or (*Char\c >= 'A' And *Char\c <= 'F')
          DigitCount + 1
          HexFlag = #True
        Else
          DigitCount = 0
          Break
        EndIf
        *Char + SizeOf(Character)
      Wend
      If DigitCount > 0
        If HexFlag And Left(Val$, 1) <> "$"
          Val$ = "$" + Val$
        EndIf
        Val$ = "$" + LCase(Hex(Val(Val$)))
      Else
        Val$ = LCase(Val$)
      EndIf
      StyleUnit$ = Tag$ + "(" + Val$ + ")"
    Else
      StyleUnit$ = Tag$ + "(" + LCase(Val$) + ")"
    EndIf
    ;
    If FontDescription$
      FontDescription$ + "," + StyleUnit$
    Else
      FontDescription$ = StyleUnit$
    EndIf
  Next
  ;
  ProcedureReturn FontDescription$
EndProcedure
;
Procedure.s FD_ExtractParameter(FontDescription$, ParameterName$, Silence = 0, *Found.Integer = #Null)
  ;
  ; Role: Retrieves the parameter enclosed in parentheses (or following a colon) that follows
  ;       'ParameterName$' in the 'FontDescription$' string.
  ; Parameters: Silence - When zero, displays an error MessageRequester if 'ParameterName$' is not found.
  ;             *Found  - Optional integer pointer receiving #True when 'ParameterName$' is present as a
  ;                       top-level attribute, or #False otherwise.
  ; Output: Extracted parameter string, or an empty string if 'ParameterName$' has no value or is absent.
  ;
  Protected *Char.Character, *Start.Character, *Scan.Character
  Protected ParamLen, ParenDepth, IsMatch, HasValue
  ;
  If *Found
    *Found\i = #False
  EndIf
  ParameterName$ = RTrim(Trim(ParameterName$), "(")
  ParamLen = Len(ParameterName$)
  If FontDescription$ <> "" And ParamLen > 0
    *Char = @FontDescription$
    While *Char\c
      ; Skip leading spaces and separators between top-level attributes
      While *Char\c = ' ' Or *Char\c = ',' Or *Char\c = ';'
        *Char + SizeOf(Character)
      Wend
      If *Char\c = 0
        Break
      EndIf
      ;
      ; Read the candidate attribute identifier at the top level
      *Start = *Char
      While *Char\c And *Char\c <> ',' And *Char\c <> ';' And *Char\c <> '(' And *Char\c <> ':' And *Char\c <> ' '
        *Char + SizeOf(Character)
      Wend
      IsMatch = Bool((*Char - *Start) / SizeOf(Character) = ParamLen And CompareMemoryString(*Start, @ParameterName$, #PB_String_NoCase, ParamLen) = #PB_String_Equal)
      While *Char\c = ' '
        *Char + SizeOf(Character)
      Wend
      If IsMatch And *Char\c <> '(' And *Char\c <> ':' And *Char\c <> ',' And *Char\c <> ';' And *Char\c <> 0
        ; Reject prefix matches on multi-word untagged font names such as "Arial Bold"
        IsMatch = #False
      EndIf
      ;
      HasValue = #False
      If *Char\c = '('
        HasValue = #True
        *Char + SizeOf(Character)
        *Start = *Char
        *Scan = *Char
        ParenDepth = 1
        While *Scan\c And ParenDepth > 0
          If *Scan\c = '('
            ParenDepth + 1
          ElseIf *Scan\c = ')'
            ParenDepth - 1
          EndIf
          If ParenDepth > 0
            *Scan + SizeOf(Character)
          EndIf
        Wend
        If ParenDepth = 0
          ; Parenthesis is properly closed: return inner value or advance past ')'
          If IsMatch
            If *Found
              *Found\i = #True
            EndIf
            ProcedureReturn Trim(PeekS(*Start, (*Scan - *Start) / SizeOf(Character)))
          EndIf
          *Char = *Scan + SizeOf(Character)
        EndIf
        ; If ParenDepth > 0 (unclosed '(' from colon conversion), fall through to stop at ',' or ';'
        ParenDepth = 0
      Else
        If *Char\c = ':'
          HasValue = #True
          *Char + SizeOf(Character)
        EndIf
        *Start = *Char
        ParenDepth = 0
      EndIf
      ;
      ; Advance to the next top-level separator while preserving commas inside balanced parentheses
      While *Char\c And (ParenDepth > 0 Or (*Char\c <> ',' And *Char\c <> ';'))
        If *Char\c = '('
          ParenDepth + 1
        ElseIf *Char\c = ')' And ParenDepth > 0
          ParenDepth - 1
        EndIf
        *Char + SizeOf(Character)
      Wend
      If IsMatch
        If *Found
          *Found\i = #True
        EndIf
        If HasValue
          ProcedureReturn Trim(PeekS(*Start, (*Char - *Start) / SizeOf(Character)))
        EndIf
        ProcedureReturn ""
      EndIf
    Wend ; End of top-level attribute scan
  EndIf
  ;
  If Silence = 0
    MessageRequester("Error", "Error while extracting parameter: Wrong parameter name -> " + ParameterName$ + #CR$ + FontDescription$)
    CallDebugger
  EndIf
  ProcedureReturn ""
EndProcedure
;
Macro FD_CompleteSimpleDescription(StyleName, Description, AlwaysKeepValue = #False)
  ;
  ; Role: Appends 'StyleName' (and its optional parameter) to 'Description' when active in 'FontDescription$'.
  ;
  parameter$ = FD_ExtractParameter(FontDescription$, StyleName, #True, @Found)
  If Found And parameter$ <> "none" And parameter$ <> "0" And parameter$ <> "default" And parameter$ <> "false"
    If AlwaysKeepValue = #False And (parameter$ = "" Or parameter$ = "1" Or parameter$ = "-1" Or parameter$ = "true")
      Description + ", " + StyleName
    ElseIf parameter$
      Description + ", " + StyleName + ":" + UCase(Left(parameter$, 1)) + Mid(parameter$, 2)
    EndIf
  EndIf
EndMacro
;
Procedure.s SimplifyFontDescription(FontDescription$)
  ;
  ; Role: Converts a normalized font description (such as "name(Segoe UI),size(9),forecolor($a05080),bold"
  ;       or "Name(Segoe UI),Size(9),ForeColor($A05080),Bold") into a simplified description
  ;       such as "Segoe UI, 9, ForeColor:$A05080, Bold".
  ;       The operation is designed to be reversible by calling NormalizeFontDescription() with
  ;       a simplified font description as parameter.
  ; Output: Simplified human-readable font description string (supports partial descriptions without a font name).
  ;
  Protected *Char.Character
  Protected parameter$, NFontDescription$
  Protected Found, CapitalizeNext
  ;
  ; First normalize the string to interpret all its parameters:
  FontDescription$ = NormalizeFontDescription(FontDescription$)
  If FontDescription$ = ""
    ProcedureReturn ""
  EndIf
  ;
  ; Then simplify the string (each active field prepends ", " uniformly):
  parameter$ = FD_ExtractParameter(FontDescription$, "name", #True, @Found)
  If Found And parameter$ And parameter$ <> "default"
    ; Capitalize the initial of each word after a space or a hyphen while preserving existing uppercase letters
    CapitalizeNext = #True
    *Char = @parameter$
    While *Char\c
      If *Char\c = ' ' Or *Char\c = '-'
        CapitalizeNext = #True
      ElseIf CapitalizeNext
        If *Char\c >= 'a' And *Char\c <= 'z'
          *Char\c - 32
        ElseIf *Char\c > 127
          *Char\c = Asc(UCase(Chr(*Char\c)))
        EndIf
        CapitalizeNext = #False
      EndIf
      *Char + SizeOf(Character)
    Wend
    NFontDescription$ + ", " + parameter$
  EndIf
  ;
  parameter$ = FD_ExtractParameter(FontDescription$, "size", #True, @Found)
  If Found And parameter$ And parameter$ <> "default"
    NFontDescription$ + ", " + parameter$
  EndIf
  ;
  parameter$ = FD_ExtractParameter(FontDescription$, "forecolor", #True, @Found)
  If Found And parameter$ And parameter$ <> "default"
    NFontDescription$ + ", ForeColor:" + UCase(parameter$)
  EndIf
  ;
  parameter$ = FD_ExtractParameter(FontDescription$, "backcolor", #True, @Found)
  If Found And parameter$ And parameter$ <> "default"
    NFontDescription$ + ", BackColor:" + UCase(parameter$)
  EndIf
  ;
  ; The following is designed to be compatible with the font descriptions
  ; used for TOM (Text Object Model):
  FD_CompleteSimpleDescription("Bold", NFontDescription$)
  FD_CompleteSimpleDescription("Italic", NFontDescription$)
  FD_CompleteSimpleDescription("Light", NFontDescription$)
  FD_CompleteSimpleDescription("Emboss", NFontDescription$)
  FD_CompleteSimpleDescription("AllCaps", NFontDescription$)
  FD_CompleteSimpleDescription("SmallCaps", NFontDescription$)
  FD_CompleteSimpleDescription("Engrave", NFontDescription$)
  FD_CompleteSimpleDescription("Shadow", NFontDescription$)
  FD_CompleteSimpleDescription("Outline", NFontDescription$)
  FD_CompleteSimpleDescription("StrikeThrough", NFontDescription$)
  FD_CompleteSimpleDescription("StrikeOut", NFontDescription$)
  FD_CompleteSimpleDescription("Subscript", NFontDescription$)
  FD_CompleteSimpleDescription("Superscript", NFontDescription$)
  FD_CompleteSimpleDescription("Hidden", NFontDescription$)
  FD_CompleteSimpleDescription("Protected", NFontDescription$)
  FD_CompleteSimpleDescription("Underline", NFontDescription$)
  FD_CompleteSimpleDescription("Spacing", NFontDescription$, #True)
  FD_CompleteSimpleDescription("Position", NFontDescription$, #True)
  FD_CompleteSimpleDescription("Kerning", NFontDescription$, #True)
  FD_CompleteSimpleDescription("Weight", NFontDescription$, #True)
  FD_CompleteSimpleDescription("Style", NFontDescription$, #True)
  ;
  ProcedureReturn Mid(NFontDescription$, 3)
EndProcedure
;
Procedure GetNumStyleFromAttributesDescription(StyleDescription$)
  ;
  ; Role: Returns the PureBasic numeric style combination (#PB_Font_Bold, #PB_Font_Italic, etc.)
  ;       matching the textual style or font description in 'StyleDescription$'.
  ;
  Protected parameter$, FontStyle = 0, Found, ct
  ;
  For ct = 1 To 5
    parameter$ = LCase(FD_ExtractParameter(StyleDescription$, StringField("bold,italic,strikeout,strikethrough,underline", ct, ","), #True, @Found))
    If Found And parameter$ <> "none" And parameter$ <> "0" And parameter$ <> "default" And parameter$ <> "false"
      Select ct
        Case 1 : FontStyle | #PB_Font_Bold
        Case 2 : FontStyle | #PB_Font_Italic
        Case 3, 4 : FontStyle | #PB_Font_StrikeOut
        Case 5 : FontStyle | #PB_Font_Underline
      EndSelect
    EndIf
  Next
  ProcedureReturn FontStyle
EndProcedure
;
Procedure.s GetAttributesDescriptionFromNumStyle(FontStyle)
  ;
  ; Role: Returns a comma-separated textual style description from a numeric combination
  ;       of PureBasic font style flags ('FontStyle').
  ;
  Protected StyleDescription$ = ""
  ;
  If FontStyle & #PB_Font_Bold
    StyleDescription$ + ", Bold"
  EndIf
  If FontStyle & #PB_Font_Italic
    StyleDescription$ + ", Italic"
  EndIf
  If FontStyle & #PB_Font_StrikeOut
    StyleDescription$ + ", StrikeOut"
  EndIf
  If FontStyle & #PB_Font_Underline
    StyleDescription$ + ", Underline"
  EndIf
  ProcedureReturn Mid(StyleDescription$, 3)
EndProcedure
;
Procedure LoadFontFromDescription(FontDescription$)
  ;
  ; Role: Loads a PureBasic font from a textual font description and returns its PureBasic font number.
  ;       Falls back to "Arial" and size 9 when the font name or size is omitted or set to "default".
  ;
  Protected parameter$, FontName$ = "Arial", FontSize = 9
  ;
  FontDescription$ = NormalizeFontDescription(FontDescription$)
  ;
  parameter$ = FD_ExtractParameter(FontDescription$, "name", #True)
  If parameter$ And parameter$ <> "default"
    FontName$ = parameter$
  EndIf
  parameter$ = FD_ExtractParameter(FontDescription$, "size", #True)
  If parameter$ And parameter$ <> "default"
    FontSize = Round(ValF(parameter$), #PB_Round_Nearest)
    If FontSize <= 0
      FontSize = 9
    EndIf
  EndIf
  ;
  ProcedureReturn LoadFont(#PB_Any, FontName$, FontSize, GetNumStyleFromAttributesDescription(FontDescription$))
EndProcedure
;
Procedure.s GetFontName(FontID)
  ;
  ; Role: Returns the system face name of the font defined by 'FontID' (falls back to "Arial").
  ; Parameters: FontID - Either a PureBasic font number or an OS font handle (HFONT / NSFont).
  ;
  Protected FontName$
  ;
  If IsFont(FontID)
    FontID = FontID(FontID)
  EndIf
  ;
  CompilerIf #PB_Compiler_OS = #PB_OS_MacOS
    ; Return the font name of FontID:
    Protected string
    If FontID
      ; "familyName" and "fontName" are for internal use; use "displayName" for the real name
      string = CocoaMessage(0, FontID, "displayName")
      If string
        FontName$ = PeekS(CocoaMessage(0, string, "UTF8String"), -1, #PB_UTF8)
      EndIf
    EndIf
  CompilerElseIf #PB_Compiler_OS = #PB_OS_Windows
    ; Create an HDC to retrieve the font info:
    Protected hdc, oldFont
    If FontID
      hdc = GetDC_(#Null)
      If hdc
        oldFont = SelectObject_(hdc, FontID)
        If oldFont
          ; Get the real font name (#LF_FACESIZE is the maximum size for a font name):
          FontName$ = Space(#LF_FACESIZE)
          If GetTextFace_(hdc, #LF_FACESIZE, @FontName$)
            ; Rebuild a clean PureBasic string up to the null terminator written by GetTextFace_:
            FontName$ = PeekS(@FontName$)
          Else
            FontName$ = ""
          EndIf
          ; Clean up:
          SelectObject_(hdc, oldFont)
        EndIf
        ReleaseDC_(#Null, hdc)
      EndIf
    EndIf
  CompilerEndIf
  ;
  If FontName$ = ""
    FontName$ = "Arial"
  EndIf
  ProcedureReturn FontName$
EndProcedure
;
Procedure GetFontSize(FontID)
  ;
  ; Role: Returns the point size (rounded to the nearest integer) of the font defined by 'FontID'.
  ; Parameters: FontID - Either a PureBasic font number or an OS font handle (HFONT / NSFont).
  ;
  Protected FpointSize
  ;
  If IsFont(FontID)
    FontID = FontID(FontID)
  EndIf
  ;
  CompilerIf #PB_Compiler_OS = #PB_OS_MacOS
    ; Return the font size of FontID:
    Protected pointSize.CGFloat = 0.0
    If FontID
      CocoaMessage(@pointSize, FontID, "pointSize")
    EndIf
    FpointSize = Round(pointSize, #PB_Round_Nearest)
  CompilerElseIf #PB_Compiler_OS = #PB_OS_Windows
    ; Get the font height in points (rounded to the nearest integer):
    Protected finfo.LOGFONT, tm.TEXTMETRIC
    Protected hdc, oldFont, PixelHeight, LogPixelsY
    If FontID And GetObject_(FontID, SizeOf(LOGFONT), @finfo)
      ; Acquire the screen device context:
      hdc = GetDC_(#Null)
      If hdc
        LogPixelsY = GetDeviceCaps_(hdc, #LOGPIXELSY)
        If finfo\lfHeight < 0
          PixelHeight = -finfo\lfHeight
        Else
          ; Positive or zero lfHeight represents cell or default height: extract character em-height via metrics
          oldFont = SelectObject_(hdc, FontID)
          If oldFont
            If GetTextMetrics_(hdc, @tm)
              PixelHeight = tm\tmHeight - tm\tmInternalLeading
            EndIf
            SelectObject_(hdc, oldFont)
          EndIf
          If PixelHeight <= 0
            PixelHeight = finfo\lfHeight
          EndIf
        EndIf
        ReleaseDC_(#Null, hdc)
        If LogPixelsY > 0
          FpointSize = Round(PixelHeight * 72.0 / LogPixelsY, #PB_Round_Nearest)
        EndIf
      EndIf
    EndIf
  CompilerEndIf
  ;
  ProcedureReturn FpointSize
EndProcedure
;
Procedure.s GetFontStyles(FontID)
  ;
  ; Role: Returns a comma-separated textual description of the styles applied to 'FontID'.
  ; Parameters: FontID - Either a PureBasic font number or an OS font handle (HFONT).
  ;
  Protected StyleDescription$
  ;
  If IsFont(FontID)
    FontID = FontID(FontID)
  EndIf
  ;
  CompilerIf #PB_Compiler_OS = #PB_OS_Windows
    Protected lf.LOGFONT, tm.TEXTMETRIC, Puretm.TEXTMETRIC
    Protected PureFont, hdc, oldFont
    ;
    ; Retrieve the font attributes:
    If FontID And GetObject_(FontID, SizeOf(LOGFONT), @lf)
      If lf\lfFaceName[0]
        hdc = GetDC_(#Null)
        If hdc
          oldFont = SelectObject_(hdc, FontID)
          If oldFont
            If GetTextMetrics_(hdc, @tm)
              PureFont = LoadFont(#PB_Any, PeekS(@lf\lfFaceName), 9)
              If PureFont
                ; Compare the physical weight of the font to the normal physical weight of its family:
                If SelectObject_(hdc, FontID(PureFont)) And GetTextMetrics_(hdc, @Puretm)
                  If tm\tmWeight > Puretm\tmWeight
                    StyleDescription$ + ", Bold"
                  ElseIf tm\tmWeight < Puretm\tmWeight
                    StyleDescription$ + ", Light"
                  EndIf
                EndIf
              EndIf
            EndIf
            ; Restore the original DC font before freeing PureFont so it is no longer selected in the DC:
            SelectObject_(hdc, oldFont)
            If PureFont
              FreeFont(PureFont)
            EndIf
          EndIf
          ReleaseDC_(#Null, hdc)
        EndIf
      EndIf
      ;
      If lf\lfItalic
        StyleDescription$ + ", Italic"
      EndIf
      If lf\lfUnderline
        StyleDescription$ + ", Underline"
      EndIf
      If lf\lfStrikeOut
        StyleDescription$ + ", StrikeOut"
      EndIf
    EndIf
  CompilerEndIf
  ;
  ProcedureReturn Mid(StyleDescription$, 3)
EndProcedure
;
Procedure.s GetFontDescription(FontID)
  ;
  ; Role: Builds a simplified textual description ("Name, Size[, Styles]") for an existing font.
  ; Parameters: FontID - Either a PureBasic font number or an OS font handle (HFONT / NSFont).
  ;
  Protected FontDescription$ = GetFontName(FontID) + ", " + Str(GetFontSize(FontID))
  Protected StyleDescription$ = GetFontStyles(FontID)
  ;
  If StyleDescription$
    FontDescription$ + ", " + StyleDescription$
  EndIf
  ProcedureReturn FontDescription$
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  ; The following code will not run when this file is included by another source file.
  ;
  Define Font, Description$
  ;
  Font = LoadFontFromDescription("arial,12,bold")
  Description$ = GetFontDescription(Font)
  FreeFont(Font)
  Debug "Simple description:        " + Description$
  Debug "Normalized description:    " + NormalizeFontDescription(Description$)
  Debug "Re-simplified description: " + SimplifyFontDescription(NormalizeFontDescription(Description$))
  ;
CompilerEndIf
;

; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 58
; FirstLine = 13
; Folding = OYg
; EnableXP
; DPIAware