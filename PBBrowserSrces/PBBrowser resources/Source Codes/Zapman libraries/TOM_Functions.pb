;***********************************************************
;
;                      TOM Library
;             v1.1 - By Zapman - Sept 2024
;
;        For text formatting or Image inserting
;                   in EditorGadgets.
;
;            Pour la mise en forme de textes
;                et l'insertion d'images
;                 dans un EditorGadget.
;
;***********************************************************
;
;     Example of using the Text Object Model (TOM)
;                 For Windows only.
;           Works on PureBasic 4.5 -> 6.11
;
; One of the advantages of TOM is that it allows modifying
; the formatting of text in an EditorGadget without changing
; (and thus, without losing) the current selection.
;
; The two main procedures in this library are:
; - TOM_SetFontStyles(), which allows applying a specific style
;   (bold, italic, underline, etc.) to a given text range.
; - TOM_SetParaStyles(), which allows applying a specific paragraph
;   style (indentation, spacing, etc.) to a given text range.
;
; The formatting commands are provided to these procedures
; in text mode.
; For example: TOM_SetFontStyles(GadgetID, StartPos, EndPos, "Bold, Size(12)"),
; so that they can be used by programmers of all levels.
;
; The possible commands are numerous and should cover almost
; all possible needs.
;
; A more complete usage example is provided at the end of the page.
;
; The list of possible commands is provided at the beginning of the code for
; the TOM_SetFontStyles() and TOM_SetParaStyles() procedures.
;
;***********************************************************
;
;     Exemple d'utilisation du Text Object Model (TOM)
;                 Pour Windows uniquement.
;           Fonctionne sur PureBasic 4.5 -> 6.11
;
; L'un des avantages du TOM est qu'il permet de modifier la
; mise en forme d'un texte dans un EditorGadget sans modifier
; (et donc, sans perdre) la sélection courante.
;
; Les deux principales procédures de la présente librairie sont
; - TOM_SetFontStyles() qui permet d'appliquer un style particulier
;   (gras, italique, souligné, etc.) à une plage de texte donnée.
; - TOM_SetParaStyles() qui permet d'appliquer un style de paragraphe
;   particulier (indentation, interparagraphe, etc.) à une plage
;   de texte donnée.
;
; Les commandes de mise en forme sont fournies à ces procédures
; en mode texte.
; Par exemple : TOM_SetFontStyles(GadgetID, StartPos, EndPos, "Bold, Size(12)"),
; afin d'être utilisables par les programmeurs de tous niveaux.
;
; Les commandes possibles sont très nombreuses et devraient permettre de
; répondre à peu près à tous les besoins possibles.
;
; Un exemple d'utilisation plus complet figure en fin de page.
;
; La liste des commandes possibles et fournies au début du code des
; procédures TOM_SetFontStyles() et TOM_SetParaStyles()
;
;
; Le fichier include "IDataObject.pb" est nécessaire
; pour les procédures TOM_InsertImage() et  TOM_InsertText()
XIncludeFile("IDataObject.pb")
;
;
#TomTrue      = -1
#TomFalse     = 0
#TomDefault   = -9999996
#TomAutoColor = -9999997
;
Enumeration Tom_UnderlineStyles
  #TomNone
  #TomSingle
  #TomWords
  #TomDouble
  #TomDotted
  #TomDash
  #TomDashDot
  #TomDashDotDot
  #TomWave
  #TomThick
  #TomHair
  #TomDoubleWave
  #TomHeavyWave
  #TomLongDash
  #TomThickDash
  #TomThickDashDot
  #TomThickDashDotDot
  #TomThickDotted
  #TomThickLongDash
EndEnumeration
;
Enumeration Tom_AlignmentStyles
  #TomAlignLeft       = 0
  #TomAlignCenter     = 1
  #TomAlignRight      = 2
  #TomAlignJustify    = 3
  #TomAlignDecimal    = 3
  #TomAlignBar        = 4
  #TomAlignInterWord  = 3
  #TomAlignNewspaper  = 4
  #TomAlignInterLetter = 5
  #TomAlignScaled     = 6
EndEnumeration
;
Enumeration Tom_SpaceLineRules
  #TomLineSpaceSingle
  #TomLineSpace1pt5
  #TomLineSpaceDouble
  #TomLineSpaceAtLeast
  #TomLineSpaceExactly
  #TomLineSpaceMultiple
  #TomLineSpacePercent
EndEnumeration
;
Enumeration Tom_Animations
  #TomNoAnimation
  #TomLasVegasLights
  #TomBlinkingBackground
  #TomSparkleText
  #TomMarchingBlackAnts
  #TomMarchingRedAnts
  #TomShimmer
  #TomWipeDown
  #TomWipeRight
EndEnumeration
;
; The resident Interface of ITextFont has bad parameters with
; version 6.11 (and olders) of PureBasic.
; So, a fixed interface must be set. Thanks to Justin (PB Forum)
; for the fixed interface:
Interface ITextFont_Fixed Extends IDispatch
  GetDuplicate(prop.i)
  SetDuplicate(Duplicate.i)
  CanChange(prop.i)
  IsEqual(pFont.i, prop.i)
  Reset(Value.l)
  GetStyle(prop.i)
  SetStyle(Style.l)
  GetAllCaps(prop.i)
  SetAllCaps(AllCaps.l)
  GetAnimation(prop.i)
  SetAnimation(Animation.l)
  GetBackColor(prop.i)
  SetBackColor(BackColor.l)
  GetBold(prop.i)
  SetBold(Bold.l)
  GetEmboss(prop.i)
  SetEmboss(Emboss.l)
  GetForeColor(prop.i)
  SetForeColor(ForeColor.l)
  GetHidden(prop.i)
  SetHidden(Hidden.l)
  GetEngrave(prop.i)
  SetEngrave(Engrave.l)
  GetItalic(prop.i)
  SetItalic(Italic.l)
  GetKerning(prop.i)
  SetKerning(Kerning.f)
  GetLanguageID(prop.i)
  SetLanguageID(LanguageID.l)
  GetName(prop.i)
  SetName(Name.p - bstr)
  GetOutline(prop.i)
  SetOutline(Outline.l)
  GetPosition(prop.i)
  SetPosition(Position.f)
  GetProtected(prop.i)
  SetProtected(Protected.l)
  GetShadow(prop.i)
  SetShadow(Shadow.l)
  GetSize(prop.i)
  SetSize(Size.f)
  GetSmallCaps(prop.i)
  SetSmallCaps(SmallCaps.l)
  GetSpacing(prop.i)
  SetSpacing(Spacing.f)
  GetStrikeThrough(prop.i)
  SetStrikeThrough(StrikeThrough.l)
  GetSubscript(prop.i)
  SetSubscript(Subscript.l)
  GetSuperscript(prop.i)
  SetSuperscript(Superscript.l)
  GetUnderline(prop.i)
  SetUnderline(Underline.l)
  GetWeight(prop.i)
  SetWeight(Weight.l)
EndInterface 
;
; The resident Interface of ITextPara has bad parameters with
; version 6.11 (and olders) of PureBasic.
; So, a fixed interface must be set. Thanks to Justin (PB Forum)
; for the fixed interface:
Interface ITextPara_Fixed Extends IDispatch
  GetDuplicate(prop.i)
  SetDuplicate(Duplicate.i)
  CanChange(prop.i)
  IsEqual(pPara.i, prop.i)
  Reset(Value.l)
  GetStyle(prop.i)
  SetStyle(Style.l)
  GetAlignment(prop.i)
  SetAlignment(Alignment.l)
  GetHyphenation(prop.i)
  SetHyphenation(Hyphenation.l)
  GetFirstLineIndent(prop.i)
  GetKeepTogether(prop.i)
  SetKeepTogether(KeepTogether.l)
  GetKeepWithNext(prop.i)
  SetKeepWithNext(KeepWithNext.l)
  GetLeftIndent(prop.i)
  GetLineSpacing(prop.i)
  GetLineSpacingRule(prop.i)
  GetListAlignment(prop.i)
  SetListAlignment(ListAlignment.l)
  GetListLevelIndex(prop.i)
  SetListLevelIndex(ListLevelIndex.l)
  GetListStart(prop.i)
  SetListStart(ListStart.l)
  GetListTab(prop.i)
  SetListTab(ListTab.f)
  GetListType(prop.i)
  SetListType(ListType.l)
  GetNoLineNumber(prop.i)
  SetNoLineNumber(NoLineNumber.l)
  GetPageBreakBefore(prop.i)
  SetPageBreakBefore(PageBreakBefore.l)
  GetRightIndent(prop.i)
  SetRightIndent(RightIndent.f)
  SetIndents(First.f, Left.f, Right.f)
  SetLineSpacing(Rule.l, Spacing.f)
  GetSpaceAfter(prop.i)
  SetSpaceAfter(SpaceAfter.f)
  GetSpaceBefore(prop.i)
  SetSpaceBefore(SpaceBefore.f)
  GetWidowControl(prop.i)
  SetWidowControl(WidowControl.l)
  GetTabCount(prop.i)
  AddTab(tbPos.f, tbAlign.l, tbLeader.l)
  ClearAllTabs()
  DeleteTab(tbPos.f)
  GetTab(iTab.l, ptbPos.i, ptbAlign.i, ptbLeader.i)
EndInterface
;
Interface ITextRange2 Extends IDispatch ; Thanks to Justin (PB Forum)
                                        ; for the interface:
  GetText(prop.i)
  SetText(Text.p - bstr)
  GetChar(prop.i)
  SetChar(Char.l)
  GetDuplicate(prop.i)
  GetFormattedText(prop.i)
  SetFormattedText(FormattedText.i)
  GetStart(prop.i)
  SetStart(Start.l)
  GetEnd(prop.i)
  SetEnd(End.l)
  GetFont(prop.i)
  SetFont(Font.i)
  GetPara(prop.i)
  SetPara(Para.i)
  GetStoryLength(prop.i)
  GetStoryType(prop.i)
  Collapse(bStart.l)
  Expand(Unit.l, prop.i)
  GetIndex(Unit.l, prop.i)
  SetIndex(Unit.l, Index.l, Extend.l)
  SetRange(cpAnchor.l, cpActive.l)
  InRange(*pTextRange.i, prop.i)
  InStory(*pTextRange.i, prop.i)
  IsEqual(*pTextRange.i, prop.i)
  Select ()
  StartOf(Unit.l, Extend.l, prop.i)
  EndOf(Unit.l, Extend.l, prop.i)
  Move(Unit.l, Count.l, prop.i)
  MoveStart(Unit.l, Count.l, prop.i)
  MoveEnd(Unit.l, Count.l, prop.i)
  MoveWhile(Cset.i, Count.l, prop.i)
  MoveStartWhile(Cset.i, Count.l, prop.i)
  MoveEndWhile(Cset.i, Count.l, prop.i)
  MoveUntil(Cset.i, Count.l, prop.i)
  MoveStartUntil(Cset.i, Count.l, prop.i)
  MoveEndUntil(Cset.i, Count.l, prop.i)
  FindText(bstr.p - bstr, Count.l, Flags.l, prop.i)
  FindTextStart(bstr.p - bstr, Count.l, Flags.l, prop.i)
  FindTextEnd(bstr.p - bstr, Count.l, Flags.l, prop.i)
  Delete(Unit.l, Count.l, prop.i)
  Cut(pVar.i)
  Copy(pVar.i)
  Paste(pVar.i, Format.l)
  CanPaste(pVar.i, Format.l, prop.i)
  CanEdit(prop.i)
  ChangeCase(Type.l)
  GetPoint(Type.l, px.i, py.i)
  SetPoint(x.l, y.l, Type.l, Extend.l)
  ScrollIntoView(Value.l)
  GetEmbeddedObject(prop.i)
  GetFlags(prop.i)
  SetFlags(Flags.l)
  GetType(prop.i)
  MoveLeft(Unit.l, Count.l, Extend.l, prop.i)
  MoveRight(Unit.l, Count.l, Extend.l, prop.i)
  MoveUp(Unit.l, Count.l, Extend.l, prop.i)
  MoveDown(Unit.l, Count.l, Extend.l, prop.i)
  HomeKey(Unit.l, Extend.l, prop.i)
  EndKey(Unit.l, Extend.l, prop.i)
  TypeText(bstr.p - bstr)
  GetCch(prop.i)
  GetCells(prop.i)
  GetColumn(prop.i)
  GetCount(prop.i)
  GetDuplicate2(prop.i)
  GetFont2(prop.i)
  SetFont2(Font2.i)
  GetFormattedText2(prop.i)
  SetFormattedText2(FormattedText2.i)
  GetGravity(prop.i)
  SetGravity(Gravity.l)
  GetPara2(prop.i)
  SetPara2(Para2.i)
  GetRow(prop.i)
  GetStartPara(prop.i)
  GetTable(prop.i)
  GetURL(prop.i)
  SetURL(URL.p - bstr)
  AddSubrange(cp1.l, cp2.l, Activate.l)
  BuildUpMath(Flags.l)
  DeleteSubrange(cpFirst.l, cpLim.l)
  Find(*pTextRange.i, Count.l, Flags.l, pDelta.i)
  GetChar2(pChar.i, Offset.l)
  GetDropCap(pcLine.i, pPosition.i)
  GetInlineObject(pType.i, pAlign.i, pChar.i, pChar1.i, pChar2.i, pCount.i, pTeXStyle.i, pcCol.i, pLevel.i)
  GetProperty(Type.l, pValue.i)
  GetRect(Type.l, pLeft.i, pTop.i, pRight.i, pBottom.i, pHit.i)
  GetSubrange(iSubrange.l, pcpFirst.i, pcpLim.i)
  GetText2(Flags.l, pbstr.i)
  HexToUnicode()
  InsertTable(cCol.l, cRow.l, AutoFit.l)
  Linearize(Flags.l)
  SetActiveSubrange(cpAnchor.l, cpActive.l)
  SetDropCap(cLine.l, Position.l)
  SetProperty(Type.l, Value.l)
  SetText2(Flags.l, bstr.p - bstr)
  UnicodeToHex()
  SetInlineObject(Type.l, Align.l, Char.l, Char1.l, Char2.l, Count.l, TeXStyle.l, cCol.l)
  GetMathFunctionType(bstr.p - bstr, pValue.i)
  InsertImage(width.l, Height.l, ascent.l, Type.l, bstrAltText.p - bstr, *pStream)
EndInterface
;
Procedure TOM_PrintErrorMessage(result)
  If result <> #S_OK
    Select Result
      Case #E_INVALIDARG: Debug("E_INVALIDARG- Invalid argument")
      Case #E_ACCESSDENIED:  Debug("E_ACCESSDENIED - write access denied")
      Case #E_OUTOFMEMORY:   Debug("E_OUTOFMEMORY - out of memory")
      Case #CO_E_RELEASED:   Debug("CO_E_RELEASED - The paragraph formatting object is attached to a range that has been deleted.")
      Default: Debug "Some other error occurred"
    EndSelect
  Else
    Debug "No error"
  EndIf
EndProcedure
;
Procedure TOM_GetTextRangeObj(GadgetID, StartPos = -2, EndPos = -2)
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ; After usage, the ITextRange object obtained must be free
  ; like this : *pTextRange\Release()
  ;
  Protected RichEditOleObject.IRichEditOle
  Protected    *pTextDocument.ITextDocument2
  Protected       *pTextRange.ITextRange2
  Protected               SC = #S_OK
  ;
  If StartPos   = -1   : StartPos = #MAXLONG : EndIf
  If EndPos     = -1   :   EndPos = #MAXLONG : EndIf
  If StartPos > EndPos :   EndPos = StartPos : EndIf
  ;
  If GadgetID 
    If IsGadget(GadgetID) : GadgetID = GadgetID(GadgetID) : EndIf
    ; Get the RichOLEInterface on our EditorGadget:
    SendMessage_(GadgetID, #EM_GETOLEINTERFACE, 0, @RichEditOleObject)
    ; Get the ITextDocument interface:
    SC = RichEditOleObject\QueryInterface(?IID_ITextDocument2, @*pTextDocument)
    RichEditOleObject\Release()
    If SC <>  #S_OK
      Debug "Unable to create ITextDocument: " + Str(SC)
    Else
      ; Get the ITextRange interface:
      If StartPos = -2
        SC = *pTextDocument\GetSelection(@*pTextRange)
      Else
        SC = *pTextDocument\Range(StartPos, EndPos, @*pTextRange)
      EndIf
      *pTextDocument\Release()
      If SC <> #S_OK
        Debug "Unable to get the ITextRange interface. Error : " + Str(SC)
        *pTextRange = 0
      EndIf
    EndIf
  EndIf
  ProcedureReturn *pTextRange
EndProcedure
;
Procedure TOM_GetSelectionPos(GadgetID, *Selrange.CHARRANGE)
  ;
  Protected *pTextRange.ITextRange2 = TOM_GetTextRangeObj(GadgetID)
  ;
  If *Selrange And *pTextRange
    *pTextRange\GetStart(@*Selrange\cpMin)
    *pTextRange\GetEnd(@*Selrange\cpMax)
    *pTextRange\Release()
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure TOM_SetSelectionPos(GadgetID, StartPos, EndPos)
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; If StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; If EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  Protected *pTextRange.ITextRange2 = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  ;
  If *pTextRange
    *pTextRange\Select()
    *pTextRange\Release()
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure TOM_GetStartPos(GadgetID, StartPos = -2)
  ;
  ; If StartPos is omitted or StartPos = -2, return value is the start of the current selection.
  ; If StartPos = -1, return value is the end of GadgetID content.
  ;
  Protected *pTextRange.ITextRange2 = TOM_GetTextRangeObj(GadgetID, StartPos)
  Protected RetValue = 0
  ;
  If *pTextRange
    *pTextRange\GetStart(@RetValue)
    *pTextRange\Release()
    ProcedureReturn RetValue
  EndIf
EndProcedure
;
Procedure TOM_GetEndPos(GadgetID, EndPos = -2)
   ;
  ; If EndPos is omitted or EndPos = -2, return value is the end of the current selection.
  ; IF EndPos = -1, return value is the end of GadgetID content.
  ;
  Protected *pTextRange.ITextRange2 = TOM_GetTextRangeObj(GadgetID, EndPos)
  Protected RetValue
  ;
  If *pTextRange
    *pTextRange\GetEnd(@RetValue)
    *pTextRange\Release()
    ProcedureReturn RetValue
  EndIf
EndProcedure
;
Procedure TOM_Copy(GadgetID, StartPos = -2, EndPos = -2)
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; If StartPos = -1, nothing will be copied.
  ; If EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  Protected *pTextRange.ITextRange2 = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  ;
  If *pTextRange
    *pTextRange\Copy(0)
    *pTextRange\Release()
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure TOM_Cut(GadgetID, StartPos = -2, EndPos = -2)
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; If StartPos = -1, nothing will be cut.
  ; If EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  Protected *pTextRange.ITextRange2 = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  ;
  If *pTextRange
    *pTextRange\Cut(0)
    *pTextRange\Release()
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure TOM_Paste(GadgetID, StartPos = -2, EndPos = -2)
  ;
  ; If StartPos is omitted or StartPos = -2, clipboard content will replace the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; If StartPos = -1, clipboard content is paste at the end of GadgetID content.
  ; If EndPos = -1, clipboard content will replace the range from StatPos to the end of GadgetID content.
  ;
  Protected *pTextRange.ITextRange2 = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  ;
  If *pTextRange
    *pTextRange\Paste(0, 0)
    *pTextRange\Release()
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure TOM_GetTextFontObj(GadgetID, StartPos = -2, EndPos = -2, Duplicate = #TomFalse)
  ;
  ; This procedure sets up a 'TextFont' interface for the 'GadgetID' GadgetID.
  ;
  ; It returns an ITextFont object that can be:
  ; - The ITextFont of the character range StartPos->EndPos, if Duplicate = #TomFalse
  ; - A copy of this ITextFont, if Duplicate = #TomTrue
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ; This ITextFont object should be cleaned up after use by calling: *TextFontObject\Release().
  ;
  ; Example of usage:
  ;
  ; We will copy the styles from the tenth character contained in the GadgetID:
  ; *TextFontObjet.ITextFont = TOM_GetTextFontObj(EGadget, 10, 11, #TomTrue)
  ; We apply the same styles to the character range from 20 to 26:
  ; TOM_ApplyTextFont(EGadget, *TextFontObjet, 20, 27)
  ; Then we free the memory:
  ; *TextFontObjet\Release()
  ;
  ; The last parameter of this procedure ('Duplicate') allows obtaining
  ; an active ITextFont object (when Duplicate = #TomFalse), with which you
  ; can later play to modify the style of the text corresponding
  ; to the provided character range. As long as this ITextFont object
  ; is not deleted by *TextFontObjet\Release(), it continues to reflect
  ; the style changes made to the corresponding range, and it can be used
  ; to modify these styles.
  ; If Duplicate = #TomTrue, the obtained ITextFont object is just a snapshot
  ; taken at a given moment. If you modify its content (with *TextFontObjet\Reset(),
  ; for example), it does not affect the character range that was used to create it.
  ; However, you can use TOM_ApplyTextFont() to reapply this set of
  ; styles to any character range.
  ;
  Protected *pTextRange.ITextRange
  Protected *pTextFont.ITextFont_Fixed
  Protected *DTextFont.ITextFont_Fixed
  Protected Result = #S_FALSE ; Valeur de retour.
                                 ;
  *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  If *pTextRange
    If *pTextRange\GetFont(@*pTextFont) = #S_OK And *pTextFont
      If Duplicate = #TomTrue
        *pTextFont\GetDuplicate(@*DTextFont)
        Result = *DTextFont
        *pTextFont\Release()
      Else
        Result = *pTextFont
      EndIf
    EndIf
    *pTextRange\Release()
  EndIf

  ProcedureReturn Result
EndProcedure
;
Procedure TOM_GetTextParaObj(GadgetID, StartPos = -2, EndPos = -2, Duplicate = #TomFalse)
  ;
  ; This procedure sets up a 'TextPara' interface for the 'GadgetID' GadgetID.
  ;
  ; It returns an ITextPara object that can be:
  ; - The ITextPara of the character range StartPos->EndPos, if Duplicate = #TomFalse
  ; - A copy of this ITextPara, if Duplicate = #TomTrue
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ; This ITextPara object should be cleaned up after use by calling: *TextParaObject\Release().
  ;
  ; Example of usage:
  ;
  ; We will copy the paragraph styles from the tenth character contained in the GadgetID:
  ; *TextParaObjet.ITextPara_Fixed = TOM_GetTextParaObj(EGadget, 10, 11, #TomTrue)
  ; We apply the same styles to the character range from 20 to 26:
  ; TOM_ApplyParaFont(EGadget, 20, 27, *TextFontObjet)
  ; Then we free the memory:
  ; *TextParaObjet\Release()
  ;
  ; Refer to the notes of 'TOM_GetTextFontObj()' for more details on the usage
  ; of the 'Duplicate' parameter.
  ;
  Protected *pTextRange.ITextRange
  Protected *pTextPara.ITextPara_Fixed
  Protected *DTextPara.ITextPara_Fixed
  Protected Result = #S_FALSE ; Valeur de retour.
                                 ;
  *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  If *pTextRange
    ; Get the ITextPara:
    If *pTextRange\GetPara(@*pTextPara) = #S_OK And *pTextPara
      If Duplicate = #TomTrue
        *pTextPara\GetDuplicate(@*DTextPara)
        Result = *DTextPara
        *pTextPara\Release()
      Else
        Result = *pTextPara
      EndIf
    EndIf
    *pTextRange\Release()
  EndIf
  ProcedureReturn Result
EndProcedure
;
Procedure TOM_ApplyTextFont(GadgetID, *pTextFont.ITextFont_Fixed, StartPos = -2, EndPos = -2)
  ;
  ; This procedure applies to a text range defined by StartPos->EndPos
  ; the set of styles contained in the '*pTextFont' object.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ; Example of usage:
  ;
  ; We will copy the styles from the tenth character contained in the 'GadgetID' GadgetID:
  ; *TextFontObjet.ITextFont_Fixed = TOM_GetTextFontObj(EGadget, 10, 11, #TomTrue)
  ; We apply the same styles to the character range from 20 to 26:
  ; TOM_ApplyTextFont(EGadget, *TextFontObjet, 20, 27)
  ; 
  ; Then we free the memory:
  ; *TextFontObjet\Release()
  ;
  Protected *pTextRange.ITextRange2
  Protected Result = #S_FALSE ; Return value
  ;
  *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  If *pTextRange
    ; Apply:
    Result = *pTextRange\SetFont(*pTextFont)
    *pTextRange\Release()
  EndIf
  ProcedureReturn Result
EndProcedure
;
Procedure TOM_ApplyTextPara(GadgetID, *pTextPara.ITextPara_Fixed, StartPos = -2, EndPos = -2)
  ;
  ; This procedure applies to a text range defined by StartPos->EndPos
  ; the set of paragraph styles contained in the '*pTextPara' object.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ; Example of usage:
  ;
  ; We will copy the paragraph styles from the tenth character contained in the 'GadgetID' GadgetID:
  ; *TextParaObjet.ITextPara_Fixed = TOM_GetTextParaObj(EGadget, 10, 11, #TomTrue)
  ; We apply the same styles to the character range from 20 to 26:
  ; TOM_ApplyTextPara(EGadget, *TextParaObjet, 20, 27)
  ; 
  ; Then we free the memory:
  ; *TextParaObjet\Release()
  ;
  Protected *pTextRange.ITextRange2
  Protected Result = #S_FALSE ; Return value
  ;
  *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, EndPos)
  If *pTextRange
    ; Apply:
    Result = *pTextRange\SetPara(*pTextPara)
    *pTextRange\Release()
  EndIf
  ProcedureReturn Result
EndProcedure
;
Procedure.s ExtractParameterForTOM(Style$, ParameterName$)
  ;
  ; This procedure, used by 'TOM_SetFontStyles()'
  ; retrieves the parameter in parenthesis that follows 'ParameterName$'
  ; in the 'Style$' string.
  ;
  Protected pa, pas, limp
  ;
  pa = FindString(Style$, LCase(ParameterName$), 1)
  If pa
    pa + Len(ParameterName$)
    If PeekC(@Style$ + (pa - 1) * SizeOf(CHARACTER)) = Asc("(")
      pa + 1
    EndIf
    If PeekC(@Style$ + (pa - 2) * SizeOf(CHARACTER)) <> Asc("(")
      ;MessageRequester("Error", "Error with 'TOM_SetFontStyles':  No parenthesis after " + ParameterName$ + Chr(13) + Style$)
      ProcedureReturn
    Else
      pas = pa
      limp = Len(Style$)
      While pa <= limp And PeekC(@Style$ + (pa - 1) * SizeOf(CHARACTER)) <> Asc(")")
        pa + 1
      Wend
      ProcedureReturn PeekS(@Style$ + (pas - 1) * SizeOf(CHARACTER), pa - pas)
    EndIf
  Else
    MessageRequester("Error", "Error with 'TOM_SetFontStyles':  Wrong parameter name -> " + ParameterName$ + Chr(13) + Style$)
  EndIf
EndProcedure
;
Macro FontSetUnset(StyleName, GetName, SetName)
  If FindString(Style$, StyleName)
    parameter$ = ExtractParameterForTOM(Style$, StyleName)
    If parameter$ And LCase(parameter$) <> "default"
      *pTextFont\SetName(Val(parameter$))
    ElseIf SetUnset = #TomDefault Or parameter$ = "default"
      *pFontDefault\GetName(@pl)
      If pl
        pl = #TomTrue
      endif
      *pTextFont\SetName(pl)
    Else 
      *pTextFont\SetName(SetUnset)
    EndIf
  EndIf
EndMacro 
;
Macro FontSetValue(StyleName, GetName, SetName)
  If FindString(Style$, StyleName)
    parameter$ = ExtractParameterForTOM(Style$, StyleName)
    If SetUnset = #TomTrue And parameter$ <> "default"
      *pTextFont\SetName(ValF(parameter$))
    Else
      *pTextFont\GetName(@pf)
      If pf = ValF(parameter$) Or parameter$ = "default" Or SetUnset = #TomFalse
        *pFontDefault\GetName(@pf)
        *pTextFont\SetName(pf)
      EndIf
    EndIf
  EndIf
EndMacro
;
Procedure TOM_SetFontStyles(GadgetID, Style$, StartPos = -2, EndPos = -2, SetUnset = #TomTrue)
  ;
  ; This procedure applies various styles to the text range
  ; defined by StartPos->EndPos.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ; GadgetID must be the GadgetID number of an EditorGadget containing text.
  ;
  ; Example content for the 'Style$' string:
  ; "Bold, Italic, BackColor($F08050)"
  ; -> will apply bold-italic to the text range with the background color $F08050.
  ; 
  ; 'SetUnset' can be omitted or can contain : #TomTrue, #TomFalse or #TomDefault
  ; With '#TomTrue', the commands contained by 'Style$' will be applied to the text range.
  ; With '#TomDefault', the text range will be set to default values, wathever the values
  ; evenually specified in parenthesis.
  ; With '#TomFalse', all the precise specified styles will be set to defaut values.
  ; For exemples:
  ; TOM_SetFontStyles(GadgetID, "Bold", StartPos, EndPos)
  ; -> Range is set to bold
  ; TOM_SetFontStyles(GadgetID, "Bold", StartPos, EndPos, #TomDefault)
  ; -> Range is set to bold if default style is bold or non-bold if default style is non-bold
  ; TOM_SetFontStyles(GadgetID, "Bold", StartPos, EndPos, #TomFalse)
  ; -> Range is set to non-bold
  ;
    ; TOM_SetFontStyles(GadgetID, "Size(12.5)", StartPos, EndPos)
  ; -> Range is set to size 12.5 pts
  ; TOM_SetFontStyles(GadgetID, "Size()", StartPos, EndPos, #TomDefault) or TOM_SetFontStyles(GadgetID, "Size(xxx)", StartPos, EndPos, #TomDefault)
  ; -> Range is set to default size.
  ; TOM_SetFontStyles(GadgetID, "Size(12.5)", StartPos, EndPos, #TomFalse)
  ; -> Only the characters having a size of 12.5 into the range will be set to default size.
  ;
  ; 'Style$' can contain some of the following commands (separated by comma):
  ; Bold, Italic, Emboss, AllCaps, SmallCaps, Engrave, Shadow, OutLine, Underline(value),
  ; StrikeThrough, Hidden, Protected, Size(value.f), Spacing(value.f), Position(value.f), Kerning(value.f),
  ; BackColor(value), ForeColor(value), Weight(value), Style(value), Name(value).
  ;
  ; The possible values for Underline are:

  ; Underline(Single)
  ; Underline(Words)
  ; Underline(Double)
  ; Underline(Dotted)
  ; Underline(Dash)
  ; Underline(DashDot)
  ; Underline(DashDotDot)
  ; Underline(Wave)
  ; Underline(Thick)
  ; Underline(Hair)
  ; Underline(DoubleWave)
  ; Underline(HeavyWave)
  ; Underline(LongDash)
  ; Underline(ThickDash)
  ; Underline(ThickDashDot)
  ; Underline(ThickDashDotDot)
  ; Underline(ThickDotted)
  ; Underline(ThickLongDash)
  ;  
  Protected *pTextFont.ITextFont_Fixed
  Protected *pFontDefault.ITextFont_Fixed
  Protected pl.l = 0, gpl.l = 0
  Protected pf.f = 0
  Protected ps.s = "", *BSTRString = 0
  Protected parameter$
  ;
  ; To simplify the parsing of the parameter string,
  ; the spaces it contains are removed.
  While FindString(Style$, " ", 1)
    Style$ = ReplaceString(Style$, " ", "")
  Wend
  Style$ = LCase(Style$)
  ;
  ; Get an active TextFontObj for the range:
  *pTextFont = TOM_GetTextFontObj(GadgetID, StartPos, EndPos)
  If *pTextFont <> #S_FALSE
    ; Get a TextFontObj copy for the range:
    *pTextFont\GetDuplicate(@*pFontDefault)
    ; Set the copy's styles to default:
    *pFontDefault\Reset(#TomDefault)
    ;
    FontSetUnset("bold", GetBold, SetBold)
    FontSetUnset("italic", GetItalic, SetItalic)
    FontSetUnset("emboss", GetEmboss, SetEmboss)
    FontSetUnset("allcaps", GetAllcaps, SetAllcaps)
    FontSetUnset("smallcaps", GetSmallcaps, SetSmallcaps)
    FontSetUnset("engrave", GetEngrave, SetEngrave)
    FontSetUnset("shadow", GetShadow, SetShadow)
    FontSetUnset("outline", GetOutline, SetOutline)
    FontSetUnset("strikethrough", GetStrikethrough, SetStrikethrough)
    FontSetUnset("subscript", GetSubscript, SetSubscript)
    FontSetUnset("superscript", GetSuperscript, SetSuperscript)
    FontSetUnset("hidden", GetHidden, SetHidden)
    FontSetUnset("protected", GetProtected, SetProtected)
    
    If FindString(Style$, "underline(")
      parameter$ = ExtractParameterForTOM(Style$, "underline(")
      If parameter$ = "none" Or parameter$ = "0"
        pl = #TomNone
      ElseIf parameter$ = "words"
        pl = #TomWords
      ElseIf parameter$ = "double"
        pl = #TomDouble
      ElseIf parameter$ = "dotted"
        pl = #TomDotted
      ElseIf parameter$ = "dash"
        pl = #TomDash
      ElseIf parameter$ = "dashdot"
        pl = #TomDashDot
      ElseIf parameter$ = "dashdotdot"
        pl = #TomDashDotDot
      ElseIf parameter$ = "wave"
        pl = #TomWave
      ElseIf parameter$ = "thick"
        pl = #TomThick
      ElseIf parameter$ = "hair"
        pl = #TomHair
      ElseIf parameter$ = "doublewave"
        pl = #TomDoubleWave
      ElseIf parameter$ = "heavywave"
        pl = #TomHeavyWave
      ElseIf parameter$ = "longdash"
        pl = #TomLongDash
      ElseIf parameter$ = "thickdash"
        pl = #TomThickDash
      ElseIf parameter$ = "thickdashdot"
        pl = #TomThickDashDot
      ElseIf parameter$ = "thickdashdotdot"
        pl = #TomThickDashDotDot
      ElseIf parameter$ = "thickdotted"
        pl = #TomThickDotted
      ElseIf parameter$ = "thicklongdash"
        pl = #TomThickLongDash
      Else
        pl = #TomSingle
      EndIf

      If SetUnset = #TomDefault Or parameter$ = "default"
        *pFontDefault\GetUnderline(@pl)
        *pTextFont\SetUnderline(pl)
      ElseIf SetUnset = #TomTrue
        *pTextFont\SetUnderline(pl)
      Else
        *pTextFont\GetUnderline(@gpl)
        If gpl = pl Or parameter$ = ""
          *pTextFont\SetUnderline(#TomNone)
        EndIf
      EndIf
    EndIf
    ;
    FontSetValue("size(", GetSize, SetSize)
    FontSetValue("spacing(", GetSpacing, SetSpacing)
    FontSetValue("position(", GetPosition, SetPosition)
    FontSetValue("kerning(", GetKerning, SetKerning)
    FontSetValue("backcolor(", GetBackcolor, SetBackcolor)
    FontSetValue("forecolor(", GetForecolor, SetForecolor)
    FontSetValue("weight(", GetWeight, SetWeight)
    FontSetValue("style(", GetStyle, SetStyle)

    If FindString(Style$, "name(")
      parameter$ = ExtractParameterForTOM(Style$, "name(")
      If SetUnset = #TomTrue
        *pTextFont\SetName(parameter$)
      Else
        *pTextFont\GetName(@*BSTRString)
        ps = PeekS(*BSTRString, -1, #PB_Unicode)
        SysFreeString_(*BSTRString)
        If ps = parameter$ Or SetUnset = #TomDefault
          *pFontDefault\GetName(@*BSTRString)
          ps = PeekS(*BSTRString, -1, #PB_Unicode)
          SysFreeString_(*BSTRString)
          *pTextFont\SetName(ps)
        ElseIf parameter$ = ""
          *pTextFont\SetName("")
        EndIf
      EndIf
    EndIf
    *pTextFont\Release()
  EndIf
EndProcedure
;
Macro ParaSetValue(StyleName, GetName, SetName)
  If FindString(Style$, StyleName)
    parameter$ = ExtractParameterForTOM(Style$, StyleName)
    If SetUnset = #TomTrue And parameter$ <> "default"
      *pTextPara\SetName(ValF(parameter$))
    Else
      *pTextPara\GetName(@pf)
      If pf = ValF(parameter$) Or parameter$ = "default" Or SetUnset = #TomFalse
        *pParaDefault\GetName(@pf)
        *pTextPara\SetName(pf)
      EndIf
    EndIf
  EndIf
EndMacro
;
Procedure TOM_SetParaStyles(GadgetID, Style$, StartPos = -2, EndPos = -2, SetUnset = #TomTrue)
  ;
  ; This procedure applies various paragraphe styles to the text range
  ; defined by StartPos->EndPos.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ; GadgetID must be the GadgetID number of an EditorGadget containing text.
  ;
  ; Example content for the 'Style$' string:
  ; "Align(left), FirstLineIndent(20)"
  ; 
  ; Explanations for the use of last parameter ('SetUnset') can be found
  ; into the code of procedure TOM_SetFontStyles()'.
  ;
  ; 'Style$' can contain some of the following commands (separated by comma):
  ; Align(value), SpaceBefore(Value.f), SpaceAfter(Value.f)
  ; RightIndent(value.f), LeftIndent(value.f), FirstLineIndent(value.f)
  ; Style(value), LineSpacing(SpacingRule, value.f)
  ;
  ; For LineSpacing, the SpacingRule value can contain:
  ; "Single", "1pt5", "Double", "AtLeast", "Exactly", "Multiple" or "Percent"
  ; The second parameter is unused with "Single", "1pt5" and "Double".
  
  Protected *pTextPara.ITextPara_Fixed
  Protected *pParaDefault.ITextPara_Fixed
  Protected pl.l = 0, gpl.l
  Protected pf.f = 0, pf1.f = 0, pf2.f = 0, pf3.f = 0
  Protected parameter$, param1$, param2$
  ;
  ; To simplify the parsing of the parameter string,
  ; the spaces it contains are removed.
  While FindString(Style$, " ", 1)
    Style$ = ReplaceString(Style$, " ", "")
  Wend
  Style$ = LCase(Style$)
  ;
  ; Get an active TextParaObj for the range:
  *pTextPara = TOM_GetTextParaObj(GadgetID, StartPos, EndPos)
  If *pTextPara <> #TomFalse
    ; Get a TextParaObj copy for the range:
    *pTextPara\GetDuplicate(@*pParaDefault)
    ; Set the copy's styles to default:
    *pParaDefault\Reset(#TomDefault)
    ;
    If FindString(Style$, "align(")
      If FindString(Style$, "align(left")
        pl = #TomAlignLeft
      ElseIf FindString(Style$, "align(center")
        pl = #TomAlignCenter
      ElseIf FindString(Style$, "align(right")
        pl = #TomAlignRight
      ElseIf FindString(Style$, "align(justify")
        pl = #TomAlignJustify
      ElseIf FindString(Style$, "align(decimal")
        pl = #TomAlignDecimal
      ElseIf FindString(Style$, "align(bar")
        pl = #TomAlignBar
      ElseIf FindString(Style$, "align(interword")
        pl = #TomAlignInterWord
      ElseIf FindString(Style$, "align(newspaper")
        pl = #TomAlignNewspaper
      ElseIf FindString(Style$, "align(interletter")
        pl = #TomAlignInterLetter
      ElseIf FindString(Style$, "align(scaled")
        pl = #TomAlignScaled
      EndIf
      If SetUnset = #TomTrue
        *pTextPara\SetAlignment(pl)
      Else
        *pTextPara\GetAlignment(@gpl)
        If gpl = pl Or SetUnset = #TomDefault
          *pParaDefault\GetAlignment(@pl)
          *pTextPara\SetAlignment(pl)
        EndIf
      EndIf
    EndIf
    ;
    ParaSetValue("rightindent", GetRightIndent, SetRightIndent)
    ;
    If FindString(Style$, "leftindent")
      parameter$ = ExtractParameterForTOM(Style$, "leftindent")
      If SetUnset = #TomTrue And parameter$ <> "default"
        *pTextPara\GetFirstLineIndent(@pf1)
        *pTextPara\GetRightIndent(@pf3)
        *pTextPara\SetIndents(pf1, ValF(parameter$), pf3)
      Else
        *pTextPara\GetLeftIndent(@pf2)
        If pf2 = ValF(parameter$) Or parameter$ = "default" Or SetUnset = #TomDefault
          *pParaDefault\GetLeftIndent(@pf2)
          *pTextPara\GetFirstLineIndent(@pf1)
          *pTextPara\GetRightIndent(@pf3)
          *pTextPara\SetIndents(pf1, pf2, pf3)
        EndIf
      EndIf
    EndIf
    If FindString(Style$, "firstlineindent")
      parameter$ = ExtractParameterForTOM(Style$, "firstlineindent")
      If SetUnset = #TomTrue And parameter$ <> "default"
        *pTextPara\GetLeftIndent(@pf2)
        *pTextPara\GetRightIndent(@pf3)
        *pTextPara\SetIndents(ValF(parameter$), pf2, pf3)
      Else
        *pTextPara\GetFirstLineIndent(@pf1)
        If pf1 = ValF(parameter$) Or parameter$ = "default" Or SetUnset = #TomDefault
          *pParaDefault\GetFirstLineIndent(@pf1)
          *pTextPara\GetLeftIndent(@pf2)
          *pTextPara\GetRightIndent(@pf3)
          *pTextPara\SetIndents(pf1, pf2, pf3)
        EndIf
      EndIf
    EndIf
    ;
    ParaSetValue("spacebefore", GetSpaceBefore, SetSpaceBefore)
    ParaSetValue("spaceafter", GetSpaceAfter, SetSpaceAfter)
    ParaSetValue("style", GetStyle, SetStyle)
    ;
    If FindString(Style$, "linespacing")
      parameter$ = ExtractParameterForTOM(Style$, "linespacing")
      If parameter$ <> "default"
        param1$ = StringField(parameter$, 1, ",")
        param2$ = StringField(parameter$, 2, ",")
        If param1$ = "single"
          pl = #TomLineSpaceSingle
        ElseIf param1$ = "1pt5"
          pl = #TomLineSpace1pt5
        ElseIf param1$ = "double"
          pl = #TomLineSpaceDouble
        ElseIf param1$ = "atleast"
          pl = #TomLineSpaceAtLeast
        ElseIf param1$ = "exactly"
          pl = #TomLineSpaceExactly
        ElseIf param1$ = "multiple"
          pl = #TomLineSpaceMultiple
        ElseIf param1$ = "percent"
          pl = #TomLineSpacePercent
        EndIf
      EndIf
      If SetUnset = #TomTrue And parameter$ <> "default" 
        *pTextPara\SetLineSpacing(pl, ValF(param2$))
      Else
        *pTextPara\GetLineSpacing(@pf)
        *pTextPara\GetLineSpacingRule(gpl)
        If (pf = ValF(param2$) And gpl = pl) Or parameter$ = "default" Or SetUnset = #TomDefault
          *pParaDefault\GetLineSpacing(@pf)
          *pParaDefault\GetLineSpacingRule(@pl)
          *pTextPara\SetLineSpacing(pf, pl)
        EndIf
      EndIf
    EndIf
    *pTextPara\Release()
  EndIf
EndProcedure
;
Procedure.s TOM_GetFontStyles(GadgetID, StartPos = -2, EndPos = -2)
  ;
  ; GadgetID must be the number of an EditorGadget.
  ; This procedure examines the styles of the text range
  ; defined by StartPos->EndPos and returns a descriptive
  ; string.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;  
  Protected *pTextFont.ITextFont_Fixed
  ;
  Protected pl.l = 0
  Protected pf.f = 0
  Protected ps.s = "", *BSTRString = 0
  ;
  Protected Style$ = "" ; Return value.
  ;
  ; Get a TextFont object for the range:
  *pTextFont = TOM_GetTextFontObj(GadgetID, StartPos, EndPos)
  ;
  If *pTextFont
    *pTextFont\GetBold(@pl)
    If pl = #TomTrue
      Style$ + "Bold, "
    EndIf
    *pTextFont\GetItalic(@pl)
    If pl = #TomTrue
      Style$ + "Italic, "
    EndIf
    *pTextFont\GetEmboss(@pl)
    If pl = #TomTrue
      Style$ + "Emboss, "
    EndIf
    *pTextFont\GetAllCaps(@pl)
    If pl = #TomTrue
      Style$ + "AllCaps, "
    EndIf
    *pTextFont\GetSmallCaps(@pl)
    If pl = #TomTrue
      Style$ + "SmallCaps, "
    EndIf
    *pTextFont\GetEngrave(@pl)
    If pl = #TomTrue
      Style$ + "Engrave, "
    EndIf
    *pTextFont\GetShadow(@pl)
    If pl = #TomTrue
      Style$ + "Shadow, "
    EndIf
    *pTextFont\GetOutline(@pl)
    If pl = #TomTrue
      Style$ + "OutLine, "
    EndIf
    *pTextFont\GetUnderline(@pl)
    If pl = #TomSingle
      Style$ + "Underline(Single), "
    ElseIf pl = #TomWords
      Style$ + "Underline(Words), "
    ElseIf pl = #TomDouble
      Style$ + "Underline(Double), "
    ElseIf pl = #TomDotted
      Style$ + "Underline(Dotted), "
    ElseIf pl = #TomDash
      Style$ + "Underline(Dash), "
    ElseIf pl = #TomDashDot
      Style$ + "Underline(DashDot), "
    ElseIf pl = #TomDashDotDot
      Style$ + "Underline(DashDotDot), "
    ElseIf pl = #TomWave
      Style$ + "Underline(Wave), "
    ElseIf pl = #TomThick
      Style$ + "Underline(Thick), "
    ElseIf pl = #TomHair
      Style$ + "Underline(Hair), "
    ElseIf pl = #TomDoubleWave
      Style$ + "Underline(DoubleWave), "
    ElseIf pl = #TomHeavyWave
      Style$ + "Underline(HeavyWave), "
    ElseIf pl = #TomLongDash
      Style$ + "Underline(LongDash), "
    ElseIf pl = #TomThickDash
      Style$ + "Underline(ThickDash), "
    ElseIf pl = #TomThickDashDot
      Style$ + "Underline(ThickDashDot), "
    ElseIf pl = #TomThickDashDotDot
      Style$ + "Underline(ThickDashDotDot), "
    ElseIf pl = #TomThickDotted
      Style$ + "Underline(ThickDotted), "
    ElseIf pl = #TomThickLongDash
      Style$ + "Underline(ThickLongDash), "
    EndIf
    *pTextFont\GetStrikeThrough(@pl)
    If pl = #TomTrue
      Style$ + "StrikeThrough, "
    EndIf
    *pTextFont\GetSubscript(@pl)
    If pl = #TomTrue
      Style$ + "Subscript, "
    EndIf
    *pTextFont\GetSuperscript(@pl)
    If pl = #TomTrue
      Style$ + "Superscript, "
    EndIf
    *pTextFont\GetHidden(@pl)
    If pl = #TomTrue
      Style$ + "Hidden, "
    EndIf
    *pTextFont\GetProtected(@pl)
    If pl = #TomTrue
      Style$ + "Protected, "
    EndIf
    *pTextFont\GetSize(@pf)
    Style$ + "Size(" + StrF(pf) + "), "
    *pTextFont\GetSpacing(@pf)
    If pf
      Style$ + "Spacing(" + StrF(pf) + "), "
    EndIf
    *pTextFont\GetPosition(@pf)
    If pf
      Style$ + "Position(" + StrF(pf) + "), "
    EndIf
    *pTextFont\GetKerning(@pf)
    If pf
      Style$ + "Kerning(" + StrF(pf) + "), "
    EndIf
    *pTextFont\GetBackColor(@pl)
    If pl <> #TomAutoColor
      Style$ + "BackColor(" + Str(pl) + "), "
    EndIf
    *pTextFont\GetForeColor(@pl)
    If pl <> #TomAutoColor
      Style$ + "ForeColor(" + Str(pl) + "), "
    EndIf
    *pTextFont\GetWeight(@pl)
    If pl <> 400
      Style$ + "Weight(" + Str(pl) + "), "
    EndIf
    *pTextFont\GetStyle(@pl)
    If pl
      Style$ + "Style(" + Str(pl) + "), "
    EndIf
    *pTextFont\GetName(@*BSTRString)
    ps = PeekS(*BSTRString, -1, #PB_Unicode)
    SysFreeString_(*BSTRString)
    Style$ + "Name(" + ps + ")"
    ;
    *pTextFont\Release()
  EndIf
  If Right(Style$, 2) = ", "
    Style$ = Left(Style$, Len(Style$) - 2)
  EndIf
  ProcedureReturn Style$
EndProcedure
;
Procedure.s TOM_GetParaStyles(GadgetID, StartPos = -2, EndPos = -2)
  ;
  ; GadgetID must be the number of an EditorGadget.
  ; This procedure examines the styles of the paragraphe(s)
  ; containing the text range defined by StartPos->EndPos
  ; and returns a descriptive string.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;  
  Protected *pTextPara.ITextPara_Fixed
  ;
  Protected pl.l = 0
  Protected pf.f = 0
  ;
  Protected Style$ = "" ; Return value.
  ;
  ; Get a TextPara object for the range:
  *pTextPara = TOM_GetTextParaObj(GadgetID, StartPos, EndPos)
  ;
  If *pTextPara
    *pTextPara\GetAlignment(@pl)
    ;
    If pl = #TomAlignLeft
      Style$ + "Align(Left), "
    ElseIf pl = #TomAlignCenter
      Style$ + "Align(Center), "
    ElseIf pl = #TomAlignRight
      Style$ + "Align(Right), "
    ElseIf pl = #TomAlignJustify
      Style$ + "Align(Justify), "
    ElseIf pl = #TomAlignBar
      Style$ + "Align(Bar), "
    ElseIf pl = #TomAlignInterLetter
      Style$ + "Align(InterLetter), "
    ElseIf pl = #TomAlignScaled
      Style$ + "Align(Scaled), "    
    EndIf
    ;
    *pTextPara\GetLeftIndent(@pf)
    If pf
      Style$ + "LeftIndent(" + StrF(pf) + "), "
    EndIf
    *pTextPara\GetRightIndent(@pf)
    If pf
      Style$ + "RightIndent(" + StrF(pf) + "), "
    EndIf
    *pTextPara\GetFirstLineIndent(@pf)
    If pf <> 0
      Style$ + "FirstLineIndent(" + StrF(pf) + "), "
    EndIf
    ;
    *pTextPara\GetSpaceBefore(@pf)
    If pf
      Style$ + "SpaceBefore(" + StrF(pf) + "), "
    EndIf
    *pTextPara\GetSpaceAfter(@pf)
    If pf
      Style$ + "SpaceAfter(" + StrF(pf) + "), "
    EndIf
    ;
    *pTextPara\GetStyle(@pl)
    If pl <> -1
      Style$ + "Style(" + Str(pl) + "), "
    EndIf
    ;
    *pTextPara\GetLineSpacingRule(@pl)
    *pTextPara\GetLineSpacing(@pf)
    If pl = #TomLineSpace1pt5
      Style$ + "LineSpacing(1pt5), "
    ElseIf pl = #TomLineSpaceDouble
      Style$ + "LineSpacing(Double), "
    ElseIf pl = #TomLineSpaceAtLeast
      Style$ + "LineSpacing(AtLeast," + StrF(pf) + "), "
    ElseIf pl = #TomLineSpaceExactly
      Style$ + "LineSpacing(Exactly," + StrF(pf) + "), "
    ElseIf pl = #TomLineSpaceMultiple
      Style$ + "LineSpacing(Multiple," + StrF(pf) + "), "
    ElseIf pl = #TomLineSpacePercent
      Style$ + "LineSpacing(Percent," + StrF(pf) + "), "
    EndIf
    ;
    *pTextPara\Release()
  EndIf
  If Right(Style$, 2) = ", "
    Style$ = Left(Style$, Len(Style$) - 2)
  EndIf
  ProcedureReturn Style$
EndProcedure
;
CompilerIf Not Defined(GetDIBHandleFromImage, #PB_Procedure)
  ; Cette procédure est également définie dans IDataObject_Helper.pb
  ;
  Procedure GetDIBHandleFromImage(HBitmap)
    ;
    ; This procedure works with a PureBasic image ID
    ; or with a handle to the image.
    ; It encapsulates the Bitmap data into a DIB, which is
    ; a format handled by many Windows functions, 
    ; especially the GDI API functions.
    ;
    If IsImage(HBitmap)
      HBitmap = ImageID(HBitmap)
    EndIf
    ;
    Protected TemporaryDC.L, TemporaryBitmap.BITMAP, TemporaryBitmapInfo.BITMAPINFO
    Protected hDib, *Buffer
    Protected BitmapSize
    
    ; Create a temporary device context (DC):
    TemporaryDC = CreateDC_("DISPLAY", #Null, #Null, #Null)
    
    ; Retrieve information about the bitmap (HBitmap):
    GetObject_(HBitmap, SizeOf(BITMAP), @TemporaryBitmap)
    
    ; Initialize the BITMAPINFOHEADER information:
    TemporaryBitmapInfo\bmiHeader\biSize        = SizeOf(BITMAPINFOHEADER)
    TemporaryBitmapInfo\bmiHeader\biWidth       = TemporaryBitmap\bmWidth
    TemporaryBitmapInfo\bmiHeader\biHeight      = TemporaryBitmap\bmHeight
    TemporaryBitmapInfo\bmiHeader\biPlanes      = 1
    TemporaryBitmapInfo\bmiHeader\biBitCount    = 32                         ; 32 bits / pixel
    TemporaryBitmapInfo\bmiHeader\biCompression = #BI_RGB
    
    ; Calculate the required size for the DIB:
    BitmapSize = TemporaryBitmap\bmWidth * TemporaryBitmap\bmHeight * (TemporaryBitmapInfo\bmiHeader\biBitCount / 8)
    
    ; Allocate memory for the DIB (Device Independent Bitmap)
    hDib = GlobalAlloc_(#GMEM_MOVEABLE, BitmapSize + SizeOf(BITMAPINFOHEADER))
    
    If hDib
      *Buffer = GlobalLock_(hDib)
      If *Buffer
        ; Copy the BITMAPINFOHEADER header into memory
        CopyMemory(@TemporaryBitmapInfo\bmiHeader, *Buffer, SizeOf(BITMAPINFOHEADER))
        
        ; Copy the bitmap bits into memory after the header
        GetDIBits_(TemporaryDC, HBITMAP, 0, TemporaryBitmap\bmHeight, *Buffer + SizeOf(BITMAPINFOHEADER), TemporaryBitmapInfo, #DIB_RGB_COLORS)
        
        GlobalUnlock_(hDib)
      Else
        ; Lock failed, free the memory
        GlobalFree_(hDib)
        hDib = 0
      EndIf
    EndIf
    
    ; Free and delete the device context (DC)
    DeleteDC_(TemporaryDC)
    ;
    ProcedureReturn hDib
    ;
  EndProcedure
CompilerEndIf
;
Procedure SetIDataObjectImage(*MyDataObject.IDataObject, Image)
  ;
  ; Feed an IDataObject with a DIB version of an image.
  ;
  Protected MyFormatEtc.FormatEtc, MyStgMed.StgMedium
  Protected hBitmap, hDib, SC
  ;
  If *MyDataObject = 0 Or Image = 0
    ProcedureReturn #False
  EndIf
  ; ________________________________________________
  ;             Set Values for FORMATETC
  ;
  MyFormatEtc\tymed = #TYMED_HGLOBAL
  MyFormatEtc\cfFormat = #CF_DIB            ; Set the format
  MyFormatEtc\ptd = #Null                   ; Target Device = Screen
  MyFormatEtc\dwAspect = #DVASPECT_CONTENT  ; Level of detail = Full content
  MyFormatEtc\lindex = -1                   ; Index = Not applicable
  ; ________________________________________________
  ;             Set Values for STGMEDIUM
  ;
  MyStgMed\pUnkForRelease = #Null        ; Use ReleaseStgMedium_() APIFunction
  MyStgMed\tymed = MyFormatEtc\tymed     ; MyStgMed and MyFormatEtc must have the same value
                                         ;   in their respective 'tymed' fields.
  ; ________________________________________________
  ;             Handle Data
  ;
  If IsImage(Image)
    Image = ImageID(Image)
  EndIf
  hBitmap = CopyImage_(Image, #IMAGE_BITMAP, 0, 0, #LR_COPYRETURNORG)
  If hBitmap = 0
    ProcedureReturn #False
  EndIf
  ;
  hDib = GetDIBHandleFromImage(hBitmap)
  MyStgMed\hGlobal = hDib
  DeleteObject_(hBitmap)
  If MyStgMed\hGlobal = 0
    ProcedureReturn #False
  EndIf
  ; _________________________________________________________________
  ;
  ;             Finally, call SetData on the IDataObject
  ;
  SC = *MyDataObject\SetData(@MyFormatEtc, @MyStgMed, #True)
  ; The IDataObject make a copy of the given data and clean the original ones, because
  ; last parametre is set to '#True'
  ;
  If SC <> #S_OK
    ProcedureReturn #False
  EndIf
  ;
  ProcedureReturn #True
EndProcedure
;
Procedure SetIDataObjectText(*MyDataObject.IDataObject, StringData$)
  ;
  ; Feed an IDataObject with Text or RTF content
  ;
  Protected MyFormatEtc.FormatEtc, MyStgMed.StgMedium
  Protected formatName$, formatValue, StringSize, hGlobal, *DataBuffer
  Protected SC
  ;
  If *MyDataObject = 0
    ProcedureReturn #False
  EndIf
  ;
  ; ________________________________________________
  ;             Set Values for FORMATETC
  ;
  MyFormatEtc\tymed = #TYMED_HGLOBAL
  If Left(StringData$, 5) = "{\rtf" Or Left(StringData$, 6) = "{\urtf"
    formatName$ = "RTF in UTF8"
    formatValue = RegisterClipboardFormat_(@formatName$)
    If formatValue = 0
      ProcedureReturn #False
    EndIf
    MyFormatEtc\cfFormat = formatValue      ; Set the format
    StringSize = Len(StringData$) + 1
  Else
    MyFormatEtc\cfFormat = #CF_UNICODETEXT  ; Set the format
    StringSize = (Len(StringData$) + 1) * 2 ; Each Unicode Char is 2 bytes long
  EndIf
  MyFormatEtc\ptd = #Null                   ; Target Device = Screen
  MyFormatEtc\dwAspect = #DVASPECT_CONTENT  ; Level of detail = Full content
  MyFormatEtc\lindex = -1                   ; Index = Not applicable
  ; ________________________________________________
  ;             Set Values for STGMEDIUM
  ;
  MyStgMed\pUnkForRelease = #Null        ; Use ReleaseStgMedium_() APIFunction
  MyStgMed\tymed = MyFormatEtc\tymed     ; MyStgMed and MyFormatEtc must have the same value
                                         ;   in their respective 'tymed' fields.
  ; ________________________________________________
  ;             Handle Data
  ;
  hGlobal = GlobalAlloc_(#GMEM_MOVEABLE, StringSize)
  If hGlobal = 0
    ProcedureReturn #False
  Else
    *DataBuffer = GlobalLock_(hGlobal)
    If *DataBuffer
      If MyFormatEtc\cfFormat = #CF_UNICODETEXT
        PokeS(*DataBuffer, StringData$)
      Else
        PokeS(*DataBuffer, StringData$, -1, #PB_UTF8)
      EndIf
      GlobalUnlock_(hGlobal)
      MyStgMed\hGlobal = hGlobal
    Else
      GlobalFree_(hGlobal)
      ProcedureReturn #False
    EndIf
  EndIf
  ; _________________________________________________________________
  ;
  ;             Finally, call SetData on the IDataObject
  ;
  SC = *MyDataObject\SetData(@MyFormatEtc, @MyStgMed, #True)
  ; The IDataObject make a copy of the given data and clean the original ones, because
  ; last parametre is set to '#True'
  ;
  If SC <> #S_OK
    ProcedureReturn #False
  EndIf
  ;
  ProcedureReturn #True
EndProcedure
;
Procedure TOM_InsertImage(GadgetID, Image, StartPos = -2, PosEnd = -2)
  ; 'GadgetID' must be an EditorGadget number or a handle (pointer) to a RichEdit Control.
  ; 'Image' can be a PureBasic image number or a handle to it.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  ;
  Protected       *pTextRange.ITextRange
  Protected               var.VARIANT
  Protected         *mDataObj.IDataObject
  Protected                SC
  ;
  If Image
    ;
    *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, PosEnd)
    ;
    If *pTextRange
      ; Create an IDataObject:
      *mDataObj = CreateIDataObject()
      ; Fill the IDataObject with image:
      If SetIDataObjectImage(*mDataObj, Image)
        ;
        ; Create a VARIANT with a pointer to the IDataObject
        var.VARIANT
        var\vt = #VT_UNKNOWN ; | #VT_BYREF
        var\ppunkVal = *mDataObj
        ;
        ; Paste the IDataObject into the range:
        ; Clipboard in NOT used. The paste operation is done from the IDataobject.
        SC = *pTextRange\Paste(@var, 0)
        ; Because a IDataObject\Release is asked by ITextRange just after Paste,
        ; the IDataObject is destroyed during this operation. It's not necessary
        ; to clean it.
      EndIf
      *pTextRange\Release()
    EndIf
  Else
    SC = 1
  EndIf
  ;
  ProcedureReturn SC
EndProcedure
;
Procedure TOM_InsertText(GadgetID, Text$, StartPos = -2, PosEnd = -2)
  ; 'GadgetID' must be an EditorGadget number or a handle (pointer) to a RichEdit Control.
  ; 'Text$' can contain simple texte or RTF text or RTF UTF8 Text.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  Protected       *pTextRange.ITextRange
  Protected               var.VARIANT
  Protected         *mDataObj.IDataObject
  Protected                SC
  ;
  *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, PosEnd)
  ;
  If *pTextRange
    ; Create an IDataObject:
    *mDataObj = CreateIDataObject()
    ; Fill the IDataObject with image:
    If SetIDataObjectText(*mDataObj, Text$)
      ;
      ; Create a VARIANT with a pointer to the IDataObject
      var.VARIANT
      var\vt = #VT_UNKNOWN ; | #VT_BYREF
      var\ppunkVal = *mDataObj
      ;
      ; Paste the IDataObject into the range:
      ; Clipboard in NOT used. The paste operation is done from the IDataobject.
      SC = *pTextRange\Paste(@var, 0)
      ; Because a IDataObject\Release is asked by ITextRange just after Paste,
      ; the IDataObject is destroyed during this operation. It's not necessary
      ; to clean it.
    EndIf
    *pTextRange\Release()
  EndIf
  ;
  ProcedureReturn SC
EndProcedure
;
Procedure.s TOM_GetAvailableFormats(GadgetID, StartPos = -2, PosEnd = -2)
  ;
  ; 'GadgetID' must be an EditorGadget number or a handle (pointer) to a RichEdit Control.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  Protected       *pTextRange.ITextRange
  Protected               var.VARIANT
  Protected          mDataObj.IDataObject
  Protected                SC
  Protected        FormatEtc.FORMATETC
  Protected       enumFormat.IEnumFORMATETC
  Protected         Formats$
  ;
  *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, PosEnd)
  ;
  If *pTextRange
    ; Create a VARIANT with a pointer to an empty IDataObject
    var.VARIANT
    var\vt = #VT_UNKNOWN | #VT_BYREF
    var\ppunkVal = @mDataObj
    ;
    ; Fill the IDataObject with the range content
    ; Clipboard in NOT used. The copy operation is done with the IDataobject.

    SC = *pTextRange\Copy(@var)
    ;
    ; Read available formats for this range:
    If SC = #S_OK
      SC = mDataObj\EnumFormatEtc(#DATADIR_GET,  @enumFormat)
      If SC = #S_OK And enumFormat
        ; Parcourir tous les formats disponibles
        While enumFormat\Next(1, @formatEtc, #Null) = #S_OK
          Formats$ + GetFormatName(formatEtc\cfFormat) + Chr(13)
        Wend
        ; Free enumerator
        enumFormat\Release()
      EndIf
    EndIf
    ;
  EndIf
  ;
  ProcedureReturn Formats$
EndProcedure
;
Procedure.s TOM_GetText(GadgetID, Format$ = "", StartPos = -2, PosEnd = -2)
  ;
  ; 'GadgetID' must be an EditorGadget number or a handle (pointer) to a RichEdit Control.
  ; 'Format$' can be null or can contain:
  ;     • "text", "texte", "cf_text", "sf_text" or "CF_UNICODETEXT" for simple texte
  ;     • "RTF", "sf_rtf", "cf_rtf" or "Rich Text Format" to get RTF content
  ;     • "RTF in UTF8" to get RTF content in UTF8
  ;     You can call 'TOM_GetAvailableFormats()' to explore other possibilities.
  ;
  ; If StartPos is omitted or StartPos = -2, the range is the current selection.
  ; If EndPos is omitted or EndPos = -2 or EndPos < StartPos, --> EndPos = StartPos.
  ; IF StartPos = -1, the range is set as an insertion point at the end of GadgetID content.
  ; IF   EndPos = -1, the range is set from StatPos to the end of GadgetID content.
  ;
  Protected       *pTextRange.ITextRange
  Protected               var.VARIANT
  Protected          mDataObj.IDataObject
  Protected        FormatEtc.FORMATETC
  Protected       enumFormat.IEnumFORMATETC
  Protected             stgm.STGMEDIUM
  ;
  Protected *pGlobalMemory, BufferSize, Text$, SC
  ;
  If Format$ = "" Or LCase(Format$) = "text" Or LCase(Format$) = "texte" Or FindString(LCase(Format$), "sf_text") Or FindString(LCase(Format$), "cf_text")
    Format$ = "CF_UNICODETEXT"
  EndIf
  If LCase(Format$) = "rtf" Or FindString(LCase(Format$), "sf_rtf") Or FindString(LCase(Format$), "cf_rtf")
    Format$ = "Rich Text Format"
  EndIf
  ;
  *pTextRange = TOM_GetTextRangeObj(GadgetID, StartPos, PosEnd)
  ;
  If *pTextRange
    ; Create a VARIANT with a pointer to an empty IDataObject
    var.VARIANT
    var\vt = #VT_UNKNOWN | #VT_BYREF
    var\ppunkVal = @mDataObj
    ;
    ; Fill the IDataObject with the range content
    ; Clipboard in NOT used. The copy operation is done with the IDataobject.
    ;
    SC = *pTextRange\Copy(@var)
    ;
    ; Find the requested format into IDataObject data:
    If SC = #S_OK
      SC = mDataObj\EnumFormatEtc(#DATADIR_GET,  @enumFormat)
      If SC = #S_OK And enumFormat
        ; Parcourir tous les formats disponibles
        SC = 1
        While enumFormat\Next(1, @formatEtc, #Null) = #S_OK
          If LCase(GetFormatName(formatEtc\cfFormat)) = LCase(Format$)
            SC = #S_OK
            Break
          EndIf
        Wend
        ; Free enumerator
        enumFormat\Release()
      Else
        SC = 1
      EndIf
    EndIf
    ;
    ; Get content from IDataObject
    If SC = #S_OK
      SC = mDataObj\getData(formatEtc, @stgm)
      If SC = #S_OK
        If stgm\tymed = #TYMED_HGLOBAL
          ; Verrouiller la mémoire globale pour y accéder
          *pGlobalMemory = GlobalLock_(stgm\hGlobal)
          If *pGlobalMemory = 0
            SC = 1
          Else
            ; Obtenir la taille de la mémoire globale
            BufferSize = GlobalSize_(stgm\hGlobal)
            If LCase(Format$) = "cf_unicodetext"
              Text$ = PeekS(*pGlobalMemory, BufferSize / 2)
            ElseIf FindString(LCase(Format$), "utf8")
              Text$ = PeekS(*pGlobalMemory, BufferSize, #PB_UTF8)
            Else
              Text$ = PeekS(*pGlobalMemory, BufferSize, #PB_Ascii)
            EndIf
          EndIf
          GlobalUnlock_(stgm\hGlobal)
        EndIf
      EndIf
    EndIf
    ;
    mDataObj\Release()
    *pTextRange\Release()
  EndIf
  ;
  ProcedureReturn Text$
EndProcedure
;
Procedure TOM_InsertTaggedJPGImageFromFile(GadgetID, FileAdr$, marker$ = "", StartPos = -2, EndPos = -2)
  ;
  Protected Img
  ;
  UseJPEGImageDecoder()
  Img = LoadImage(#PB_Any, FileAdr$)
  If Img
    StartPos = TOM_GetStartPos(GadgetID, StartPos)
    TOM_InsertImage(GadgetID, Img, StartPos, EndPos)
    FreeImage(Img)
    If marker$
      marker$ = "{\rtf1\v " + marker$ + "}"
      TOM_InsertText(GadgetID, marker$, StartPos)
    EndIf
  Else
    MessageRequester("Oops!", "Can't load image!" + FileAdr$)
  EndIf
EndProcedure
;
Procedure TOM_ComputeWordPosition(GadgetID, MyWord$, StartPos = 0)
  ; Look for the position of 'MyWord$' inside the GadgetID's content.
  ;
  Protected EditorText$, Result
  ;
  ; Get the GadgetID's content
  EditorText$ = GetGadgetText(GadgetID)
  ; An ajustment is necessary to be able to compute position from the text obtained,
  ; because The TOM system, as all other RichEdit interfaces, count only one
  ; character for the EndOfLine (Carriage return). But the text we have now has
  ; two characters for the EndOfLine: Chr(10) + Chr(13)    (CRLF).
  ; So, we delete Chr(10) to keep only the carriage return (one sole character).
  EditorText$ = ReplaceString(EditorText$, Chr(10), "")
  ; Now, the positions which we'll get from FindString will be compatible with
  ; our needs.
  Result = FindString(EditorText$, MyWord$, StartPos)
  ;
  ; The returned value is Result less one, because PureBasic attribute position '1' 
  ; to the first character, while Windows's functions attribute position '0' to it.
  ;
  ; We set the result to Windows needs:
  ProcedureReturn Result - 1
EndProcedure
;
Procedure TOM_SetGadgetAsRichEdit(GadgetID)
  SendMessage_(GadgetID(GadgetID), #EM_SETTEXTMODE, #TM_RICHTEXT, 0)
  SendMessage_(GadgetID(GadgetID), #EM_SETTARGETDEVICE, #Null, 0);<<--- Automatic carriage return.
  SendMessage_(GadgetID(GadgetID), #EM_LIMITTEXT, -1, 0)             ; Set unlimited content size.
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  ; The following won't run when this file is used as 'Included'.
  ;
  If OpenWindow(0, 200, 200, 600, 400, "TOM Example")
    EGadget = EditorGadget(#PB_Any, 10, 10, 580, 300)
    TGadget = TextGadget(#PB_Any, 10, 320, 580, 70, "")
    ; TOM_SetFontStyles() works on any
    ; EditorGadget without any special configuration.
    ; However, TOM_SetParamStyles() requires that the
    ; GadgetID be set up as a RichEdit GadgetID:
    TOM_SetGadgetAsRichEdit(EGadget)
    ;
    ; Note that the TOM library can't be used with TextGadgets or StringGadgets.
    
    AddGadgetItem(EGadget, -1, "This is a sample text with image:")
    ;
    MyImage = CreateImage(#PB_Any, 80, 80)
    StartDrawing(ImageOutput(MyImage))
    Box(0, 0, 80, 80, RGB(255, 255, 255))
    Circle(40, 40, 30, RGB(255, 0, 0))
    Box(50, 10, 60, 40, RGB(0, 0, 255))
    DrawText(18, 40, "Image")
    StopDrawing()
    TOM_InsertImage(EGadget, MyImage, 33)
  
    ; Apply styles (bold, italic, underline, size: 15, position on line: 4, Times font) to characters from 10 to 15
    TOM_SetFontStyles(EGadget, "Size(15), Bold, Italic, Underline(), Name(Times), position(4)", 10, 16)
    ;
    ; Apply Wave underline to characters from 0 to 4
    TOM_SetFontStyles(EGadget, "Underline(Wave)", 0, 5)
    ;
    ; Center first line:
    TOM_SetParaStyles(EGadget, "Align(Center)", 10, 16)
    ;
    ; Describe styles of character 11 :
    Info$ = "Character 11: " + TOM_GetFontStyles(EGadget, 11, 12) + Chr(13)
    ; Describe styles of character 3:
    Info$ + "Character 3: " + TOM_GetFontStyles(EGadget, 3, 4) + Chr(13)
    ;
    ; Copy style from character 10:
    *TextFontObjet.ITextFont_Fixed = TOM_GetTextFontObj(EGadget, 10, 11, #TomTrue)
    ; Apply this style to characters from 18 to 19:
    TOM_ApplyTextFont(EGadget, *TextFontObjet, 18, 20)
    ; Free memory:
    *TextFontObjet\Release()
    ;
    ; Unset styles applied to character 18:
    TOM_SetFontStyles(EGadget, "Size(15), Bold, Italic, Underline(), Name(Times), position(4)", 15, 16, #TomFalse)
    
    AddGadgetItem(EGadget, -1, "")
    AddGadgetItem(EGadget, -1, "This is another sample text with more words to see other possibilities of setting for paragraphe, including FirstLineIndent for this one.")
    TOM_SetFontStyles(EGadget, "ForeColor($0000D0), Bold", 143, 159)
    TOM_SetParaStyles(EGadget, "Align(Left), FirstLineIndent(10)", 143, 159)
    ;
    AddGadgetItem(EGadget, -1, "")
    AddGadgetItem(EGadget, -1, "This is another sample text with more words to see other possibilities of setting for paragraphe, including LeftIndent for this one.")
    TOM_SetFontStyles(EGadget, "ForeColor($0000D0), Bold", 282, 293)
    TOM_SetParaStyles(EGadget, "FirstLineIndent(0), LeftIndent(10)", 282, 293)
    ;
    AddGadgetItem(EGadget, -1, "This is another sample text with more words to see other possibilities of setting for paragraphe, including RightIndent, Justify, LineSpacing and SpaceBefore for this one. Qui sommes-nous ? Quelle est notre essence, notre véritable identité ? Ces questions nous préoccupent depuis toujours.")
    TOM_SetFontStyles(EGadget, "ForeColor($DE7723), Bold", 415, 428)
    TOM_SetParaStyles(EGadget, "LeftIndent(0), RightIndent(40), Align(Justify), LineSpacing(exactly,16), SpaceBefore(3)", 415, 428)
    ;
      ; Describe styles of paragraphe including character 411:
    Info$ + "Character 411: " + TOM_GetParaStyles(EGadget, 411, 412) + Chr(13)
    ;
    ; If you’re as bored as I am, calculating the character positions to determine which range to apply styles to,
    ; you can do it this way:
    StartPos = TOM_ComputeWordPosition(EGadget, "Justify")
    EndPos = StartPos + Len("justify")
    TOM_SetFontStyles(EGadget, "BackColor($00D0D0), Bold", StartPos, EndPos)
    StartPos = TOM_ComputeWordPosition(EGadget, "LineSpacing", StartPos)
    EndPos = StartPos + Len("LineSpacing")
    TOM_SetFontStyles(EGadget, "ForeColor($0000D0), Bold", StartPos, EndPos)
    StartPos = TOM_ComputeWordPosition(EGadget, "SpaceBefore", StartPos)
    EndPos = StartPos + Len("SpaceBefore")
    TOM_SetFontStyles(EGadget, "ForeColor($0000D0), Bold", StartPos, EndPos)
    ;
    SetGadgetText(TGadget, Info$)
    ;
    Repeat
    Until WaitWindowEvent() = #PB_Event_CloseWindow
  EndIf
CompilerEndIf


DataSection
  IID_ITextDocument2: 
    Data.l $01C25500
    Data.w $4268, $11D1
    Data.b $88, $3A, $3C, $8B, $00, $C1, $00, $00

  IID_ITextRange2: 
    Data.l $3103CCC3
    Data.w $004D, $4B37
    Data.b $98, $73, $26, $AE, $7C, $0B, $56, $8F
EndDataSection

; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; CursorPosition = 737
; FirstLine = 724
; Folding = ------
; EnableXP
; DPIAware