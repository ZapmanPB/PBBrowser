;***********************************************************************
;
;                            PrefWindow.pb
;                      Part of PBBrowser project
;                          Zapman - Dec 2024
;
;          Open the preferences window and manage user choices.
;
;***********************************************************************
;
Structure ParametersStruct
  ParamWindow.i ; Window number
  ;
  BorderColor.i
  ;
  ; Layout parameters:
  Margins.i
  HorizSpaceInterGadget.i
  OKCancelButtonsWidth.i
  ButtonsHeight.i
  RightButtonsWidth.i
  ScrollBarWidth.i
  InnerWidth.i
  HPosFields.i
  LegendWidth.i
  InputWidth.i
  TextHeight.i
  LineHeight.i
  ButtonHeight.i
  ;
  ; Gadget numbers:
  AreaGadget.i
  TColorTheme.i
  SColorTheme.i
  BColorTheme.i
  TNotRecommended.i
  LowCanvasBar.i
  BSave.i
  BCancel.i
  ;
  IWhiteOver.i
  ;
EndStructure
;
Global ParamWAG.ParametersStruct
;
Procedure PW_SetFontsToPWGadgets()
  ;
  Shared ParamWAG.ParametersStruct
  Static AllGadgetsFont
  Static TitleFont
  ;
  ForEach PBBParameters()
    If PBBParameters()\Name$ = "PBBAllGadgetsFont"
      If IsFont(AllGadgetsFont)
        FreeFont(AllGadgetsFont)
      EndIf
      AllGadgetsFont = FontID(GetFontFromDescription(GetGadgetText(PBBParameters()\InputGadget)))
    ElseIf PBBParameters()\Name$ = "PBBTitleFont"
      If IsFont(TitleFont)
        FreeFont(TitleFont)
      EndIf
      TitleFont = FontID(GetFontFromDescription(GetGadgetText(PBBParameters()\InputGadget)))
    EndIf
  Next
  ForEach PBBParameters()
    If IsGadget(PBBParameters()\TitleGadget)
      SetGadgetFont(PBBParameters()\TitleGadget, TitleFont)
    EndIf
    SetGadgetFont(PBBParameters()\LegendGadget, AllGadgetsFont)
    SetGadgetFont(PBBParameters()\StandardButtonGadget, AllGadgetsFont)
    If PBBParameters()\Type = #PBBP_FontStyle And IsGadget(PBBParameters()\InputGadget)
      If IsFont(PBBParameters()\InputGadgetFont)
        FreeFont(PBBParameters()\InputGadgetFont)
      EndIf
      PBBParameters()\InputGadgetFont = GetFontFromDescription(GetGadgetText(PBBParameters()\InputGadget))
      SetGadgetFont(PBBParameters()\InputGadget, FontID(PBBParameters()\InputGadgetFont))
    ElseIf PBBParameters()\Type <> #PBBP_Color 
      SetGadgetFont(PBBParameters()\InputGadget, AllGadgetsFont)
    EndIf
  Next
  ;
  SetGadgetFont(ParamWAG\TNotRecommended, AllGadgetsFont)
  ;
  SetGadgetFont(ParamWAG\BSave, AllGadgetsFont)
  SetGadgetFont(ParamWAG\BCancel, AllGadgetsFont)
  SetGadgetFont(ParamWAG\TColorTheme, AllGadgetsFont)
  SetGadgetFont(ParamWAG\SColorTheme, AllGadgetsFont)
  SetGadgetFont(ParamWAG\BColorTheme, AllGadgetsFont)
EndProcedure
;
Procedure PW_ResizeWindowAndGadgets()
  ;
  Shared ParamWAG.ParametersStruct
  ;
  Protected WWidth = WindowWidth(ParamWAG\ParamWindow)
  Protected WHeight = WindowHeight(ParamWAG\ParamWindow)
  Protected ButtonExtend = 1.5, InputWidthMax, Vpos, LastCategory$, TNotRecommended
  ;
  ResizeGadget(ParamWAG\AreaGadget, 0, 0, WWidth, WHeight - 2 * ParamWAG\Margins - ParamWAG\ButtonsHeight)
  ;
  SetGadgetAttribute(ParamWAG\AreaGadget, #PB_ScrollArea_InnerWidth, WWidth)
  ParamWAG\InnerWidth = WWidth
  If IsVertScrollBarVisible(ParamWAG\AreaGadget)
    ParamWAG\InnerWidth = WWidth - ParamWAG\ScrollBarWidth
  Else
    ParamWAG\InnerWidth = WWidth
  EndIf
  SetGadgetAttribute(ParamWAG\AreaGadget, #PB_ScrollArea_InnerWidth, ParamWAG\InnerWidth)
  ;
  InputWidthMax = ParamWAG\InnerWidth - ParamWAG\HPosFields - ParamWAG\RightButtonsWidth - ParamWAG\HorizSpaceInterGadget - ParamWAG\Margins
  If ParamWAG\InputWidth > InputWidthMax : ParamWAG\InputWidth = InputWidthMax : EndIf
  ;
  Vpos = ParamWAG\Margins
  ;
  LastCategory$ = ""
  TNotRecommended = 0
  ForEach PBBParameters()
    If PBBParameters()\Category$
      If PBBParameters()\Category$ <> LastCategory$
        If LastCategory$
          If TNotRecommended
            TNotRecommended = 0
            Vpos + ParamWAG\LineHeight + 1
            ResizeGadget(ParamWAG\TNotRecommended, ParamWAG\Margins, Vpos + 2, ParamWAG\InnerWidth, ParamWAG\TextHeight)
            Vpos - 8
          EndIf
          Vpos + ParamWAG\LineHeight + 4
          ResizeGadget(PBBParameters()\CanvasLineGadget, ParamWAG\Margins, Vpos, ParamWAG\InnerWidth - ParamWAG\Margins * 2, 1)
          FillCanvasWithColor(PBBParameters()\CanvasLineGadget, ParamWAG\BorderColor, ParamWAG\BorderColor)
          Vpos + 8
        Else
          Vpos - 2
          TNotRecommended = 1
        EndIf
        ResizeGadget(PBBParameters()\TitleGadget, ParamWAG\Margins, Vpos + 2, ParamWAG\InnerWidth - ParamWAG\Margins * 2, ParamWAG\TextHeight)
        Vpos + ParamWAG\LineHeight
        
        If LastCategory$ = ""
          ResizeGadget(ParamWAG\TColorTheme, ParamWAG\Margins, Vpos + 2, 280, ParamWAG\TextHeight)
          ResizeGadget(ParamWAG\SColorTheme, ParamWAG\HPosFields, Vpos - ButtonExtend + 1, ParamWAG\InputWidth, ParamWAG\TextHeight + ButtonExtend * 2 - 2)
          ResizeGadget(ParamWAG\BColorTheme, ParamWAG\HPosFields + ParamWAG\InputWidth + ParamWAG\HorizSpaceInterGadget, Vpos - ButtonExtend, ParamWAG\RightButtonsWidth, ParamWAG\TextHeight + ButtonExtend * 2)
          Vpos + ParamWAG\LineHeight
        EndIf
        LastCategory$ = PBBParameters()\Category$
      Else
        Vpos + ParamWAG\LineHeight
      EndIf
      
      ResizeGadget(PBBParameters()\LegendGadget, ParamWAG\Margins, Vpos + 2, 280, ParamWAG\TextHeight)
      If PBBParameters()\Type = #PBBP_FontStyle
        ResizeGadget(PBBParameters()\InputGadget, ParamWAG\HPosFields, Vpos - ButtonExtend + 1, ParamWAG\InputWidth, ParamWAG\TextHeight + ButtonExtend * 2 - 2)
      ElseIf PBBParameters()\Type = #PBBP_Color
        ResizeGadget(PBBParameters()\InputGadget, ParamWAG\HPosFields, Vpos - ButtonExtend, ParamWAG\TextHeight + ButtonExtend * 2, ParamWAG\TextHeight + ButtonExtend * 2)
      Else
        ResizeGadget(PBBParameters()\InputGadget , ParamWAG\HPosFields, Vpos, ParamWAG\TextHeight, ParamWAG\TextHeight)
      EndIf
      ResizeGadget(PBBParameters()\StandardButtonGadget, ParamWAG\HPosFields + ParamWAG\InputWidth + ParamWAG\HorizSpaceInterGadget, Vpos - ButtonExtend, ParamWAG\RightButtonsWidth, ParamWAG\TextHeight + ButtonExtend * 2)
    EndIf    
  Next
  ;
  Vpos  = WHeight - 2 * ParamWAG\Margins - ParamWAG\ButtonsHeight
  ResizeGadget(ParamWAG\LowCanvasBar, 0, Vpos, WWidth, 1)
  FillCanvasWithColor(ParamWAG\LowCanvasBar, ParamWAG\BorderColor, ParamWAG\BorderColor)
  Vpos + ParamWAG\Margins
  ;
  ResizeGadget(ParamWAG\BSave, WWidth - ParamWAG\OKCancelButtonsWidth - ParamWAG\Margins, Vpos, ParamWAG\OKCancelButtonsWidth, ParamWAG\ButtonHeight)
  ResizeGadget(ParamWAG\BCancel, WWidth - 2 * ParamWAG\OKCancelButtonsWidth - ParamWAG\HorizSpaceInterGadget - ParamWAG\Margins, Vpos, ParamWAG\OKCancelButtonsWidth, ParamWAG\ButtonHeight)
  ;
  ; Redraw the WhiteBox with the new window dimensions:
  ParamWAG\IWhiteOver = WhiteBoxOverWindow(ParamWAG\ParamWindow)
  ;
EndProcedure
;
Procedure PW_FontRequester(Gadget)
  ;
  Protected GadgetText$ = GetGadgetText(Gadget)
  Protected Font$
  ;
  If FontRequester(Trim(StringField(GadgetText$, 1, ",")), Val(Trim(StringField(GadgetText$, 2, ","))), 0, 0, AttributesTextDescriptionToNumAttributes(GadgetText$))
    Font$ = SelectedFontName() + ", " + Str(SelectedFontSize())
    Font$ + NumAttributesToAttributesTextDescription(SelectedFontStyle())
    SetGadgetText(Gadget, Font$)
    PW_SetFontsToPWGadgets()
  EndIf
EndProcedure
;
Procedure PW_SetColorsToPWGadgets()
  ;
  Shared ParamWAG.ParametersStruct
  Protected WarningColor = RGB(150,120,20)
  Protected TitleColor
  ;
  ForEach PBBParameters()
    If PBBParameters()\Name$ = "PBBTitleColor"
      TitleColor = Val(PBBParameters()\ColorTempValue$)
      Break
    EndIf
  Next
  ;
  ForEach PBBParameters()
    If IsGadget(PBBParameters()\TitleGadget)
      SetGadgetColor(PBBParameters()\TitleGadget, #PB_Gadget_FrontColor, TitleColor)
    EndIf
    If PBBParameters()\Type = #PBBP_Color
      FillCanvasWithColor(PBBParameters()\InputGadget, Val(PBBParameters()\ColorTempValue$), ParamWAG\BorderColor)
    EndIf
    If IsGadget((PBBParameters()\LegendGadget)) And FindString(Right(GetGadgetText(PBBParameters()\LegendGadget), 4), "*")
      SetGadgetColor(PBBParameters()\LegendGadget, #PB_Gadget_FrontColor, WarningColor)
    Else
    EndIf
  Next
  ;
  SetGadgetColorEx(ParamWAG\TNotRecommended, #PB_Gadget_FrontColor, WarningColor)
  ;
EndProcedure
;
Procedure PBBSetParameters(WWidth = 550, WHeight = 485)
  ;
  Shared ParamWAG.ParametersStruct
  ;
  Protected *GadgetAdress, GadgetList$, LastCategory$, InnerHeight
  Protected Event, EventType, EventGadget, MouseButton, EventMenu, GetOut
  Protected TNotRecommended, Vpos
  ;
  ParamWAG\BorderColor = RGB(150, 150, 150)
  ParamWAG\ParamWindow = OpenWindow(#PB_Any, 0, 0, WWidth, WHeight, GetTextFromCatalogPB("PBBrowserParameters"), #PB_Window_SystemMenu | #PB_Window_MinimizeGadget | #PB_Window_MaximizeGadget | #PB_Window_SizeGadget | #PB_Window_ScreenCentered | #PB_Window_Invisible)
  If ParamWAG\ParamWindow
    ApplyDarkModeToWindow(ParamWAG\ParamWindow)
    StickyWindow(ParamWAG\ParamWindow, #True)
    WindowBounds(ParamWAG\ParamWindow, WWidth, 220, WWidth, 540) 
    BindEvent(#PB_Event_SizeWindow, @PW_ResizeWindowAndGadgets(), ParamWAG\ParamWindow)
    ;
    Protected ColorTheme = GetColorThemesFromPreferences(PBBrowserPrefile$)
    ;
    ParamWAG\Margins = 15      ; For top, left, bottom and right.
    ParamWAG\HorizSpaceInterGadget = 5   ; Horizontal space between gadgets.
    ParamWAG\OKCancelButtonsWidth = 80 ; Ok and Cancel button's width.
    ParamWAG\RightButtonsWidth = 80    ; 'Modify' and 'Standard' button's width.
    ParamWAG\ButtonsHeight = 25        ; Button's height.
    ParamWAG\ScrollBarWidth = GetSystemMetrics_(#SM_CXVSCROLL)
    ;
    ParamWAG\HPosFields = 300  ; Horiz position for input fields.
    ParamWAG\LegendWidth = ParamWAG\HPosFields - ParamWAG\Margins - ParamWAG\HorizSpaceInterGadget
    ParamWAG\InputWidth = 140
    ParamWAG\TextHeight = 20
    ParamWAG\LineHeight = 27
    ParamWAG\ButtonHeight = ParamWAG\LineHeight
    ;
    ; -----------  Compute Area inner height: ------------
    LastCategory$ = ""
    TNotRecommended = 0
    Vpos = ParamWAG\Margins
    ForEach PBBParameters()
      If PBBParameters()\Category$
        If PBBParameters()\Category$ <> LastCategory$
          If LastCategory$
            If TNotRecommended
              TNotRecommended = 0
              Vpos + ParamWAG\LineHeight + 9
            EndIf
            Vpos + ParamWAG\LineHeight + 12
          Else
            Vpos - 2
            TNotRecommended = 1
          EndIf
          Vpos + ParamWAG\LineHeight
          ;
          If LastCategory$ = ""
            Vpos + ParamWAG\LineHeight
          EndIf
          LastCategory$ = PBBParameters()\Category$
        Else
          Vpos + ParamWAG\LineHeight
        EndIf
      EndIf    
    Next
    InnerHeight = Vpos + ParamWAG\Margins
    ; -----------------------------------------------------
    ;
    ParamWAG\AreaGadget = ScrollAreaGadget(#PB_Any, 1, 1, 1, 1, WWidth, InnerHeight, 20, #PB_ScrollArea_BorderLess)
      ;
      GadgetList$ = ""
      LastCategory$ = ""
      ForEach PBBParameters()
        If PBBParameters()\Category$
          If PBBParameters()\Category$ <> LastCategory$
            If LastCategory$
              PBBParameters()\CanvasLineGadget = CanvasGadget(#PB_Any, 1, 1, 1, 1)
              GadgetList$ + Str(PBBParameters()\CanvasLineGadget) + ","
            EndIf
            LastCategory$ = PBBParameters()\Category$
            PBBParameters()\TitleGadget = TextGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB(PBBParameters()\Category$), #PB_Text_Center)
            GadgetList$ + Str(PBBParameters()\TitleGadget) + ","
          EndIf
          PBBParameters()\LegendGadget = TextGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB(PBBParameters()\Name$), #PB_Text_Right)
          GadgetList$ + Str(PBBParameters()\LegendGadget) + ","
          If PBBParameters()\Type = #PBBP_FontStyle
            PBBParameters()\InputGadget = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetPBBStringParameter(PBBParameters()\Name$))
            GadgetList$ + Str(PBBParameters()\InputGadget) + ","
          ElseIf PBBParameters()\Type = #PBBP_Color
            PBBParameters()\InputGadget = CanvasGadget(#PB_Any, 1, 1, 1, 1)
            PBBParameters()\ColorTempValue$ = PBBParameters()\Value$
          Else
            PBBParameters()\InputGadget = StringGadget(#PB_Any, 1, 1, 1, 1, GetPBBStringParameter(PBBParameters()\Name$))
            GadgetList$ + Str(PBBParameters()\InputGadget) + ","
          EndIf
          PBBParameters()\StandardButtonGadget = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("Standard"))
          GadgetList$ + Str(PBBParameters()\StandardButtonGadget) + ","
        EndIf
      Next
      ;
      ParamWAG\TColorTheme = TextGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("PBBrowserColorTheme"), #PB_Text_Right)
      ParamWAG\SColorTheme = ComboBoxGadget(#PB_Any, 1, 1, 1, 1)
      ForEach InterfaceColorPresets()
        AddGadgetItem(ParamWAG\SColorTheme, -1, InterfaceColorPresets()\PresetName)
      Next
      SelectElement(InterfaceColorPresets(), ColorTheme)
      SetGadgetState(ParamWAG\SColorTheme, ColorTheme)
      ParamWAG\BColorTheme = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("Modify"))
      ;
      ParamWAG\TNotRecommended = TextGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("NotRecommendedToModify"), #PB_Text_Center)
      ;
      CloseGadgetList()
    ;
    ParamWAG\LowCanvasBar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
    ;
    ParamWAG\BSave   = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("Save"))
    ParamWAG\BCancel = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("Cancel"))
    ;
    ; Make a list of Gadget numbers
    *GadgetAdress = @ParamWAG\AreaGadget
    Repeat
      GadgetList$ + Str(PeekI(*GadgetAdress)) + ","
      *GadgetAdress + SizeOf(Integer)
    Until *GadgetAdress > @ParamWAG\BCancel
    ;
    PW_ResizeWindowAndGadgets()
    PW_SetFontsToPWGadgets()
    ;
    SetGadgetsColorsFromTheme(ParamWAG\ParamWindow, @InterfaceColorPresets(), GadgetList$)
    ParamWAG\BorderColor = CalculateBorderColor(InterfaceColorPresets()\TextColor, InterfaceColorPresets()\BackgroundColor)
    PW_SetColorsToPWGadgets()
    ;
    #PW_Escape_Cmd = 1
    AddKeyboardShortcut(ParamWAG\ParamWindow, #PB_Shortcut_Escape, #PW_Escape_Cmd)
    #PW_Enter = 2
    AddKeyboardShortcut(ParamWAG\ParamWindow, #PB_Shortcut_Return, #PW_Enter)
    ;
    ; The window was invisible until now, because we created it with #PB_Window_Invisible.
    ; We make it visible now.
    HideWindow(ParamWAG\ParamWindow, #False)
    ;
    Repeat
      Event = WaitWindowEvent()
      EventType = EventType()
      EventGadget = EventGadget()
      ;
      If EventType = #PB_EventType_LeftButtonDown Or EventType = #PB_EventType_LeftButtonUp
        MouseButton = EventType
      EndIf
      Select Event
         Case #PB_Event_Menu
          EventMenu = EventMenu()
          If EventMenu = #PW_Escape_Cmd
            GetOut = -1
            Break
          ElseIf EventMenu = #PW_Enter
            GetOut = 1
            Break
          EndIf
        Case #PB_Event_CloseWindow
          GetOut = -1
          Break
        Case #PB_Event_Gadget
          ForEach PBBParameters()
            If PBBParameters()\InputGadget = EventGadget
              If PBBParameters()\Type = #PBBP_FontStyle
                HideGadget(ParamWAG\IWhiteOver, #False)
                DisableWindow(ParamWAG\ParamWindow, #True)
                PW_FontRequester(EventGadget)
                DisableWindow(ParamWAG\ParamWindow, #False)
                HideGadget(ParamWAG\IWhiteOver, #True)
              ElseIf PBBParameters()\Type = #PBBP_Color And MouseButton = #PB_EventType_LeftButtonDown
                HideGadget(ParamWAG\IWhiteOver, #False)
                DisableWindow(ParamWAG\ParamWindow, #True)
                PBBParameters()\ColorTempValue$ = Str(ZapmanColorRequester(Val(PBBParameters()\ColorTempValue$), InterfaceColorPresets()\TextColor, InterfaceColorPresets()\BackgroundColor))
                DisableWindow(ParamWAG\ParamWindow, #False)
                HideGadget(ParamWAG\IWhiteOver, #True)
                PW_SetColorsToPWGadgets()
                MouseButton = 0
              EndIf
              Break
            ElseIf PBBParameters()\StandardButtonGadget = EventGadget
              If PBBParameters()\Type = #PBBP_Color
                PBBParameters()\ColorTempValue$ = PBBParameters()\DefaultValue$
                PW_SetColorsToPWGadgets()
              Else
                SetGadgetText(PBBParameters()\InputGadget, PBBParameters()\DefaultValue$)
                PW_SetFontsToPWGadgets()
              EndIf
              Break
            EndIf
          Next
          ;
          Select EventGadget
            Case ParamWAG\SColorTheme
              If GetGadgetState(ParamWAG\SColorTheme) > -1
                SelectElement(InterfaceColorPresets(), GetGadgetState(ParamWAG\SColorTheme))
                SetGadgetsColorsFromTheme(ParamWAG\ParamWindow, @InterfaceColorPresets(), GadgetList$)
                ParamWAG\BorderColor = CalculateBorderColor(InterfaceColorPresets()\TextColor, InterfaceColorPresets()\BackgroundColor)
                PW_SetColorsToPWGadgets()
              EndIf
            Case ParamWAG\BColorTheme
              HideGadget(ParamWAG\IWhiteOver, #False)
              DisableWindow(ParamWAG\ParamWindow, #True)
              EditThemesColors()
              DisableWindow(ParamWAG\ParamWindow, #False)
              HideGadget(ParamWAG\IWhiteOver, #True)
              ClearGadgetItems(ParamWAG\SColorTheme)
              ColorTheme = ListIndex(InterfaceColorPresets())
              ForEach InterfaceColorPresets()
                AddGadgetItem(ParamWAG\SColorTheme, -1, InterfaceColorPresets()\PresetName)
              Next
              SelectElement(InterfaceColorPresets(), ColorTheme)
              SetGadgetState(ParamWAG\SColorTheme, ColorTheme)
              ;
              SetGadgetsColorsFromTheme(ParamWAG\ParamWindow, @InterfaceColorPresets(), GadgetList$)
              ParamWAG\BorderColor = CalculateBorderColor(InterfaceColorPresets()\TextColor, InterfaceColorPresets()\BackgroundColor)
              PW_SetColorsToPWGadgets()
              
            Case ParamWAG\BCancel
              GetOut = -1
            Case ParamWAG\BSave
              GetOut = 1
          EndSelect
          ;
      EndSelect
    Until GetOut
    ;
    If GetOut = 1
      ;
      SetColorThemesToPreferences(PBBrowserPrefile$, GetGadgetState(ParamWAG\SColorTheme))
      ;
      ForEach PBBParameters()
        If PBBParameters()\Type = #PBBP_Color
          PBBParameters()\Value$ = PBBParameters()\ColorTempValue$
        Else
          If IsGadget(PBBParameters()\InputGadget)
            PBBParameters()\Value$ = GetGadgetText(PBBParameters()\InputGadget)
          EndIf
        EndIf
        PBBParameters()\ColorTempValue$ = ""
        PBBParameters()\LegendGadget = 0
        PBBParameters()\InputGadget = 0
        If IsFont(PBBParameters()\InputGadgetFont)
          FreeFont(PBBParameters()\InputGadgetFont)
        EndIf
        PBBParameters()\InputGadgetFont = 0
        PBBParameters()\StandardButtonGadget = 0
        PBBParameters()\TitleGadget = 0
        PBBParameters()\CanvasLineGadget = 0
      Next
      ;
      If OpenPreferences(PBBrowserPrefile$)
        ;
        ForEach PBBParameters()
          WritePreferenceString(PBBParameters()\Name$, PBBParameters()\Value$)
        Next
        ClosePreferences()
      EndIf
    EndIf
    ;
    PBBAllGadgetsFont     = FontID(GetFontFromDescription(GetPBBStringParameter("PBBAllGadgetsFont")))
    PBBTitleFont          = FontID(GetFontFromDescription(GetPBBStringParameter("PBBTitleFont")))
    CloseWindow(ParamWAG\ParamWindow)
  EndIf
EndProcedure

CompilerIf #PB_Compiler_IsMainFile

  PBBSetParameters()

CompilerEndIf

; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; Folding = b-
; EnableXP
; DPIAware
; Executable = PrefWindow.exe