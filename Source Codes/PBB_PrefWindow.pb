; *******************************************************************************
;
;                                PrefWindow.pb
;                          Part of PBBrowser project
;                             Zapman - nov. 2026
;
;         Preferences dialog layout, live theme/font preview, and saving.
;
; *******************************************************************************
;
; ===============================================================================
; 1. DIALOG CONSTANTS, STATE STRUCTURE & LAYOUT/STYLING HELPERS
; ===============================================================================
;- Dialog state and styling helpers
;{
Enumeration PW_Shortcuts
  #PW_Escape_Cmd = 1
  #PW_Enter      = 2
EndEnumeration
;
Structure ParametersStruct
  ParamWindow.i ; Preferences window number.
  BorderColor.i ; Separator line and swatch border color derived from the active theme.
  ;
  ; Layout metrics:
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
  ; Contiguous gadget numbers (from AreaGadget through BCancel for pointer traversal):
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
EndStructure
;
Global ParamWAG.ParametersStruct
;
Procedure PW_ToggleModalOverlay(DisableState)
  ;
  ; Role: Shows or hides the semi-transparent overlay and disables/enables the Preferences
  ;       window while a child requester (font, color, or theme editor) is open.
  ;
  If DisableState
    If IsGadget(ParamWAG\IWhiteOver)
      HideGadget(ParamWAG\IWhiteOver, #False)
    EndIf
    DisableWindow(ParamWAG\ParamWindow, #True)
  Else
    DisableWindow(ParamWAG\ParamWindow, #False)
    If IsGadget(ParamWAG\IWhiteOver)
      HideGadget(ParamWAG\IWhiteOver, #True)
    EndIf
    SetActiveWindow(ParamWAG\ParamWindow)
  EndIf
EndProcedure
;
Procedure PW_SetFontsToPWGadgets()
  ;
  ; Role: Loads preview fonts for all '#PBBP_FontStyle' buttons (freeing any previous preview
  ;       handles) and applies the live 'PBBAllGadgetsFont' and 'PBBTitleFont' previews to
  ;       every gadget inside the Preferences window without mutating the main window's globals.
  ;
  Protected PreviewAllFont   = PBBAllGadgetsFont
  Protected PreviewTitleFont = PBBTitleFont
  Protected FontHnd
  ;
  ; 1. Refresh per-parameter preview fonts and capture live handles for the two master UI fonts:
  ForEach PBBParameters()
    If PBBParameters()\Category$ <> "" And PBBParameters()\Type = #PBBP_FontStyle And IsGadget(PBBParameters()\InputGadget)
      If IsFont(PBBParameters()\InputGadgetFont)
        FreeFont(PBBParameters()\InputGadgetFont)
      EndIf
      PBBParameters()\InputGadgetFont = LoadFontFromDescription(GetGadgetText(PBBParameters()\InputGadget))
      If IsFont(PBBParameters()\InputGadgetFont)
        FontHnd = FontID(PBBParameters()\InputGadgetFont)
        SetGadgetFont(PBBParameters()\InputGadget, FontHnd)
        If PBBParameters()\Name$ = "PBBAllGadgetsFont"
          PreviewAllFont = FontHnd
        ElseIf PBBParameters()\Name$ = "PBBTitleFont"
          PreviewTitleFont = FontHnd
        EndIf
      EndIf
    EndIf
  Next
  ;
  ; 2. Apply the live preview fonts across all parameter labels, inputs, and dialog controls:
  ForEach PBBParameters()
    If IsGadget(PBBParameters()\TitleGadget)
      SetGadgetFont(PBBParameters()\TitleGadget, PreviewTitleFont)
    EndIf
    If IsGadget(PBBParameters()\LegendGadget)
      SetGadgetFont(PBBParameters()\LegendGadget, PreviewAllFont)
    EndIf
    If IsGadget(PBBParameters()\StandardButtonGadget)
      SetGadgetFont(PBBParameters()\StandardButtonGadget, PreviewAllFont)
    EndIf
    If IsGadget(PBBParameters()\InputGadget) And PBBParameters()\Type <> #PBBP_FontStyle And PBBParameters()\Type <> #PBBP_Color
      SetGadgetFont(PBBParameters()\InputGadget, PreviewAllFont)
    EndIf
  Next
  ;
  SetGadgetFont(ParamWAG\TNotRecommended, PreviewAllFont)
  SetGadgetFont(ParamWAG\BSave,           PreviewAllFont)
  SetGadgetFont(ParamWAG\BCancel,         PreviewAllFont)
  SetGadgetFont(ParamWAG\TColorTheme,     PreviewAllFont)
  SetGadgetFont(ParamWAG\SColorTheme,     PreviewAllFont)
  SetGadgetFont(ParamWAG\BColorTheme,     PreviewAllFont)
EndProcedure
;
Procedure PW_SetColorsToPWGadgets()
  ;
  ; Role: Applies live section title colors, color swatch fills, separator line colors,
  ;       and warning highlights across the Preferences dialog.
  ;
  Protected WarningColor = RGB(150, 120, 20)
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
      SetGadgetColorEx(PBBParameters()\TitleGadget, #PB_Gadget_FrontColor, TitleColor)
    EndIf
    If IsGadget(PBBParameters()\CanvasLineGadget)
      FillCanvasWithColor(PBBParameters()\CanvasLineGadget, ParamWAG\BorderColor, ParamWAG\BorderColor)
    EndIf
    If PBBParameters()\Type = #PBBP_Color And IsGadget(PBBParameters()\InputGadget)
      FillCanvasWithColor(PBBParameters()\InputGadget, Val(PBBParameters()\ColorTempValue$), ParamWAG\BorderColor)
    EndIf
    If IsGadget(PBBParameters()\LegendGadget) And FindString(Right(GetGadgetText(PBBParameters()\LegendGadget), 4), "*")
      SetGadgetColorEx(PBBParameters()\LegendGadget, #PB_Gadget_FrontColor, WarningColor)
    EndIf
  Next
  ;
  If IsGadget(ParamWAG\LowCanvasBar)
    FillCanvasWithColor(ParamWAG\LowCanvasBar, ParamWAG\BorderColor, ParamWAG\BorderColor)
  EndIf
  SetGadgetColorEx(ParamWAG\TNotRecommended, #PB_Gadget_FrontColor, WarningColor)
EndProcedure
;
Procedure PW_ApplyTheme(GadgetList$)
  ;
  ; Role: Applies the currently selected InterfaceColorPresets() entry to the Preferences
  ;       window, updates border/separator colors, and refreshes the modal overlay tint.
  ;
  SetGadgetsColorsFromTheme(ParamWAG\ParamWindow, @InterfaceColorPresets(), GadgetList$)
  ParamWAG\BorderColor = CalculateBorderColor(InterfaceColorPresets()\TextColor, InterfaceColorPresets()\BackGroundColor)
  PW_SetColorsToPWGadgets()
  ParamWAG\IWhiteOver  = WhiteBoxOverWindow(ParamWAG\ParamWindow)
EndProcedure
;
Procedure PW_PopulateThemeCombo(SelectedThemeIndex)
  ;
  ; Role: Populates (or refreshes) the color theme ComboBox from InterfaceColorPresets()
  ;       and selects 'SelectedThemeIndex'.
  ;
  ClearGadgetItems(ParamWAG\SColorTheme)
  ForEach InterfaceColorPresets()
    AddGadgetItem(ParamWAG\SColorTheme, -1, InterfaceColorPresets()\PresetName)
  Next
  If SelectedThemeIndex >= 0 And SelectedThemeIndex < ListSize(InterfaceColorPresets())
    SelectElement(InterfaceColorPresets(), SelectedThemeIndex)
    SetGadgetState(ParamWAG\SColorTheme, SelectedThemeIndex)
  EndIf
EndProcedure
;
Procedure PW_ResizeWindowAndGadgets()
  ;
  ; Role: Dynamically repositions and resizes the ScrollAreaGadget, parameter rows,
  ;       separator bars, and action buttons whenever the Preferences window is resized.
  ;
  Protected WWidth       = WindowWidth(ParamWAG\ParamWindow)
  Protected WHeight      = WindowHeight(ParamWAG\ParamWindow)
  Protected ButtonExtend = 1
  Protected InputWidthMax, InputWidth, RightBtnX, Vpos, TNotRecommended
  Protected LastCategory$ = ""
  ;
  ResizeGadget(ParamWAG\AreaGadget, 0, 0, WWidth, WHeight - 2 * ParamWAG\Margins - ParamWAG\ButtonsHeight)
  ;
  ; Probe full width first, then subtract scrollbar width if the vertical scrollbar is visible:
  SetGadgetAttribute(ParamWAG\AreaGadget, #PB_ScrollArea_InnerWidth, WWidth)
  If IsVertScrollBarVisible(ParamWAG\AreaGadget)
    ParamWAG\InnerWidth = WWidth - ParamWAG\ScrollBarWidth
    SetGadgetAttribute(ParamWAG\AreaGadget, #PB_ScrollArea_InnerWidth, ParamWAG\InnerWidth)
  Else
    ParamWAG\InnerWidth = WWidth
  EndIf
  ;
  ; Clamp local InputWidth without mutating the configured ParamWAG\InputWidth baseline:
  InputWidthMax = ParamWAG\InnerWidth - ParamWAG\HPosFields - ParamWAG\RightButtonsWidth - ParamWAG\HorizSpaceInterGadget - ParamWAG\Margins
  InputWidth    = ParamWAG\InputWidth
  If InputWidth > InputWidthMax
    InputWidth = InputWidthMax
  EndIf
  RightBtnX = ParamWAG\HPosFields + InputWidth + ParamWAG\HorizSpaceInterGadget
  ;
  Vpos = ParamWAG\Margins
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
        ;
        If LastCategory$ = ""
          ResizeGadget(ParamWAG\TColorTheme, ParamWAG\Margins,    Vpos + 2,                ParamWAG\LegendWidth,       ParamWAG\TextHeight)
          ResizeGadget(ParamWAG\SColorTheme, ParamWAG\HPosFields, Vpos - ButtonExtend + 1, InputWidth,                 ParamWAG\TextHeight + ButtonExtend * 2 - 2)
          ResizeGadget(ParamWAG\BColorTheme, RightBtnX,           Vpos - ButtonExtend,     ParamWAG\RightButtonsWidth, ParamWAG\TextHeight + ButtonExtend * 2)
          Vpos + ParamWAG\LineHeight
        EndIf
        LastCategory$ = PBBParameters()\Category$
      Else
        Vpos + ParamWAG\LineHeight
      EndIf
      ;
      ResizeGadget(PBBParameters()\LegendGadget, ParamWAG\Margins, Vpos + 2, ParamWAG\LegendWidth, ParamWAG\TextHeight)
      If PBBParameters()\Type = #PBBP_FontStyle
        ResizeGadget(PBBParameters()\InputGadget, ParamWAG\HPosFields, Vpos - ButtonExtend + 1, InputWidth, ParamWAG\TextHeight + ButtonExtend * 2 - 2)
      ElseIf PBBParameters()\Type = #PBBP_Color
        ResizeGadget(PBBParameters()\InputGadget, ParamWAG\HPosFields, Vpos - ButtonExtend, ParamWAG\TextHeight + ButtonExtend * 2, ParamWAG\TextHeight + ButtonExtend * 2)
      Else
        ResizeGadget(PBBParameters()\InputGadget, ParamWAG\HPosFields, Vpos, ParamWAG\TextHeight, ParamWAG\TextHeight)
      EndIf
      ResizeGadget(PBBParameters()\StandardButtonGadget, RightBtnX, Vpos - ButtonExtend, ParamWAG\RightButtonsWidth, ParamWAG\TextHeight + ButtonExtend * 2)
    EndIf
  Next
  ;
  Vpos = WHeight - 2 * ParamWAG\Margins - ParamWAG\ButtonsHeight
  ResizeGadget(ParamWAG\LowCanvasBar, 0, Vpos, WWidth, 1)
  FillCanvasWithColor(ParamWAG\LowCanvasBar, ParamWAG\BorderColor, ParamWAG\BorderColor)
  Vpos + ParamWAG\Margins
  ;
  ResizeGadget(ParamWAG\BSave,   WWidth - ParamWAG\OKCancelButtonsWidth - ParamWAG\Margins,                                               Vpos, ParamWAG\OKCancelButtonsWidth, ParamWAG\ButtonHeight)
  ResizeGadget(ParamWAG\BCancel, WWidth - 2 * ParamWAG\OKCancelButtonsWidth - ParamWAG\HorizSpaceInterGadget - ParamWAG\Margins,          Vpos, ParamWAG\OKCancelButtonsWidth, ParamWAG\ButtonHeight)
  ;
  ; Synchronize the semi-transparent overlay rectangle with the updated window dimensions:
  ParamWAG\IWhiteOver = WhiteBoxOverWindow(ParamWAG\ParamWindow)
EndProcedure
;}
;
; ===============================================================================
; 2. MAIN PREFERENCES DIALOG & EVENT LOOP
; ===============================================================================
;- Main preferences dialog
;{
Procedure PBBSetParameters(WWidth = 550, WHeight = 485)
  ;
  ; Role: Opens the modal Preferences dialog, manages live theme, color, and font previews,
  ;       persists changes to PBBrowserPrefile$ on 'Save', or cleanly reverts previews on 'Cancel'.
  ;
  Static ManagedAllFontPB, ManagedTitleFontPB
  Protected *GadgetAddress, GadgetList$, LastCategory$, PickedFont$
  Protected Event, EventType, EventGadget, MouseButton, EventMenu, GetOut
  Protected ColorTheme, NewFontPB, InnerHeight, TNotRecommended, Vpos, OX, OY
  ;
  ParamWAG\BorderColor = RGB(150, 150, 150)
  Protected WParam         = #PB_Window_SystemMenu | #PB_Window_MinimizeGadget | #PB_Window_MaximizeGadget | #PB_Window_SizeGadget | #PB_Window_Invisible
  Protected ParentWindowID = ComputeWinOrigins(@OX, @OY, WWidth, WHeight, GPBBGadgets\PBBWindow)
  ;
  ParamWAG\ParamWindow = OpenWindow(#PB_Any, OX, OY, WWidth, WHeight, GetTextFromCatalogPB("PBBrowserParameters"), WParam, ParentWindowID)
  If ParamWAG\ParamWindow
    ApplyDarkModeToWindow(ParamWAG\ParamWindow)
    StickyWindow(ParamWAG\ParamWindow, #True)
    WindowBounds(ParamWAG\ParamWindow, WWidth, 220, WWidth, 540)
    BindEvent(#PB_Event_SizeWindow, @PW_ResizeWindowAndGadgets(), ParamWAG\ParamWindow)
    ;
    ColorTheme = GetColorThemesFromPreferences(PBBrowserPrefile$)
    ;
    ParamWAG\Margins                = 15 ; Outer top, left, bottom, and right margin.
    ParamWAG\HorizSpaceInterGadget  = 5  ; Horizontal gap between adjacent controls.
    ParamWAG\OKCancelButtonsWidth   = 80 ; Width of 'Save' and 'Cancel' buttons.
    ParamWAG\RightButtonsWidth      = 80 ; Width of 'Modify' and 'Standard' buttons.
    ParamWAG\ButtonsHeight          = 25
    ParamWAG\ScrollBarWidth         = GetSystemMetrics_(#SM_CXVSCROLL)
    ParamWAG\HPosFields             = 300 ; Horizontal X offset of parameter input controls.
    ParamWAG\LegendWidth            = ParamWAG\HPosFields - ParamWAG\Margins - ParamWAG\HorizSpaceInterGadget
    ParamWAG\InputWidth             = 140
    ParamWAG\TextHeight             = 20
    ParamWAG\LineHeight             = 27
    ParamWAG\ButtonHeight           = ParamWAG\LineHeight
    ;
    ; --- 1. Compute ScrollArea inner height ---
    LastCategory$   = ""
    TNotRecommended = 0
    Vpos            = ParamWAG\Margins
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
    ;
    ; --- 2. Create scrollable parameter controls ---
    ParamWAG\AreaGadget = ScrollAreaGadget(#PB_Any, 1, 1, 1, 1, WWidth, InnerHeight, 20, #PB_ScrollArea_BorderLess)
      GadgetList$   = ""
      LastCategory$ = ""
      ForEach PBBParameters()
        If PBBParameters()\Category$
          If PBBParameters()\Category$ <> LastCategory$
            If LastCategory$
              PBBParameters()\CanvasLineGadget = CanvasGadget(#PB_Any, 1, 1, 1, 1)
              GadgetList$ + Str(PBBParameters()\CanvasLineGadget) + ","
            EndIf
            LastCategory$               = PBBParameters()\Category$
            PBBParameters()\TitleGadget = TextGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB(PBBParameters()\Category$), #PB_Text_Center)
            GadgetList$ + Str(PBBParameters()\TitleGadget) + ","
          EndIf
          PBBParameters()\LegendGadget = TextGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB(PBBParameters()\Name$), #PB_Text_Right)
          GadgetList$ + Str(PBBParameters()\LegendGadget) + ","
          If PBBParameters()\Type = #PBBP_FontStyle
            PBBParameters()\InputGadget = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetPBBStringParameter(PBBParameters()\Name$))
            GadgetList$ + Str(PBBParameters()\InputGadget) + ","
          ElseIf PBBParameters()\Type = #PBBP_Color
            PBBParameters()\InputGadget     = CanvasGadget(#PB_Any, 1, 1, 1, 1)
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
      ParamWAG\TColorTheme = TextGadget    (#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("PBBrowserColorTheme"), #PB_Text_Right)
      ParamWAG\SColorTheme = ComboBoxGadget(#PB_Any, 1, 1, 1, 1)
      PW_PopulateThemeCombo(ColorTheme)
      ParamWAG\BColorTheme     = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("Modify"))
      ParamWAG\TNotRecommended = TextGadget  (#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("NotRecommendedToModify"), #PB_Text_Center)
    CloseGadgetList()
    ;
    ParamWAG\LowCanvasBar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
    ParamWAG\BSave        = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("Save"))
    ParamWAG\BCancel      = ButtonGadget(#PB_Any, 1, 1, 1, 1, GetTextFromCatalogPB("Cancel"))
    ;
    ; Append fixed dialog gadgets via contiguous structure traversal:
    *GadgetAddress = @ParamWAG\AreaGadget
    Repeat
      GadgetList$ + Str(PeekI(*GadgetAddress)) + ","
      *GadgetAddress + SizeOf(Integer)
    Until *GadgetAddress > @ParamWAG\BCancel
    ;
    PW_ResizeWindowAndGadgets()
    PW_SetFontsToPWGadgets()
    PW_ApplyTheme(GadgetList$)
    ;
    AddKeyboardShortcut(ParamWAG\ParamWindow, #PB_Shortcut_Escape, #PW_Escape_Cmd)
    AddKeyboardShortcut(ParamWAG\ParamWindow, #PB_Shortcut_Return, #PW_Enter)
    ClearKeyboardBuffer(ParamWAG\ParamWindow)
    HideWindow(ParamWAG\ParamWindow, #False)
    ;
    ; --- 3. Modal event loop ---
    Repeat
      Event = WaitWindowEvent()
      If EventWindow() <> ParamWAG\ParamWindow
        SetActiveWindow(ParamWAG\ParamWindow)
        Continue
      EndIf
      ;
      EventType   = EventType()
      EventGadget = EventGadget()
      If EventType = #PB_EventType_LeftButtonDown Or EventType = #PB_EventType_LeftButtonUp
        MouseButton = EventType
      EndIf
      ;
      Select Event
        Case #PB_Event_Menu
          EventMenu = EventMenu()
          If EventMenu = #PW_Escape_Cmd
            GetOut = -1
          ElseIf EventMenu = #PW_Enter
            GetOut = 1
          EndIf
          ;
        Case #PB_Event_CloseWindow
          GetOut = -1
          ;
        Case #PB_Event_Gadget
          ForEach PBBParameters()
            If PBBParameters()\InputGadget = EventGadget
              If PBBParameters()\Type = #PBBP_FontStyle
                PW_ToggleModalOverlay(#True)
                PickedFont$ = FontRequesterEx(GetGadgetText(EventGadget), #ZFR_FontRequester_Default, 0, ParamWAG\ParamWindow)
                PW_ToggleModalOverlay(#False)
                If PickedFont$ <> ""
                  SetGadgetText(EventGadget, PickedFont$)
                  PW_SetFontsToPWGadgets()
                EndIf
              ElseIf PBBParameters()\Type = #PBBP_Color And MouseButton = #PB_EventType_LeftButtonDown
                PW_ToggleModalOverlay(#True)
                PBBParameters()\ColorTempValue$ = Str(ZapmanColorRequester(Val(PBBParameters()\ColorTempValue$), ParamWAG\ParamWindow))
                PW_ToggleModalOverlay(#False)
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
                PW_ApplyTheme(GadgetList$)
              EndIf
            Case ParamWAG\BColorTheme
              PW_ToggleModalOverlay(#True)
              EditThemesColors(ParamWAG\ParamWindow)
              PW_ToggleModalOverlay(#False)
              PW_PopulateThemeCombo(ListIndex(InterfaceColorPresets()))
              PW_ApplyTheme(GadgetList$)
            Case ParamWAG\BCancel
              GetOut = -1
            Case ParamWAG\BSave
              GetOut = 1
          EndSelect
      EndSelect
    Until GetOut
    ;
    ; --- 4. Persist settings on Save, or restore persisted theme on Cancel ---
    If GetOut = 1
      SetColorThemesToPreferences(PBBrowserPrefile$, GetGadgetState(ParamWAG\SColorTheme))
      ForEach PBBParameters()
        If PBBParameters()\Category$ <> ""
          If PBBParameters()\Type = #PBBP_Color
            PBBParameters()\Value$ = PBBParameters()\ColorTempValue$
          ElseIf IsGadget(PBBParameters()\InputGadget)
            PBBParameters()\Value$ = GetGadgetText(PBBParameters()\InputGadget)
          EndIf
        EndIf
      Next
      ;
      If OpenPreferencesWithPatience(PBBrowserPrefile$)
        ForEach PBBParameters()
          WritePreferenceString(PBBParameters()\Name$, PBBParameters()\Value$)
        Next
        ClosePreferences()
      EndIf
      ;
      ; Update global UI font handles (freeing any previous PB font slot owned by this dialog) and bullet line break:
      NewFontPB = LoadFontFromDescription(GetPBBStringParameter("PBBAllGadgetsFont"))
      If IsFont(NewFontPB)
        PBBAllGadgetsFont = FontID(NewFontPB)
        If IsFont(ManagedAllFontPB)
          FreeFont(ManagedAllFontPB)
        EndIf
        ManagedAllFontPB = NewFontPB
      EndIf
      ;
      NewFontPB = LoadFontFromDescription(GetPBBStringParameter("PBBTitleFont"))
      If IsFont(NewFontPB)
        PBBTitleFont = FontID(NewFontPB)
        If IsFont(ManagedTitleFontPB)
          FreeFont(ManagedTitleFontPB)
        EndIf
        ManagedTitleFontPB = NewFontPB
      EndIf
      ;
      PBB_LineBreak$ = #CR$ + GetPBBStringParameter("PBB_Bullet") + #NonBreakableSpace$
    Else
      ; Revert any unsaved theme selection back to the stored preference preset:
      GetColorThemesFromPreferences(PBBrowserPrefile$)
    EndIf
    ;
    ; --- 5. Unconditional cleanup of transient gadget references and preview fonts ---
    ForEach PBBParameters()
      PBBParameters()\ColorTempValue$      = ""
      PBBParameters()\LegendGadget         = 0
      PBBParameters()\InputGadget          = 0
      If IsFont(PBBParameters()\InputGadgetFont)
        FreeFont(PBBParameters()\InputGadgetFont)
      EndIf
      PBBParameters()\InputGadgetFont      = 0
      PBBParameters()\StandardButtonGadget = 0
      PBBParameters()\TitleGadget          = 0
      PBBParameters()\CanvasLineGadget     = 0
    Next
    ;
    UnbindEvent(#PB_Event_SizeWindow, @PW_ResizeWindowAndGadgets(), ParamWAG\ParamWindow)
    ClearKeyboardBuffer(ParamWAG\ParamWindow)
    CloseWindow(ParamWAG\ParamWindow)
  EndIf
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  PBBSetParameters()
CompilerEndIf
;}
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 10
; Folding = --
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb
; Executable = PrefWindow.exe