#NomProg = "PBBrowser"

XIncludeFile "PBBrowserDeclarations.pb"
XIncludeFile "PBBrowserInitialize.pb"
;
Procedure FillCanvasWithColor(Canvas, Color, BorderColor)
  If StartDrawing(CanvasOutput(Canvas))
    Box(0, 0, DesktopScaledX(GadgetWidth(Canvas)), DesktopScaledY(GadgetHeight(Canvas)), BorderColor)
    Box(1, 1, DesktopScaledX(GadgetWidth(Canvas)) - 2, DesktopScaledY(GadgetHeight(Canvas)) - 2, Color)
    StopDrawing()
  EndIf
EndProcedure

Procedure Open_Window_0(X = 0, Y = 0, WWidth = 470, WHeight = 390)
  BorderColor = RGB(120, 120, 120)
  Window_0 = OpenWindow(#PB_Any, X, Y, WWidth, WHeight, "Title", #PB_Window_SystemMenu | #PB_Window_MinimizeGadget | #PB_Window_MaximizeGadget | #PB_Window_SizeGadget | #PB_Window_ScreenCentered)
  If Window_0
    Margins = 10
    LineHeight = 25
    Vpos = 10
    TPBBrowserWin = TextGadget(#PB_Any, Margins, Vpos + 2, WWidth - Margins * 2, 20, "PBBrowser windows:", #PB_Text_Center)
    SetGadgetColor(TPBBrowserWin, #PB_Gadget_FrontColor, PBBTitleColor)
    SetGadgetFont(TPBBrowserWin, PBBTitleFont)
    Vpos + LineHeight
    TDefFont = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "PBBrowser default font:", #PB_Text_Right)
    SDefFont = StringGadget(#PB_Any, 300, Vpos, 140, 20, "String_1")
    Vpos + LineHeight
    TColorTheme = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "PBBrowser color theme:", #PB_Text_Right)
    SColorTheme = StringGadget(#PB_Any, 300, Vpos, 140, 20, "String_1")
    Vpos + LineHeight
    TMWFont = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Font for the titles in windows:", #PB_Text_Right)
    SMWFont = StringGadget(#PB_Any, 300, Vpos, 140, 20, "String_1")
    Vpos + LineHeight
    TMWTitle = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Color for the titles in windows:", #PB_Text_Right)
    CanvMWTitle = CanvasGadget(#PB_Any, 300, Vpos, 20, 20)
    FillCanvasWithColor(CanvMWTitle, PBBTitleColor, BorderColor)
    Vpos + LineHeight + 5
    TTitles = TextGadget(#PB_Any, Margins, Vpos + 2, WWidth - Margins * 2, 20, "'Details' and 'Found in...' pannels:", #PB_Text_Center)
    SetGadgetColor(TTitles, #PB_Gadget_FrontColor, PBBTitleColor)
    SetGadgetFont(TTitles, PBBTitleFont)
    Vpos + LineHeight
    TColorElementTitle = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Color for Element types and names:", #PB_Text_Right)
    CanvElementTitle = CanvasGadget(#PB_Any, 300, Vpos, 20, 20)
    FillCanvasWithColor(CanvElementTitle, PBBDarkRedColor, BorderColor)
    Vpos + LineHeight
    TFontElementDetails = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Font for Element details:", #PB_Text_Right)
    SFontElementDetails = StringGadget(#PB_Any, 300, Vpos, 140, 20, "String_1")
    Vpos + LineHeight
    TBulletChar = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Bullet character:", #PB_Text_Right)
    SBulletChar = StringGadget(#PB_Any, 300, Vpos, 20, 20, Left(PBB_Bullet$,1))
    Vpos + LineHeight + 5
    TTitles = TextGadget(#PB_Any, Margins, Vpos + 2, WWidth - Margins * 2, 20, "'Details' pannel:", #PB_Text_Center)
    SetGadgetColor(TTitles, #PB_Gadget_FrontColor, PBBTitleColor)
    SetGadgetFont(TTitles, PBBTitleFont)

    Vpos + LineHeight
    TColorElementDetails = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Color for Element details:", #PB_Text_Right)
    CanvColorElementDetails = CanvasGadget(#PB_Any, 300, Vpos, 20, 20)
    FillCanvasWithColor(CanvColorElementDetails, PBBGreyColor, BorderColor)
    Vpos + LineHeight
    TFontElementDetails = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Font for the code:", #PB_Text_Right)
    SFontElementDetails = StringGadget(#PB_Any, 300, Vpos, 140, 20, "String_1")
    Vpos + LineHeight + 5
    TTitles = TextGadget(#PB_Any, Margins, Vpos + 2, WWidth - Margins * 2, 20, "'Found in...' pannel:", #PB_Text_Center)
    SetGadgetColor(TTitles, #PB_Gadget_FrontColor, PBBTitleColor)
    SetGadgetFont(TTitles, PBBTitleFont)
    Vpos + LineHeight
    TColorSetRead = TextGadget(#PB_Any, Margins, Vpos + 2, 280, 20, "Color for the 'Set', 'Read', 'Param',... mentions:", #PB_Text_Right)
    CanvColorSetRead = CanvasGadget(#PB_Any, 300, Vpos, 20, 20)
    FillCanvasWithColor(CanvColorSetRead, PBBSetValueColor, BorderColor)
    Vpos + LineHeight
    Repeat
      Select WaitWindowEvent()
        Case #PB_Event_CloseWindow
          Break
  
          ;-> Event Gadget
        Case #PB_Event_Gadget
          Select EventGadget()
          EndSelect
  
      EndSelect
    ForEver
  EndIf
EndProcedure

CompilerIf #PB_Compiler_IsMainFile
;- Programme Principal

If Open_Window_0()

EndIf
CompilerEndIf

; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; CursorPosition = 13
; Folding = -
; EnableXP
; DPIAware