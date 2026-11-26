;
; ********************************************************************
;
;                       Zapman Color Requester
;                           CrossPlatform
;                          Nov. 2026 - 6
;
; This file should be saved under the name "ZapmanColorRequester.pbi".
;
;    This library implements an intuitive and professional color
;    requester for your PureBasic applications.
;    Its main function is 'ZapmanColorRequester()'.
;    Run the demo at the bottom of this file to see it in action.
;
;    The two blocks of three gadgets used to configure a color
;    are laid out with configurable top, bottom, left, right,
;    and inter-block margins so they can easily be embedded inside
;    custom windows hosting additional gadgets.
;    An example of reusing these gadgets is provided in the
;    'EditThemesColors()' procedure of 'ApplyColorThemes.pbi'.
;
; This library is compatible with the Zapman 'ApplyColorThemes.pbi' library,
; allowing your windows to be displayed in Dark Mode or custom color themes.
;
; ********************************************************************
;
CompilerIf #PB_Compiler_IsMainFile
  EnableExplicit
CompilerEndIf
;
;- 1. Language settings
;
; All Zapman libraries are designed to integrate smoothly into multilingual applications.
; You can override 'GetTextFromCatalog()' in a file included prior to this one if you
; prefer to store and manage your UI translations differently.
; Otherwise, this section provides a lightweight, portable translation catalog; adding
; a new language only requires a few lines at the end of this section.
;
Global MyLanguage$
If MyLanguage$ = ""
  MyLanguage$ = "EN" ; Supported out of the box: "EN", "RU", "DE", "IT", "ES", "ZH", "FR"
EndIf
;
CompilerIf Not Defined(LanguageListStructure, #PB_Structure)
  Structure LanguageListStructure
    Language$
    LanguageEntry$
    LanguageTranslation$
  EndStructure
  ;
  Global NewList LanguageList.LanguageListStructure()
CompilerEndIf
;
CompilerIf Not Defined(GetTextFromCatalog, #PB_Procedure)
  Procedure.s GetTextFromCatalog(SName$)
    ;
    ; Role: Looks up the localized string matching 'SName$' for the active 'MyLanguage$'.
    ; Output: Returns the translated text, or 'SName$' itself as fallback if not found.
    ;
    ForEach LanguageList()
      If LanguageList()\Language$ = MyLanguage$ And LCase(LanguageList()\LanguageEntry$) = LCase(SName$)
        ProcedureReturn LanguageList()\LanguageTranslation$
      EndIf
    Next
    ProcedureReturn SName$
  EndProcedure
CompilerEndIf
;
CompilerIf Not Defined(FillLanguageList, #PB_Procedure)
  Procedure FillLanguageList(Language$, LanguageEntry$, LanguageTranslation$)
    ;
    ; Role: Populates 'LanguageList()' from comma-separated key and translation lists for 'Language$'.
    ;
    Protected PosInList = 1, KeyWord$, Found
    ;
    KeyWord$ = StringField(LanguageEntry$, PosInList, ",")
    While KeyWord$
      Found = #False
      ForEach LanguageList()
        If LanguageList()\Language$ = Language$ And LCase(LanguageList()\LanguageEntry$) = LCase(KeyWord$)
          Found = #True
          Break
        EndIf
      Next
      If Found = #False
        AddElement(LanguageList())
        LanguageList()\Language$ = Language$
        LanguageList()\LanguageEntry$ = KeyWord$
        LanguageList()\LanguageTranslation$ = StringField(LanguageTranslation$, PosInList, ",")
      EndIf
      PosInList + 1
      KeyWord$ = StringField(LanguageEntry$, PosInList, ",")
    Wend
  EndProcedure
CompilerEndIf
;
Define LanguageEntry$ = "ThemeColors,Presets,BackgroundColor,TextColor,ResetToDefaultColor,"
LanguageEntry$ + "EditColor,Red,Green,Blue,Hue,Saturation,Lightness,HexValue,Cancel,OK,"
LanguageEntry$ + "ClickToAdjustColor"
;
; ------------- English list of expressions ---------------
Define LanguageTranslation$ = "Theme colors,Preset:,Background Color,Text Color,Reset to original color,"
LanguageTranslation$ + "Edit Color,Red:,Green:,Blue:,Hue:,Saturation:,Lightness:,Hex value:,Cancel,OK,"
LanguageTranslation$ + "Click to choose a color. Use the mouse wheel to adjust saturation."
FillLanguageList("EN", LanguageEntry$, LanguageTranslation$)
;
; ------------- Spanish list of expressions ---------------
LanguageTranslation$ = "Colores del tema,Predefinido:,Color de fondo,Color de texto,Restablecer al color original,"
LanguageTranslation$ + "Editar color,Rojo:,Verde:,Azul:,Matiz:,Saturación:,Luminosidad:,Valor hexadecimal:,Cancelar,OK,"
LanguageTranslation$ + "Haz clic para elegir un color. Usa la rueda del ratón para ajustar la saturación."
FillLanguageList("ES", LanguageEntry$, LanguageTranslation$)
;
; ------------- Mandarin list of expressions ---------------
LanguageTranslation$ = "主题颜色,预设:,背景颜色,文字颜色,重置为原始颜色,"
LanguageTranslation$ + "编辑颜色,红:,绿:,蓝:,色相:,饱和度:,亮度:,十六进制值:,取消,确定,"
LanguageTranslation$ + "点击选择颜色。使用鼠标滚轮调整饱和度。"
FillLanguageList("ZH", LanguageEntry$, LanguageTranslation$)
;
; ------------- French list of expressions ---------------
LanguageTranslation$ = "Couleurs du thème,Préréglages :,Couleur de fond,Couleur du texte,Couleur d'origine,"
LanguageTranslation$ + "Éditeur de couleur,Rouge :,Vert :,Bleu :,Teinte :,Saturation :,Luminosité :,Valeur hexa :,Annuler,OK,"
LanguageTranslation$ + "Cliquez pour choisir une couleur. Utilisez la roue de souris pour régler la saturation."
FillLanguageList("FR", LanguageEntry$, LanguageTranslation$)
;
; ------------- German list of expressions ---------------
LanguageTranslation$ = "Themenfarben,Voreinstellungen:,Hintergrundfarbe,Textfarbe,Ursprungsfarbe,"
LanguageTranslation$ + "Farbeditor,Rot:,Grün:,Blau:,Farbton:,Sättigung:,Helligkeit:,Hex-Wert:,Abbrechen,OK,"
LanguageTranslation$ + "Klicken Sie, um eine Farbe auszuwählen. Verwenden Sie das Mausrad, um die Sättigung einzustellen."
FillLanguageList("DE", LanguageEntry$, LanguageTranslation$)
;
; ------------- Russian list of expressions ---------------
LanguageTranslation$ = "Цвета темы,Предустановки:,Цвет фона,Цвет текста,Исходный цвет,"
LanguageTranslation$ + "Редактор цветов,Красный:,Зеленый:,Синий:,Оттенок:,Насыщенность:,Яркость:,Hex-значение:,Отмена,OK,"
LanguageTranslation$ + "Щелкните, чтобы выбрать цвет. Используйте колесо мыши для регулировки насыщенности."
FillLanguageList("RU", LanguageEntry$, LanguageTranslation$)
;
; ------------- Italian list of expressions ---------------
LanguageTranslation$ = "Colori del tema,Predefinito:,Colore di sfondo,Colore del testo,Reimposta al colore originale,"
LanguageTranslation$ + "Modifica colore,Rosso:,Verde:,Blu:,Tonalità:,Saturazione:,Luminosità:,Valore esadecimale:,Annulla,OK,"
LanguageTranslation$ + "Clicca per scegliere un colore. Usa la rotella del mouse per regolare la saturazione."
FillLanguageList("IT", LanguageEntry$, LanguageTranslation$)
;
; ------------- Add your own language here --------------
; LanguageTranslation$ = "Expression1,Expression2,etc."
; FillLanguageList("XX", LanguageEntry$, LanguageTranslation$) ; Replace "XX" with your language code.
;
; ------------------- END OF LANGUAGE SETTINGS -------------------
;
CompilerIf Not Defined(CWO_Positioning, #PB_Enumeration)
  Enumeration CWO_Positioning
    #CWO_ActiveWindowPos = -1
    #CWO_AbsolutePos     = -2
    #CWO_MonitorPos      = -3
  EndEnumeration
CompilerEndIf
;
CompilerIf Not Defined(CWO_PositionAnchor, #PB_Enumeration)
  Enumeration CWO_PositionAnchor
    #CWO_Center = 0
    #CWO_TopLeft
    #CWO_TopRight
    #CWO_BottomLeft
    #CWO_BottomRight
  EndEnumeration
CompilerEndIf
;
CompilerIf Not Defined(ComputeWinOrigins, #PB_Procedure)
  Procedure ComputeWinOrigins(*OX.Integer, *OY.Integer, WWidth, WHeight, ParentWindow = #CWO_ActiveWindowPos, XShiftOrPos = 0, YShiftOrPos = 0, ParentAnchor = #CWO_Center, WindowAnchor = #CWO_Center)
    ;
    ; Role: Computes the X and Y logical screen origins for a new window from its client dimensions
    ;       (WWidth, WHeight) relative to a parent window or to the target monitor's work area,
    ;       clamped so the window stays within visible work-area bounds.
    ; Output: Writes computed coordinates to '*OX' and '*OY', and returns the OS HWND of 'ParentWindow'
    ;         (or 0 when positioned relative to a monitor or at absolute coordinates).
    ;
    ; Positioning rules:
    ; • If ParentWindow is a valid PureBasic window number, the new window is positioned relative
    ;   to it and then shifted by XShiftOrPos and YShiftOrPos.
    ; • If ParentWindow = #CWO_ActiveWindowPos and GetActiveWindow() returns a valid window number,
    ;   GetActiveWindow() is used as the parent window; otherwise the main monitor is used.
    ; • If ParentWindow = #CWO_AbsolutePos, XShiftOrPos and YShiftOrPos are used directly as the
    ;   absolute coordinates of the new window.
    ; • If ParentWindow = #CWO_MonitorPos, the position is calculated relative to the monitor
    ;   containing GetActiveWindow() (or the main monitor if no PureBasic window is active).
    ; • 'ParentAnchor' selects the reference point on the parent window or monitor (#CWO_Center,
    ;   #CWO_TopLeft, #CWO_TopRight, #CWO_BottomLeft, or #CWO_BottomRight).
    ; • 'WindowAnchor' selects the matching anchor point on the new window itself.
    ; • In all cases except #CWO_AbsolutePos, 'XShiftOrPos' and 'YShiftOrPos' offset the result.
    ;
    Protected DesktopLeft, DesktopRight, DesktopTop, DesktopBottom
    Protected ParentWindowID, MFWindow, hMonitor, TotalHeight, mi.MONITORINFO
    ;
    If ParentWindow = #CWO_AbsolutePos
      *OX\i = XShiftOrPos
      *OY\i = YShiftOrPos
      ProcedureReturn 0
    EndIf
    ;
    If ParentWindow = #CWO_ActiveWindowPos And IsWindow(GetActiveWindow())
      ParentWindow = GetActiveWindow()
    EndIf
    ;
    ; Retrieve the work-area bounds of the monitor hosting the parent (or active) window:
    MFWindow = ParentWindow
    If Not IsWindow(MFWindow)
      MFWindow = GetActiveWindow()
    EndIf
    If IsWindow(MFWindow)
      MFWindow = WindowID(MFWindow)
    Else
      MFWindow = 0
    EndIf
    ;
    hMonitor  = MonitorFromWindow_(MFWindow, #MONITOR_DEFAULTTONEAREST)
    mi\cbSize = SizeOf(MONITORINFO)
    GetMonitorInfo_(hMonitor, @mi)
    ;
    DesktopLeft   = DesktopUnscaledX(mi\rcWork\left)
    DesktopRight  = DesktopUnscaledX(mi\rcWork\right)
    DesktopTop    = DesktopUnscaledY(mi\rcWork\top)
    DesktopBottom = DesktopUnscaledY(mi\rcWork\bottom)
    ;
    ; Compute the reference point on the parent window or monitor:
    If IsWindow(ParentWindow)
      ParentWindowID = WindowID(ParentWindow)
      *OX\i = WindowX(ParentWindow)
      *OY\i = WindowY(ParentWindow)
      ;
      If ParentAnchor = #CWO_Center
        *OX\i + WindowWidth(ParentWindow) / 2
        *OY\i + WindowHeight(ParentWindow) / 2
      ElseIf ParentAnchor = #CWO_TopRight
        *OX\i + WindowWidth(ParentWindow)
      ElseIf ParentAnchor = #CWO_BottomLeft
        *OY\i + WindowHeight(ParentWindow)
      ElseIf ParentAnchor = #CWO_BottomRight
        *OX\i + WindowWidth(ParentWindow)
        *OY\i + WindowHeight(ParentWindow)
      EndIf
    Else
      *OX\i = DesktopLeft
      *OY\i = DesktopTop
      ;
      If ParentAnchor = #CWO_Center
        *OX\i = (DesktopLeft + DesktopRight) / 2
        *OY\i = (DesktopTop + DesktopBottom) / 2
      ElseIf ParentAnchor = #CWO_TopRight
        *OX\i = DesktopRight
      ElseIf ParentAnchor = #CWO_BottomLeft
        *OY\i = DesktopBottom
      ElseIf ParentAnchor = #CWO_BottomRight
        *OX\i = DesktopRight
        *OY\i = DesktopBottom
      EndIf
    EndIf
    ;
    ; Account for the title bar and horizontal frame borders added by OpenWindow() around WHeight:
    TotalHeight = WHeight + DesktopUnscaledY(GetSystemMetrics_(#SM_CYCAPTION) + GetSystemMetrics_(#SM_CYSIZEFRAME) * 2)
    If WindowAnchor = #CWO_Center
      *OX\i - WWidth / 2
      *OY\i - TotalHeight / 2
    ElseIf WindowAnchor = #CWO_TopRight
      *OX\i - WWidth
    ElseIf WindowAnchor = #CWO_BottomLeft
      *OY\i - TotalHeight
    ElseIf WindowAnchor = #CWO_BottomRight
      *OX\i - WWidth
      *OY\i - TotalHeight
    EndIf
    ;
    ; Apply caller offset and clamp to the monitor work area (checking right/bottom before left/top
    ; so the title bar remains reachable even if the window exceeds small screen dimensions):
    *OX\i + XShiftOrPos
    *OY\i + YShiftOrPos
    If *OX\i + WWidth > DesktopRight
      *OX\i = DesktopRight - WWidth - 10
    EndIf
    If *OX\i < DesktopLeft
      *OX\i = DesktopLeft + 10
    EndIf
    If *OY\i + TotalHeight > DesktopBottom
      *OY\i = DesktopBottom - TotalHeight - 10
    EndIf
    If *OY\i < DesktopTop
      *OY\i = DesktopTop + 10
    EndIf
    ;
    ProcedureReturn ParentWindowID
  EndProcedure
CompilerEndIf
;
CompilerIf Not Defined(WindowNumFromHandle, #PB_Procedure)
  Procedure WindowNumFromHandle(handle)
    ;
    ; Role: Retrieves the PureBasic window number associated with an OS window handle.
    ; Output: Returns the valid PureBasic Window ID, or -1 if 'handle' is not a PureBasic window.
    ;
    Protected Window = -1
    ;
    If handle
      CompilerSelect #PB_Compiler_OS
        CompilerCase #PB_OS_Linux
          Window = g_object_get_data_(handle, "pb_id")
        CompilerCase #PB_OS_Windows
          Window = GetProp_(handle, "PB_WINDOWID") - 1
        CompilerCase #PB_OS_MacOS
          Window = PB_Window_GetID(handle)
      CompilerEndSelect
      ;
      If Window >= 0 And IsWindow(Window) And WindowID(Window) = handle
        ProcedureReturn Window
      EndIf
    EndIf
    ProcedureReturn -1
  EndProcedure
CompilerEndIf
;
Structure InterfaceColorsStruct
  PresetName.s
  Editable.b
  BackgroundColor.l
  TextColor.l
  DefaultBackgroundColor.l
  DefaultTextColor.l
EndStructure
;
Global NewList InterfaceColorPresets.InterfaceColorsStruct()
;
Structure EditColorGadgetsStruct
  EditColorWindow.i
  BackColor.i
  TextColor.i
  BorderColor.i
  Color.i ; Updated on the fly as the user edits the color
  ;
  ; Layout parameters:
  MarginLeft.i
  MarginRight.i
  MarginTop.i
  MarginBottom.i
  MarginButtonsBottom.i
  MarginButtonsTop.i
  MarginInterBlocks.i
  InterBlocksVerticalPos.i
  HorizGadgetsMargin.i
  LegendWidth.i
  InputWidth.i
  HexColorWidth.i
  HexFieldVPos.i
  ButtonsWidth.i
  ;
  CanvasDim.i
  TargetDim.i
  CursorWidth.i
  LineHeight.i
  TrackBarsWidth.i
  ;
  ; Gadget IDs -> Square canvases
  PreviewCanvas.i
  RainbowCanvas.i
  ; Cached off-screen images (32-bit overlays and 2D rainbow picker buffer)
  TargetImage.i
  CursorImage.i
  RainbowImage.i
  ; Last rendered state used to skip redundant redraws
  PrevColor.i
  PrevBackColor.i
  PrevBorderColor.i
  PrevRGBGrad.i
  PrevCurR.i
  PrevCurG.i
  PrevCurB.i
  PrevInR.i
  PrevInG.i
  PrevInB.i
  PrevHueSat.i
  PrevHueLum.i
  PrevInH.i
  PrevInS.i
  PrevInL.i
  PrevLumRef.i
  PrevSatRef.i
  PrevRainbowSat.i
  PrevRainbowBorder.i
  PrevTargetH.i
  PrevTargetL.i
  ; Gadget IDs -> TrackBars
  Red_Trackbar.i
  Green_Trackbar.i
  Blue_Trackbar.i
  Hue_Trackbar.i
  Sat_Trackbar.i
  Lum_Trackbar.i
  ; Gadget IDs -> Legends (must remain contiguous down to BCancel for pointer iteration)
  Red_Legend.i
  Green_Legend.i
  Blue_Legend.i
  Hue_Legend.i
  Sat_Legend.i
  Lum_Legend.i
  HexColor_Legend.i
  ; Gadget IDs -> Input fields
  Red_Input.i
  Green_Input.i
  Blue_Input.i
  Hue_Input.i
  Sat_Input.i
  Lum_Input.i
  HexColor_Input.i
  ; Gadget IDs -> Buttons
  BOk.i
  BCancel.i
  ; Gadget IDs -> Separators
  SeparatorGadget1.i
  SeparatorGadget2.i
  SeparatorGadget3.i
  ;
  ; Parameter used when embedded in ApplyColorThemes
  SGC_ColorType$
EndStructure
;
EnumerationBinary DontSetOptions
  #DontSetRGBString
  #DontSetHSLString
  #DontSetHexField
EndEnumeration
;
Procedure MinOf3(value1, value2, value3)
  ;
  ; Role: Returns the smallest of three integer values.
  ;
  Protected result = value1
  ;
  If value2 < result : result = value2 : EndIf
  If value3 < result : result = value3 : EndIf
  ProcedureReturn result
EndProcedure
;
Procedure MaxOf3(value1, value2, value3)
  ;
  ; Role: Returns the largest of three integer values.
  ;
  Protected result = value1
  ;
  If value2 > result : result = value2 : EndIf
  If value3 > result : result = value3 : EndIf
  ProcedureReturn result
EndProcedure
;
Procedure Hue(rgbColor)
  ;
  ; Role: Extracts the Hue component of 'rgbColor' scaled to the [0..239] range.
  ;
  Protected red = Red(rgbColor)
  Protected green = Green(rgbColor)
  Protected blue = Blue(rgbColor)
  Protected maxValue = MaxOf3(red, green, blue)
  Protected minValue = MinOf3(red, green, blue)
  Protected delta = maxValue - minValue
  Protected hue.f = 0.0
  Protected result
  ;
  If delta
    If maxValue = red
      hue = (green - blue) / delta
    ElseIf maxValue = green
      hue = 2.0 + (blue - red) / delta
    Else
      hue = 4.0 + (red - green) / delta
    EndIf
    ;
    hue * 40.0
    If hue < 0.0
      hue + 240.0
    EndIf
  EndIf
  ;
  result = Round(hue, #PB_Round_Nearest)
  If result >= 240 : result - 240 : EndIf
  ProcedureReturn result
EndProcedure
;
Procedure Saturation(rgbColor)
  ;
  ; Role: Extracts the HSL Saturation component of 'rgbColor' scaled to the [0..240] range.
  ;
  Protected red = Red(rgbColor)
  Protected green = Green(rgbColor)
  Protected blue = Blue(rgbColor)
  Protected maxValue = MaxOf3(red, green, blue)
  Protected minValue = MinOf3(red, green, blue)
  Protected delta = maxValue - minValue
  Protected saturation.f = 0.0
  Protected lightness.f = (maxValue + minValue) / 510.0
  ;
  If delta
    If lightness <= 0.5
      saturation = delta / (maxValue + minValue)
    Else
      saturation = delta / (510 - maxValue - minValue)
    EndIf
  EndIf
  ;
  ProcedureReturn Round(saturation * 240.0, #PB_Round_Nearest)
EndProcedure
;
Procedure Lightness(rgbColor)
  ;
  ; Role: Extracts the HSL Lightness component of 'rgbColor' scaled to the [0..240] range.
  ;
  Protected red = Red(rgbColor)
  Protected green = Green(rgbColor)
  Protected blue = Blue(rgbColor)
  Protected maxValue = MaxOf3(red, green, blue)
  Protected minValue = MinOf3(red, green, blue)
  ;
  ProcedureReturn Round(240.0 * (maxValue + minValue) / 510.0, #PB_Round_Nearest)
EndProcedure
;
Procedure.i HSLToRGB(hue.i, saturation.i, lightness.i)
  ;
  ; Role: Converts HSL values (each in the [0..240] range) to a 24-bit RGB color.
  ;
  Protected r.f, g.f, b.f, c.f, x.f, m.f, sector.f
  Protected normalizedSaturation.f, normalizedLightness.f
  ;
  If hue < 0 : hue = 0 : ElseIf hue >= 240 : hue = hue % 240 : EndIf
  If saturation < 0 : saturation = 0 : ElseIf saturation > 240 : saturation = 240 : EndIf
  If lightness < 0 : lightness = 0 : ElseIf lightness > 240 : lightness = 240 : EndIf
  ;
  normalizedSaturation = saturation / 240.0
  normalizedLightness  = lightness / 240.0
  ;
  c = (1.0 - Abs(2.0 * normalizedLightness - 1.0)) * normalizedSaturation
  sector = hue / 40.0
  x = c * (1.0 - Abs(Mod(sector, 2.0) - 1.0))
  m = normalizedLightness - c / 2.0
  ;
  Select Int(sector)
    Case 0 : r = c : g = x : b = 0
    Case 1 : r = x : g = c : b = 0
    Case 2 : r = 0 : g = c : b = x
    Case 3 : r = 0 : g = x : b = c
    Case 4 : r = x : g = 0 : b = c
    Default: r = c : g = 0 : b = x
  EndSelect
  ;
  ProcedureReturn RGB(Round((r + m) * 255.0, #PB_Round_Nearest), Round((g + m) * 255.0, #PB_Round_Nearest), Round((b + m) * 255.0, #PB_Round_Nearest))
EndProcedure
;
Procedure LimitGadgetValue(Gadget, Min, Max)
  ;
  ; Role: Reads the integer value of 'Gadget', clamps it to [Min..Max], and updates the text if out of bounds.
  ;
  Protected Value = Val(GetGadgetText(Gadget))
  If Value > Max Or Value < Min
    If Value > Max : Value = Max : EndIf
    If Value < Min : Value = Min : EndIf
    SetGadgetText(Gadget, Str(Value))
  EndIf
  ProcedureReturn Value
EndProcedure
;
Procedure ColorShiftLum(CUnit)
  ;
  ; Role: Shifts a lightness value [0..240] away from mid-tones to ensure visible contrast for graduations.
  ;
  If CUnit < 127
    CUnit + 80
  Else
    CUnit - 100
  EndIf
  If CUnit > 200 : CUnit = 200 : ElseIf CUnit < 60 : CUnit = 60 : EndIf
  ProcedureReturn CUnit
EndProcedure
;
Procedure InvertedColor(Color)
  ;
  ; Role: Computes a high-contrast complementary color suitable for drawing markers over 'Color'.
  ;
  ProcedureReturn HSLToRGB((Hue(Color) + 120) % 240, 240, ColorShiftLum(Lightness(Color)))
EndProcedure
;
Macro DrawTrackBarBox()
  ;
  ; Role: Clears the trackbar canvas background, draws the outer border box, and resets graduation step 'tl'.
  ;
  Box(0, 0, DesktopScaledX(*ECGadgets\TrackBarsWidth), *ECGadgets\LineHeight, *ECGadgets\BackColor)
  Box(GStart - 1, 0, GWidth + 1, *ECGadgets\LineHeight, *ECGadgets\BorderColor)
  tl = stl
EndMacro
;
Macro DrawOneTrackBarLine(Color)
  ;
  ; Role: Draws one vertical color slice, center guide dots, and tick marks inside a trackbar canvas.
  ;
  x = ct + GStart - 1
  Line(x, 1, 1, *ECGadgets\LineHeight - 2, Color)
  Plot(x, GHalfHeight - 1, GradColor)
  Plot(x, GHalfHeight + 1, GradColor)
  If ct = Int(tl)
    tl + stl
    Line(x, 1, 1, 3, GradColor)
  EndIf
EndMacro
;
CompilerIf Not Defined(MixColors, #PB_Procedure)
  Procedure MixColors(Color1, Color2, Ratio.f)
    ;
    ; Role: Blends two RGB colors according to 'Ratio' (0.0 = Color2, 1.0 = Color1).
    ;
    If Ratio > 1 : Ratio = 1 : ElseIf Ratio < 0 : Ratio = 0 : EndIf
    Protected Red = Red(Color1) * Ratio + Red(Color2) * (1 - Ratio)
    Protected Green = Green(Color1) * Ratio + Green(Color2) * (1 - Ratio)
    Protected Blue = Blue(Color1) * Ratio + Blue(Color2) * (1 - Ratio)
    ProcedureReturn RGB(Red, Green, Blue)
  EndProcedure
CompilerEndIf
;
Procedure RepaintCanvasSeparator(CanvasGadget, Color)
  ;
  ; Role: Fills a 1-pixel separator CanvasGadget with 'Color'.
  ;
  If IsGadget(CanvasGadget) And GadgetType(CanvasGadget) = #PB_GadgetType_Canvas And StartDrawing(CanvasOutput(CanvasGadget))
    Box(0, 0, DesktopScaledX(GadgetWidth(CanvasGadget)), DesktopScaledY(GadgetHeight(CanvasGadget)), Color)
    StopDrawing()
  EndIf
EndProcedure
;
Procedure GetRealColorFromType(ColorType$, Color)
  ;
  ; Role: Resolves #PB_Default into an actual RGB system color for 'BackgroundColor' or 'TextColor'.
  ;
  If Color = #PB_Default
    CompilerIf #PB_Compiler_OS = #PB_OS_Windows
      If ColorType$ = "BackgroundColor"
        Color = GetSysColor_(#COLOR_BTNFACE)
      ElseIf ColorType$ = "TextColor"
        Color = GetSysColor_(#COLOR_WINDOWTEXT)
      EndIf
    CompilerElse
      If ColorType$ = "BackgroundColor"
        Color = $F0F0F0
      ElseIf ColorType$ = "TextColor"
        Color = $000000
      EndIf
    CompilerEndIf
  EndIf
  ProcedureReturn Color
EndProcedure
;
Procedure CalculateBorderColor(TextColor, BackgroundColor)
  ;
  ; Role: Blends TextColor (40%) and BackgroundColor (60%) to produce a subtle border shade.
  ;
  ProcedureReturn MixColors(GetRealColorFromType("TextColor", TextColor), GetRealColorFromType("BackgroundColor", BackgroundColor), 0.4)
EndProcedure
;
Procedure ZCR_SetGadgetTextIfChanged(Gadget, NewText$)
  ;
  ; Role: Updates 'Gadget' text only when its content actually differs from 'NewText$'.
  ;
  If GetGadgetText(Gadget) <> NewText$
    SetGadgetText(Gadget, NewText$)
  EndIf
EndProcedure
;
Procedure ZCR_ResetRenderCache(*ECGadgets.EditColorGadgetsStruct)
  ;
  ; Role: Invalidates all cached render states so the next update repaints every canvas completely.
  ;
  *ECGadgets\PrevColor         = -1
  *ECGadgets\PrevBackColor     = -1
  *ECGadgets\PrevBorderColor   = -1
  *ECGadgets\PrevRGBGrad       = -1
  *ECGadgets\PrevCurR          = -1
  *ECGadgets\PrevCurG          = -1
  *ECGadgets\PrevCurB          = -1
  *ECGadgets\PrevInR           = -1
  *ECGadgets\PrevInG           = -1
  *ECGadgets\PrevInB           = -1
  *ECGadgets\PrevHueSat        = -1
  *ECGadgets\PrevHueLum        = -1
  *ECGadgets\PrevInH           = -1
  *ECGadgets\PrevInS           = -1
  *ECGadgets\PrevInL           = -1
  *ECGadgets\PrevLumRef        = -1
  *ECGadgets\PrevSatRef        = -1
  *ECGadgets\PrevRainbowSat    = -1
  *ECGadgets\PrevRainbowBorder = -1
  *ECGadgets\PrevTargetH       = -1
  *ECGadgets\PrevTargetL       = -1
EndProcedure
;
Procedure ZCR_CalculateBorderColor(*ECGadgets.EditColorGadgetsStruct, RepaintCase$)
  ;
  ; Role: Normalizes theme colors, updates *ECGadgets\BorderColor, and repaints separator lines on change.
  ;
  Protected NewBorderColor
  ;
  If RepaintCase$ = "BackgroundColor"
    ; Used by ApplyColorThemes when live-previewing background color changes
    *ECGadgets\BackColor = *ECGadgets\Color
  ElseIf RepaintCase$ = "TextColor"
    ; Used by ApplyColorThemes when live-previewing text color changes
    *ECGadgets\TextColor = *ECGadgets\Color
  EndIf
  ;
  *ECGadgets\BackColor = GetRealColorFromType("BackgroundColor", *ECGadgets\BackColor)
  *ECGadgets\TextColor = GetRealColorFromType("TextColor", *ECGadgets\TextColor)
  NewBorderColor       = CalculateBorderColor(*ECGadgets\TextColor, *ECGadgets\BackColor)
  ;
  If RepaintCase$ = "Resize" Or NewBorderColor <> *ECGadgets\BorderColor
    *ECGadgets\BorderColor = NewBorderColor
    RepaintCanvasSeparator(*ECGadgets\SeparatorGadget1, NewBorderColor)
    RepaintCanvasSeparator(*ECGadgets\SeparatorGadget2, NewBorderColor)
    RepaintCanvasSeparator(*ECGadgets\SeparatorGadget3, NewBorderColor)
  EndIf
EndProcedure
;
Procedure FillCanvasWithColor(Canvas, Color, BorderColor)
  ;
  ; Role: Fills 'Canvas' with 'Color' framed by a 1-pixel 'BorderColor' outline.
  ;
  If StartDrawing(CanvasOutput(Canvas))
    Box(0, 0, DesktopScaledX(GadgetWidth(Canvas)), DesktopScaledY(GadgetHeight(Canvas)), BorderColor)
    Box(1, 1, DesktopScaledX(GadgetWidth(Canvas)) - 2, DesktopScaledY(GadgetHeight(Canvas)) - 2, Color)
    StopDrawing()
  EndIf
EndProcedure
;
Macro DrawTrackBarCursor(CursorValue, CursorMax)
  ;
  ; Role: Alpha-blends the slider cursor image directly into the active trackbar CanvasOutput().
  ;
  If IsImage(*ECGadgets\CursorImage)
    CursorClamped = CursorValue
    If CursorClamped < 0 : CursorClamped = 0 : ElseIf CursorClamped > CursorMax : CursorClamped = CursorMax : EndIf
    DrawAlphaImage(ImageID(*ECGadgets\CursorImage), (CursorClamped * (DesktopScaledX(*ECGadgets\TrackBarsWidth) - *ECGadgets\CursorWidth)) / CursorMax, (OutputHeight() - ImageHeight(*ECGadgets\CursorImage)) / 2)
  EndIf
EndMacro
;
Procedure ZCR_DrawTrackbarsAndPreviewCanvas(*ECGadgets.EditColorGadgetsStruct)
  ;
  ; Role: Selectively repaints the preview square and only those RGB/HSL trackbar canvases
  ;       whose gradient parameters or cursor position have changed since the last frame.
  ;
  Protected x, GWidth, GStart, GHalfHeight, GHalfWidth, ThemeChanged, NeedDraw, CursorVal, CursorClamped
  Protected DecreasingGrey, IncreasingWhite, RefLum
  Protected CurR, CurG, CurB, InR, InG, InB, InH, InS, InL
  Protected Red, Green, Blue, NRed, NGreen, NBlue, RGBGradColor, GradColor, GreyColor
  Protected HueSat, HueLum, LumRefColor, SatRefColor, Hue
  Protected stl.f, tl.f, ct, ct2, ci, Trackbar
  ;
  ZCR_CalculateBorderColor(*ECGadgets, *ECGadgets\SGC_ColorType$)
  ThemeChanged = Bool(*ECGadgets\BackColor <> *ECGadgets\PrevBackColor Or *ECGadgets\BorderColor <> *ECGadgets\PrevBorderColor)
  ;
  If ThemeChanged Or *ECGadgets\Color <> *ECGadgets\PrevColor
    FillCanvasWithColor(*ECGadgets\PreviewCanvas, *ECGadgets\Color, *ECGadgets\BorderColor)
  EndIf
  ;
  ; Gather current color components and active slider values
  CurR = Red(*ECGadgets\Color)
  CurG = Green(*ECGadgets\Color)
  CurB = Blue(*ECGadgets\Color)
  InR  = LimitGadgetValue(*ECGadgets\Red_Input, 0, 255)
  InG  = LimitGadgetValue(*ECGadgets\Green_Input, 0, 255)
  InB  = LimitGadgetValue(*ECGadgets\Blue_Input, 0, 255)
  InH  = LimitGadgetValue(*ECGadgets\Hue_Input, 0, 240)
  InS  = LimitGadgetValue(*ECGadgets\Sat_Input, 0, 240)
  InL  = LimitGadgetValue(*ECGadgets\Lum_Input, 0, 240)
  ;
  ; Compute shared trackbar geometry and graduation spacing
  GWidth      = DesktopScaledX(*ECGadgets\TrackBarsWidth) - *ECGadgets\CursorWidth + 1
  If GWidth < 2 : GWidth = 2 : EndIf
  GStart      = *ECGadgets\CursorWidth / 2
  GHalfHeight = *ECGadgets\LineHeight / 2
  stl         = (GWidth + 1) / 6.0
  ;
  ; Draw RGB trackbars (each bar only repaints if its own cursor, cross-channels, or tick contrast changed)
  If Lightness(*ECGadgets\Color) < 60
    RGBGradColor = RGB(160, 160, 160)
  ElseIf Lightness(*ECGadgets\Color) > 130
    RGBGradColor = RGB(140, 140, 50)
  Else
    RGBGradColor = RGB(255, 255, 255)
  EndIf
  GradColor = RGBGradColor
  ;
  For ci = 1 To 3
    Select ci
      Case 1
        NeedDraw  = Bool(ThemeChanged Or CurG <> *ECGadgets\PrevCurG Or CurB <> *ECGadgets\PrevCurB Or RGBGradColor <> *ECGadgets\PrevRGBGrad Or InR <> *ECGadgets\PrevInR)
        Trackbar  = *ECGadgets\Red_Trackbar
        CursorVal = InR
      Case 2
        NeedDraw  = Bool(ThemeChanged Or CurR <> *ECGadgets\PrevCurR Or CurB <> *ECGadgets\PrevCurB Or RGBGradColor <> *ECGadgets\PrevRGBGrad Or InG <> *ECGadgets\PrevInG)
        Trackbar  = *ECGadgets\Green_Trackbar
        CursorVal = InG
      Case 3
        NeedDraw  = Bool(ThemeChanged Or CurR <> *ECGadgets\PrevCurR Or CurG <> *ECGadgets\PrevCurG Or RGBGradColor <> *ECGadgets\PrevRGBGrad Or InB <> *ECGadgets\PrevInB)
        Trackbar  = *ECGadgets\Blue_Trackbar
        CursorVal = InB
    EndSelect
    ;
    If NeedDraw And StartDrawing(CanvasOutput(Trackbar))
      Red   = CurR
      Green = CurG
      Blue  = CurB
      DrawTrackBarBox()
      For ct = 1 To GWidth - 1
        If ci = 1
          Red = (ct * 255) / GWidth
        ElseIf ci = 2
          Green = (ct * 255) / GWidth
        Else
          Blue = (ct * 255) / GWidth
        EndIf
        DrawOneTrackBarLine(RGB(Red, Green, Blue))
      Next
      DrawTrackBarCursor(CursorVal, 255)
      StopDrawing()
    EndIf
  Next
  ;
  ; Draw Hue trackbar
  HueSat = Saturation(*ECGadgets\Color)
  HueLum = Lightness(*ECGadgets\Color)
  If ThemeChanged Or HueSat <> *ECGadgets\PrevHueSat Or HueLum <> *ECGadgets\PrevHueLum Or InH <> *ECGadgets\PrevInH
    If StartDrawing(CanvasOutput(*ECGadgets\Hue_Trackbar))
      DrawTrackBarBox()
      GradColor = HSLToRGB(120, 120, ColorShiftLum(HueLum))
      For ct = 1 To GWidth - 1
        Hue = (ct * 240) / GWidth
        DrawOneTrackBarLine(HSLToRGB(Hue, HueSat, HueLum))
      Next
      DrawTrackBarCursor(InH, 240)
      StopDrawing()
    EndIf
  EndIf
  ;
  ; Draw Lightness trackbar
  LumRefColor = HSLToRGB(InH, InS, 120)
  If ThemeChanged Or LumRefColor <> *ECGadgets\PrevLumRef Or InL <> *ECGadgets\PrevInL
    If StartDrawing(CanvasOutput(*ECGadgets\Lum_Trackbar))
      DrawTrackBarBox()
      Red        = Red(LumRefColor)
      Green      = Green(LumRefColor)
      Blue       = Blue(LumRefColor)
      GradColor  = RGB(127, 127, 127)
      GHalfWidth = GWidth / 2
      If GHalfWidth < 1 : GHalfWidth = 1 : EndIf
      ;
      For ct = 1 To GWidth - 1
        If ct < GHalfWidth
          NRed   = (Red   * ct) / GHalfWidth
          NGreen = (Green * ct) / GHalfWidth
          NBlue  = (Blue  * ct) / GHalfWidth
        Else
          ct2             = ct - GHalfWidth
          IncreasingWhite = (255 * ct2) / GHalfWidth
          NRed            = (Red   * (GHalfWidth - ct2)) / GHalfWidth + IncreasingWhite
          NGreen          = (Green * (GHalfWidth - ct2)) / GHalfWidth + IncreasingWhite
          NBlue           = (Blue  * (GHalfWidth - ct2)) / GHalfWidth + IncreasingWhite
        EndIf
        DrawOneTrackBarLine(RGB(NRed, NGreen, NBlue))
      Next
      DrawTrackBarCursor(InL, 240)
      StopDrawing()
    EndIf
  EndIf
  ;
  ; Draw Saturation trackbar
  If InL = 0 : RefLum = 1 : Else : RefLum = InL : EndIf
  SatRefColor = HSLToRGB(InH, 240, RefLum)
  If ThemeChanged Or SatRefColor <> *ECGadgets\PrevSatRef Or InL <> *ECGadgets\PrevInL Or InS <> *ECGadgets\PrevInS
    If StartDrawing(CanvasOutput(*ECGadgets\Sat_Trackbar))
      DrawTrackBarBox()
      Red       = Red(SatRefColor)
      Green     = Green(SatRefColor)
      Blue      = Blue(SatRefColor)
      GradColor = InvertedColor(SatRefColor)
      GreyColor = (InL * 255) / 240
      ;
      For ct = 1 To GWidth - 1
        DecreasingGrey = (GreyColor * (GWidth - ct)) / GWidth
        NRed           = (Red   * ct) / GWidth + DecreasingGrey
        NGreen         = (Green * ct) / GWidth + DecreasingGrey
        NBlue          = (Blue  * ct) / GWidth + DecreasingGrey
        DrawOneTrackBarLine(RGB(NRed, NGreen, NBlue))
      Next
      DrawTrackBarCursor(InS, 240)
      StopDrawing()
    EndIf
  EndIf
  ;
  ; Record rendered state for the next frame
  *ECGadgets\PrevColor       = *ECGadgets\Color
  *ECGadgets\PrevBackColor   = *ECGadgets\BackColor
  *ECGadgets\PrevBorderColor = *ECGadgets\BorderColor
  *ECGadgets\PrevRGBGrad     = RGBGradColor
  *ECGadgets\PrevCurR        = CurR
  *ECGadgets\PrevCurG        = CurG
  *ECGadgets\PrevCurB        = CurB
  *ECGadgets\PrevInR         = InR
  *ECGadgets\PrevInG         = InG
  *ECGadgets\PrevInB         = InB
  *ECGadgets\PrevHueSat      = HueSat
  *ECGadgets\PrevHueLum      = HueLum
  *ECGadgets\PrevInH         = InH
  *ECGadgets\PrevInS         = InS
  *ECGadgets\PrevInL         = InL
  *ECGadgets\PrevLumRef      = LumRefColor
  *ECGadgets\PrevSatRef      = SatRefColor
EndProcedure
;
Procedure ZCR_SetHexFieldFromColor(*ECGadgets.EditColorGadgetsStruct)
  ;
  ; Role: Formats *ECGadgets\Color as a 6-digit PureBasic BGR hexadecimal string into HexColor_Input.
  ;
  Protected HStr$ = RSet(Hex(*ECGadgets\Color & $FFFFFF, #PB_Long), 6, "0")
  If GetGadgetText(*ECGadgets\HexColor_Input) <> HStr$
    SetGadgetText(*ECGadgets\HexColor_Input, HStr$)
    CompilerIf #PB_Compiler_OS = #PB_OS_Windows
      SendMessage_(GadgetID(*ECGadgets\HexColor_Input), #EM_SETSEL, Len(HStr$), Len(HStr$))
    CompilerEndIf
  EndIf
EndProcedure
;
Procedure ZCR_DrawRainbowCanvas(*ECGadgets.EditColorGadgetsStruct)
  ;
  ; Role: Caches the 2D Hue/Lightness rainbow gradient in *ECGadgets\RainbowImage (recomputed only
  ;       when Saturation, BorderColor, or size changes) and composites the target reticle onto RainbowCanvas.
  ;
  Protected CanvasW, CanvasH, InnerW, InnerH, MaxPx, HalfH, BottomH, RemH, WhitePart
  Protected px, py, InH, InS, InL, RebuildRainbow, TargetX, TargetY
  Protected Hue, RefColor, Red, Green, Blue, NRed, NGreen, NBlue
  ;
  CanvasW = DesktopScaledX(*ECGadgets\CanvasDim)
  CanvasH = DesktopScaledY(*ECGadgets\CanvasDim)
  If CanvasW <= 2 Or CanvasH <= 2
    ProcedureReturn
  EndIf
  ;
  InH = LimitGadgetValue(*ECGadgets\Hue_Input, 0, 240)
  InS = LimitGadgetValue(*ECGadgets\Sat_Input, 0, 240)
  InL = LimitGadgetValue(*ECGadgets\Lum_Input, 0, 240)
  ;
  If Not IsImage(*ECGadgets\RainbowImage) Or ImageWidth(*ECGadgets\RainbowImage) <> CanvasW Or ImageHeight(*ECGadgets\RainbowImage) <> CanvasH
    If IsImage(*ECGadgets\RainbowImage)
      FreeImage(*ECGadgets\RainbowImage)
    EndIf
    *ECGadgets\RainbowImage = CreateImage(#PB_Any, CanvasW, CanvasH, 24)
    RebuildRainbow = #True
  ElseIf InS <> *ECGadgets\PrevRainbowSat Or *ECGadgets\BorderColor <> *ECGadgets\PrevRainbowBorder
    RebuildRainbow = #True
  EndIf
  ;
  If RebuildRainbow And IsImage(*ECGadgets\RainbowImage)
    If StartDrawing(ImageOutput(*ECGadgets\RainbowImage))
      Box(0, 0, CanvasW, CanvasH, *ECGadgets\BorderColor)
      InnerW  = CanvasW - 2
      InnerH  = CanvasH - 2
      MaxPx   = InnerW - 1 : If MaxPx < 1 : MaxPx = 1 : EndIf
      HalfH   = InnerH / 2 : If HalfH < 1 : HalfH = 1 : EndIf
      BottomH = InnerH - HalfH : If BottomH < 1 : BottomH = 1 : EndIf
      ;
      For px = 0 To InnerW - 1
        Hue      = (px * 239) / MaxPx
        RefColor = HSLToRGB(Hue, InS, 120)
        Red      = Red(RefColor)
        Green    = Green(RefColor)
        Blue     = Blue(RefColor)
        ;
        For py = 0 To InnerH - 1
          If py < HalfH
            WhitePart = (255 * (HalfH - py)) / HalfH
            NRed      = (Red   * py) / HalfH + WhitePart
            NGreen    = (Green * py) / HalfH + WhitePart
            NBlue     = (Blue  * py) / HalfH + WhitePart
          Else
            RemH   = InnerH - 1 - py
            NRed   = (Red   * RemH) / BottomH
            NGreen = (Green * RemH) / BottomH
            NBlue  = (Blue  * RemH) / BottomH
          EndIf
          Plot(px + 1, py + 1, RGB(NRed, NGreen, NBlue))
        Next
      Next
      StopDrawing()
    EndIf
    *ECGadgets\PrevRainbowSat    = InS
    *ECGadgets\PrevRainbowBorder = *ECGadgets\BorderColor
  EndIf
  ;
  If (RebuildRainbow Or InH <> *ECGadgets\PrevTargetH Or InL <> *ECGadgets\PrevTargetL) And IsImage(*ECGadgets\RainbowImage)
    If StartDrawing(CanvasOutput(*ECGadgets\RainbowCanvas))
      DrawImage(ImageID(*ECGadgets\RainbowImage), 0, 0)
      If IsImage(*ECGadgets\TargetImage)
        TargetX = (InH * CanvasW) / 240 - *ECGadgets\TargetDim / 2
        TargetY = ((240 - InL) * CanvasH) / 240 - *ECGadgets\TargetDim / 2
        DrawAlphaImage(ImageID(*ECGadgets\TargetImage), TargetX, TargetY)
      EndIf
      StopDrawing()
    EndIf
    *ECGadgets\PrevTargetH = InH
    *ECGadgets\PrevTargetL = InL
  EndIf
EndProcedure
;
Procedure ZCR_SetRGBStringsFromColor(*ECGadgets.EditColorGadgetsStruct)
  ;
  ; Role: Updates the Red, Green, and Blue numeric input fields from *ECGadgets\Color when modified.
  ;
  ZCR_SetGadgetTextIfChanged(*ECGadgets\Red_Input, Str(Red(*ECGadgets\Color)))
  ZCR_SetGadgetTextIfChanged(*ECGadgets\Green_Input, Str(Green(*ECGadgets\Color)))
  ZCR_SetGadgetTextIfChanged(*ECGadgets\Blue_Input, Str(Blue(*ECGadgets\Color)))
EndProcedure
;
Procedure ZCR_SetHSLStringsFromColor(*ECGadgets.EditColorGadgetsStruct)
  ;
  ; Role: Updates the Hue, Saturation, and Lightness numeric input fields from *ECGadgets\Color when modified.
  ;
  ZCR_SetGadgetTextIfChanged(*ECGadgets\Hue_Input, Str(Hue(*ECGadgets\Color)))
  ZCR_SetGadgetTextIfChanged(*ECGadgets\Sat_Input, Str(Saturation(*ECGadgets\Color)))
  ZCR_SetGadgetTextIfChanged(*ECGadgets\Lum_Input, Str(Lightness(*ECGadgets\Color)))
EndProcedure
;
Procedure ZCR_SetGadgetsFromColor(*ECGadgets.EditColorGadgetsStruct, DontSet = 0)
  ;
  ; Role: Synchronizes text inputs (unless masked by 'DontSet') and selectively refreshes modified canvases.
  ;
  If DontSet & #DontSetRGBString = 0
    ZCR_SetRGBStringsFromColor(*ECGadgets)
  EndIf
  If DontSet & #DontSetHSLString = 0
    ZCR_SetHSLStringsFromColor(*ECGadgets)
  EndIf
  If DontSet & #DontSetHexField = 0
    ZCR_SetHexFieldFromColor(*ECGadgets)
  EndIf
  ZCR_DrawTrackbarsAndPreviewCanvas(*ECGadgets)
  ZCR_DrawRainbowCanvas(*ECGadgets)
EndProcedure
;
Define ECGadgets.EditColorGadgetsStruct
;
Procedure ZCR_ResizeGadgets()
  ;
  ; Role: Computes responsive layout coordinates, resizes all requester gadgets, and repaints canvases.
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  Protected HexFieldVPos, HexFieldHPos
  Protected TextHeight = ECGadgets\LineHeight / DesktopResolutionY()
  Protected ButtonHeight = TextHeight + 3
  ;
  ; Reserve bottom margin space for the Cancel and OK buttons if not explicitly provided
  If ECGadgets\MarginBottom = 0
    ECGadgets\MarginBottom = 2 * ECGadgets\MarginButtonsBottom + ButtonHeight + ECGadgets\MarginButtonsTop
  EndIf
  ;
  Protected AreaHeight = WindowHeight(ECGadgets\EditColorWindow) - ECGadgets\MarginTop - ECGadgets\MarginBottom - ECGadgets\MarginInterBlocks
  Protected AreaWidth  = WindowWidth(ECGadgets\EditColorWindow) - ECGadgets\MarginLeft - ECGadgets\MarginRight
  ;
  Protected HorizLineShift = 3
  Protected InterLine = (AreaHeight - 2 * TextHeight + (2 * HorizLineShift)) / 4.75
  ;
  ECGadgets\CanvasDim      = InterLine * 2 + TextHeight
  ECGadgets\TrackBarsWidth = AreaWidth - ECGadgets\LegendWidth - ECGadgets\InputWidth - ECGadgets\CanvasDim - (ECGadgets\HorizGadgetsMargin * 3)
  ;
  Protected VPos         = ECGadgets\MarginTop
  Protected LegendHPos   = ECGadgets\MarginLeft
  Protected InputHPos    = LegendHPos + ECGadgets\LegendWidth + ECGadgets\HorizGadgetsMargin
  Protected TrackBarHPos = InputHPos + ECGadgets\InputWidth + ECGadgets\HorizGadgetsMargin
  Protected CanvasHPos   = TrackBarHPos + ECGadgets\TrackBarsWidth + ECGadgets\HorizGadgetsMargin
  ;
  ResizeGadget(ECGadgets\PreviewCanvas, CanvasHPos, VPos - 2, ECGadgets\CanvasDim * DesktopResolutionX() / DesktopResolutionY(), ECGadgets\CanvasDim)
  ;
  ResizeGadget(ECGadgets\Red_Legend, LegendHPos, VPos, ECGadgets\LegendWidth, TextHeight)
  ResizeGadget(ECGadgets\Red_Input, InputHPos, VPos - 2, ECGadgets\InputWidth, TextHeight)
  ResizeGadget(ECGadgets\Red_Trackbar, TrackBarHPos, VPos - 2, ECGadgets\TrackBarsWidth, TextHeight)
  ;
  VPos + InterLine
  ResizeGadget(ECGadgets\Green_Legend, LegendHPos, VPos, ECGadgets\LegendWidth, TextHeight)
  ResizeGadget(ECGadgets\Green_Input, InputHPos, VPos - 2, ECGadgets\InputWidth, TextHeight)
  ResizeGadget(ECGadgets\Green_Trackbar, TrackBarHPos, VPos - 2, ECGadgets\TrackBarsWidth, TextHeight)
  ;
  VPos + InterLine
  ResizeGadget(ECGadgets\Blue_Legend, LegendHPos, VPos, ECGadgets\LegendWidth, TextHeight)
  ResizeGadget(ECGadgets\Blue_Input, InputHPos, VPos - 2, ECGadgets\InputWidth, TextHeight)
  ResizeGadget(ECGadgets\Blue_Trackbar, TrackBarHPos, VPos - 2, ECGadgets\TrackBarsWidth, TextHeight)
  ;
  VPos + TextHeight - HorizLineShift + InterLine * 0.25
  ResizeGadget(ECGadgets\SeparatorGadget1, ECGadgets\MarginLeft, VPos, AreaWidth, 1)
  ;
  If ECGadgets\MarginInterBlocks
    ECGadgets\InterBlocksVerticalPos = VPos
    VPos + ECGadgets\MarginInterBlocks
    ResizeGadget(ECGadgets\SeparatorGadget2, ECGadgets\MarginLeft, VPos, AreaWidth, 1)
  EndIf
  ;
  VPos + InterLine * 0.25 + HorizLineShift
  ResizeGadget(ECGadgets\RainbowCanvas, CanvasHPos, VPos - 2, ECGadgets\CanvasDim * DesktopResolutionX() / DesktopResolutionY(), ECGadgets\CanvasDim)
  ;
  ResizeGadget(ECGadgets\Hue_Legend, LegendHPos, VPos, ECGadgets\LegendWidth, TextHeight)
  ResizeGadget(ECGadgets\Hue_Input, InputHPos, VPos - 2, ECGadgets\InputWidth, TextHeight)
  ResizeGadget(ECGadgets\Hue_Trackbar, TrackBarHPos, VPos - 2, ECGadgets\TrackBarsWidth, TextHeight)
  ;
  VPos + InterLine
  ResizeGadget(ECGadgets\Lum_Legend, LegendHPos, VPos, ECGadgets\LegendWidth, TextHeight)
  ResizeGadget(ECGadgets\Lum_Input, InputHPos, VPos - 2, ECGadgets\InputWidth, TextHeight)
  ResizeGadget(ECGadgets\Lum_Trackbar, TrackBarHPos, VPos - 2, ECGadgets\TrackBarsWidth, TextHeight)
  ;
  VPos + InterLine
  ResizeGadget(ECGadgets\Sat_Legend, LegendHPos, VPos, ECGadgets\LegendWidth, TextHeight)
  ResizeGadget(ECGadgets\Sat_Input, InputHPos, VPos - 2, ECGadgets\InputWidth, TextHeight)
  ResizeGadget(ECGadgets\Sat_Trackbar, TrackBarHPos, VPos - 2, ECGadgets\TrackBarsWidth, TextHeight)
  ;
  VPos + TextHeight - HorizLineShift + InterLine * 0.25
  ResizeGadget(ECGadgets\SeparatorGadget3, ECGadgets\MarginLeft, VPos, AreaWidth, 1)
  ;
  ; Position Hex field and Cancel / OK buttons
  VPos = WindowHeight(ECGadgets\EditColorWindow) - ECGadgets\MarginButtonsBottom - ButtonHeight
  ;
  HexFieldHPos = ECGadgets\MarginLeft
  If ECGadgets\HexFieldVPos = -1
    HexFieldVPos = VPos + 1
  ElseIf ECGadgets\HexFieldVPos = -2
    HexFieldVPos = ECGadgets\InterBlocksVerticalPos + 6
    HexFieldHPos = WindowWidth(ECGadgets\EditColorWindow) - ECGadgets\LegendWidth - ECGadgets\HorizGadgetsMargin - ECGadgets\HexColorWidth - ECGadgets\MarginRight
  Else
    HexFieldVPos = ECGadgets\HexFieldVPos
  EndIf
  ;
  ResizeGadget(ECGadgets\HexColor_Legend, HexFieldHPos, HexFieldVPos + 1, ECGadgets\LegendWidth, TextHeight)
  ResizeGadget(ECGadgets\HexColor_Input, HexFieldHPos + ECGadgets\LegendWidth + ECGadgets\HorizGadgetsMargin, HexFieldVPos, ECGadgets\HexColorWidth, TextHeight)
  ResizeGadget(ECGadgets\BCancel, WindowWidth(ECGadgets\EditColorWindow) - ECGadgets\MarginRight - 2 * ECGadgets\ButtonsWidth - ECGadgets\HorizGadgetsMargin, VPos, ECGadgets\ButtonsWidth, ButtonHeight)
  ResizeGadget(ECGadgets\BOk, WindowWidth(ECGadgets\EditColorWindow) - ECGadgets\MarginRight - ECGadgets\ButtonsWidth, VPos, ECGadgets\ButtonsWidth, ButtonHeight)
  ;
  ZCR_ResetRenderCache(ECGadgets)
  ZCR_CalculateBorderColor(ECGadgets, "Resize")
  ZCR_DrawTrackbarsAndPreviewCanvas(ECGadgets)
  ZCR_DrawRainbowCanvas(ECGadgets)
EndProcedure
;
Procedure ZCR_SetText()
  ;
  ; Role: Applies localized labels and tooltips to all requester gadgets.
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  ;
  SetGadgetText(ECGadgets\Red_Legend,   GetTextFromCatalog("Red"))
  SetGadgetText(ECGadgets\Green_Legend, GetTextFromCatalog("Green"))
  SetGadgetText(ECGadgets\Blue_Legend,  GetTextFromCatalog("Blue"))
  ;
  SetGadgetText(ECGadgets\Hue_Legend, GetTextFromCatalog("Hue"))
  SetGadgetText(ECGadgets\Sat_Legend, GetTextFromCatalog("Saturation"))
  SetGadgetText(ECGadgets\Lum_Legend, GetTextFromCatalog("Lightness"))
  ;
  SetGadgetText(ECGadgets\HexColor_Legend, GetTextFromCatalog("HexValue"))
  SetGadgetText(ECGadgets\BCancel, GetTextFromCatalog("Cancel"))
  SetGadgetText(ECGadgets\BOk, GetTextFromCatalog("OK"))
  ;
  GadgetToolTip(ECGadgets\RainbowCanvas, GetTextFromCatalog("ClickToAdjustColor"))
EndProcedure
;
Procedure ZCR_CreateGadgets()
  ;
  ; Role: Instantiates all canvases, text legends, numeric inputs, separators, and action buttons,
  ;       and resets the render state cache for the new requester instance.
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  ;
  ZCR_ResetRenderCache(ECGadgets)
  ;
  ECGadgets\PreviewCanvas = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Red_Legend   = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Red_Input    = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Red_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Green_Legend   = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Green_Input    = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Green_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Blue_Legend   = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Blue_Input    = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Blue_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\SeparatorGadget1 = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  If ECGadgets\MarginInterBlocks
    ECGadgets\SeparatorGadget2 = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  EndIf
  ;
  ECGadgets\RainbowCanvas = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Hue_Legend   = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Hue_Input    = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Hue_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Lum_Legend   = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Lum_Input    = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Lum_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Sat_Legend   = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Sat_Input    = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Sat_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\SeparatorGadget3 = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\HexColor_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\HexColor_Input  = StringGadget(#PB_Any, 1, 1, 1, 1, "")
  ;
  ECGadgets\BCancel = ButtonGadget(#PB_Any, 1, 1, 1, 22, "")
  ECGadgets\BOk     = ButtonGadget(#PB_Any, 1, 1, 1, 22, "")
EndProcedure
;
Procedure ZCR_CreateTargetAndCursors()
  ;
  ; Role: Builds (or reuses cached) 32-bit transparent vector images for the rainbow reticle
  ;       and trackbar slider cursors drawn onto their host canvases.
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  Protected ct
  ;
  If Not IsImage(ECGadgets\TargetImage) Or ImageWidth(ECGadgets\TargetImage) <> ECGadgets\TargetDim
    If IsImage(ECGadgets\TargetImage) : FreeImage(ECGadgets\TargetImage) : EndIf
    ECGadgets\TargetImage = CreateImage(#PB_Any, ECGadgets\TargetDim, ECGadgets\TargetDim, 32, #PB_Image_Transparent)
    If StartVectorDrawing(ImageVectorOutput(ECGadgets\TargetImage))
      AddPathCircle(ECGadgets\TargetDim / 2, ECGadgets\TargetDim / 2, ECGadgets\TargetDim / 2 - 2)
      ClosePath()
      VectorSourceColor(RGBA(120, 120, 120, 255))
      StrokePath(1)
      AddPathCircle(ECGadgets\TargetDim / 2, ECGadgets\TargetDim / 2, ECGadgets\TargetDim / 2 - 3)
      ClosePath()
      VectorSourceColor(RGBA(255, 255, 255, 255))
      StrokePath(1)
      StopVectorDrawing()
    EndIf
  EndIf
  ;
  If Not IsImage(ECGadgets\CursorImage) Or ImageWidth(ECGadgets\CursorImage) <> ECGadgets\CursorWidth Or ImageHeight(ECGadgets\CursorImage) <> ECGadgets\LineHeight
    If IsImage(ECGadgets\CursorImage) : FreeImage(ECGadgets\CursorImage) : EndIf
    ECGadgets\CursorImage = CreateImage(#PB_Any, ECGadgets\CursorWidth, ECGadgets\LineHeight, 32, #PB_Image_Transparent)
    If StartVectorDrawing(ImageVectorOutput(ECGadgets\CursorImage))
      For ct = 1 To 3
        MovePathCursor(ECGadgets\CursorWidth / 2 - ct + 2, 3)
        AddPathLine(ECGadgets\CursorWidth / 2 - ct + 2, ECGadgets\LineHeight - 3)
        ClosePath()
        If ct = 2
          VectorSourceColor(RGBA(255, 255, 255, 255))
        Else
          VectorSourceColor(RGBA(120, 120, 120, 255))
        EndIf
        StrokePath(1)
      Next
      AddPathCircle(ECGadgets\CursorWidth / 2, ECGadgets\LineHeight / 2, ECGadgets\TargetDim / 2 - 2)
      ClosePath()
      VectorSourceColor(RGBA(120, 120, 120, 255))
      StrokePath(1)
      AddPathCircle(ECGadgets\CursorWidth / 2, ECGadgets\LineHeight / 2, ECGadgets\TargetDim / 2 - 3)
      ClosePath()
      VectorSourceColor(RGBA(255, 255, 255, 255))
      StrokePath(1)
      StopVectorDrawing()
    EndIf
  EndIf
EndProcedure
;
Procedure ZCR_SetFieldsSizeAndMargins()
  ;
  ; Role: Measures the active UI font to size legends, numeric inputs, and line height dynamically.
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  Protected LegendHeight, HexColorWidth, InputWidth, tx$, LegendWidth
  ;
  ECGadgets\TargetDim   = 10
  ECGadgets\CursorWidth = 10
  ;
  If StartDrawing(WindowOutput(ECGadgets\EditColorWindow))
    DrawingFont(GetGadgetFont(#PB_Default))
    LegendHeight  = TextHeight("A")
    HexColorWidth = TextWidth("1234567890ABCDEF") * 6 / 16
    InputWidth    = TextWidth("1234567890") * 3 / 10
    ;
    ; Determine the widest localized label among all active legends
    LegendWidth = TextWidth(GetGadgetText(ECGadgets\Red_Legend))
    tx$ = GetGadgetText(ECGadgets\Green_Legend) : If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Blue_Legend)  : If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Hue_Legend)   : If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Sat_Legend)   : If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Lum_Legend)   : If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    If ECGadgets\HexFieldVPos = -1
      tx$ = GetGadgetText(ECGadgets\HexColor_Legend)
      If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    EndIf
    StopDrawing()
  EndIf
  ;
  ECGadgets\LegendWidth   = LegendWidth / DesktopResolutionX() + 5
  ECGadgets\LineHeight    = LegendHeight + 5
  ECGadgets\InputWidth    = InputWidth / DesktopResolutionX() + 10
  ECGadgets\HexColorWidth = HexColorWidth / DesktopResolutionX() + 10
  CompilerIf #PB_Compiler_OS = #PB_OS_Linux
    ECGadgets\InputWidth + 20
    ECGadgets\HexColorWidth + 20
  CompilerEndIf
  ECGadgets\ButtonsWidth = 80
EndProcedure
;
Procedure ZCR_TrackBarEvents(MouseButton, EventType, InputGadget, CanvasGadget, *ECGadgets.EditColorGadgetsStruct)
  ;
  ; Role: Handles left-click dragging and mouse-wheel adjustments on any RGB or HSL trackbar canvas.
  ;
  Protected Value, localX, TrackSpan, DontSet, VMax
  ;
  If EventType = #PB_EventType_MouseWheel Or MouseButton = #PB_EventType_LeftButtonDown
    Select InputGadget
      Case *ECGadgets\Red_Input, *ECGadgets\Green_Input, *ECGadgets\Blue_Input
        VMax    = 255
        DontSet = #DontSetRGBString
      Case *ECGadgets\Hue_Input
        VMax    = 239
        DontSet = #DontSetHSLString
      Case *ECGadgets\Lum_Input, *ECGadgets\Sat_Input
        VMax    = 240
        DontSet = #DontSetHSLString
    EndSelect
    ;
    If EventType = #PB_EventType_MouseWheel
      Value = Val(GetGadgetText(InputGadget)) + GetGadgetAttribute(CanvasGadget, #PB_Canvas_WheelDelta) * 4
      If Value > VMax : Value = VMax : ElseIf Value < 0 : Value = 0 : EndIf
    EndIf
    If MouseButton = #PB_EventType_LeftButtonDown
      TrackSpan = DesktopScaledX(GadgetWidth(CanvasGadget)) - *ECGadgets\CursorWidth
      If TrackSpan < 1 : TrackSpan = 1 : EndIf
      localX = GetGadgetAttribute(CanvasGadget, #PB_Canvas_MouseX) - (*ECGadgets\CursorWidth / 2)
      If localX > TrackSpan
        localX = TrackSpan
      ElseIf localX < 0
        localX = 0
      EndIf
      Value = (localX * VMax + TrackSpan / 2) / TrackSpan
    EndIf
    If Val(GetGadgetText(InputGadget)) <> Value
      SetGadgetText(InputGadget, Str(Value))
      If DontSet = #DontSetHSLString
        *ECGadgets\Color = HSLToRGB(Val(GetGadgetText(*ECGadgets\Hue_Input)), Val(GetGadgetText(*ECGadgets\Sat_Input)), Val(GetGadgetText(*ECGadgets\Lum_Input)))
      Else
        *ECGadgets\Color = RGB(Val(GetGadgetText(*ECGadgets\Red_Input)), Val(GetGadgetText(*ECGadgets\Green_Input)), Val(GetGadgetText(*ECGadgets\Blue_Input)))
      EndIf
      ZCR_SetGadgetsFromColor(*ECGadgets, DontSet)
    EndIf
  EndIf
EndProcedure
;
Procedure ZCR_Events(Event, EventGadget, EventType)
  ;
  ; Role: Dispatches gadget interactions for the color requester.
  ; Output: Returns 1 on OK, -1 on Cancel, or 0 while editing continues.
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  Static MouseButton, CanvasXDim, CanvasYDim
  Protected Red, Green, Blue, Saturation, Hue, Lightness
  Protected RawStr$, HStr$, ct, CChar$, localX, localY, GetOut
  ;
  If EventType = #PB_EventType_LeftButtonDown Or EventType = #PB_EventType_LeftButtonUp
    MouseButton = EventType
  ElseIf MouseButton = #PB_EventType_LeftButtonDown
    CompilerIf #PB_Compiler_OS = #PB_OS_Windows
      ; Guard against missed #PB_EventType_LeftButtonUp when the cursor is released outside the window
      If (GetAsyncKeyState_(#VK_LBUTTON) & $8000) = 0
        MouseButton = #PB_EventType_LeftButtonUp
      EndIf
    CompilerEndIf
  EndIf
  ;
  If Event = #PB_Event_Gadget
    Select EventGadget
      Case ECGadgets\HexColor_Input
        If EventType = #PB_EventType_Change
          RawStr$ = GetGadgetText(ECGadgets\HexColor_Input)
          HStr$   = UCase(RawStr$)
          For ct = 1 To Len(HStr$)
            CChar$ = Mid(HStr$, ct, 1)
            If FindString("0123456789ABCDEF", CChar$) = 0
              HStr$ = ReplaceString(HStr$, CChar$, "")
              ct - 1
            EndIf
          Next
          HStr$ = Left(HStr$, 6)
          If HStr$ <> RawStr$
            SetGadgetText(ECGadgets\HexColor_Input, HStr$)
            CompilerIf #PB_Compiler_OS = #PB_OS_Windows
              SendMessage_(GadgetID(ECGadgets\HexColor_Input), #EM_SETSEL, Len(HStr$), Len(HStr$))
            CompilerEndIf
          EndIf
          If HStr$
            ECGadgets\Color = Val("$" + HStr$) & $FFFFFF
            ZCR_SetGadgetsFromColor(ECGadgets, #DontSetHexField)
          EndIf
        EndIf
        ;
      Case ECGadgets\Red_Trackbar
        ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Red_Input, ECGadgets\Red_Trackbar, ECGadgets)
      Case ECGadgets\Green_Trackbar
        ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Green_Input, ECGadgets\Green_Trackbar, ECGadgets)
      Case ECGadgets\Blue_Trackbar
        ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Blue_Input, ECGadgets\Blue_Trackbar, ECGadgets)
      Case ECGadgets\Hue_Trackbar
        ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Hue_Input, ECGadgets\Hue_Trackbar, ECGadgets)
      Case ECGadgets\Lum_Trackbar
        ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Lum_Input, ECGadgets\Lum_Trackbar, ECGadgets)
      Case ECGadgets\Sat_Trackbar
        ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Sat_Input, ECGadgets\Sat_Trackbar, ECGadgets)
        ;
      Case ECGadgets\RainbowCanvas
        If EventType = #PB_EventType_MouseWheel Or MouseButton = #PB_EventType_LeftButtonDown
          Saturation = LimitGadgetValue(ECGadgets\Sat_Input, 0, 240)
          If EventType = #PB_EventType_MouseWheel
            Saturation + GetGadgetAttribute(ECGadgets\RainbowCanvas, #PB_Canvas_WheelDelta) * 10
            If Saturation > 240 : Saturation = 240 : ElseIf Saturation < 0 : Saturation = 0 : EndIf
          EndIf
          If MouseButton = #PB_EventType_LeftButtonDown
            localX = GetGadgetAttribute(ECGadgets\RainbowCanvas, #PB_Canvas_MouseX)
            localY = GetGadgetAttribute(ECGadgets\RainbowCanvas, #PB_Canvas_MouseY)
            If localX > DesktopScaledX(ECGadgets\CanvasDim)
              localX = DesktopScaledX(ECGadgets\CanvasDim)
            ElseIf localX < 0
              localX = 0
            EndIf
            If localY > DesktopScaledY(ECGadgets\CanvasDim)
              localY = DesktopScaledY(ECGadgets\CanvasDim)
            ElseIf localY < 0
              localY = 0
            EndIf
            CanvasXDim = ECGadgets\CanvasDim * DesktopResolutionX()
            CanvasYDim = ECGadgets\CanvasDim * DesktopResolutionY()
            Hue        = localX * 239 / CanvasXDim
            Lightness  = (CanvasYDim - localY) * 240 / CanvasYDim
            If Saturation = 0 : Saturation = 1 : EndIf
          Else
            Hue       = LimitGadgetValue(ECGadgets\Hue_Input, 0, 240)
            Lightness = LimitGadgetValue(ECGadgets\Lum_Input, 0, 240)
          EndIf
          ;
          ECGadgets\Color = HSLToRGB(Hue, Saturation, Lightness)
          SetGadgetText(ECGadgets\Sat_Input, Str(Saturation))
          SetGadgetText(ECGadgets\Hue_Input, Str(Hue))
          SetGadgetText(ECGadgets\Lum_Input, Str(Lightness))
          ZCR_SetGadgetsFromColor(ECGadgets, #DontSetHSLString)
        EndIf
        ;
      Case ECGadgets\Red_Input, ECGadgets\Green_Input, ECGadgets\Blue_Input
        If EventType = #PB_EventType_Change
          Red   = LimitGadgetValue(ECGadgets\Red_Input, 0, 255)
          Green = LimitGadgetValue(ECGadgets\Green_Input, 0, 255)
          Blue  = LimitGadgetValue(ECGadgets\Blue_Input, 0, 255)
          ECGadgets\Color = RGB(Red, Green, Blue)
          ZCR_SetGadgetsFromColor(ECGadgets, #DontSetRGBString)
        EndIf
        ;
      Case ECGadgets\Hue_Input, ECGadgets\Sat_Input, ECGadgets\Lum_Input
        If EventType = #PB_EventType_Change
          Hue        = LimitGadgetValue(ECGadgets\Hue_Input, 0, 240)
          Saturation = LimitGadgetValue(ECGadgets\Sat_Input, 0, 240)
          Lightness  = LimitGadgetValue(ECGadgets\Lum_Input, 0, 240)
          ECGadgets\Color = HSLToRGB(Hue, Saturation, Lightness)
          ZCR_SetGadgetsFromColor(ECGadgets, #DontSetHSLString)
        EndIf
        ;
      Case ECGadgets\BOk
        GetOut = 1
      Case ECGadgets\BCancel
        GetOut = -1
    EndSelect
  EndIf
  ProcedureReturn GetOut
EndProcedure
;
CompilerIf Not Defined(ZCR_Escape_Cmd, #PB_Constant)
  #ZCR_Escape_Cmd = 1
  #ZCR_Enter      = 2
CompilerEndIf
;
Procedure ZapmanColorRequester(InitialColor, ParentID = #PB_Default, XShiftOrPos = 0, YShiftOrPos = 0, ParentAnchor = #CWO_Center, WindowAnchor = #CWO_Center)
  ;
  ; Role: Opens a modal, theme-aware RGB/HSL color picker dialog initialized with 'InitialColor'.
  ; Parameters: ParentID     - Valid PB Window ID, OS HWND, or a CWO_Positioning mode:
  ;                            • #CWO_ActiveWindowPos (#PB_Default): relative to GetActiveWindow() (or main monitor).
  ;                            • #CWO_AbsolutePos                  : places dialog at (XShiftOrPos, YShiftOrPos).
  ;                            • #CWO_MonitorPos                   : relative to GetActiveWindow()'s monitor.
  ;             ParentAnchor - Anchor on the parent window or monitor (#CWO_Center, #CWO_TopLeft, etc.).
  ;             WindowAnchor - Anchor on the dialog window itself (#CWO_Center, #CWO_TopLeft, etc.).
  ; Output: Returns the selected 24-bit RGB color on OK, or 'InitialColor' if canceled.
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  Protected WWidth, WHeight, GetOut, Event, EventType, EventMenu
  Protected WParam, OX, OY, MarginCircular, *GadgetAddress, EventGadget
  Protected TextColor, BackGroundColor, ParentHasBeenDisabled
  Protected ParentWindow = #PB_Default
  ;
  If ParentID >= 0 And IsWindow(ParentID)
    ParentWindow = ParentID
    ParentID     = WindowID(ParentID)
  ElseIf ParentID > 0
    ParentWindow = WindowNumFromHandle(ParentID)
  ElseIf ParentID = #PB_Default And GetActiveWindow() >= 0 And IsWindow(GetActiveWindow())
    ParentWindow = GetActiveWindow()
    ParentID     = WindowID(ParentWindow)
  EndIf
  If ParentWindow = #PB_Default
    ParentWindow = ParentID
  EndIf
  ;
  ECGadgets\Color          = InitialColor
  ECGadgets\SGC_ColorType$ = ""
  ;
  If ListSize(InterfaceColorPresets()) > 0 And ListIndex(InterfaceColorPresets()) <> -1
    TextColor       = InterfaceColorPresets()\TextColor
    BackGroundColor = InterfaceColorPresets()\BackgroundColor
  Else
    TextColor       = #PB_Default
    BackGroundColor = #PB_Default
  EndIf
  ;
  WWidth  = 420
  WHeight = 280
  WParam  = #PB_Window_SystemMenu | #PB_Window_SizeGadget | #PB_Window_Invisible
  Protected ActiveWindow = ParentWindow
  If ActiveWindow < 0 Or Not IsWindow(ActiveWindow)
    ActiveWindow = GetActiveWindow()
  EndIf
  ;
  Protected ParentWindowID = ComputeWinOrigins(@OX, @OY, WWidth, WHeight, ParentWindow, XShiftOrPos, YShiftOrPos, ParentAnchor, WindowAnchor)
  ECGadgets\EditColorWindow = OpenWindow(#PB_Any, OX, OY, WWidth, WHeight, GetTextFromCatalog("EditColor"), WParam, ParentWindowID)
  If ECGadgets\EditColorWindow
    If ActiveWindow >= 0 And IsWindow(ActiveWindow)
      DisableWindow(ActiveWindow, #True)
      ParentHasBeenDisabled = #True
    EndIf
    CompilerIf Defined(SetGadgetColorEx, #PB_Procedure)
      ApplyDarkModeToWindow(ECGadgets\EditColorWindow)
    CompilerEndIf
    SetWindowColor(ECGadgets\EditColorWindow, BackGroundColor)
    StickyWindow(ECGadgets\EditColorWindow, #True)
    ;
    WindowBounds(ECGadgets\EditColorWindow, 340, 220, 600, 340)
    BindEvent(#PB_Event_SizeWindow, @ZCR_ResizeGadgets(), ECGadgets\EditColorWindow)
    ;
    ; Margin settings allow embedding this gadget layout inside larger composite windows
    ; (see 'ApplyColorThemes.pbi' at https://www.editions-humanis.com/downloads/PureBasic/ZapmanDowloads_EN.htm).
    ; Here, a uniform margin is applied to all four sides of the dialog.
    MarginCircular               = 15
    ECGadgets\MarginBottom       = 0
    ECGadgets\MarginButtonsBottom = MarginCircular
    ECGadgets\MarginButtonsTop   = 0
    ECGadgets\MarginLeft         = MarginCircular
    ECGadgets\MarginRight        = MarginCircular
    ECGadgets\MarginTop          = MarginCircular
    ECGadgets\HexFieldVPos       = #PB_Default
    ECGadgets\HorizGadgetsMargin = 5
    ECGadgets\MarginInterBlocks  = 0
    ;
    ECGadgets\BackColor = GetRealColorFromType("BackgroundColor", BackGroundColor)
    ECGadgets\TextColor = GetRealColorFromType("TextColor", TextColor)
    ;
    ZCR_CreateGadgets()
    ZCR_SetText()
    ZCR_SetFieldsSizeAndMargins()
    ZCR_CreateTargetAndCursors()
    ZCR_ResizeGadgets()
    ZCR_SetGadgetsFromColor(ECGadgets)
    ;
    AddKeyboardShortcut(ECGadgets\EditColorWindow, #PB_Shortcut_Escape, #ZCR_Escape_Cmd)
    AddKeyboardShortcut(ECGadgets\EditColorWindow, #PB_Shortcut_Return, #ZCR_Enter)
    ;
    CompilerIf Defined(SetGadgetColorEx, #PB_Procedure)
      ; Include 'SetGadgetColorEx.pbi' before this file to enable full gadget color theming
      If BackGroundColor <> #PB_Default Or TextColor <> #PB_Default
        *GadgetAddress = @ECGadgets\Red_Legend
        Repeat
          If TextColor <> #PB_Default
            SetGadgetColorEx(PeekI(*GadgetAddress), #PB_Gadget_FrontColor, TextColor, 1)
          EndIf
          If BackGroundColor <> #PB_Default
            SetGadgetColorEx(PeekI(*GadgetAddress), #PB_Gadget_BackColor, BackGroundColor, 1)
          EndIf
          *GadgetAddress + SizeOf(Integer)
        Until *GadgetAddress > @ECGadgets\BCancel
      EndIf
    CompilerEndIf
    ;
    HideWindow(ECGadgets\EditColorWindow, #False)
    ;
    Repeat
      Event       = WaitWindowEvent()
      EventGadget = EventGadget()
      EventType   = EventType()
      EventMenu   = EventMenu()
      ;
      If Event = #PB_Event_CloseWindow
        GetOut = -1
      ElseIf Event = #PB_Event_Menu
        If EventMenu = #ZCR_Escape_Cmd
          GetOut = -1
        ElseIf EventMenu = #ZCR_Enter
          GetOut = 1
        EndIf
      Else
        GetOut = ZCR_Events(Event, EventGadget, EventType)
      EndIf
    Until GetOut
    ;
    If GetOut = -1
      ECGadgets\Color = InitialColor
    EndIf
    ;
    UnbindEvent(#PB_Event_SizeWindow, @ZCR_ResizeGadgets(), ECGadgets\EditColorWindow)
    CloseWindow(ECGadgets\EditColorWindow)
  EndIf
  ;
  If ParentHasBeenDisabled And IsWindow(ActiveWindow)
    DisableWindow(ActiveWindow, #False)
  EndIf
  ProcedureReturn ECGadgets\Color
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  ; Demo execution when run standalone
  Debug Hex(ZapmanColorRequester($EA5D8))
CompilerEndIf
;
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 96
; FirstLine = 35
; Folding = AEAAAAgQA9
; EnableXP
; DPIAware