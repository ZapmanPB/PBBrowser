;
; ********************************************************************
;
;                       Zapman Color Requester
;                           CrossPlatform
;                          March 2025 - 4
;
; This file should be saved under the name "ZapmanColorRequester.pbi".
;
;    This library is intended to implemente an intuitive and
;    professionnal color requester into your applications.
;    Its main function is 'ZapmanColorRequester()'.
;    Run the demo to see what it does.
;
;    The two blocks of three gadgets allowing to set a color
;    are arranged in the window respecting top, bottom, left
;    right and interblocks margins, in order to be able to be 
;    installed inside other windows containing additional gadgets.
;    An example of reuse of these gadgets is proposed in the
;    procedure 'EditThemesColors()' of file SetGadgetColorEx.pbi
;
;
; This library is compatible with the Zapman 'ApplyColorThemes.pb' library
; that allows you to display your windows in 'Dark Mode' or with other color themes.
;
; ********************************************************************
;
;- 1. Language settings
;
; All Zapman libraries are now designed to be integrated into multilingual applications.
; It is entirely possible to redefine the 'GetTextFromCatalog()' function in a file
; that precedes this one in your 'IncludedFile' list. This way, you can store the
; different translations of the texts used in your application's gadgets in a
; different manner.
; However, the functions in this section provide a simple (and portable) way to do it,
; and if your language is not included in the list of translated languages, only a few
; lines need to be added to address this omission (look at the end of this section).
;
; The translations were provided by ChatGPT from the English version.
;
Global MyLanguage$
If MyLanguage$ = ""
  MyLanguage$ = "EN" ; You can set MyLanguage$ to "EN", "RU", "DE", "IT", "ES", "ZH" or "FR".
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
    Protected PosInList = 0, KeyWord$, Found
    ;
    Repeat
      PosInList + 1
      KeyWord$ = StringField(LanguageEntry$, PosInList, ",")
      Found = 0
      ForEach LanguageList()
        If LanguageList()\Language$ = Language$ And LCase(LanguageList()\LanguageEntry$) = LCase(KeyWord$)
          Found = 1
          Break
        EndIf
      Next
      If Found = 0
        AddElement(LanguageList())
        LanguageList()\Language$ = Language$
        LanguageList()\LanguageEntry$ = KeyWord$
        LanguageList()\LanguageTranslation$ = StringField(LanguageTranslation$, PosInList, ",")
      EndIf
    Until KeyWord$ = ""
  EndProcedure
CompilerEndIf
;
Define LanguageEntry$ = "ThemeColors,Preset,BackgroundColor,TextColor,ResetToDefaultColor,"
       LanguageEntry$ + "EditColor,Red,Green,Blue,Hue,Saturation,Lightness,HexValue,Cancel,OK,"
       LanguageEntry$ + "ClickToAdjustColor"
;
; ------------- English list of expressions ---------------
Define LanguageTranslation$ = "Theme colors,Preset:,Background Color,Text Color,Reset to original color,"
       LanguageTranslation$ + "Edit Color,Red:,Green:,Blue:,Hue:,Saturation:,Lightness:,Hex value:,Cancel,OK,"
       LanguageTranslation$ + "Click to choose a color. Use the mouse wheel to adjust saturation."
FillLanguageList("EN", LanguageEntry$, LanguageTranslation$)

; ------------- Spanish list of expressions ---------------
       LanguageTranslation$ = "Colores del tema,Predefinido:,Color de fondo,Color de texto,Restablecer al color original,"
       LanguageTranslation$ + "Editar color,Rojo:,Verde:,Azul:,Matiz:,Saturación:,Luminosidad:,Valor hexadecimal:,Cancelar,OK,"
       LanguageTranslation$ + "Haz clic para elegir un color. Usa la rueda del ratón para ajustar la saturación."
FillLanguageList("ES", LanguageEntry$, LanguageTranslation$)

; ------------- Mandarin list of expressions ---------------
       LanguageTranslation$ = "主题颜色,预设:,背景颜色,文字颜色,重置为原始颜色,"
       LanguageTranslation$ + "编辑颜色,红:,绿:,蓝:,色相:,饱和度:,亮度:,十六进制值:,取消,确定,"
       LanguageTranslation$ + "点击选择颜色。使用鼠标滚轮调整饱和度。"
FillLanguageList("ZH", LanguageEntry$, LanguageTranslation$)

; ------------- French list of expressions ---------------
       LanguageTranslation$ = "Couleurs du thème,Préréglages :,Couleur de fond,Couleur du texte,Couleur d'origine,"
       LanguageTranslation$ + "Editeur de couleur,Rouge :,Vert :,Bleu :,Teinte :,Saturation :,Luminosité :,Valeur hexa :,Annuler,OK,"
       LanguageTranslation$ + "Cliquez pour choisir une couleur. Utilisez la roue de souris pour régler la saturation."
FillLanguageList("FR", LanguageEntry$, LanguageTranslation$)
;
; ------------- German list of expressions ---------------
       LanguageTranslation$ = "Themenfarben,Voreinstellungen:,Hintergrundfarbe,Textfarbe,Ursprungsfarbe,"
       LanguageTranslation$ + "Farbeditor,Rot:,Grün:,Blau :,Farbton:,Sättigung:,Helligkeit:,Hex-Wert:,Abbrechen,OK,"
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
Define LanguageTranslation$ = "Colori del tema,Predefinito:,Colore di sfondo,Colore del testo,Reimposta al colore originale,"
       LanguageTranslation$ + "Modifica colore,Rosso:,Verde:,Blu:,Tonalità:,Saturazione:,Luminosità:,Valore esadecimale:,Annulla,OK,"
       LanguageTranslation$ + "Clicca per scegliere un colore. Usa la rotella del mouse per regolare la saturazione."
FillLanguageList("IT", LanguageEntry$, LanguageTranslation$)
;
; ------------- Add your own language here --------------
;      LanguageTranslation$ = "Expression1,Expression2,etc."
;FillLanguageList("XX", LanguageEntry$, LanguageTranslation$) ; and replace "XX" by the abbreviation of your language name.
;
; ------------------- END OF LANGUAGE SETTING -------------------
;
Enumeration CWO_Positioning
  #CWO_ActiveWindowPos = -1
  #CWO_AbsolutePos     = -2
  #CWO_MonitorPos      = -3
EndEnumeration
;
Enumeration CWO_PositionAnchor
  #CWO_Center = 0
  #CWO_TopLeft
  #CWO_TopRight
  #CWO_BottomLeft
  #CWO_BottomRight
EndEnumeration
;
CompilerIf Not Defined(ComputeWinOrigins, #PB_Procedure)
  Procedure ComputeWinOrigins(*OX.Integer, *OY.Integer, WWidth, WHeight, ParentWindow = #CWO_ActiveWindowPos, XShiftOrPos = 0, YShiftOrPos = 0, ParentAnchor = #CWO_Center, WindowAnchor = #CWO_Center)
    ;
    ; Compute the X and Y origins of a window from its width (WWidth) and height (WHeight)
    ; and from the parent window position or from the monitor of the parent window.
    ;
    ; • If ParentWindow is a valid window number, the new window will be positionned relatively to it,
    ;   and then shifted by XShiftOrPos and YShiftOrPos.
    ; • If ParentWindow = #CWO_ActiveWindowPos and GetActiveWindow() returns a valid window number,
    ;   GetActiveWindow() will be considered as the parent window.
    ; • If ParentWindow = #CWO_ActiveWindowPos and GetActiveWindow() does NOT return a valid window number,
    ;   the position will be calculated relatively to the main screen.
    ; • It ParentWindow = #CWO_AbsolutePos, XShiftOrPos and YShiftOrPos will be the absolute coordinates
    ;   of the new window.
    ; • If ParentWindow = #CWO_MonitorPos, the position will be calculated relatively to the screen where
    ;   GetActiveWindow() is found or relatively to the main screen if GetActiveWindow()
    ;   is not a valid window.
    ;
    ; • The new window will be positionned relatively to the center, top-left, top-right, bottom-left
    ;   or bottom-right of the parent window (or monitor), depending on the 'ParentAnchor' parameter.
    ;
    ; • The anchor of the new window can be the center, top-left, top-right, bottom-left or bottom-right
    ;   of the new window, depending on the 'WindowAnchor' parameter.
    ;
    ; • In all case, except if ParentWindow = #CWO_AbsolutePos, the position of the new window will be
    ;   shifted regarding the 'XShiftOrPos' and 'YShiftOrPos' parameters.
    ;
    Protected DesktopLeft, DesktopRight, DesktopTop, DesktopBottom
    Protected ParentWindowID, MFWindow, hMonitor, mi.MONITORINFO
    ;
    If ParentWindow = #CWO_AbsolutePos
      *OX\i = XShiftOrPos
      *OY\i = YShiftOrPos
      ProcedureReturn 0
    EndIf
    ;
    If ParentWindow = #CWO_ActiveWindowPos And IsWindow(GetActiveWindow())
      ; If ParentWindow = #CWO_ActiveWindowPos, use GetActiveWindow() as parent:
      ParentWindow = GetActiveWindow()
    EndIf
    ;
    ; Get the screen coordinates for the monitor where the parent window is found:
    MFWindow = ParentWindow
    If Not(IsWindow(MFWindow))
      MFWindow = GetActiveWindow()
    EndIf
    If IsWindow(MFWindow)
      MFWindow = WindowID(MFWindow)
    Else
      ; If there is no active-window, use the main monitor:
      MFWindow = 0
    EndIf
    ;
    CompilerIf #PB_Compiler_OS = #PB_OS_Windows
      ;
      hMonitor = MonitorFromWindow_(MFWindow, #MONITOR_DEFAULTTONEAREST)
      mi\cbSize = SizeOf(MONITORINFO)
      GetMonitorInfo_(hMonitor, @mi)
      ;
      DesktopLeft   = DesktopUnscaledX(mi\rcWork\left)
      DesktopRight  = DesktopUnscaledX(mi\rcWork\Right)
      DesktopTop    = DesktopUnscaledY(mi\rcWork\top)
      DesktopBottom = DesktopUnscaledY(mi\rcWork\Bottom)
      ;
    CompilerElse
      ;
      ExamineDesktops()
      DesktopLeft   = 0
      DesktopRight  = DesktopWidth(0)
      DesktopTop    = 0
      DesktopBottom = DesktopHeight(0)
    CompilerEndIf
    ;
    ;____________________________________________
    ;
    ; Calculate the starting point coordinates:
    If IsWindow(ParentWindow)
      ParentWindowID = WindowID(ParentWindow)
      ;
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
      ;
      *OX\i = DeskTopLeft
      *OY\i = DesktopTop
      ;
      If ParentAnchor = #CWO_Center
        *OX\i / 2 + DesktopRight / 2
        *OY\i / 2 + DesktopBottom / 2
      ElseIf ParentAnchor = #CWO_TopRight
        *OX\i = DesktopRight
      ElseIf ParentAnchor = #CWO_BottomLeft
        *OY\i = DesktopBottom
      ElseIf ParentAnchor = #CWO_BottomRight
        *OX\i = DesktopRight
        *OY\i = DesktopBottom
      EndIf
    EndIf
    ;____________________________________________
    ;
    ; Adjust the final coordinates regarding the new window size:
    ;
    WHeight + MenuHeight() + 4 ; <-- Calculate the real total height of the new window
                               ;     that will be created by "OpenWindow(#PB_Any, OX, OY, WWidth, WHeight..."
    If WindowAnchor = #CWO_Center
      *OX\i - WWidth / 2
      *OY\i - WHeight / 2
    ElseIf WindowAnchor = #CWO_TopRight
      *OX\i - WWidth
    ElseIf WindowAnchor = #CWO_BottomLeft
      *OY\i - WHeight
    ElseIf WindowAnchor = #CWO_BottomRight
      *OX\i - WWidth
      *OY\i - WHeight
    EndIf
    ;
    ;____________________________________________
    ;
    ; Shift the final position:
    *OX\i + XShiftOrPos
    *OY\i + YShiftOrPos
    ;
    ; Test if window extends beyond screen bounds due to ParentWindow position or XShiftOrPos or YShiftOrPos:
    If *OX\i < DesktopLeft
      *OX\i = DesktopLeft + 10
    ElseIf *OX\i + WWidth > DesktopRight
      *OX\i =  DesktopRight - WWidth - 10
    EndIf
    If *OY\i < DesktopTop
      *OY\i = DesktopTop + 10
    ElseIf *OY\i + WHeight > DesktopBottom
      *OY\i =  DesktopBottom - WHeight - 10
    EndIf
    ;
    ProcedureReturn ParentWindowID
  EndProcedure
CompilerEndIf
;
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
  Color.i ; This value will be updated on the fly.
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
  ; Gadgets numbers -> Square Canvases
  PreviewCanvas.i
  RainbowCanvas.i
  ; Gadgets numbers -> Target and cursors
  TargetGadget.i
  Red_Cursor.i
  Green_Cursor.i
  Blue_Cursor.i
  Hue_Cursor.i
  Sat_Cursor.i
  Lum_Cursor.i
  ; Gadgets numbers -> TrackBars
  Red_Trackbar.i
  Green_Trackbar.i
  Blue_Trackbar.i
  Hue_Trackbar.i
  Sat_Trackbar.i
  Lum_Trackbar.i
  ; Gadgets numbers -> Legends
  Red_Legend.i
  Green_Legend.i
  Blue_Legend.i
  Hue_Legend.i
  Sat_Legend.i
  Lum_Legend.i
  HexColor_Legend.i
  ; Gadgets numbers -> Input fields
  Red_Input.i
  Green_Input.i
  Blue_Input.i
  Hue_Input.i
  Sat_Input.i
  Lum_Input.i
  HexColor_Input.i
  ; Gadgets numbers -> Buttons
  BOk.i
  BCancel.i
  ; Gadgets numbers -> Separators
  SeparatorGadget1.i
  SeparatorGadget2.i
  SeparatorGadget3.i
  ;
  ; Parameter for SetGadgetColorEx()
  SGC_ColorType$
  ;
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
  ; Procedure to find the minimum value among the given arguments
  ;
  Protected result = value1
  ;
  If value2 < result
    result = value2
  EndIf
  If value3 < result
    result = value3
  EndIf
  ProcedureReturn result
EndProcedure
;
Procedure MaxOf3(value1, value2, value3)
  ;
  ; Procedure to find the maximum value among the given arguments
  ;
  Protected result = value1
  ;
  If value2 > result
    result = value2
  EndIf
  If value3 > result
    result = value3
  EndIf
  ProcedureReturn result
EndProcedure
;
Procedure Hue(rgbColor)
  ; Extract the red, green, and blue components  
  Protected red = Red(rgbColor)
  Protected green = Green(rgbColor)
  Protected blue = Blue(rgbColor)
  ;
  ; Find minimum and maximum values among the components
  Protected maxValue = MaxOf3(red, green, blue)
  Protected minValue = MinOf3(red, green, blue)
  ;
  ; Calculate the range (delta)
  Protected delta = maxValue - minValue
  ;
  ; Initialize hue
  Protected hue.f = 0
  ;
  If delta
    ; Calculate hue based on which component is the maximum
    If maxValue = red
      hue = (green - blue) / delta
    ElseIf maxValue = green
      hue = 2.0 + (blue - red) / delta
    Else
      hue = 4.0 + (red - green) / delta
    EndIf
    ;
    ; Scale hue to 0-240 range
    hue * 40
    If hue < 0
      hue + 240
    EndIf
  EndIf
  ;
  ProcedureReturn hue
EndProcedure
;
Procedure Saturation(rgbColor)
  ; Extract the red, green, and blue components  
  Protected red = Red(rgbColor)
  Protected green = Green(rgbColor)
  Protected blue = Blue(rgbColor)
  ;
  Protected saturation.f
  ;
  ; Find minimum and maximum values among the components
  Protected maxValue = MaxOf3(red, green, blue)
  Protected minValue = MinOf3(red, green, blue)
  ;
  ; Calculate the range (delta)
  Protected delta = maxValue - minValue
  ; Calculate lightness (average of max and min)
  Protected lightness.f = (maxValue + minValue) / (2 * 255)
  ;
  If delta
    If lightness < 0.5
      saturation = delta / (maxValue + minValue)
    Else
      saturation = delta / (510 - maxValue - minValue)
    EndIf
  EndIf
  ;
  ProcedureReturn saturation * 240
EndProcedure
;
Procedure Lightness(rgbColor)
  ; Extract the red, green, and blue components  
  Protected red = Red(rgbColor)
  Protected green = Green(rgbColor)
  Protected blue = Blue(rgbColor)
  ;
  ; Find minimum and maximum values among the components
  Protected maxValue = MaxOf3(red, green, blue)
  Protected minValue = MinOf3(red, green, blue)
  ;
  ; Calculate lightness (average of max and min)
  ProcedureReturn 240 * (maxValue + minValue) / (2 * 255)
EndProcedure
;
Procedure.i HSLToRGB(hue.i, saturation.i, lightness.i)
  ; Hue: 0 - 240
  ; Saturation: 0 - 240
  ; Lightness: 0 - 240
  Protected MaxHue = 240
  Protected MaxSL = 240
  ;
  Protected r.f, g.f, b.f
  Protected c.f, x.f, m.f
  ;
  ; Normalize SL values to [0, 1] range
  Protected normalizedSaturation.f = saturation / MaxSL
  Protected normalizedLightness.f = Round(lightness * 100 / MaxSL, #PB_Round_Down) / 100
  ;
  Protected HueSlice = MaxHue / 6
  ;
  ; Calculate intermediary values
  c = (1 - Abs(2 * normalizedLightness - 1)) * normalizedSaturation
  Protected vi = Int(Hue * 1000 / HueSlice) % 2000
  x = c * (1 - Abs(vi / 1000 - 1))
  m = normalizedLightness - c / 2
  ;
  ; Determine RGB components based on the hue sector
  Select Int((Hue - 1) / HueSlice)
    Case 0 : r = c : g = x : b = 0
    Case 1 : r = x : g = c : b = 0
    Case 2 : r = 0 : g = c : b = x
    Case 3 : r = 0 : g = x : b = c
    Case 4 : r = x : g = 0 : b = c
    Case 5 : r = c : g = 0 : b = x
  EndSelect
  ;
  ; Adjust RGB components to final range [0, 255]
  r = (r + m) * 255
  g = (g + m) * 255
  b = (b + m) * 255
  ;
  ; Return the RGB color
  ProcedureReturn RGB(r, g, b)
EndProcedure
;
Procedure LimitGadgetValue(Gadget, Min, Max)
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
  If CUnit < 127
    CUnit + 80
  Else
    CUnit - 100
  EndIf
  If CUnit > 200
    CUnit = 200
  EndIf
  If CUnit < 60
    CUnit = 60
  EndIf
  ProcedureReturn Cunit
EndProcedure
;
Procedure InvertedColor(Color)
  ProcedureReturn HSLToRGB(Abs(120 - hue(Color)), 240, ColorShiftLum(lightness(Color)))
EndProcedure
;
Macro DrawTrackBarBox()
  ; Repaint box background:
  Box(0, 0, DesktopScaledX(*ECGadgets\TrackBarsWidth), DesktopScaledY(*ECGadgets\LineHeight), *ECGadgets\BackColor)
  ; Draw trackbar box
  Box(GStart - 1, 0, GWidth + 1, *ECGadgets\LineHeight, *ECGadgets\BorderColor)
  ;
  ; Init graduation counter
  tl = stl
EndMacro
;
Macro DrawOneTrackBarLine(Color)
  x = ct + GStart - 1
  ; Draw color vertical line:
  Line(x, 1, 1, *ECGadgets\LineHeight - 2, Color)
  ; Draw center lines points:
  Plot(x, GHalfHeight - 1, GradColor)
  Plot(x, GHalfHeight + 1, GradColor)
  ; Draw graduations:
  If ct = Int(tl)
    tl + stl
    Line(x, 1, 1, 3, GradColor)
  EndIf
EndMacro
;
CompilerIf Not(Defined(MixColors, #PB_Procedure))
  Procedure MixColors(Color1, Color2, Ratio.f)
    If Ratio > 1 : Ratio = 1 : EndIf
    Protected Red = Red(Color1) * Ratio + Red(Color2) * (1 - Ratio)
    Protected Green = Green(Color1) * Ratio + Green(Color2) * (1 - Ratio)
    Protected Blue = Blue(Color1) * Ratio + Blue(Color2) * (1 - Ratio)
    ProcedureReturn RGB(Red, Green, Blue)
  EndProcedure
CompilerEndIf
;
Procedure RepaintCanvasSeparator(CanvasGadget, Color)
  If IsGadget(CanvasGadget) And GadgetType(CanvasGadget) = #PB_GadgetType_Canvas And StartDrawing(CanvasOutput(CanvasGadget))
    Box(0, 0, DesktopScaledX(GadgetWidth(CanvasGadget)), DesktopScaledY(GadgetHeight(CanvasGadget)), Color)
    StopDrawing()
  EndIf
EndProcedure
;
Procedure GetRealColorFromType(ColorType$, Color)
  ;
  If Color = #PB_Default
    If ColorType$ = "BackgroundColor"
      Color = GetSysColor_(#COLOR_BTNFACE)
    ElseIf ColorType$ = "TextColor"
      Color = GetSysColor_(#COLOR_WINDOWTEXT)
    EndIf
  EndIf
  ProcedureReturn Color
EndProcedure
;
Procedure CalculateBorderColor(TextColor, BackgroundColor)
  ProcedureReturn MixColors(GetRealColorFromType("TextColor", TextColor), GetRealColorFromType("BackgroundColor", BackgroundColor), 0.4)
EndProcedure
;
Procedure ZCR_CalculateBorderColor(*ECGadgets.EditColorGadgetsStruct, RepaintCase$)
  ;
  If RepaintCase$
    If RepaintCase$ = "BackgroundColor"
      ; This particular case is used by the SegGadgetColorEx.pb library
      ; where the gadget's background or text colors need to be set to the actual color.
      *ECGadgets\BackColor = *ECGadgets\Color
    ElseIf RepaintCase$ = "TextColor"
      ; Also.
      *ECGadgets\TextColor = *ECGadgets\Color
    EndIf
    ;
    *ECGadgets\BorderColor = CalculateBorderColor(*ECGadgets\TextColor, *ECGadgets\BackColor)
    ;
    RepaintCanvasSeparator(*ECGadgets\SeparatorGadget1, *ECGadgets\BorderColor)
    RepaintCanvasSeparator(*ECGadgets\SeparatorGadget2, *ECGadgets\BorderColor)
    RepaintCanvasSeparator(*ECGadgets\SeparatorGadget3, *ECGadgets\BorderColor)
  EndIf
  ;
EndProcedure
;
Procedure FillCanvasWithColor(Canvas, Color, BorderColor)
  If StartDrawing(CanvasOutput(Canvas))
    Box(0, 0, DesktopScaledX(GadgetWidth(Canvas)), DesktopScaledY(GadgetHeight(Canvas)), BorderColor)
    Box(1, 1, DesktopScaledX(GadgetWidth(Canvas)) - 2, DesktopScaledY(GadgetHeight(Canvas)) - 2, Color)
    StopDrawing()
  EndIf
EndProcedure
;
Procedure ZCR_DrawTrackbarsAndPreviewCanvas(*ECGadgets.EditColorGadgetsStruct)
  ;
  Protected x, y, GWidth, GStart, GHalfHeight, GHalfWidth
  Protected Increase.f, Decrease.f, DecreasingGrey, IncreasingWhite
  Protected Red, Green, Blue, Nred, Ngreen, Nblue, RefColor, GradColor, GreyColor
  Protected Saturation, Lightness, Hue
  Protected stl.f, tl.f, ct, ct2, ci
  Protected Trackbar
  ;
  ZCR_CalculateBorderColor(*ECGadgets, *ECGadgets\SGC_ColorType$)
  ;
  ; DRAW PREVIEW CANVAS:
  FillCanvasWithColor(*ECGadgets\PreviewCanvas, *ECGadgets\Color, *ECGadgets\BorderColor)
  ;
  ; *********************************************************************
  ;
  ; Initialize some values for trackbars drawing:
  ;
  ; Values for drawing area:
  GWidth  = DesktopScaledX(*ECGadgets\TrackBarsWidth) - *ECGadgets\CursorWidth + 1
  GStart = *ECGadgets\CursorWidth / 2
  ; Value for center horizontal lines
  GHalfHeight = *ECGadgets\LineHeight / 2
  ; Values for graduations:
  stl = (GWidth + 1) / 6
  ;
  ; *********************************************************************
  ;
  ; DRAW RGB_Trackbars:
  ;
  If Lightness(*ECGadgets\Color) < 60
    GradColor = RGB(160, 160, 160)
  ElseIf Lightness(*ECGadgets\Color) > 130
    GradColor = RGB(140, 140, 50)
  Else
    GradColor = RGB(255, 255, 255)
  EndIf
  ;
  For ci = 1 To 3
    Red = Red(*ECGadgets\Color)
    Green = Green(*ECGadgets\Color)
    Blue = Blue(*ECGadgets\Color)
    If ci = 1 : Trackbar = *ECGadgets\Red_Trackbar
    ElseIf ci = 2 : Trackbar = *ECGadgets\Green_Trackbar
    ElseIf ci = 3 : Trackbar = *ECGadgets\Blue_Trackbar
    EndIf
    If StartDrawing(CanvasOutput(Trackbar))
      ;
      DrawTrackBarBox()
      ;
      For ct = 1 To GWidth - 1
        If ci = 1 : Red = ct * 255 / GWidth
        ElseIf ci = 2 : Green = ct * 255 / GWidth
        ElseIf ci = 3 : Blue = ct * 255 / GWidth
        EndIf
        DrawOneTrackBarLine(RGB(Red, Green, Blue))
      Next
      StopDrawing()
    EndIf
  Next
  ;
  ; *********************************************************************
  ;
  ; DRAW Hue_Trackbar:
  ;
  If StartDrawing(CanvasOutput(*ECGadgets\Hue_Trackbar))
    ;
    Saturation = Saturation(*ECGadgets\Color)
    Lightness = Lightness(*ECGadgets\Color)
    ;
    DrawTrackBarBox()

    GradColor = HSLToRGB(120, 120, ColorShiftLum(Lightness))
    ;
    For ct = 1 To GWidth - 1
      Hue = ct * 240 / GWidth
      DrawOneTrackBarLine(HSLToRGB(Hue, Saturation, Lightness))
    Next
    StopDrawing()
  EndIf
  ;
  ; *********************************************************************
  ;
  ; DRAW Lum_Trackbar:
  ;
  If StartDrawing(CanvasOutput(*ECGadgets\Lum_Trackbar))
    ;
    ; Calculate color with a half luminosity:
    RefColor = HSLToRGB(LimitGadgetValue(*ECGadgets\Hue_Input, 0, 240), LimitGadgetValue(*ECGadgets\Sat_Input, 0, 240), 120)
    ;
    DrawTrackBarBox()

    Red = Red(RefColor)
    Green = Green(RefColor)
    Blue = Blue(RefColor)
    GradColor = RGB(127, 127, 127)
    GHalfWidth = GWidth / 2
    ;
    For ct = 1 To GWidth - 1
      If ct < GHalfWidth
        Increase.f = ct / GHalfWidth
        NRed   = Red   * Increase
        NGreen = Green * Increase
        NBlue  = Blue  * Increase
      Else
        ct2 = ct - GHalfWidth
        IncreasingWhite = 255 * ct2 / GHalfWidth
        Decrease.f = (GHalfWidth - ct2) / GHalfWidth
        NRed   = Red   * Decrease + IncreasingWhite
        NGreen = Green * Decrease + IncreasingWhite
        NBlue  = Blue  * Decrease + IncreasingWhite
      EndIf
      ;
      DrawOneTrackBarLine(RGB(Nred, Ngreen, Nblue))
      ;
    Next
    StopDrawing()
  EndIf
  ;
  ; *********************************************************************
  ;
  ; DRAW Sat_Trackbar:
  ;
  If StartDrawing(CanvasOutput(*ECGadgets\Sat_Trackbar))
    ;
    ; Calculate color with a full saturation:
    RefColor = HSLToRGB(LimitGadgetValue(*ECGadgets\Hue_Input, 0, 240), 240, LimitGadgetValue(*ECGadgets\Lum_Input, 0, 240) + 1)
    ;
    DrawTrackBarBox()
    
    Red = Red(RefColor)
    Green = Green(RefColor)
    Blue = Blue(RefColor)
    GradColor = InvertedColor(RefColor)
    GreyColor = LimitGadgetValue(*ECGadgets\Lum_Input, 0, 240) * 255 / 240
    ;
    For ct = 1 To GWidth - 1
      Increase.f = ct / GWidth
      DecreasingGrey = GreyColor * (GWidth - ct) / GWidth
      Nred   = (Red   * Increase) + DecreasingGrey
      Ngreen = (Green * Increase) + DecreasingGrey
      Nblue  = (Blue  * Increase) + DecreasingGrey
      ;
      DrawOneTrackBarLine(RGB(Nred, Ngreen, Nblue))
      ;
    Next
    StopDrawing()
  EndIf
  ;
  x = GadgetX(*ECGadgets\RainbowCanvas) + LimitGadgetValue(*ECGadgets\Hue_Input, 0, 240) * *ECGadgets\CanvasDim / 240 - DesktopUnscaledX(*ECGadgets\TargetDim / 2)
  y = GadgetY(*ECGadgets\RainbowCanvas) + (240 - LimitGadgetValue(*ECGadgets\Lum_Input, 0, 240)) * *ECGadgets\CanvasDim / 240 - DesktopUnscaledY(*ECGadgets\TargetDim / 2)
  ResizeGadget(*ECGadgets\TargetGadget, x, y, #PB_Ignore, #PB_Ignore)
EndProcedure
;
Procedure ZCR_SetHexFieldFromColor(*ECGadgets.EditColorGadgetsStruct)
  Protected HStr$ = Hex(*ECGadgets\Color, #PB_Long)
  While Len(HStr$) < 6 : HStr$ = "0" + HStr$ : Wend
  SetGadgetText(*ECGadgets\HexColor_Input, HStr$)
  CompilerIf #PB_Compiler_OS = #PB_OS_Windows
    SendMessage_(GadgetID(*ECGadgets\HexColor_Input), #EM_SETSEL, Len(HStr$), Len(HStr$))
  CompilerEndIf
EndProcedure
;
Procedure ZCR_DrawRainbowCanvas(*ECGadgets.EditColorGadgetsStruct)
  ;
  Protected ColorWidth, ColorHeight, HalfColorHeight
  Protected cx, x, cy, cy2, y
  Protected DecreasingWhite, RefColor, Decrease.f, Increase.f
  Protected NRed, NGreen, NBlue
  Protected Saturation, Lightness, Hue, Red, Green, Blue
  ;
  Saturation = Val(GetGadgetText(*ECGadgets\Sat_Input))
  Lightness = 120
  ;
  If StartDrawing(CanvasOutput(*ECGadgets\RainbowCanvas))
    Box(0, 0, DesktopScaledX(*ECGadgets\CanvasDim), DesktopScaledY(*ECGadgets\CanvasDim), *ECGadgets\BorderColor)
    ColorWidth = (DesktopScaledX(*ECGadgets\CanvasDim) - 2) * 20 - 20
    ColorHeight = (DesktopScaledY(*ECGadgets\CanvasDim) - 2) * 20
    HalfColorHeight = ColorHeight / 2
    ;
    For cx = 0 To ColorWidth Step 20
      ;
      Hue = cx * 239 / ColorWidth
      RefColor = HSLToRGB(Hue, Saturation, Lightness)
      Red = Red(RefColor)
      Green = Green(RefColor)
      Blue = Blue(RefColor)
      ;
      x = (cx / 20) + 1
      For cy = 0 To ColorHeight Step 20
        y = (cy / 20) + 1
        If cy < HalfColorHeight
          DecreasingWhite = 255 * (HalfColorHeight - cy) / HalfColorHeight
          Increase.f = cy / HalfColorHeight
          NRed   = Red   * Increase + DecreasingWhite
          NGreen = Green * Increase + DecreasingWhite
          NBlue  = Blue  * Increase + DecreasingWhite
        Else
          cy2 = cy - HalfColorHeight
          Decrease.f = (HalfColorHeight - cy2) / HalfColorHeight
          NRed   = Red   * Decrease
          NGreen = Green * Decrease
          NBlue  = Blue  * Decrease
        EndIf
        Plot(x, y, RGB(Nred, Ngreen, Nblue))
      Next
    Next
    StopDrawing()
  EndIf
EndProcedure
;
Procedure ZCR_SetRGBStringsFromColor(*ECGadgets.EditColorGadgetsStruct)
  SetGadgetText(*ECGadgets\Red_Input, Str(Red(*ECGadgets\Color)))
  SetGadgetText(*ECGadgets\Green_Input, Str(Green(*ECGadgets\Color)))
  SetGadgetText(*ECGadgets\Blue_Input, Str(Blue(*ECGadgets\Color)))
EndProcedure
;
Procedure ZCR_SetHSLStringsFromColor(*ECGadgets.EditColorGadgetsStruct)
  SetGadgetText(*ECGadgets\Hue_Input, Str(Hue(*ECGadgets\Color)))
  SetGadgetText(*ECGadgets\Sat_Input, Str(Saturation(*ECGadgets\Color)))
  SetGadgetText(*ECGadgets\Lum_Input, Str(Lightness(*ECGadgets\Color)))
EndProcedure
;
Procedure ZCR_SetCursor(CanvasGadget, Value, *ECGadgets.EditColorGadgetsStruct)
  ;
  Protected CursorGadget, MaxValue, UWidth, x, y
  ;
  If CanvasGadget = *ECGadgets\Red_Trackbar
    CursorGadget = *ECGadgets\Red_Cursor
    MaxValue = 255
  ElseIf CanvasGadget = *ECGadgets\Green_Trackbar
    CursorGadget = *ECGadgets\Green_Cursor
    MaxValue = 255
  ElseIf CanvasGadget = *ECGadgets\Blue_Trackbar
    CursorGadget = *ECGadgets\Blue_Cursor
    MaxValue = 255
  ElseIf CanvasGadget = *ECGadgets\Hue_Trackbar
    CursorGadget = *ECGadgets\Hue_Cursor
    MaxValue = 240
  ElseIf CanvasGadget = *ECGadgets\Sat_Trackbar
    CursorGadget = *ECGadgets\Sat_Cursor
    MaxValue = 240
  ElseIf CanvasGadget = *ECGadgets\Lum_Trackbar
    CursorGadget = *ECGadgets\Lum_Cursor
    MaxValue = 240
  EndIf
  UWidth = GadgetWidth(CanvasGadget) - DesktopUnscaledX(*ECGadgets\CursorWidth)
  x = GadgetX(CanvasGadget) + (Value * UWidth / MaxValue)
  y = GadgetY(CanvasGadget)
  ResizeGadget(CursorGadget, x, y, #PB_Ignore, #PB_Ignore)
EndProcedure
;
Procedure ZCR_SetTrackBarsFromStrings(*ECGadgets.EditColorGadgetsStruct)
  ZCR_SetCursor(*ECGadgets\Red_Trackbar, Val(GetGadgetText(*ECGadgets\Red_Input)), *ECGadgets)
  ZCR_SetCursor(*ECGadgets\Green_Trackbar, Val(GetGadgetText(*ECGadgets\Green_Input)), *ECGadgets)
  ZCR_SetCursor(*ECGadgets\Blue_Trackbar, Val(GetGadgetText(*ECGadgets\Blue_Input)), *ECGadgets)
  ZCR_SetCursor(*ECGadgets\Hue_Trackbar, Val(GetGadgetText(*ECGadgets\Hue_Input)), *ECGadgets)
  ZCR_SetCursor(*ECGadgets\Lum_Trackbar, Val(GetGadgetText(*ECGadgets\Lum_Input)), *ECGadgets)
  ZCR_SetCursor(*ECGadgets\Sat_Trackbar, Val(GetGadgetText(*ECGadgets\Sat_Input)), *ECGadgets)
EndProcedure
;
Procedure ZCR_SetGadgetsFromColor(*ECGadgets.EditColorGadgetsStruct, DontSet = 0)
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
  ZCR_SetTrackBarsFromStrings(*ECGadgets)
EndProcedure
;
Define ECGadgets.EditColorGadgetsStruct
;
Procedure ZCR_ResizeGadgets()
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  ;
  Protected HexFieldVPos, HexFieldHPos
  ;
  Protected TextHeight = ECGadgets\LineHeight / DesktopResolutionY()
  Protected ButtonHeight = TextHeight + 3
  ;
  ; Reserve room for the Cancel & OK buttons:
  ;
  If ECGadgets\MarginBottom = 0
    ECGadgets\MarginBottom = 2 * ECGadgets\MarginButtonsBottom + ButtonHeight + ECGadgets\MarginButtonsTop
  EndIf
  ;
  ; Calculate available area for all gadgets:
  ;
  Protected AreaHeight = WindowHeight(ECGadgets\EditColorWindow) - ECGadgets\MarginTop - ECGadgets\MarginBottom - ECGadgets\MarginInterBlocks
  Protected AreaWidth = WindowWidth(ECGadgets\EditColorWindow) - ECGadgets\MarginLeft - ECGadgets\MarginRight
  ;
  ; Calculate general values for gadgets:
  ;
  Protected HorizLineShift = 3 ; Shift value for horizontal lines.
  Protected InterLine = (AreaHeight - 2 * TextHeight + (2 * HorizLineShift)) / 4.75
  ;
  ECGadgets\CanvasDim = InterLine * 2 + TextHeight
  ECGadgets\TrackBarsWidth = AreaWidth - ECGadgets\LegendWidth - ECGadgets\InputWidth - ECGadgets\CanvasDim - (ECGadgets\HorizGadgetsMargin * 3)
  ;
  Protected VPos = ECGadgets\MarginTop
  ;
  Protected LegendHPos = ECGadgets\MarginLeft
  Protected InputHPos = LegendHPos + ECGadgets\LegendWidth + ECGadgets\HorizGadgetsMargin
  Protected TrackBarHPos = InputHPos + ECGadgets\InputWidth + ECGadgets\HorizGadgetsMargin
  Protected CanvasHPos = TrackBarHPos + ECGadgets\TrackBarsWidth + ECGadgets\HorizGadgetsMargin
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
  ResizeGadget(ECGadgets\SeparatorGadget1, ECGadgets\MarginLeft, VPos, AreaWidth, 1) ; Draw a horizontal line
  ;
  If ECGadgets\MarginInterBlocks
    ECGadgets\InterBlocksVerticalPos = VPos
    VPos + ECGadgets\MarginInterBlocks
    ResizeGadget(ECGadgets\SeparatorGadget2, ECGadgets\MarginLeft, VPos, AreaWidth, 1) ; Draw a horizontal line
  EndIf
  ;
  VPos + InterLine * 0.25 + HorizLineShift
  ResizeGadget(ECGadgets\RainbowCanvas, TrackBarHPos + ECGadgets\TrackBarsWidth + ECGadgets\HorizGadgetsMargin, VPos - 2, ECGadgets\CanvasDim * DesktopResolutionX() / DesktopResolutionY(), ECGadgets\CanvasDim)
  ResizeGadget(ECGadgets\TargetGadget, TrackBarHPos + ECGadgets\TrackBarsWidth + ECGadgets\HorizGadgetsMargin, VPos, ECGadgets\TargetDim, ECGadgets\TargetDim)
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
  ;
  ResizeGadget(ECGadgets\SeparatorGadget3, ECGadgets\MarginLeft, VPos, AreaWidth, 1) ; Draw a horizontal line
  ;
  ; 'Cancel' and 'OK' buttons:
  ;
  VPos = WindowHeight(ECGadgets\EditColorWindow) - ECGadgets\MarginButtonsBottom - TextHeight - 3 * DesktopResolutionX()
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
  ZCR_CalculateBorderColor(ECGadgets, "Resize")
  ZCR_DrawRainbowCanvas(ECGadgets)
  ZCR_DrawTrackbarsAndPreviewCanvas(ECGadgets)
  ZCR_SetTrackBarsFromStrings(ECGadgets)
  ;
EndProcedure
;
Procedure ZCR_SetText()
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
  ;
EndProcedure
;
Procedure ZCR_CreateGadgets()
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  ;
  ECGadgets\PreviewCanvas = CanvasGadget(#PB_Any, 1, 1, 1, 1) ; Preview Color
  ;
  ECGadgets\Red_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Red_Input = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Red_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Green_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Green_Input = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Green_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Blue_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Blue_Input = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Blue_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\SeparatorGadget1 = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  If ECGadgets\MarginInterBlocks
    ECGadgets\SeparatorGadget2 = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  EndIf
  ;
  ECGadgets\RainbowCanvas = CanvasGadget(#PB_Any, 1, 1, 1, 1) ; Preview Color
  ;
  ECGadgets\Hue_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Hue_Input = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Hue_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Lum_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Lum_Input = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Lum_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\Sat_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\Sat_Input = StringGadget(#PB_Any, 1, 1, 1, 1, "", #PB_String_Numeric)
  ECGadgets\Sat_Trackbar = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\SeparatorGadget3 = CanvasGadget(#PB_Any, 1, 1, 1, 1)
  ;
  ECGadgets\HexColor_Legend = TextGadget(#PB_Any, 1, 1, 1, 1, "")
  ECGadgets\HexColor_Input = StringGadget(#PB_Any, 1, 1, 1, 1, "")
  ;
  ECGadgets\BCancel = ButtonGadget(#PB_Any, 1, 1, 1, 22, "")
  ECGadgets\BOk = ButtonGadget(#PB_Any, 1, 1, 1, 22, "")
  ;
EndProcedure
;
Procedure ZCR_CreateTargetAndCursors()
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  ;
  Protected TargetImage, CursorImage, ct
  ;
  TargetImage  = CreateImage(#PB_Any, ECGadgets\TargetDim, ECGadgets\TargetDim, 32, #PB_Image_Transparent)
  If StartVectorDrawing(ImageVectorOutput(TargetImage))
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
  ECGadgets\TargetGadget = ImageGadget(#PB_Any, 1, 1, 1, 1, ImageID(TargetImage))
  ;
  CursorImage  = CreateImage(#PB_Any, ECGadgets\CursorWidth, ECGadgets\LineHeight, 32, #PB_Image_Transparent)
  If StartVectorDrawing(ImageVectorOutput(CursorImage))
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
  ECGadgets\Red_Cursor   = ImageGadget(#PB_Any, 1, 1, 1, 1, ImageID(CursorImage))
  ECGadgets\Green_Cursor = ImageGadget(#PB_Any, 1, 1, 1, 1, ImageID(CursorImage))
  ECGadgets\Blue_Cursor  = ImageGadget(#PB_Any, 1, 1, 1, 1, ImageID(CursorImage))
  ECGadgets\Hue_Cursor   = ImageGadget(#PB_Any, 1, 1, 1, 1, ImageID(CursorImage))
  ECGadgets\Lum_Cursor   = ImageGadget(#PB_Any, 1, 1, 1, 1, ImageID(CursorImage))
  ECGadgets\Sat_Cursor   = ImageGadget(#PB_Any, 1, 1, 1, 1, ImageID(CursorImage))
EndProcedure
;
Procedure ZCR_SetFieldsSizeAndMargins()
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  ;
  Protected LegendHeight, HexColorWidth, InputWidth, tx$, LegendWidth, MarginCircular
  ;
  ECGadgets\TargetDim = 10     ; Width of the moving target over the rainbow canvas.
  ECGadgets\CursorWidth = 10   ; Width of the moving cursors over trackbars.
  ;
  ; Examine texts width and height
  If StartDrawing(WindowOutput(ECGadgets\EditColorWindow))
    DrawingFont(GetGadgetFont(#PB_Default))
    LegendHeight  = TextHeight("A")                        ; Memorize the text height.
    HexColorWidth = TextWidth("1234567890ABCDEF") * 6 / 16 ; Calculate the average width for 6 hexa chars.
    InputWidth = TextWidth("1234567890") * 3 / 10          ; Calculate the average width for 3 numeric chars.
    ;
    ; Look for the maximum legend width
    tx$ = GetGadgetText(ECGadgets\Red_Legend)
    LegendWidth   = TextWidth(tx$)
    tx$ = GetGadgetText(ECGadgets\Green_Legend)
    If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Blue_Legend)
    If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Hue_Legend)
    If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Sat_Legend)
    If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    tx$ = GetGadgetText(ECGadgets\Lum_Legend)
    If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    If ECGadgets\HexFieldVPos = -1
      tx$ = GetGadgetText(ECGadgets\HexColor_Legend)
      If TextWidth(tx$) > LegendWidth : LegendWidth = TextWidth(tx$) : EndIf
    EndIf
    StopDrawing()
  EndIf
  ;
  ECGadgets\LegendWidth = LegendWidth / DesktopResolutionX() + 5      ; This value set the width of 'Red', 'Green', Blue', 'Hue', Lightness', 'Saturation' and 'Hex value' legends.
  ECGadgets\LineHeight = LegendHeight + 5                             ; Height of one line.
  ECGadgets\InputWidth = InputWidth / DesktopResolutionX() + 10       ; Width of all input fields (except the 'Hex value' input field).
  ECGadgets\HexColorWidth = HexColorWidth / DesktopResolutionX() + 10 ; Width of the 'Hex value' input field.
  CompilerIf #PB_Compiler_OS = #PB_OS_Linux
    ECGadgets\InputWidth + 20
    ECGadgets\HexColorWidth + 20
  CompilerEndIf
  ECGadgets\ButtonsWidth = 80                                         ; Change this value to modify the width of 'OK' and 'Cancel' buttons.
  ;
EndProcedure
;
Procedure ZCR_TrackBarEvents(MouseButton, EventType, InputGadget, CanvasGadget, *ECGadgets.EditColorGadgetsStruct)
  ;
  Protected Value, localX, DontSet, VMax
  ;
  If EventType = #PB_EventType_MouseWheel Or MouseButton = #PB_EventType_LeftButtonDown
    Select InputGadget
      Case *ECGadgets\Red_Input, *ECGadgets\Green_Input, *ECGadgets\Blue_Input
        VMax = 255
        DontSet = #DontSetRGBString
      Case *ECGadgets\Hue_Input
        VMax = 239
        DontSet = #DontSetHSLString
      Case *ECGadgets\Lum_Input, *ECGadgets\Sat_Input
        VMax = 240
        DontSet = #DontSetHSLString
    EndSelect
    ;
    If EventType = #PB_EventType_MouseWheel
      Value = Val(GetGadgetText(InputGadget)) + GetGadgetAttribute(CanvasGadget, #PB_Canvas_WheelDelta) * 4
      If Value > VMax : Value = VMax : EndIf
      If Value < 0 : Value = 0 : EndIf
    EndIf
    If MouseButton = #PB_EventType_LeftButtonDown
      localX = GetGadgetAttribute(CanvasGadget, #PB_Canvas_MouseX)
      ;
      If localX > DesktopScaledX(GadgetWidth(CanvasGadget))
        localX = DesktopScaledX(GadgetWidth(CanvasGadget))
      ElseIf localX < 0
        localX = 0
      EndIf
      ;
      Value = localX / DesktopResolutionX() * VMax / GadgetWidth(CanvasGadget)
    EndIf
    If Val(GetGadgetText(InputGadget)) <> Value
      SetGadgetText(InputGadget, Str(Value))
      If DontSet = #DontSetHSLString
        *ECGadgets\Color = HSLToRGB(Val(GetGadgetText(*ECGadgets\Hue_Input)), Val(GetGadgetText(*ECGadgets\Sat_Input)), Val(GetGadgetText(*ECGadgets\Lum_Input)))
      Else
        *ECGadgets\Color = RGB(Val(GetGadgetText(*ECGadgets\Red_Input)), Val(GetGadgetText(*ECGadgets\Green_Input)), Val(GetGadgetText(*ECGadgets\Blue_Input)))
      EndIf
      ;
      ZCR_SetGadgetsFromColor(*ECGadgets, DontSet)
    EndIf
  EndIf
EndProcedure
;
Procedure ZCR_Events(Event, EventGadget, EventType)
  ;
  Shared ECGadgets.EditColorGadgetsStruct
  ;
  Static MouseButton, CanvasXDim, CanvasYDim
  ;
  Protected Red, Green, Blue, Saturation, Hue, Lightness
  Protected HStr$, mHStr$, ct, CChar$, localX, localY, GetOut
  ;
  If EventType = #PB_EventType_LeftButtonDown Or EventType = #PB_EventType_LeftButtonUp
    MouseButton = EventType
  ElseIf MouseButton = #PB_EventType_LeftButtonDown
    CompilerIf #PB_Compiler_OS = #PB_OS_Windows
      ; #PB_EventType_LeftButtonUp is sometimes missed under Windows.
      ; Check if button is still pressed:
      If Not(GetAsyncKeyState_(#VK_LBUTTON) >> 15)
        MouseButton = #PB_EventType_LeftButtonUp
      EndIf
    CompilerEndIf
  EndIf
  
  Select Event
    Case #PB_Event_Gadget
      Select EventGadget
        Case ECGadgets\HexColor_Input
          HStr$ = UCase(GetGadgetText(ECGadgets\HexColor_Input))
          If HStr$
            mHStr$ = HStr$
            For ct = 1 To Len(HStr$)
              CChar$ = Mid(HStr$, ct, 1)
              If FindString("0123456789ABCDEF", CCHar$) = 0
                HStr$ = ReplaceString(HStr$, CChar$, "")
                ct - 1
              EndIf
            Next
            If HStr$ <> mHStr$
              SetGadgetText(ECGadgets\HexColor_Input, HStr$)
              CompilerIf #PB_Compiler_OS = #PB_OS_Windows
                SendMessage_(GadgetID(ECGadgets\HexColor_Input), #EM_SETSEL, Len(HStr$), Len(HStr$))
              CompilerEndIf
            EndIf
            ECGadgets\Color = Val("$" + HStr$)
            ZCR_SetGadgetsFromColor(ECGadgets, #DontSetHexField)
          EndIf
          ;
        Case ECGadgets\Red_Trackbar
          ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Red_Input, ECGadgets\Red_Trackbar, ECGadgets)
          ;
        Case ECGadgets\Green_Trackbar
          ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Green_Input, ECGadgets\Green_Trackbar, ECGadgets)
          ;
        Case ECGadgets\Blue_Trackbar
          ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Blue_Input, ECGadgets\Blue_Trackbar, ECGadgets)
          ;
        Case ECGadgets\Hue_Trackbar
          ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Hue_Input, ECGadgets\Hue_Trackbar, ECGadgets)
          ;
        Case ECGadgets\Lum_Trackbar
          ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Lum_Input, ECGadgets\Lum_Trackbar, ECGadgets)
          ;
        Case ECGadgets\Sat_Trackbar
          ZCR_TrackBarEvents(MouseButton, EventType, ECGadgets\Sat_Input, ECGadgets\Sat_Trackbar, ECGadgets)
          ;
        Case ECGadgets\RainbowCanvas
          If EventType = #PB_EventType_MouseWheel Or MouseButton = #PB_EventType_LeftButtonDown
            Saturation = LimitGadgetValue(ECGadgets\Sat_Input, 0, 240)
            If EventType = #PB_EventType_MouseWheel
              Saturation + GetGadgetAttribute(ECGadgets\RainbowCanvas, #PB_Canvas_WheelDelta) * 10
              If Saturation > 240 : Saturation = 240 : EndIf
              If Saturation < 0 : Saturation = 0 : EndIf
            EndIf
            If MouseButton = #PB_EventType_LeftButtonDown
              localX = GetGadgetAttribute(ECGadgets\RainbowCanvas, #PB_Canvas_MouseX)
              localY = GetGadgetAttribute(ECGadgets\RainbowCanvas, #PB_Canvas_MouseY) 
              ;
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
              ;
              CanvasXDim = ECGadgets\CanvasDim * DesktopResolutionX()
              CanvasYDim = ECGadgets\CanvasDim * DesktopResolutionY()
              Hue = localX * 240 / CanvasXDim
              Lightness = (CanvasYDim - localY) * 240 / CanvasYDim
              If Saturation = 0 : Saturation = 1 : EndIf
            Else
              Hue = LimitGadgetValue(ECGadgets\Hue_Input, 0, 240)
              Lightness = LimitGadgetValue(ECGadgets\Lum_Input, 0, 240)
            EndIf

            ECGadgets\Color = HSLToRGB(Hue, Saturation, Lightness)
            ;
            SetGadgetText(ECGadgets\Sat_Input, Str(Saturation))
            SetGadgetText(ECGadgets\Hue_Input, Str(Hue))
            SetGadgetText(ECGadgets\Lum_Input, Str(Lightness))
            ;
            ZCR_SetGadgetsFromColor(ECGadgets, #DontSetHSLString)
            ;
          EndIf
          ;
        Case ECGadgets\Red_Input, ECGadgets\Green_Input, ECGadgets\Blue_Input
          Red = LimitGadgetValue(ECGadgets\Red_Input, 0, 255)
          Green = LimitGadgetValue(ECGadgets\Green_Input, 0, 255)
          Blue = LimitGadgetValue(ECGadgets\Blue_Input, 0, 255)
          ECGadgets\Color = RGB(Red, Green, Blue)
          ;
          ZCR_SetGadgetsFromColor(ECGadgets, #DontSetRGBString)
          ;
        Case ECGadgets\Hue_Input, ECGadgets\Sat_Input, ECGadgets\Lum_Input
          Hue = LimitGadgetValue(ECGadgets\Hue_Input, 0, 240)
          Saturation = LimitGadgetValue(ECGadgets\Sat_Input, 0, 240)
          Lightness = LimitGadgetValue(ECGadgets\Lum_Input, 0, 240)
          ECGadgets\Color = HSLToRGB(Hue, Saturation, Lightness)
          ;
          ZCR_SetGadgetsFromColor(ECGadgets, #DontSetHSLString)
          ;
        Case ECGadgets\BOk ; OK button
          GetOut = 1
        Case ECGadgets\BCancel ; Cancel button
          GetOut = -1
      EndSelect
      ;
  EndSelect
  ProcedureReturn GetOut
EndProcedure
;
Procedure ZapmanColorRequester(InitialColor, ParentID = #PB_Default, XShiftOrPos = 0, YShiftOrPos = 0, ParentAnchor = #CWO_Center, WindowAnchor = #CWO_Center)
  ;
  ; Return an edited color from InitialColor.
  ;
  ; • If ParentID is a valid window number, the dialog window will be positionned relatively to it,
  ;   and then shifted by XShiftOrPos and YShiftOrPos.
  ; • If ParentID = #CWO_ActiveWindowPos and GetActiveWindow() returns a valid window number,
  ;   GetActiveWindow() will be considered as the parent window.
  ; • If ParentID = #CWO_ActiveWindowPos and GetActiveWindow() does NOT return a valid window number,
  ;   the position will be calculated relatively to the main screen.
  ; • It ParentID = #CWO_AbsolutePos, XShiftOrPos and YShiftOrPos will be the absolute coordinates
  ;   of the new window.
  ; • If ParentID = #CWO_MonitorPos, the position will be calculated relatively to the screen where
  ;   GetActiveWindow() is found or relatively to the main screen if GetActiveWindow()
  ;   is not a valid window.
  ;
  ; • The new window will be positionned relatively to the center, top-left, top-right, bottom-left
  ;   or bottom-right of the parent window (or monitor), depending on the 'ParentAnchor' parameter.
  ;
  ; • The anchor of the new window can be the center, top-left, top-right, bottom-left or bottom-right
  ;   of the new window, depending on the 'WindowAnchor' parameter.
  ;
  ; • In all case, except if ParentID = #CWO_AbsolutePos, the position of the new window will be
  ;   shifted regarding the 'XShiftOrPos' and 'YShiftOrPos' parameters.
  ;
  ; Window and gadget numbers:
  Shared ECGadgets.EditColorGadgetsStruct
  ; Other variables
  Protected WWidth, WHeight, GetOut
  Protected Event, EventType, EventMenu
  Protected WParam, OX, OY, MarginCircular, *GadgetAdress, EventGadget
  ;
  Protected ParentWindow = #PB_Default
  If IsWindow(ParentID)
    ParentWindow = ParentID
    ParentID = WindowID(ParentID)
  ElseIf ParentID <> #PB_Default And IsWindow_(ParentID)
    ParentWindow = GetProp_(ParentID, "PB_WindowID") - 1
  ElseIf IsWindow(GetActiveWindow()) And ParentID = #PB_Default
    ParentWindow = GetActiveWindow()
    ParentID = WindowID(ParentWindow)
  EndIf
  If ParentWindow = #PB_Default
    ParentWindow = ParentID
  EndIf
  ;
  ECGadgets\Color = InitialColor
  ECGadgets\SGC_ColorType$ = ""
  ;
  If ListSize(InterfaceColorPresets()) > 0 And ListIndex(InterfaceColorPresets()) <> -1
    TextColor = InterfaceColorPresets()\TextColor
    BackGroundColor = InterfaceColorPresets()\BackGroundColor
  Else
    TextColor = #PB_Default
    BackGroundColor = #PB_Default
  EndIf
  ;
  WWidth = 420
  WHeight = 280
  WParam = #PB_Window_SystemMenu | #PB_Window_SizeGadget | #PB_Window_Invisible
  Protected ActiveWindow = ParentWindow
  If Not(IsWindow(ActiveWindow))
    ActiveWindow = GetActiveWindow()
  EndIf
  ;
  Protected ParentWindowID = ComputeWinOrigins(@OX, @OY, WWidth, WHeight, ParentWindow, XShiftOrPos, YShiftOrPos, ParentAnchor, WindowAnchor)
  ECGadgets\EditColorWindow = OpenWindow(#PB_Any, OX, OY, WWidth, WHeight, GetTextFromCatalog("EditColor"), WParam, ParentWindowID)
  If ECGadgets\EditColorWindow
    If IsWindow(ActiveWindow) And IsWindowEnabled_(WindowID(ActiveWindow))
      DisableWindow(ActiveWindow, #True)
      Protected ParentHasBeenDisabled = #True
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
    ; The following parameters allow to reuse the main part of the interface
    ; inside a windows comprising other gadgets.
    ;
    ; You can configure various types of margins in order to position the gadgets
    ; on the left, on the right, with an offset at the top or bottom.
    ;
    ; See the SetGadgetColorEx.pb example available at 
    ; https://www.editions-humanis.com/downloads/PureBasic/ZapmanDowloads_EN.htm
    ;
    ; Here, an identical margin is set for the forth sides of the window.
    ; But you can adjust them.
    MarginCircular = 15
    ECGadgets\MarginButtonsBottom = MarginCircular
    ECGadgets\MarginButtonsTop = 0
    ECGadgets\MarginLeft = MarginCircular
    ECGadgets\MarginRight = MarginCircular
    ECGadgets\MarginTop = MarginCircular
    ECGadgets\HexFieldVPos = #PB_Default ; Vertical position of the 'Hexa value' fields (#PB_Default is at the bottom of the window).
    ECGadgets\HorizGadgetsMargin = 5     ; Vertical margins between gadgets.
    ECGadgets\MarginInterBlocks = 0      ; Space between the two blocks of gadgets.
    ;
    ECGadgets\BackColor = GetRealColorFromType("BackgroundColor", BackGroundColor)  ; Background color for all canvas cursors.
    ECGadgets\TextColor = GetRealColorFromType("TextColor", TextColor)              ; Color used to calculate gadgets borders color.
    ;
    ZCR_CreateGadgets()
    ;
    ZCR_SetText()
    ZCR_SetFieldsSizeAndMargins()
    ZCR_CreateTargetAndCursors()
    ZCR_ResizeGadgets()
    ;
    ; Initialize cursor's positions for trackbars and gadget's colors:
    ZCR_SetGadgetsFromColor(ECGadgets)
    ;
    #ZCR_Escape_Cmd = 1
    #ZCR_Enter = 2
    AddKeyboardShortcut(ECGadgets\EditColorWindow, #PB_Shortcut_Escape, #ZCR_Escape_Cmd)
    AddKeyboardShortcut(ECGadgets\EditColorWindow, #PB_Shortcut_Return, #ZCR_Enter)
    ;
    CompilerIf Defined(SetGadgetColorEx, #PB_Procedure)
      ; Apply colors to gadgets if it is possible.
      ; The SetGadgetColorEx() is defined into the Zapman 'SetGadgetColorEx.pb' library
      ; and this file must be included by your program before this file to benefit from
      ; color theme support.
      ;
      If BackGroundColor <> #PB_Default Or TextColor <> #PB_Default
        ; Colorize gadgets
        *GadgetAdress = @ECGadgets\Red_Legend
        Repeat
          If TextColor <> #PB_Default
            SetGadgetColorEx(PeekI(*GadgetAdress), #PB_Gadget_FrontColor, TextColor, 1)
          EndIf
          If BackGroundColor <> #PB_Default
            SetGadgetColorEx(PeekI(*GadgetAdress), #PB_Gadget_BackColor , BackGroundColor, 1)
          EndIf
          *GadgetAdress + SizeOf(Integer)
        Until *GadgetAdress > @ECGadgets\BCancel
      EndIf
    CompilerEndIf
    ;
    ;
    ; The window was invisible until now, because we created it with #PB_Window_Invisible.
    ; We make it visible now.
    HideWindow(ECGadgets\EditColorWindow, #False)
    ;
    Repeat
      Event       = WaitWindowEvent()
      EventGadget = EventGadget()
      EventType   = EventType()
      EventMenu   = EventMenu()
      ; Manage trackbars aned input events:
      
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
    CloseWindow(ECGadgets\EditColorWindow)
  EndIf
  ;
  If ParentHasBeenDisabled
    DisableWindow(ActiveWindow, #False)
  EndIf
  ProcedureReturn ECGadgets\Color
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  ; The following won't run when this file is used as 'Included'.
  ;
  Debug Hex(ZapmanColorRequester($EA5D8))
CompilerEndIf
; IDE Options = PureBasic 6.20 (Windows - x64)
; CursorPosition = 24
; Folding = ---------
; EnableXP
; DPIAware