;
; ****************************************************************************************
;
;                                   SetGadgetColorEx()
;                                      Windows only
;                                 Zapman - March 2025 - 6
;
;            This file should be saved under the name "SetGadgetColorEx.pbi".
;
;            The first part of this library offers miscellaneous functions like:
; • IsDarkModeEnabled() allows to know if the running computer is set to dark mode or not.
; • ApplyDarkModeToWindow() matches the title bar of a window with the computer theme.
; • GetProperty() can retrieve a PureBasic ID from a handle (the contrary of GadgetID()).
;
; The second part of this library is a set of specialized functions used by the third part.
;
;              The third part of this library extends the SetGadgetColor()
;              PureBasic native function to all of the PureBasic gadgets.
;    You can now call SetGadgetColor() for buttons, panels, combos, stringGadgets, etc.
;        You can also call ApplyColorsToAllGadgets() to attribute a couple of colors
;                to all of the gadgets created into a particular window.
;
;                        A forth part contains a demo procedure.
;
; IMPORTANT NOTE: the 'ApplyColorThemes.pbi' library is also available on the Zapman website:
;        https://www.editions-humanis.com/downloads/PureBasic/ZapmanDowloads_EN.htm
;         It offers you a cool and sophisticated interface to manage colors themes
;                             and apply them to your windows.
;
; ****************************************************************************************
;
;-   1--- FIRST PART: MISCELLANEOUS FUNCTIONS (possibly reusable for other needs) ---
;
CompilerIf Not(Defined(DwmSetWindowAttribute, #PB_Prototype))
  Prototype.i DwmSetWindowAttribute(hWnd.i, dwAttribute.i, pvAttribute.i, cbAttribute.i)
CompilerEndIf
;
CompilerIf Not(Defined(IsDarkModeEnabled, #PB_Procedure))
  Procedure IsDarkModeEnabled()
    ;
    ; Detects if dark mode is enabled in Windows
    ;
    Protected key = 0
    Protected darkModeEnabled = 0
    ;
    If RegOpenKeyEx_(#HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", 0, #KEY_READ, @key) = #ERROR_SUCCESS
      Protected value = 1
      Protected valueSize = SizeOf(value)
      If RegQueryValueEx_(key, "AppsUseLightTheme", 0, #Null, @value, @valueSize) = #ERROR_SUCCESS
        darkModeEnabled = Abs(value - 1) ; 0 = dark, 1 = light
      EndIf
      RegCloseKey_(key)
    EndIf
    ;
    ProcedureReturn darkModeEnabled
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(ApplyDarkModeToWindow, #PB_Procedure))
  Procedure ApplyDarkModeToWindow(Window = 0)
    ;
    ; Applies dark theme to a window if dark theme is enabled in Windows.
    ;
    Protected hWnd = WindowID(Window)
    ;
    If hWnd And OSVersion() >= #PB_OS_Windows_10
      Protected hDwmapi = OpenLibrary(#PB_Any, "dwmapi.dll")
      ;
      If hDwmapi
        Protected DwmSetWindowAttribute_.DwmSetWindowAttribute = GetFunction(hDwmapi, "DwmSetWindowAttribute")
        ; Enable dark mode if possible
        If DwmSetWindowAttribute_
          Protected darkModeEnabled = IsDarkModeEnabled()
          If darkModeEnabled
            #DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DwmSetWindowAttribute_(hWnd, #DWMWA_USE_IMMERSIVE_DARK_MODE, @darkModeEnabled, SizeOf(darkModeEnabled))
            SetWindowColor(Window, $202020)
            ;
            ; Force the window to repaint:
            If IsWindowVisible_(hWnd)
              HideWindow(Window, #True)
              HideWindow(Window, #False)
            EndIf
          EndIf
        EndIf
        ;
        CloseLibrary(hDwmapi)
      EndIf
    EndIf
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(GetProperty, #PB_Procedure))
  ;
  Procedure GP_EnumPropsProc(hWnd, lpszString, hData, lParam)
    If lpszString > 65536  ; When lpszString < 65536, it is an ATOM.
      If PeekS(lpszString) = PeekS(lParam)
        PokeS(lParam, "*")
        ProcedureReturn #False ; Stop searching.
      EndIf
    EndIf
    ProcedureReturn #True ; Continue searching.
  EndProcedure
  ;
  Procedure GetProperty(hWnd, propName$)
    ; Check if the window designated by 'hWnd' has a property named propName$
    ; and return its value if it exists.
    ; When it doesn't exist the returned value is -1.
    ; This allows to make a difference between a non existent property and
    ; a property that does exist but have a value of zero.
    ;
    ; You can call GetProperty() with "PB_ID" as second parameter
    ; to retreive a gadget PurBasic number from the gadget handle:
    ; #MyGadgetNum = 1
    ; StringGadget(#MyGadgetNum, 10, 10, 100, 25, "My string")
    ; handle = GadgetID(#MyGadgetNum)
    ; PBNum = GetProperty(handle, "PB_ID")
    ; ----> Now, PBNum = #MyGadgetNum
    ;
    If propName$
      Protected TString$ = propName$
      EnumPropsEx_(hWnd, @GP_EnumPropsProc(), @TString$)
      If TString$ = "*"
        ProcedureReturn GetProp_(hWnd, propName$)
      Else
        ProcedureReturn -1
      EndIf
    Else
      ProcedureReturn -1
    EndIf
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(GetGadgetParentWindow, #PB_Procedure))
  Procedure GetGadgetParentWindow(Gadget)
    ;
    ; This function is by 'mk-soft', english forum.
    ;
    Protected ID, r1
    ;
    If IsGadget(Gadget)
      CompilerSelect #PB_Compiler_OS
        CompilerCase #PB_OS_MacOS
          Protected *Gadget.sdkGadget = IsGadget(Gadget)
          If *Gadget
            ID = WindowID(*Gadget\Window)
            r1 = PB_Window_GetID(ID)
          Else
            r1 = -1
          EndIf
        CompilerCase #PB_OS_Linux
          ID = gtk_widget_get_toplevel_(GadgetID(Gadget))
          If ID
            r1 = g_object_get_data_(ID, "pb_id")
          Else
            r1 = -1
          EndIf
        CompilerCase #PB_OS_Windows           
          ID = GetAncestor_(GadgetID(Gadget), #GA_ROOT)
          r1 = GetProp_(ID, "PB_WINDOWID")
          If r1 > 0
            r1 - 1
          Else
            r1 = -1
          EndIf
      CompilerEndSelect
    Else
      r1 = -1
    EndIf
    ProcedureReturn r1
  EndProcedure
CompilerEndIf
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
CompilerIf Not(Defined(ShiftColorComponent, #PB_Procedure))
  Procedure ShiftColorComponent(CComponent, Shift)
    CComponent + Shift
    If CComponent < 0 : CComponent = 0 : EndIf
    If CComponent > 256 : CComponent = 256 : EndIf
    ProcedureReturn CComponent
  EndProcedure
CompilerEndIf
;
;
; ****************************************************************************************
;-              2--- SECOND PART: SPECIALIZED FUNCTIONS FOR THIS LIBRARY ---
;
EnumerationBinary EdgesStyle
  #SGCE_NoEdges = 0
  #SGCE_HoverSensitive = 1
  #SGCE_Flat
  #SGCE_Single
  #SGCE_Double
  #SGCE_WithTitle
EndEnumeration
;
Structure SGCE_GadgetsInfosStruct
  ID.i
  Handle.i
  FrontColor.i
  BackColor.i
  EdgesColor.l
  HighlightedEdgesColor.l
  EdgesStyle.i
  Enabled.i
  MainWindow.i
EndStructure
;
#SGCE_StandardColors = -1
;
NewList SGCE_GadgetsInfos.SGCE_GadgetsInfosStruct()
;
Procedure SGCE_IsCustomColorGadget(GadgetID)
  ;
  Protected CustomColor = 1
  ;
  Select GadgetType(GadgetID)
    Case #PB_GadgetType_Button, #PB_GadgetType_CheckBox, #PB_GadgetType_Option, #PB_GadgetType_Panel
      ; Colors are managed by this library.
      CustomColor = 1
    Case #PB_GadgetType_ComboBox, #PB_GadgetType_ExplorerCombo, #PB_GadgetType_Frame, #PB_GadgetType_Container
      ; Colors are managed by this library.
      CustomColor = 1
    Case #PB_GadgetType_Calendar, #PB_GadgetType_Date, #PB_GadgetType_ListIcon
      ; Colors are managed by this library.
      CustomColor = 1
    Case #PB_GadgetType_ExplorerList
      ; Colors are managed by this library.
      CustomColor = 1
      ;
    Case #PB_GadgetType_Editor, #PB_GadgetType_String, #PB_GadgetType_ListView, #PB_GadgetType_ButtonImage
      ; Colors are managed by PureBasic, but Edges need to be redrawn.
      CustomColor = -1
    Case #PB_GadgetType_Spin, #PB_GadgetType_Image, #PB_GadgetType_ExplorerTree, #PB_GadgetType_Tree
      ; Colors are managed by PureBasic, but Edges need to be redrawn.
      CustomColor = -1
    Case #PB_GadgetType_MDI, #PB_GadgetType_ScrollArea, #PB_GadgetType_Canvas
      ; Colors are managed by PureBasic, but Edges need to be redrawn.
      CustomColor = -1
      ;
      ; The following gadgets are natively managed by PureBasic and don't need to be redrawn:
      ;
    Case #PB_GadgetType_HyperLink, #PB_GadgetType_Text, #PB_GadgetType_ProgressBar
      ; Colors are managed by PureBasic.
      CustomColor = 0
  EndSelect
  ProcedureReturn CustomColor
EndProcedure
;
CompilerIf Not(Defined(RGBAColor, #PB_Structure))
  Structure RGBAColor
    ; Field are 'a' type and not 'b' type,
    ; because 'a' is unsigned and 'b' is signed.
    red.a
    green.a
    blue.a
    alpha.a
  EndStructure
CompilerEndIf
;
Procedure SGCE_PartialReplaceColors(StartX, StartY, EndX, EndY, Width, *BitmapAddress, maxlum, LowContrast.f, *backColor.RGBAColor, *FrontColor.RGBAColor, DoBlowOut = 1)
  ;
  Structure tagRGBQUAD
    ; Field are 'a' type and not 'b' type,
    ; because 'a' is unsigned and 'b' is signed.
    blue.a
    green.a
    red.a
    alpha.a
  EndStructure
  ;
  Protected y, x, *pixelColor.tagRGBQUAD, PixelDarkness.f, PixelLuminosity.f
  Protected BlowOut.f, CrushBlacks.f, mColor.l, mConvColor.l
  ;
  For y = StartY To EndY - 1
    Protected yXWidth = y * Width
    For x = StartX To EndX - 1
      ; Calculate the address of the current pixel in 32 bits (BGRA)
      *pixelColor = *BitmapAddress + (yXWidth + x) * 4
      If mColor = PeekL(*pixelColor)
        ; If the pixel has the same color as the last one, simply apply last result:
        PokeL(*pixelColor, mConvColor)
      Else
        mColor = PeekL(*pixelColor)
        ;
        PixelLuminosity = (*pixelColor\red + *pixelColor\green + *pixelColor\blue) / maxlum
        ;
        ; The following correction is intended to improve text readability when
        ; the contrast between the background color (*backColor) and the stroke color (*FrontColor)
        ; is medium or low.
        ;
        ; BlowOut brings high luminosities closer to white.
        ; (In photography, this is called "blowing out" the whites).
        If DoBlowOut And PixelLuminosity > 0.7
          Protected Emphasys.f = 3.7
          BlowOut = (PixelLuminosity + 1) * (PixelLuminosity + 1) / Emphasys
          If BlowOut > 1 : BlowOut = 1 : EndIf
          ; The lower the contrast between the background and stroke color,
          ; the more the BlowOut correction is applied.
          PixelLuminosity = PixelLuminosity * (1 - LowContrast) + BlowOut * LowContrast

        EndIf
        If DoBlowOut = 2 And PixelLuminosity < 0.99
          ; Boost light gray for PanelGadget, because its lightgray value
          ; is not the same than other gadgets.
          PixelLuminosity * 0.95
        EndIf
        If PixelLuminosity < 0.3
          ; CrushBlacks darkens the dark tones.
          ; (In photography, this is called "crushing" the blacks).
          CrushBlacks = PixelLuminosity * PixelLuminosity * PixelLuminosity
          ; The lower the contrast between the background and stroke color,
          ; the more the CrushBlacks correction is applied.
          PixelLuminosity = PixelLuminosity * (1 - LowContrast) + CrushBlacks * LowContrast
        EndIf
        ;
        If PixelLuminosity > 1 : PixelLuminosity = 1 : EndIf
        PixelDarkness = 1 - PixelLuminosity
        ;
        ; The principle of color modification is as follows:
        ; - The lighter the pixel, the less its color is retained and the more it is replaced by *backColor:
        *pixelColor\red   * PixelDarkness + *backColor\red   * PixelLuminosity
        *pixelColor\green * PixelDarkness + *backColor\green * PixelLuminosity
        *pixelColor\blue  * PixelDarkness + *backColor\blue  * PixelLuminosity
        ;
        ; - The darker the pixel, the less its color is retained and the more it is replaced by *FrontColor:
        *pixelColor\red   * PixelLuminosity + *FrontColor\red   * PixelDarkness
        *pixelColor\green * PixelLuminosity + *FrontColor\green * PixelDarkness
        *pixelColor\blue  * PixelLuminosity + *FrontColor\blue  * PixelDarkness
        mConvColor = PeekL(*pixelColor)
      EndIf
      ;
    Next
  Next
EndProcedure
;
Procedure SGCE_CreateBitmapWithAddress(hDCSrce, Width, Height, *backColor.RGBAColor, *FrontColor.RGBAColor, *memDC.Integer, *BitmapAddress.Integer, *oldBitmap.Integer, *maxlum.Integer, *LowContrast.float)
  ;
  ; Initialize a BITMAPINFO structure:
  Protected bmi.BITMAPINFO
  bmi\bmiHeader\biSize = SizeOf(BITMAPINFOHEADER)
  bmi\bmiHeader\biWidth = Width
  bmi\bmiHeader\biHeight = Height
  bmi\bmiHeader\biPlanes = 1
  bmi\bmiHeader\biBitCount = 32     ; 32 bits par pixel (RGBA).
  bmi\bmiHeader\biCompression = #BI_RGB
  ;
  ; Calculate the maximum brightness of a pixel:
  *maxlum\i = Red(GetSysColor_(#COLOR_BTNFACE)) + Green(GetSysColor_(#COLOR_BTNFACE)) + Blue(GetSysColor_(#COLOR_BTNFACE))
  ; LowContrast is calculated to be 1 when the contrast between the asked background color (BackColor)
  ; and the asked drawing color (EdgesColor) is zero, and to be zero when this contrast is maximal:
  *LowContrast\f = (*backColor\red + *backColor\green + *backColor\blue - *FrontColor\red - *FrontColor\green - *FrontColor\blue) / *maxlum\i
  *LowContrast\f = Abs(1 - Abs(*LowContrast\f))
  ;
  ; Create a DIBSection and retrieve the pointer to the pixels:
  Protected hBitmap = CreateDIBSection_(hDCSrce, bmi, #DIB_RGB_COLORS, *BitmapAddress, 0, 0)
  ;
  ; Create a compatible memory context and copy the pixels to it:
  *memDC\i = CreateCompatibleDC_(hDCSrce)
  ;
  *oldBitmap\i = SelectObject_(*memDC\i, hBitmap)
  BitBlt_(*memDC\i, 0, 0, Width, Height, hDCSrce, 0, 0, #SRCCOPY)
  ProcedureReturn hBitmap
EndProcedure
;
Procedure SGCE_RepaintEdges(*GadgetsInfos.SGCE_GadgetsInfosStruct, Highlight = #False, hDCType = 1)
  ;
  Protected gRect.Rect, pt.point, ps.PAINTSTRUCT
  Protected hdc, hOldPen, hOldBrush, hBrush, hPen
  Protected gHandle = *GadgetsInfos\Handle
  Protected EdgeShift, EdgeWidth, EdgesColor, PaintBackground
  ;
  If *GadgetsInfos\EdgesStyle & #SGCE_Double
    EdgeWidth = 2
  ElseIf *GadgetsInfos\EdgesStyle & #SGCE_Single Or *GadgetsInfos\EdgesStyle & #SGCE_Flat Or *GadgetsInfos\EdgesStyle & #SGCE_WithTitle
    EdgeWidth = 1
  Else
    EdgeWidth = 0
  EndIf
  ;
  If Highlight
    EdgesColor = *GadgetsInfos\HighlightedEdgesColor
    If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Date
      PaintBackground = #True
    Else
      PaintBackground = #False
    EndIf
  Else
    If *GadgetsInfos\EdgesStyle & #SGCE_Flat
      EdgesColor = *GadgetsInfos\FrontColor
    Else
      EdgesColor = *GadgetsInfos\EdgesColor
    EndIf
    ;
    If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Container
      PaintBackground = #True
    ElseIf GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Frame And GetWindow_(GadgetID(*GadgetsInfos\ID), #GW_CHILD)
      PaintBackground = #True
    EndIf
  EndIf
  ;
  If EdgeWidth Or PaintBackground
    If Abs(hDCType) > 2
      ; Edges painting is occuring in #WM_PAINT for FrameGadget.
      ; BeginPaint is allready done and hDCType contains the hDC
      hdc = hDCType
      GetClientRect_(gHandle, @gRect)
      If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Frame
        ; Room must be reserved for the title of the frame:
        gRect\top + DesktopScaledX(10) - 3
      EndIf
    ElseIf hDCType = -1
      ; Edges painting is occuring in #WM_PAINT handling and the client rectangle
      ; is wide enough to draw the edges.
      hDC = BeginPaint_(gHandle, ps)
      GetClientRect_(gHandle, @gRect)
      gRect\right + 2 * EdgeWidth : gRect\bottom + 2 * EdgeWidth
    ElseIf hDCType > 0
      ; Edges painting is occuring in #WM_PAINT, but the client rectangle
      ; cannot be used.
      ; Prepare to paint in the parent window hdc
      hdc =  GetDC_(GetParent_(gHandle))
      GetWindowRect_(gHandle, @gRect)
      ; GetWindowRect_ returns values relative to the screen.
      ; Correct it:
      pt\x = gRect\left : pt\y = gRect\top
      ScreenToClient_(GetParent_(gHandle), @pt)
      gRect\right - gRect\left + pt\x : gRect\bottom - gRect\top + pt\y
      gRect\left = pt\x : gRect\top = pt\y
    Else ; hDCType = 0
      hdc =  GetDC_(gHandle)
      ; Edges painting is occuring in #WM_NCPAINT, after BeginPaint_()/EndPaint().
      ; WindowRect is needed, but with a zero positionning:
      GetWindowRect_(gHandle, @gRect)
      ; Set position relative to zero for client painting:
      gRect\right - gRect\left : gRect\bottom - gRect\top
      gRect\left = 0 : gRect\top = 0
      ;
    EndIf
    ;
    If EdgeWidth
      hPen = CreatePen_(#PS_SOLID, EdgeWidth, EdgesColor)
    Else
      hPen = CreatePen_(#PS_SOLID, EdgeWidth, *GadgetsInfos\BackColor)
    EndIf
    If PaintBackground
      hBrush = CreateSolidBrush_(*GadgetsInfos\BackColor)
    Else
      hBrush = GetStockObject_(#NULL_BRUSH)
    EndIf
    hOldPen = SelectObject_(hDC, hPen)
    hOldBrush = SelectObject_(hDC, hBrush)
    ;
    Rectangle_(hdc, gRect\left - EdgeShift + EdgeWidth/2, gRect\top - EdgeShift + EdgeWidth/2, gRect\right - EdgeShift, gRect\bottom - EdgeShift)
    If PaintBackground = 0 And EdgeWidth = 1
      ; Erase the inner white frame:
      DeleteObject_(hPen)
      hPen = CreatePen_(#PS_SOLID, 1, *GadgetsInfos\BackColor)
      SelectObject_(hDC, hPen)
      Rectangle_(hdc, gRect\left - EdgeShift + 1, gRect\top - EdgeShift + 1, gRect\right - EdgeShift - 1, gRect\bottom - EdgeShift - 1)
    EndIf
    SelectObject_(hDC, hOldPen)
    SelectObject_(hDC, hOldBrush)
    DeleteObject_(hBrush)
    DeleteObject_(hPen)
    If Abs(hDCType) <= 2 And hDCType <> -1
      DeleteDC_(hdc)
    EndIf
  ElseIf hDCType = -1
    BeginPaint_(gHandle, ps)
  EndIf
  If hDCType = -1
    EndPaint_(gHandle, ps)
  EndIf
EndProcedure
;
Define SGCE_KillCalendarAnimation
Procedure SGCE_DrawAndReplaceColors(gHandle, lParam, OldCBProc, BackColor, FrontColor, GType)
  ;
  Protected ps.PAINTSTRUCT, gRect.Rect, ActualTheme
  Protected hBitmap, oldBitmap, memDC, maxlum, LowContrast.f
  Protected *BitmapAddress, CenterRect.Rect, hBrush, hDC
  Shared SGCE_KillCalendarAnimation
  ;
  If GType = #PB_GadgetType_Date Or GType = #PB_GadgetType_Calendar
    If SGCE_KillCalendarAnimation
      ActualTheme = GetWindowTheme_(gHandle)
      If ActualTheme <> 0
        ; Save the actual theme of the Calendar:
        SetProp_(gHandle, "SGCE_OldTheme", ActualTheme)
        ; To kill the Calendar animations, temporarily set the theme to null.
        SetWindowTheme_(gHandle, "", "")
        ; Colorize the Untheme version of Calendar to avoid visible flickering:
        SendMessage_(gHandle, #MCM_SETCOLOR, #MCSC_BACKGROUND,  BackColor)
      EndIf
      SGCE_KillCalendarAnimation = #False
    EndIf
    ;
    ActualTheme = GetWindowTheme_(gHandle)
    If ActualTheme = 0 And GetProp_(gHandle, "SGCE_OldTheme")
      ; To kill the calendar animations, the theme of the calendar has been
      ; temporarily set to null. Restore it:
      SetWindowTheme_(gHandle, @"Explorer", #Null)
    EndIf
  EndIf
    ;
  GetClientRect_(gHandle, @gRect)
  ;
  hdc = BeginPaint_(gHandle, ps)
  ;
  Protected width = gRect\right - gRect\left
  Protected Height = gRect\bottom - gRect\top
  ;
  hBitmap = SGCE_CreateBitmapWithAddress(hDC, Width, Height, @BackColor, @FrontColor, @memDC, @*BitmapAddress, @oldBitmap, @maxlum, @LowContrast)
  ;
  ; Call the normal #WM_PAINT process to draw the gadget.
  ; This will draw inside memDC instead of hdc:
  CallWindowProc_(OldCBProc, gHandle, #WM_PAINT, memDC, lParam)
  ;
  ; Change image colors:
  ;
  If GType = #PB_GadgetType_Panel
    SelectClipRgn_(memDC, 0)
    ; If the gadget is big (this can occur with PanelGadgets),
    ; replacing colors pixel by pixel can take too much time. In that case,
    ; limit the replacement to the Edges and repaint the center with BackColor.
    ;
    Protected hSlice = DesktopScaledX(3), vSlice = DesktopScaledY(3), upSlice = DesktopScaledY(26)
    ;
    If hSlice > Width : hSlice = Width : EndIf
    If vSlice > Height : vSlice = Height : EndIf
    If vSlice + upSlice > Height : upSlice = Height - vSlice : EndIf
    ; horizontal bottom Edge
    SGCE_PartialReplaceColors(0, 0, Width, hSlice, Width, *BitmapAddress, maxlum, LowContrast, @BackColor, @FrontColor, 2)
    ; horizontal top Edge
    SGCE_PartialReplaceColors(0, Height - upSlice, Width, Height, Width, *BitmapAddress, maxlum, LowContrast, @BackColor, @FrontColor, 2)
    ; vertical left Edge
    SGCE_PartialReplaceColors(0, vslice, hSlice, Height - upSlice, Width, *BitmapAddress, maxlum, LowContrast, @BackColor, @FrontColor, 2)
    ; vertical right Edge
    SGCE_PartialReplaceColors(Width - hSlice, vslice, Width, Height - upSlice, Width, *BitmapAddress, maxlum, LowContrast, @BackColor, @FrontColor, 2)
    ;
    If Width > hSlice * 2 And Height > (vSlice + upSlice)
      hBrush = CreateSolidBrush_(PeekL(@BackColor))
      CenterRect\left = hSlice : CenterRect\right = Width - hSlice
      CenterRect\top = upSlice : CenterRect\bottom = Height - vSlice
      FillRect_(memDC, CenterRect, hBrush)
      DeleteObject_(hBrush)
    EndIf
    ;
  Else;If GType <> #PB_GadgetType_ComboBox
    SGCE_PartialReplaceColors(0, 0, Width, Height, Width, *BitmapAddress, maxlum, LowContrast, @BackColor, @FrontColor)
  EndIf
  ;
  ;
  ; Copy memDC to hdc
  BitBlt_(hdc, gRect\left, gRect\top, width, Height, memDC, 0, 0, #SRCCOPY)
  ;
  ; Clean up:
  SelectObject_(memDC, oldBitmap)
  DeleteObject_(hBitmap)
  DeleteDC_(memDC)
  ;           
  ; End the drawing process:
  EndPaint_(gHandle, ps)
EndProcedure
;
Procedure SGCE_ChildrenCallback(CTRLhandle, uMsg, wParam, lParam)
  ;
  Protected SGCE_OldCallBack = GetProp_(CTRLhandle, "SGCE_OldCallBack")
  Protected *GadgetsInfos.SGCE_GadgetsInfosStruct = GetWindowLongPtr_(CTRLhandle, #GWL_USERDATA)
  ;
  If uMsg = #WM_PAINT
    SGCE_DrawAndReplaceColors(CTRLhandle, lParam, SGCE_OldCallBack, *GadgetsInfos\BackColor, *GadgetsInfos\FrontColor, GadgetType(*GadgetsInfos\ID))
    ProcedureReturn 1
  EndIf
  ;
  ProcedureReturn CallWindowProc_(SGCE_OldCallBack, CTRLhandle, uMsg, wParam, lParam)
EndProcedure
;
Procedure SGCE_DropDownCallback(gHandle, uMsg, wParam, lParam)
  ;
  Protected SGCE_OldCallBack = GetProp_(gHandle, "SGCE_OldCallBack")
  Protected ps.PAINTSTRUCT
  Protected *GadgetsInfos.SGCE_GadgetsInfosStruct = GetWindowLongPtr_(gHandle, #GWL_USERDATA)
  ;
  If uMsg = #WM_PAINT
    BeginPaint_(gHandle, ps)
    ; Tell the system that painting is done.
    EndPaint_(gHandle, ps)
    ;
    SGCE_RepaintEdges(*GadgetsInfos, #True, 0)
    ProcedureReturn 0
  EndIf
  ;
  ProcedureReturn CallWindowProc_(SGCE_OldCallBack, gHandle, uMsg, wParam, lParam)
EndProcedure
;
Procedure SGCE_DrawContainerGadget(*GadgetsInfos.SGCE_GadgetsInfosStruct)
  ;
  ; #PB_GadgetType_Container is not managed by the system as other gadgets.
  ; Edges are repainted in #WM_PAINT but with Parent hDC.
  ; It will be entirelly drawn (and not just colorized) here.
  ;
  Protected ps.PAINTSTRUCT
  Protected gHandle = *GadgetsInfos\Handle
  ;
  BeginPaint_(gHandle, ps)
  ; Tell the system that painting is done.
  EndPaint_(gHandle, ps)
  ;
  ; Now, use Parent hDC to paint the edges:
  SGCE_RepaintEdges(*GadgetsInfos)
EndProcedure
;
Procedure SGCE_DrawFrameGadget(*GadgetsInfos.SGCE_GadgetsInfosStruct)
  ;
  ; #PB_Frame_Container needs a special handling because it can be
  ; transparent or not and will print a title or not regarding its creating flags.
  ; It will be entirelly drawn (and not just colorized) here.
  ;
  Protected gRect.Rect, Size.SIZE, ps.PAINTSTRUCT
  Protected gHandle = *GadgetsInfos\Handle, hBrush
  ;
  If *GadgetsInfos\EdgesStyle & #SGCE_WithTitle = 0
    SGCE_RepaintEdges(*GadgetsInfos, #False, -1)
  Else
    ;
    Protected hDC = BeginPaint_(gHandle, ps)
    ; Draw the edges
    SGCE_RepaintEdges(*GadgetsInfos, #False, hDC)
    ;
    ; Get frame title:
    Protected Text$ = GetGadgetText(*GadgetsInfos\ID)
    Protected hFont = SendMessage_(gHandle, #WM_GETFONT, 0, 0)
    SelectObject_(hDC, hFont)
    ;
    ; Erase the top edge behind the text:
    If Text$
      GetClientRect_(gHandle, @gRect.Rect)
      gRect\Left + DesktopScaledX(4)
      gRect\top + DesktopScaledX(10) - 3
      gRect\bottom = gRect\top + 1
      GetTextExtentPoint32_(hDC, Text$, Len(Text$), @Size) ; Get the text width
      gRect\Right = gRect\Left + Size\cx + DesktopScaledX(6)
      hBrush = CreateSolidBrush_(*GadgetsInfos\BackColor)
      FillRect_(hdc, gRect, hBrush)
      DeleteObject_(hBrush)
      ;
      ; Draw the text:
      gRect\top - DesktopScaledX(15) + 7
      gRect\bottom + 30
      SetTextColor_(hDC, *GadgetsInfos\FrontColor)
      SetBkMode_(hDC, #TRANSPARENT)
      DrawText_(hDC, Text$, Len(Text$), @gRect.Rect, #DT_CENTER | #DT_SINGLELINE)
    EndIf
    ;
    EndPaint_(gHandle, ps)
  EndIf
  ProcedureReturn 1
  ;
EndProcedure
;
Procedure SGCE_DrawExplorerComboItem(SGCE_OldCallBack, gHandle, uMsg, wParam, lparam, *GadgetsInfos.SGCE_GadgetsInfosStruct)
  ;
  Protected bitmapInfo.BITMAP
  Protected *drawItem.DRAWITEMSTRUCT = lParam
  Protected hDC = *drawItem\hDC
  Protected disWidth = *drawItem\rcItem\right - *drawItem\rcItem\left
  Protected disHeight = *drawItem\rcItem\bottom - *drawItem\rcItem\top
  ; 
  ; Recover the bitmap size:
  Protected hBitmapSrc = GetCurrentObject_(hDC, #OBJ_BITMAP)
  GetObject_(hBitmapSrc, SizeOf(BITMAP), @bitmapInfo.BITMAP)
  Protected Width = bitmapInfo\bmWidth
  Protected Height = bitmapInfo\bmHeight
  ;
  Protected hBitmap, oldBitmap, memDC, maxlum, LowContrast.f
  Protected *BitmapAddress
  ;
  hBitmap = SGCE_CreateBitmapWithAddress(hDC, Width, Height, @*GadgetsInfos\BackColor, @*GadgetsInfos\FrontColor, @memDC, @*BitmapAddress, @oldBitmap, @maxlum, @LowContrast)
  ;         
  SelectObject_(memDC, GetCurrentObject_(hDC, #OBJ_FONT))
  ;
  ; Call the normal process to draw the item.
  ; This will draw inside memDC instead of hdc:
  *drawItem\hDC = memDC
  CallWindowProc_(SGCE_OldCallBack, gHandle, uMsg, wParam, lparam)
  *drawItem\hDC = hDC
  ;
  ; Decide if the colors must be replaced or not:
  If *drawItem\itemAction <> #ODA_FOCUS
    SGCE_PartialReplaceColors(0, 0, Width, Height, Width, *BitmapAddress, maxlum, LowContrast, @*GadgetsInfos\BackColor, @*GadgetsInfos\FrontColor, 0)
  EndIf
  ;
  ; Copy memDC to hdc
  If *drawItem\itemAction = #ODA_DRAWENTIRE And *drawItem\rcItem\top < SendMessage_(gHandle, #CB_GETITEMHEIGHT, 0, 0)
    SelectClipRgn_(hDC, 0)
    ; Copy all the bitmap:
    BitBlt_(hDC, 0, 0, Width, Height, memDC, 0, 0, #SRCCOPY)
  Else
    ; Copy only the item line:
    BitBlt_(hDC, *drawItem\rcItem\left, *drawItem\rcItem\top, disWidth, disHeight, memDC, *drawItem\rcItem\left, *drawItem\rcItem\top, #SRCCOPY)
  EndIf
  ;
  ; Clean up:
  SelectObject_(memDC, oldBitmap)
  DeleteObject_(hBitmap)
  DeleteDC_(memDC)
  ;
EndProcedure
;
Procedure SGCE_InstallCallbackAndData(hWnd, *procAddr, *GadgetsInfos, ClassName$ = "")
  If ClassName$
    Protected WclassName$ = Space(256)
    GetClassName_(hWnd, @WclassName$, 255)
    If LCase(WclassName$) <> LCase(ClassName$)
      ProcedureReturn #False
    EndIf
  EndIf
  Protected SGCE_OldCallBack = GetProp_(hWnd, "SGCE_OldCallBack")
  If SGCE_OldCallBack = 0
    SGCE_OldCallBack = SetWindowLongPtr_(hWnd, #GWL_WNDPROC, *procAddr)
    SetProp_(hWnd, "SGCE_OldCallBack"   , SGCE_OldCallBack)
    SetProp_(hWnd, "SGCE_ActualCallBack", *procAddr)
    If *GadgetsInfos
      SetWindowLongPtr_(hWnd, #GWL_USERDATA, *GadgetsInfos)
    EndIf
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure SGCE_UnUnstallCallback(hWnd, ClassName$ = "")
  If hWnd
    If ClassName$
      Protected WclassName$ = Space(256)
      GetClassName_(hWnd, @WclassName$, 255)
      If LCase(WclassName$) <> LCase(ClassName$)
        ProcedureReturn #False
      EndIf
    EndIf
    Protected SGCE_OldCallBack    = GetProp_(hWnd, "SGCE_OldCallBack")
    Protected SGCE_ActualCallBack = GetProp_(hWnd, "SGCE_ActualCallBack")
    If SGCE_OldCallBack And GetWindowLongPtr_(hWnd, #GWL_WNDPROC) = SGCE_ActualCallBack
      ; Reset to standard management:
      SetWindowLongPtr_(hWnd, #GWL_WNDPROC, SGCE_OldCallBack)
      RemoveProp_(hWnd, "SGCE_OldCallBack")
      RemoveProp_(hWnd, "SGCE_ActualCallBack")
      ProcedureReturn #True
    EndIf
  EndIf
EndProcedure
;
Procedure SGCE_EnumThreadAndInstallCallback(hWnd, *GadgetsInfos.SGCE_GadgetsInfosStruct)
  ;
  Protected hCal, gRect.Rect
  Protected className$ = Space(256)
  ;
  GetClassName_(hWnd, @className$, 255)
  If LCase(className$) = "dropdown"
    ; The user has clicked onto the right arrow of a DateGadget,
    ; and a DropDown window has just been opened.
    ; Install a callback for the DropDown window
    ; in order to repaint its edges:
    If SGCE_InstallCallbackAndData(hWnd, @SGCE_DropDownCallback(), *GadgetsInfos)
      ; Retreive the handle of the calendar included into the DropDown
      hCal = GetWindow_(hWnd, #GW_CHILD)
      ; At this place, you can eventually turn the Calendar to old appearance:
      ;    SetWindowTheme_(hCal, "", "")
      ; and set colors for title, for exemple.
;       PostMessage_(hCal, #MCM_SETCOLOR, #MCSC_BACKGROUND,  *GadgetsInfos\BackColor)
;       PostMessage_(hCal, #MCM_SETCOLOR, #MCSC_MONTHBK, *GadgetsInfos\BackColor)
;       PostMessage_(hCal, #MCM_SETCOLOR, #MCSC_TITLEBK, *GadgetsInfos\BackColor)
;       etc..........
      ;
      ; If you do that you should NOT call SGCE_InstallCallbackAndData()
      ; which install a callback for the Calendar control
      ; in order to repaint its colors:
      SGCE_InstallCallbackAndData(hCal, @SGCE_ChildrenCallback(), *GadgetsInfos)
      ;
      ; Resize the DropDown to fit exactly the Calendar size:
      SendMessage_(hCal, #MCM_GETMINREQRECT, 0, @gRect.Rect)
      SetWindowPos_(hWnd, 0, 0, 0, gRect\right - gRect\left, gRect\bottom - gRect\top + 6, #SWP_NOMOVE | #SWP_NOZORDER)
      ;
      *GadgetsInfos\Handle = hWnd
    EndIf
  EndIf
  ProcedureReturn #True
EndProcedure
;
Procedure SGCE_RedrawChildWindows(hWnd, lParam)
  If GetProperty(hWnd, "PB_ID") <> -1
    InvalidateRect_(hWnd, 0, #True)
  EndIf
  ProcedureReturn #True
EndProcedure
;
Procedure SGCE_ChangeColorsCallback(gHandle, uMsg, wParam, lParam)
  ;
  Protected hCTRLWnd, BackColor, BackLuminosity, Shift
  Protected *GadgetsInfos.SGCE_GadgetsInfosStruct = GetWindowLongPtr_(gHandle, #GWL_USERDATA)
  Protected SGCE_OldCallBack = GetProp_(gHandle, "SGCE_OldCallBack")
  Static    LastHoveredGadget
  Shared    SGCE_KillCalendarAnimation
  ;
  Select uMsg
    Case #WM_CAPTURECHANGED
      If IsGadget(*GadgetsInfos\ID) And (GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Date)
        ; This will install a callback procedure for the DropDown window
        ; opened when the user click on the right arrow of a DateGadget:
        EnumThreadWindows_(GetCurrentThreadId_(), @SGCE_EnumThreadAndInstallCallback(), *GadgetsInfos)
      EndIf
      ;
    Case #WM_NOTIFY
      If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Date
        Protected *nmhdr.NMHDR = lParam
        If *nmhdr\code = #NM_RELEASEDCAPTURE
          ; This message occur when something is clicked inside the calendar of a DropDown window.
          ; A DropDown window is opened by a DateGadget.
          SGCE_KillCalendarAnimation = #True
        EndIf
      EndIf
      ;
    Case #WM_LBUTTONUP
      If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Calendar
        ; Something has been clicked inside a CalendarGadget.
        SGCE_KillCalendarAnimation = #True
      EndIf
      ;
    Case 4097
      If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Date Or GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Calendar
        ; Private message from calendar:
        ; Something has been clicked into the calendar.
        SGCE_KillCalendarAnimation = #True
      EndIf
      ;
    Case #WM_DRAWITEM
      If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_ExplorerCombo
        ; Paint ExplorerComboGadget's items:
        SGCE_DrawExplorerComboItem(SGCE_OldCallBack, gHandle, uMsg, wParam, lparam, *GadgetsInfos)
        ;        
        ProcedureReturn 0
      EndIf
      ;
    Case #WM_SETTEXT, #WM_ENABLE
      ; The text of the gadget has been modified or the gadget has been enabled.
      ; Force a repaint:
      InvalidateRect_(gHandle, 0, #True)
      ;
    Case #WM_MOUSEMOVE, #WM_NCMOUSEMOVE
      CallWindowProc_(SGCE_OldCallBack, gHandle, uMsg, wParam, lParam)
      If LastHoveredGadget <> gHandle
        ; Mouse is over the gadget --> Highlight the edges:
        If *GadgetsInfos\EdgesStyle & #SGCE_HoverSensitive
          SGCE_RepaintEdges(*GadgetsInfos, #True)
          LastHoveredGadget = gHandle
          ;
          ; Activate #WM_MOUSELEAVE
          Protected tme.TRACKMOUSEEVENT
          tme\cbSize = SizeOf(TRACKMOUSEEVENT)
          tme\hwndTrack = gHandle
          tme\dwFlags = #TME_LEAVE
          If uMsg = #WM_NCMOUSEMOVE
            tme\dwFlags | #TME_NONCLIENT
          EndIf
          TrackMouseEvent_(@tme)
        EndIf
      EndIf
      ;
    Case #WM_MOUSELEAVE, #WM_NCMOUSELEAVE
      If LastHoveredGadget
        LastHoveredGadget = 0
        ; Repaint the borders to their normal color.
        If *GadgetsInfos\EdgesStyle & #SGCE_HoverSensitive
          SGCE_RepaintEdges(*GadgetsInfos, #False)
        EndIf
      EndIf
      ;
    Case #WM_CTLCOLORLISTBOX, #WM_CTLCOLOREDIT
      ; Colorize the lists of ComboGadgets:
      SetBkColor_(wParam, *GadgetsInfos\BackColor)
      SetTextColor_(wParam, *GadgetsInfos\FrontColor)
      ProcedureReturn CreateSolidBrush_(*GadgetsInfos\BackColor)
      ;
    Case #WM_NCPAINT
      ; Draw the gadget Edges.
      ;
      CallWindowProc_(SGCE_OldCallBack, gHandle, uMsg, wParam, lParam) ; <-- needed by some gadgets (as listview) to repaint scrollbars.
      ;
      If SGCE_IsCustomColorGadget(*GadgetsInfos\ID) = -1
        If LastHoveredGadget = gHandle
          SGCE_RepaintEdges(*GadgetsInfos, #True)
        Else
          SGCE_RepaintEdges(*GadgetsInfos, #False)
        EndIf
      EndIf
      ProcedureReturn 0
      ;
    Case #WM_PAINT
      ;
      If GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Spin
        ; SpinGadget colors are managed by PureBasic.
        ; All we have to do is to install a callback for the UP-DOWN control:
        hCTRLWnd = GetWindow_(gHandle, #GW_HWNDNEXT) ; For SpinGadget
        SGCE_InstallCallbackAndData(hCTRLWnd, @SGCE_ChildrenCallback(), *GadgetsInfos, #UPDOWN_CLASS)
        InvalidateRect_(hCTRLWnd, 0, #True)
        ;
        ProcedureReturn CallWindowProc_(SGCE_OldCallBack, gHandle, uMsg, wParam, lParam)
        ;
      ElseIf SGCE_IsCustomColorGadget(*GadgetsInfos\ID) = -1
        ; Nothing to do here. The edges are painted in #WM_NCPAINT.
        ProcedureReturn CallWindowProc_(SGCE_OldCallBack, gHandle, uMsg, wParam, lParam)
        ;
      ElseIf GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Frame
        ; #PB_Frame_Container needs a special handling because it can be
        ; transparent or not and will print a title or not regarding its creating flags.
        SGCE_DrawFrameGadget(*GadgetsInfos)
        ProcedureReturn 0
        ;
      ElseIf GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_Container
        ; #PB_GadgetType_Container is not managed by the system as other gadgets.
        ; Edges are repainted in #WM_PAINT but with Parent hDC
        SGCE_DrawContainerGadget(*GadgetsInfos)
        EnumChildWindows_(gHandle, @SGCE_RedrawChildWindows(), 0)
        ProcedureReturn 0
        ;
      Else
        BackColor = *GadgetsInfos\BackColor
        If (GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_ListIcon Or GadgetType(*GadgetsInfos\ID) = #PB_GadgetType_ExplorerList) And FindWindowEx_(gHandle, 0, #WC_HEADER, #Null) = 0
          ; Alter the header's color of ListIconGadget and ExplorerListGadget:
          BackLuminosity = Red(BackColor)*0.299 + Green(BackColor)*0.587 + Blue(BackColor)*0.114
          Shift = 15
          If BackLuminosity > 128 : Shift * -1 : EndIf
          BackColor = RGB(ShiftColorComponent(Red(BackColor), Shift), ShiftColorComponent(Green(BackColor), Shift), ShiftColorComponent(Blue(BackColor), Shift))
        EndIf
        ;
        SGCE_DrawAndReplaceColors(gHandle, lParam, SGCE_OldCallBack, BackColor, *GadgetsInfos\FrontColor, GadgetType(*GadgetsInfos\ID))
        ;
        If *GadgetsInfos\Enabled <> IsWindowEnabled_(gHandle)
          ; Gadget has been enabled or desabled.
          ; Force a second repaint after this one:
          *GadgetsInfos\Enabled = IsWindowEnabled_(gHandle)
          InvalidateRect_(gHandle, 0, #True)
        EndIf
        ;
        If *GadgetsInfos\EdgesStyle & #SGCE_HoverSensitive
          SGCE_RepaintEdges(*GadgetsInfos, LastHoveredGadget)
        EndIf
        ;
      EndIf
      ;
      hCTRLWnd = FindWindowEx_(gHandle, 0, #UPDOWN_CLASS, #Null) ; For PanelGadget
      If hCTRLWnd
        SGCE_InstallCallbackAndData(hCTRLWnd, @SGCE_ChildrenCallback(), *GadgetsInfos)
        InvalidateRect_(hCTRLWnd, 0, #True)
      EndIf
      ;
      hCTRLWnd = FindWindowEx_(gHandle, 0, #WC_HEADER, #Null) ; For ListIcon and ExplorerList gadgets
      If hCTRLWnd
        SGCE_InstallCallbackAndData(hCTRLWnd, @SGCE_ChangeColorsCallback(), *GadgetsInfos)
        InvalidateRect_(hCTRLWnd, 0, #True)
      EndIf 
      ;
      ProcedureReturn 0
      ;
  EndSelect
  ;
  ; Normal callback for all other messages:
  ProcedureReturn CallWindowProc_(SGCE_OldCallBack, gHandle, uMsg, wParam, lParam)
  ;
EndProcedure
;
Procedure SGCE_MainWindowCallbackCleaner(hWnd, uMsg, wParam, lParam)
  ;
  Shared SGCE_GadgetsInfos()
  Protected *pnmh.NMHDR
  ;
  If uMsg = #WM_DESTROY
    ; Clean the memory when the main window is destroyed:
    ForEach SGCE_GadgetsInfos()
      If Not (IsGadget(SGCE_GadgetsInfos()\ID)) Or WindowID(SGCE_GadgetsInfos()\MainWindow) = hWnd
        DeleteElement(SGCE_GadgetsInfos())
      EndIf
    Next
  EndIf
  ;
  ; Normal callback:
  Protected SGCE_OldCallBack = GetProp_(hWnd, "SGCE_OldCallBack")
  ProcedureReturn CallWindowProc_(SGCE_OldCallBack, hWnd, uMsg, wParam, lParam)
  ;
EndProcedure
;
; ****************************************************************************************
;
;-                 3--- THIRD PART: MAIN FUNCTIONS OF THIS LIBRARY ---
;
; ****************************************************************************************
;
Procedure GetGadgetColorEx(GadgetID, ColorType)
  ;
  ; Extended GetGadgetColor function for buttons and other
  ; gadgets with wich GetGadgetColor doesn't work.
  ;
  ;
  Shared SGCE_GadgetsInfos()
  Protected Result = GetGadgetColor(GadgetID, ColorType)
  ;
  If IsGadget(GadgetID) And SGCE_IsCustomColorGadget(GadgetID) And ListSize(SGCE_GadgetsInfos())
    ;
    Protected *ActualElement = @SGCE_GadgetsInfos()
    ;
    ForEach SGCE_GadgetsInfos()
      If SGCE_GadgetsInfos()\Handle = GadgetID(GadgetID)
        If ColorType = #PB_Gadget_BackColor
          Result = SGCE_GadgetsInfos()\BackColor
        ElseIf ColorType = #PB_Gadget_FrontColor
          Result = SGCE_GadgetsInfos()\FrontColor
        EndIf
        Break
      EndIf
    Next
    ;
    ChangeCurrentElement(SGCE_GadgetsInfos(), *ActualElement)
    ;
  EndIf
  ProcedureReturn Result
EndProcedure
;
Procedure SetGadgetColorEx(GadgetID, ColorType, Color = 0, ThisColorOnly = 0)
  ;
  ; Extended SetGadgetColor function for all gadgets.
  ;
  ; If parameter 'ColorType' is set to '#SGCE_StandardColors',
  ; standard colors are restored for the gadget.
  ;
  ; The 'Color' parameter can be an RGB() value or can be '#PB_Default'.
  ; The 'ColorType' parameter can be #PB_Gadget_BackColor or #PB_Gadget_FrontColor.
  ;
  ; If 'ThisColorOnly' is set to zero and if #PB_Gadget_FrontColor is not allready set,
  ; the FrontColor will be adjusted in opposition from the BackColor.
  ;
  ; If 'ThisColorOnly' is set to zero and if #PB_Gadget_BackColor is not allready set,
  ; the BackColor with be set to the same color as the window or contener background color.
  ;
  ; If 'ThisColorOnly' is set to 1, only the color specified by ColorType
  ; (#PB_Gadget_BackColor or #PB_Gadget_FrontColor) is adjusted.
  ;
  ;
  Shared SGCE_GadgetsInfos()
  ;
  Protected BackLuminosity, ParentWindow, Found, hCTRLWnd, hWnd, gRect.Rect
  Protected FrontColor, hParent, ApplyDarkMode, Red, Green, Blue
  ;
  If IsGadget(GadgetID)
    ;
    ParentWindow = GetGadgetParentWindow(GadgetID)
    If IsWindow(ParentWindow)
      ; Set a callback procedure for the main window, in order to clean
      ; the memory when the window is closed:
      SGCE_InstallCallbackAndData(WindowID(ParentWindow), @SGCE_MainWindowCallbackCleaner(), 0)
    EndIf
    ;
    ; Check is a SGCE_GadgetsInfos() element allready exists for the gadget:
    Found = 0
    If ListSize(SGCE_GadgetsInfos())
      ForEach SGCE_GadgetsInfos()
        If SGCE_GadgetsInfos()\ID = GadgetID
          Found = 1
          Break
        EndIf
      Next
    EndIf
    ;
    SetStandardColor: 
    If ColorType = #SGCE_StandardColors
      If Found
        ;
        ; Clean memory --> Clean the element and its child element and their controls (if any):
        ForEach SGCE_GadgetsInfos()
          If SGCE_GadgetsInfos()\ID = GadgetID
            hCTRLWnd = FindWindowEx_(GadgetID(GadgetID), 0, #UPDOWN_CLASS, #Null) ; For PanelGadget
            SGCE_UnUnstallCallback(hCTRLWnd)
            ;
            hCTRLWnd = FindWindowEx_(GadgetID(GadgetID), 0, #WC_HEADER, #Null)    ; For ListIcon and ExplorerList
            SGCE_UnUnstallCallback(hCTRLWnd)
            ;
            hCTRLWnd = GetWindow_(GadgetID(GadgetID), #GW_HWNDNEXT)               ; For SpinGadget
            SGCE_UnUnstallCallback(hCTRLWnd, #UPDOWN_CLASS)
            ;
            SGCE_UnUnstallCallback(GadgetID(GadgetID))
            ;
            DeleteElement(SGCE_GadgetsInfos())
          EndIf
        Next
        ;
        ; Force the gadget to be redrawn:
        If IsWindowVisible_(GadgetID(GadgetID))
          HideGadget(GadgetID, #True)
          HideGadget(GadgetID, #False)
        EndIf
      EndIf
      SetGadgetColor(GadgetID, #PB_Gadget_FrontColor, #PB_Default)
      SetGadgetColor(GadgetID, #PB_Gadget_BackColor,  #PB_Default)
      ;
      ProcedureReturn
      ;
    EndIf
    ;
    If Found = 0
      Found = 1
      AddElement(SGCE_GadgetsInfos())
      SGCE_GadgetsInfos()\ID = GadgetID
      SGCE_GadgetsInfos()\Handle = GadgetID(GadgetID)
      SGCE_GadgetsInfos()\FrontColor = #PB_Default
      SGCE_GadgetsInfos()\BackColor = #PB_Default
      SGCE_GadgetsInfos()\Enabled = IsWindowEnabled_(GadgetID(GadgetID))
      SGCE_GadgetsInfos()\MainWindow = ParentWindow
      SGCE_GadgetsInfos()\EdgesStyle = 0
    EndIf
    ;
    If ColorType = #PB_Gadget_BackColor
      ;
      SGCE_GadgetsInfos()\BackColor = Color
      ;
      ; From BackColor, compute an automatic color for FrontColor:
      ;
      If ThisColorOnly = 0 And SGCE_GadgetsInfos()\FrontColor = #PB_Default And Color <> #PB_Default
        ; If gadget's front color is not allready set,
        ; setup automatic FrontColor from BackColor:
        BackLuminosity = Red(Color)*0.299 + Green(Color)*0.587 + Blue(Color)*0.114
        If BackLuminosity < 128 And BackLuminosity >80
          ; Gadget color is quite middle grey
          FrontColor = #White
        ElseIf BackLuminosity > 128
          ; Bright theme
          FrontColor = 0
        Else
          ; Dark theme
          FrontColor = RGB(220, 220, 220)
        EndIf
        SGCE_GadgetsInfos()\FrontColor = FrontColor
      EndIf
      ;
    ElseIf ColorType = #PB_Gadget_FrontColor
      ;
      SGCE_GadgetsInfos()\FrontColor = Color
      ;
      If ThisColorOnly = 0 And SGCE_GadgetsInfos()\BackColor = #PB_Default And Color <> #PB_Default
        ; If gadget's back color is not allready set,
        ; use the ParentWindow back color:
        ;
        hParent = GetAncestor_(GadgetID(GadgetID), #GA_PARENT) ; Get the parent window/gadget handle
        ;
        If hParent
          If IsWindow(ParentWindow) And hParent = WindowID(ParentWindow)
            SGCE_GadgetsInfos()\BackColor = GetWindowColor(ParentWindow)
          Else
            Protected ParentPBID = GetProp_(hParent, "PB_ID")
            If IsGadget(ParentPBID)
              SGCE_GadgetsInfos()\BackColor = GetGadgetColorEx(ParentPBID, #PB_Gadget_BackColor)
            EndIf
          EndIf
        EndIf
      EndIf
    EndIf
    ;
    ApplyDarkMode = 0
    Select GadgetType(GadgetID)
      Case #PB_GadgetType_ScrollArea, #PB_GadgetType_ScrollBar, #PB_GadgetType_Editor, #PB_GadgetType_ListIcon, #PB_GadgetType_ListView
        ApplyDarkMode = 1
      Case #PB_GadgetType_ListView, #PB_GadgetType_ComboBox, #PB_GadgetType_ExplorerCombo, #PB_GadgetType_ExplorerList
        ApplyDarkMode = 1
      Case #PB_GadgetType_ExplorerTree, #PB_GadgetType_Container, #PB_GadgetType_Web, #PB_GadgetType_WebView, #PB_GadgetType_Scintilla
        ApplyDarkMode = 1
      Case #PB_GadgetType_Tree, #PB_GadgetType_MDI
        ApplyDarkMode = 1
    EndSelect
    ;
    If SGCE_GadgetsInfos()\FrontColor = #PB_Default
      Protected RealFrontColor = GetSysColor_(#COLOR_WINDOWTEXT)
    Else
      RealFrontColor = SGCE_GadgetsInfos()\FrontColor
    EndIf
    If SGCE_GadgetsInfos()\BackColor = #PB_Default
      Protected RealBackColor = GetSysColor_(#COLOR_BTNFACE)
    Else
      RealBackColor = SGCE_GadgetsInfos()\BackColor
    EndIf
    ;
    If ApplyDarkMode
      ; If the background is darker than lighter, set the scrollbar to dark mode:
      If Red(RealBackColor)*0.299 + Green(RealBackColor)*0.587 + Blue(RealBackColor)*0.114 > 128
        SetWindowTheme_(GadgetID(GadgetID), @"Explorer", #Null)
      Else
        SetWindowTheme_(GadgetID(GadgetID), @"DarkMode_Explorer", #Null)
      EndIf
    EndIf
    ;
    If (SGCE_GadgetsInfos()\FrontColor = #PB_Default Or SGCE_GadgetsInfos()\FrontColor = GetSysColor_(#COLOR_WINDOWTEXT)) And (SGCE_GadgetsInfos()\BackColor = #PB_Default Or SGCE_GadgetsInfos()\BackColor = GetSysColor_(#COLOR_BTNFACE))
      ; Both BackColor and FrontColor are set to #PB_Default.
      ColorType = #SGCE_StandardColors
      Goto SetStandardColor
    EndIf
    ;
    ;
    Protected CustomColor = SGCE_IsCustomColorGadget(GadgetID)
    ;
    If CustomColor
      ;
      SGCE_GadgetsInfos()\FrontColor = RealFrontColor
      SGCE_GadgetsInfos()\BackColor = RealBackColor
      ;
      ; Compute the color of the Edges, for some gadgets, as EditorGadget, which need to repaint the Edges:
      If CustomColor = -1 Or GadgetType(GadgetID) = #PB_GadgetType_ListIcon Or GadgetType(GadgetID) = #PB_GadgetType_Calendar Or GadgetType(GadgetID) = #PB_GadgetType_ExplorerList
        If GetWindowLongPtr_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_CLIENTEDGE Or GetWindowLongPtr_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_WINDOWEDGE Or GetWindowLongPtr_(GadgetID(GadgetID), #GWL_STYLE) & #WS_BORDER Or GadgetType(GadgetID) = #PB_GadgetType_Editor
          SGCE_GadgetsInfos()\EdgesStyle = #SGCE_HoverSensitive | #SGCE_Single
        EndIf
        If GadgetType(GadgetID) = #PB_GadgetType_ScrollArea Or GadgetType(GadgetID) = #PB_GadgetType_Canvas
          If GetWindowLong_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_STATICEDGE
            SGCE_GadgetsInfos()\EdgesStyle = #SGCE_HoverSensitive | #SGCE_Single
          ElseIf GetWindowLong_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_WINDOWEDGE
            SGCE_GadgetsInfos()\EdgesStyle = #SGCE_HoverSensitive | #SGCE_Double
          ElseIf GetWindowLong_(GadgetID(GadgetID), #GWL_STYLE) & #WS_BORDER
            SGCE_GadgetsInfos()\EdgesStyle = #SGCE_HoverSensitive | #SGCE_Flat
          EndIf
        EndIf
        If GadgetType(GadgetID) = #PB_GadgetType_Calendar
          ; Adjust the gadget size, because CalendarGadget has a fixed size.
          SendMessage_(GadgetID(GadgetID), #MCM_GETMINREQRECT, 0, @gRect.Rect)
          SetWindowPos_(GadgetID(GadgetID), 0, 0, 0, gRect\right - gRect\left - 2, gRect\bottom - gRect\top - 2, #SWP_NOMOVE | #SWP_NOZORDER)
        EndIf
        ;
      ElseIf GadgetType(GadgetID) = #PB_GadgetType_Date
        SGCE_GadgetsInfos()\EdgesStyle = #SGCE_Single
        ;
      ElseIf GadgetType(GadgetID) = #PB_GadgetType_Frame
        If GetWindowLong_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_CLIENTEDGE
          SGCE_GadgetsInfos()\EdgesStyle | #SGCE_Double
        ElseIf GetWindowLong_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_STATICEDGE
          SGCE_GadgetsInfos()\EdgesStyle | #SGCE_Single
        ElseIf Not(GetWindowLong_(GadgetID(GadgetID), #GWL_STYLE) & #SS_ETCHEDFRAME)
          SGCE_GadgetsInfos()\EdgesStyle | #SGCE_Flat
        Else
          SGCE_GadgetsInfos()\EdgesStyle | #SGCE_WithTitle
        EndIf
        ;
      ElseIf GadgetType(GadgetID) = #PB_GadgetType_Container
        If GetWindowLong_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_CLIENTEDGE
          SGCE_GadgetsInfos()\EdgesStyle | #SGCE_Double
        ElseIf GetWindowLong_(GadgetID(GadgetID), #GWL_EXSTYLE) & #WS_EX_STATICEDGE
          SGCE_GadgetsInfos()\EdgesStyle | #SGCE_Single
        ElseIf GetWindowLong_(GadgetID(GadgetID), #GWL_STYLE) & #WS_BORDER
          SGCE_GadgetsInfos()\EdgesStyle | #SGCE_Flat
        EndIf
      EndIf
      SGCE_GadgetsInfos()\EdgesColor =  MixColors(SGCE_GadgetsInfos()\FrontColor, SGCE_GadgetsInfos()\BackColor, 0.25)
      SGCE_GadgetsInfos()\HighlightedEdgesColor = MixColors(SGCE_GadgetsInfos()\FrontColor, SGCE_GadgetsInfos()\BackColor, 0.6)
      ;
      ; Set the gadget callback if not already done:
      ;
      SGCE_InstallCallbackAndData(GadgetID(GadgetID), @SGCE_ChangeColorsCallback(), @SGCE_GadgetsInfos())
      ;
    EndIf
    If CustomColor <> 1
      If ColorType = #PB_Gadget_FrontColor
        SetGadgetColor(GadgetID, #PB_Gadget_FrontColor, SGCE_GadgetsInfos()\FrontColor)
      ElseIf ColorType = #PB_Gadget_BackColor
        SetGadgetColor(GadgetID, #PB_Gadget_BackColor, SGCE_GadgetsInfos()\BackColor)
      Else
        SetGadgetColor(GadgetID, ColorType, Color)
      EndIf
    EndIf
  EndIf
EndProcedure
;
Structure SGCE_ColorForEnum
  ColorType.i
  Color.i
EndStructure
;
Procedure SGCE_EnumGadgetsAndSetColor(hGadget, *ColorForEnum.SGCE_ColorForEnum)
  ;
  ; Retreive the PureBasic gadget ID from the gadget handle:
  Protected GadgetPBID = GetProperty(hGadget, "PB_ID")
  If GadgetPBID >= 0
    ; Apply the color if it is a valid ID:
    SetGadgetColorEx(GadgetPBID, *ColorForEnum\ColorType, *ColorForEnum\Color)
  EndIf
  ProcedureReturn #True
EndProcedure
;
Procedure ApplyColorsToAllGadgets(WindowNum, FrontColor, BackColor)
  ;
  ; Enumerate the gadgets of the 'WindowNum' window and apply the colors to them.
  ;
  Protected ColorForEnum.SGCE_ColorForEnum
  ;
  SetWindowColor(WindowNum, BackColor)
  ;
  Protected hWnd = WindowID(WindowNum)
  ;
  ColorForEnum\Color = BackColor
  ColorForEnum\ColorType = #PB_Gadget_BackColor
  EnumChildWindows_(hWnd, @SGCE_EnumGadgetsAndSetColor(), @ColorForEnum)
  ColorForEnum\Color = FrontColor
  ColorForEnum\ColorType = #PB_Gadget_FrontColor
  EnumChildWindows_(hWnd, @SGCE_EnumGadgetsAndSetColor(), @ColorForEnum)
  ;
EndProcedure
;
;
Macro GetGadgetColor(a, b)
	GetGadgetColorEx(a, b)
EndMacro

Macro SetGadgetColor(a, b, c)
	SetGadgetColorEx(a, b, c)
EndMacro
;
;
; ****************************************************************************************
;
;-                         4--- FORTH PART: DEMO PROCEDURE ---
;
; ****************************************************************************************
;
;
CompilerIf #PB_Compiler_IsMainFile
  ; The following won't run when this file is used as 'Included'.
  ;
  Macro ColorizeGadgets()
    ApplyColorsToAllGadgets(#SGCE_DemoWindow, FrontColor, BackColor)
    ;
    ; Then customize colors for some gadgets:
    ;
    ; Lighten the background of the frame contener and its buttons:
    If BackColor = -1
      LightenBackground = $D0D0D0
    Else
      If Red(BackColor)*0.299 + Green(BackColor)*0.587 + Blue(BackColor)*0.114 > 128
        LightenBackground = RGB(Red(BackColor) - 20, Green(BackColor) - 20, Blue(BackColor) - 20)
      Else
        LightenBackground = RGB(Red(BackColor) + 20, Green(BackColor) + 20, Blue(BackColor) + 20)
      EndIf
      SetGadgetColor(#SGCE_Frame, #PB_Gadget_BackColor, LightenBackground)
      SetGadgetColor(#SGCE_Option1, #PB_Gadget_BackColor, LightenBackground)
      SetGadgetColor(#SGCE_Option2, #PB_Gadget_BackColor, LightenBackground)
      ; Set fantasy colors for OptionGadgets:
      SetGadgetColor(#SGCE_Option1, #PB_Gadget_FrontColor, RGB($FF, $50, $FF))
      If Red(BackColor)*0.299 + Green(BackColor)*0.587 + Blue(BackColor)*0.114 > 128
        SetGadgetColor(#SGCE_Option2, #PB_Gadget_FrontColor, $008080)
      Else
        SetGadgetColor(#SGCE_Option2, #PB_Gadget_FrontColor, $00FFE0)
      EndIf
    EndIf
    ;
    ; Lighten the background of the right contener:
    SetGadgetColor(#SGCE_Contener, #PB_Gadget_BackColor, LightenBackground)
    ;
    ; Clear the colors of #SGCE_StandardButton (it will have default colors):
    SetGadgetColor(#SGCE_StandardButton, #PB_Gadget_BackColor, StandardBack)
    SetGadgetColor(#SGCE_StandardButton, #PB_Gadget_FrontColor, StandardFront)
    ; Set the colors of #LightBlueButton:
    SetGadgetColor(#SGCE_LightBlueButton, #PB_Gadget_BackColor, LightBlueBack)
    SetGadgetColor(#SGCE_LightBlueButton, #PB_Gadget_FrontColor, LightBlueFront)
    ; Set the colors of #SGCE_MatrixButton:
    SetGadgetColor(#SGCE_MatrixButton, #PB_Gadget_BackColor, MatrixBack)
    SetGadgetColor(#SGCE_MatrixButton, #PB_Gadget_FrontColor, MatrixFront)
    ; Set the colors of #SGCE_DarkGreyButton:
    SetGadgetColor(#SGCE_DarkGreyButton, #PB_Gadget_BackColor, DarkGreyBack)
    SetGadgetColor(#SGCE_DarkGreyButton, #PB_Gadget_FrontColor, DarkGreyFront)
    ;
  EndMacro
  ;
  Procedure MulticolorDemoWindow()
    ;
    Protected LightenBackground
    ;
    Enumeration WindowAndGadgetNum
      #SGCE_DemoWindow
      #SGCE_Frame
      #SGCE_Option1
      #SGCE_Option2
      #SGCE_CheckBox
      #SGCE_Button1
      #SGCE_Text
      #SGCE_String
      #SGCE_Editor
      #SGCE_ListView
      #SGCE_Panel
      #SGCE_Contener
      #SGCE_StandardButton
      #SGCE_LightBlueButton
      #SGCE_MatrixButton
      #SGCE_DarkGreyButton
      #SGCE_ListIcon
      #SGCE_Combo
      #SGCE_ComboEdit
      #SGCE_ExplorerCombo
      #SGCE_Date
      #SGCE_Calendar
      #SGCE_ExplorerList
      #SGCE_ExplorerTree
      #SGCE_Tree
      #SGCE_Spin
      #SGCE_MDI
      #SGCE_ChildWindow
      #SGCE_BQuit
    EndEnumeration
    ;
    ; To have a faster and cleaner color drawing, first create the window
    ; with the #PB_Window_Invisible attribute and make it visible just
    ; before the WindowEvent loop.
    ; (This is not mandatory)
    If OpenWindow(#SGCE_DemoWindow, 100, 100, 540, 310, "SetGadgetColorEx Demo", #PB_Window_ScreenCentered | #PB_Window_SystemMenu | #PB_Window_Invisible)
    ; If your computer is set with 'Dark Theme', the following will also
    ; switch the title bar of the window to dark theme:
    ApplyDarkModeToWindow(#SGCE_DemoWindow)
    ;
    ;
    ; Create two option buttons inside a frame
    FrameGadget(#SGCE_Frame, 10, 10, 120, 65, "FrameGadget",  #PB_Frame_Container)
    ;
      OptionGadget(#SGCE_Option1, 15, 15, 70, 25, "Option #1")
      SetGadgetState(#SGCE_Option1, 1)
      OptionGadget(#SGCE_Option2, 15, 35, 70, 25, "Option #2")
    CloseGadgetList()
    ;
    CheckBoxGadget(#SGCE_CheckBox, 10, 90, 120, 25, "CheckBoxGadget")
    SetGadgetState(#SGCE_CheckBox, 1)
    ButtonGadget(#SGCE_Button1, 20, 120, 100, 25, "Desabled Button")
    DisableGadget(#SGCE_Button1, #True)
    ;
    TextGadget(#PB_Any, 10, 160, 120, 25, "StringGadget")
    StringGadget(#SGCE_String, 10, 180, 120, 25, "StringGadget")
    ;
    TextGadget(#PB_Any, 10, 230, 120, 20, "EditorGadget")
    EditorGadget(#SGCE_Editor, 10, 250, 120, 50)
    SetGadgetText(#SGCE_Editor, "EditorGadget with a long text.")
    ;
    PanelWidth = 247
    PanelGadget(#SGCE_Panel, 140, 10, PanelWidth, WindowHeight(#SGCE_DemoWindow) - 19)
    ;
    AddGadgetItem(#SGCE_Panel, -1, "Combos")
      TextGadget(#PB_Any, 6, 20, PanelWidth - 20, 20, "ComboBoxGadget")
      ComboBoxGadget(#SGCE_Combo, 6, 40, PanelWidth - 20, 22)
      For ct = 1 To 5 : AddGadgetItem(#SGCE_Combo, -1, "ComboBox standard " + Str(ct)) : Next
      SetGadgetState(#SGCE_Combo, 0)
      ;
      TextGadget(#PB_Any, 6, 80, PanelWidth - 20, 20, "ComboBoxGadget - Editable")
      ComboBoxGadget(#SGCE_ComboEdit, 6, 100, PanelWidth - 20, 22, #PB_ComboBox_Editable)
      For ct = 1 To 5 : AddGadgetItem(#SGCE_ComboEdit, -1, "ComboBox Editable " + Str(ct)) : Next
      SetGadgetState(#SGCE_ComboEdit, 0)
      ;
      TextGadget(#PB_Any, 6, 160, PanelWidth - 20, 20, "ExplorerComboGadget")
      ExplorerComboGadget(#SGCE_ExplorerCombo, 6, 180, PanelWidth - 20, 22, "C:")
      ;
    AddGadgetItem(#SGCE_Panel, -1, "Lists")
      TextGadget(#PB_Any, 6, 10, PanelWidth - 20, 20, "ListIconGadget")
      ListIconGadget(#SGCE_ListIcon, 6, 30, PanelWidth - 20, 80, "ListIcon", 70)
      AddGadgetColumn(#SGCE_ListIcon, 1, "Column 2", 100)
      For ct = 1 To 10 : AddGadgetItem(#SGCE_ListIcon, -1, "Element " + Str(ct) + #LF$ + "Column 2 Element " + Str(ct))
      Next
      TextGadget(#PB_Any, 6, 130, PanelWidth - 20, 20, "ExplorerListGadget")
      ExplorerListGadget(#SGCE_ExplorerList, 6, 150, PanelWidth - 20, 100, "C:")
      ;
      AddGadgetItem(#SGCE_Panel, -1, "ListView")
      TextGadget(#PB_Any, 6, 30, PanelWidth - 20, 20, "ListViewGadget")
      ListViewGadget(#SGCE_ListView, 6, 50, PanelWidth - 20, 140)
      ;
      For ct = 1 To 20
        AddGadgetItem(#SGCE_ListView, ct - 1, "ListView Gadget " + Str(ct))
      Next
      ;
    AddGadgetItem(#SGCE_Panel, -1, "Date")
      TextGadget(#PB_Any, 6, 10, PanelWidth - 20, 20, "DateGadget")
      DateGadget(#SGCE_Date, 6, 30, PanelWidth - 20, 22)
      TextGadget(#PB_Any, 6, 70, PanelWidth - 20, 20, "CalendarGadget")
      CalendarGadget(#SGCE_Calendar, 6, 90, PanelWidth , 164)
    AddGadgetItem(#SGCE_Panel, -1, "Trees")
      TextGadget(#PB_Any, 6, 10, PanelWidth - 20, 20, "ExplorerTreeGadget")
      ExplorerTreeGadget(#SGCE_ExplorerTree, 6, 30, PanelWidth - 20, 95, "C:")
      TextGadget(#PB_Any, 6, 135, PanelWidth - 20, 20, "TreeGadget")
      TreeGadget(#SGCE_Tree, 6, 155, PanelWidth - 20, 100)
      For a = 0 To 10
        AddGadgetItem(#SGCE_Tree, -1, "Normal "+Str(a), 0, 0)
        AddGadgetItem(#SGCE_Tree, -1, "Node "+Str(a), 0, 0)
        AddGadgetItem(#SGCE_Tree, -1, "SubElement 1", 0, 1)
        AddGadgetItem(#SGCE_Tree, -1, "SubElement 2", 0, 1)
        AddGadgetItem(#SGCE_Tree, -1, "SubElement 3", 0, 1)
        AddGadgetItem(#SGCE_Tree, -1, "SubElement 4", 0, 1)
        AddGadgetItem(#SGCE_Tree, -1, "File "+Str(a), 0, 0)
      Next
    AddGadgetItem(#SGCE_Panel, -1, "Others")
      TextGadget(#PB_Any, 6, 10, PanelWidth - 20, 20, "SpinGadget")
      SpinGadget(#SGCE_Spin, 6, 30, PanelWidth - 20, 22, 0, 5, #PB_Spin_Numeric)
      SetGadgetText(#SGCE_Spin, "1")
      HyperLinkGadget(#PB_Any, 6, 70, PanelWidth - 20, 20, "HyperLinkGadget", $FFA020)
      
      TextGadget(#PB_Any, 6, 110, PanelWidth - 20, 20, "ScrollAreaGadget")
      ScrollAreaGadget(#PB_Any, 6, 130, PanelWidth - 20, 50, 220, 60)
      CloseGadgetList()
      ;
    CloseGadgetList()
    ;
    ContainerWidth = 130
    TextGadget(#PB_Any, WindowWidth(#SGCE_DemoWindow) - ContainerWidth - 10, 12, ContainerWidth, 70, "<-- Browse the panels to see how each gadget looks like when colorized.", #PB_Text_Center)
    ;
    TextGadget(#PB_Any, WindowWidth(#SGCE_DemoWindow) - ContainerWidth - 10, 100, ContainerWidth, 20, "Choose your color:", #PB_Text_Center)
    ContainerGadget(#SGCE_Contener, WindowWidth(#SGCE_DemoWindow) - ContainerWidth - 10, 120, ContainerWidth, 130)
      ;
      ButtonGadget(#SGCE_StandardButton,  10, 10,  ContainerWidth - 20, 22, "Standard")
      ButtonGadget(#SGCE_LightBlueButton, 10, 40,  ContainerWidth - 20, 22, "Light blue")
      ButtonGadget(#SGCE_MatrixButton,    10, 70,  ContainerWidth - 20, 22, "Matrix")
      ButtonGadget(#SGCE_DarkGreyButton,  10, 100, ContainerWidth - 20, 22, "Dark grey")
      CloseGadgetList()
      ;
    DisableGadget(#SGCE_MatrixButton, #True)
    ;     
    ButtonGadget(#SGCE_BQuit, WindowWidth(#SGCE_DemoWindow) - ContainerWidth - 10, WindowHeight(#SGCE_DemoWindow) - 35, 130, 25, "Exit")
    ;
    ;
    ; If you want to, you can attribute colors to your gadgets one by one.
    ; But if you want to achieve something like a color theme, which changes the colors
    ; of all the gadgets in your window (to get a dark mode, for example), there are two
    ; ways to do it:
    ;
    ; • You can include the 'ApplyColorThemes.pb' library in your project.
    ;   This library is available on the Zapman website:
    ;   https://www.editions-humanis.com/downloads/PureBasic/ZapmanDowloads_EN.htm
    ;   It allows the users to create, modify or simply choose an existing colors theme
    ;   and to apply it to the current program.
    ;
    ; • If you just need a quick method to apply a couple of colors to all of your gadgets,
    ;   you can simply call ApplyColorsToAllGadgets() as it is done by ColorizeGadgets() below.
    ;
    StandardBack   = #PB_Default
    StandardFront  = #PB_Default
    LightBlueBack  = $FFFFE0
    LightBlueFront = $A04000
    MatrixBack     = $282800
    MatrixFront    = $A0FF5A
    DarkGreyBack   = $202020
    DarkGreyFront  = $E0E0E0
    ;
    BackColor  = MatrixBack ; Matrix default setting
    FrontColor = MatrixFront; Matrix default setting
    ;
    ColorizeGadgets()
    ;
    ; If the window has been created with the #PB_Window_Invisible attribute,
    ; this is the moment to make it visible.
    HideWindow(#SGCE_DemoWindow, #False)
    ;
    Repeat
      Event = WaitWindowEvent()
      
      If Event = #PB_Event_Gadget
  
        Select EventGadget()
          Case #SGCE_CheckBox
            If GetGadgetState(#SGCE_CheckBox)
              DisableGadget(#SGCE_Button1, #True)
              SetGadgetText(#SGCE_Button1, "Desabled Button")
            Else
              DisableGadget(#SGCE_Button1, #False)
              SetGadgetText(#SGCE_Button1, "Enabled Button")
            EndIf
          Case #SGCE_StandardButton, #SGCE_LightBlueButton, #SGCE_MatrixButton, #SGCE_DarkGreyButton
            BackColor  = GetGadgetColor(EventGadget(), #PB_Gadget_BackColor)
            FrontColor = GetGadgetColor(EventGadget(), #PB_Gadget_FrontColor)
            ;
            DisableGadget(#SGCE_StandardButton, #False)
            DisableGadget(#SGCE_LightBlueButton, #False)
            DisableGadget(#SGCE_MatrixButton, #False)
            DisableGadget(#SGCE_DarkGreyButton, #False)
            ;
            DisableGadget(EventGadget(), #True)
            
            ColorizeGadgets()
            ;
          Case #SGCE_BQuit
            Break
        EndSelect
        ;
      ElseIf Event = #PB_Event_CloseWindow
        Break
      EndIf
    ForEver
    CloseWindow(#SGCE_DemoWindow)
  EndIf
  EndProcedure
  ;
  MulticolorDemoWindow()
  ;
CompilerEndIf
; IDE Options = PureBasic 6.20 (Windows - x64)
; CursorPosition = 5
; Folding = oq1vv--
; DPIAware