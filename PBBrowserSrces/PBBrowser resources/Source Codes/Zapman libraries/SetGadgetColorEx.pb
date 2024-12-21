;
; **************************************************************************
;
;                              SetGadgetColorEx()
;                                 Windows only
;                              Zapman - Dec 2024
;
;             The first part of this library offers two functions:
;    • IsDarkModeEnabled() allows to know if the running computer is set to dark mode or not.
;    • ApplyDarkModeToWindow() matches the title bar of a window with the computer theme.
;
;     The second part of this library extends the SetGadgetColor() PureBasic
;           native function to a large part of the PureBasic gadgets.
;    Instead of calling SetGadgetColor(), you can now call SetGadgetColorEx()
;            for buttons, pannels, combos, stringGadgets, etc.
;
;                   A third part contains a demo procedure.
;
; IMPORTANT NOTE: If you want to use SetGadgetColorEx to create color themes
; and apply them to your programs (to present them in 'Dark Mode', for example),
; you can use the 'ApplyColorThemes.pb' library also available on the Zapman website:
;   https://www.editions-humanis.com/downloads/PureBasic/ZapmanDowloads_EN.htm
;
;********************************************************************
;-     1--- FIRST PART: FUNCTION FOR COLORING WINDOW TITLE ---
;
; Prototype for the DwmSetWindowAttribute_ function
Prototype.i DwmSetWindowAttribute(hWnd.i, dwAttribute.i, pvAttribute.i, cbAttribute.i)
;
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
;
Procedure ApplyDarkModeToWindow(Window)
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
;
;
;********************************************************************
;-        2--- SECOND PART: FUNCTIONS FOR COLORING GADGETS ---
;
Structure SGCE_GadgetsColors
  Gadget_ID.i
  Gadget_Handle.i
  Gadget_FrontColor.i
  Gadget_BackColor.i
  Gadget_BorderColor.l
  Gadget_OverColor.l
  Gadget_OldREProc.i
  Gadget_Enabled.i
  Gadget_MainWindow.i
  Gadget_ChildOf.i
  Gadget_ParentOf.i
EndStructure
;
#SGCE_StandardColors = -1
;
NewList GadgetsColors.SGCE_GadgetsColors()
;
Structure pixelColor
  ; Field are 'a' type and not 'b' type,
  ; because 'a' is unsigned and 'b' is signed.
  red.a
  green.a
  blue.a
  alpha.a
EndStructure
;
Procedure SGCE_ReplaceColors(hdc, width, height, *backColor.pixelColor, *BorderColor.pixelColor)
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
  Protected *pixels, y, x, LowContrast.f
  Protected *pixelColor.tagRGBQUAD, PixelLuminosity.f, PixelDarkness.f
  Protected BlowOut.f, CrushBlacks.f
  ;
  ; Initialize a BITMAPINFO structure:
  Protected bmi.BITMAPINFO
  bmi\bmiHeader\biSize = SizeOf(BITMAPINFOHEADER)
  bmi\bmiHeader\biWidth = width
  bmi\bmiHeader\biHeight = height
  bmi\bmiHeader\biPlanes = 1
  bmi\bmiHeader\biBitCount = 32     ; 32 bits par pixel (RGBA).
  bmi\bmiHeader\biCompression = #BI_RGB
  ;
  ; Calculate the maximum brightness of a pixel:
  Protected maxlum = Red(GetSysColor_(#COLOR_BTNFACE)) + Green(GetSysColor_(#COLOR_BTNFACE)) + Blue(GetSysColor_(#COLOR_BTNFACE))
  ; LowContrast is calculated to be 1 when the contrast between the asked background color (BackColor)
  ; and the asked drawing color (BorderColor) is zero, and to be zero when this contrast is maximal:
  LowContrast = (*backColor\red + *backColor\green + *backColor\blue - *BorderColor\red - *BorderColor\green - *BorderColor\blue) / maxlum
  LowContrast = Abs(1 - Abs(LowContrast))
  ;
  ; Create a DIBSection and retrieve the pointer to the pixels:
  Protected hBitmap = CreateDIBSection_(hdc, bmi, #DIB_RGB_COLORS, @*pixels, 0, 0)
  ;
  ; Create a compatible memory context and copy the pixels to it:
  Protected memDC = CreateCompatibleDC_(hdc)
  ;
  Protected oldBitmap = SelectObject_(memDC, hBitmap)
  BitBlt_(memDC, 0, 0, width, height, hdc, 0, 0, #SRCCOPY)
  ;
  ; Create a compatible memory context and copy the pixels to it:
  For y = 0 To height - 1
    For x = 0 To width - 1
      ; Calculate the address of the current pixel in 32 bits (BGRA)
      *pixelColor = *pixels + (y * width + x) * 4

      PixelLuminosity = (*pixelColor\red + *pixelColor\green + *pixelColor\blue) / maxlum
      ;
      ; The following correction is intended to improve text readability when
      ; the contrast between the background color (*backColor) and the stroke color (*BorderColor)
      ; is medium or low.
      ;
      ; BlowOut brings high luminosities closer to white.
      ; (In photography, this is called "blowing out" the whites).
      If PixelLuminosity > 0.7
        BlowOut = (PixelLuminosity + 1) * (PixelLuminosity + 1) / 2.5
        If BlowOut > 1 : BlowOut = 1 : EndIf
        ; The lower the contrast between the background and stroke color,
        ; the more the BlowOut correction is applied.
        PixelLuminosity = PixelLuminosity * (1 - LowContrast) + BlowOut * LowContrast
      ElseIf PixelLuminosity < 0.3
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
      ; - The darker the pixel, the less its color is retained and the more it is replaced by *BorderColor:
      *pixelColor\red   * PixelLuminosity + *BorderColor\red   * PixelDarkness
      *pixelColor\green * PixelLuminosity + *BorderColor\green * PixelDarkness
      *pixelColor\blue  * PixelLuminosity + *BorderColor\blue  * PixelDarkness
      ;
    Next
  Next
  ;
  ; Copy the modified bitmap into the original HDC
  BitBlt_(hdc, 0, 0, width, height, memDC, 0, 0, #SRCCOPY)
  ;
  ; Release resources
  SelectObject_(memDC, oldBitmap)
  DeleteObject_(hBitmap)
  DeleteDC_(memDC)
EndProcedure
;
Procedure RepaintBordersInCallBack(gHandle, BorderColor, BackColor, BorderShift = 2, OuterBorderWidth = 0, InnerBorderWidth = 2)
  ;
  Protected windowRect.Rect, hdc, hRgn1, hRgn2, hBrush, PBorderWidth
  ;
  ; Get the position and full size of the gadget window (including borders and scrollbars)
  GetWindowRect_(gHandle, @windowRect.Rect)
  ;
  ; Include the gadget borders:
  windowRect\right - windowRect\left - BorderShift + OuterBorderWidth
  windowRect\bottom - windowRect\top - BorderShift + OuterBorderWidth
  windowRect\left = -BorderShift - OuterBorderWidth
  windowRect\top = -BorderShift - OuterBorderWidth
  ;
  hdc =  GetDC_(gHandle)
  ;
  ; Create a region to redraw only the borders
  hRgn1 = CreateRectRgn_(windowRect\left , windowRect\top, windowRect\right, windowRect\bottom)
  BorderWidth = OuterBorderWidth + InnerBorderWidth
  hRgn2 = CreateRectRgn_(windowRect\left + BorderWidth, windowRect\top + BorderWidth, windowRect\right - BorderWidth, windowRect\bottom - BorderWidth)
  CombineRgn_(hRgn1, hRgn1, hRgn2, #RGN_DIFF)
  DeleteObject_(hRgn2)

  ; Apply the clipping region
  SelectClipRgn_(hdc, hRgn1)
  DeleteObject_(hRgn1)

  ; Draw within the delimited area:
  hBrush = CreateSolidBrush_(BorderColor)
  FillRect_(hdc, windowRect, hBrush)
  DeleteObject_(hBrush)
  If BorderWidth > 1
    windowRect\left + 1
    windowRect\top  + 1
    windowRect\right  - 1
    windowRect\bottom - 1
    hBrush = CreateSolidBrush_(BackColor)
    FillRect_(hdc, windowRect, hBrush)
    DeleteObject_(hBrush)
  EndIf
  DeleteDC_(hdc)
EndProcedure
;
Procedure SGCE_ChangeColorsCallback(gHandle, uMsg, wParam, lParam)
  ;
  ; Variables for drawing:
  ;
  Protected hdc, memDC, oldBitmap
  Protected gadgetRect.Rect, width, height
  Protected ps.PAINTSTRUCT, hBitmap, retvalue
  ;
  Protected *GadgetsColors.SGCE_GadgetsColors = GetWindowLongPtr_(gHandle, #GWL_USERDATA)
  Protected OldREProc = *GadgetsColors\Gadget_OldREProc
  ;
  Select GadgetType(*GadgetsColors\Gadget_ID)
    Case  #PB_GadgetType_Editor, #PB_GadgetType_String, #PB_GadgetType_ListView, #PB_GadgetType_ButtonImage
      Protected RepaintBorder = 1
  EndSelect
  ;
  If uMsg = #WM_SETTEXT Or uMsg = #WM_ENABLE
    ; Force a repaint:
    InvalidateRect_(gHandle, 0, #True)
  EndIf
  ;
  If uMsg = #WM_MOUSEMOVE And GadgetType(*GadgetsColors\Gadget_ID) = #PB_GadgetType_ListView
    ProcedureReturn 1
  EndIf
  ;
  If uMsg = #WM_PAINT
    ;
    If RepaintBorder
      ;
      ; Call the default procedure to draw the gadget
      CallWindowProc_(OldREProc, gHandle, uMsg, wParam, lParam)
      ;
      RepaintBordersInCallBack(gHandle, *GadgetsColors\Gadget_BorderColor, *GadgetsColors\Gadget_BackColor)
      ;
      ProcedureReturn 1
    EndIf
    ;
    GetClientRect_(gHandle, @gadgetRect.Rect)
    ;
    hdc = BeginPaint_(gHandle, ps)
    width = gadgetRect\right - gadgetRect\left
    height = gadgetRect\bottom - gadgetRect\top
    ;
    ; Create a compatible DC to draw inside:
    memDC = CreateCompatibleDC_(hdc)
    ; 
    ; Create a bitmap to draw inside:
    hBitmap = CreateCompatibleBitmap_(hdc, width, height)
    ;
    oldBitmap = SelectObject_(memDC, hBitmap)
    ;
    ; Fill the bitmap with the background color:
    Rectangle_(memDC, gadgetRect\left - 1, gadgetRect\top - 1, gadgetRect\right + 1, gadgetRect\bottom + 1)
    ;
    ; Call the noram #WM_PAINT process to draw the gadget.
    ; This will draw inside memDC instead of hdc:
    retvalue = CallWindowProc_(OldREProc, gHandle, #WM_PAINT, memDC, lParam)
    ;
    ; Change image colors:
    ;
    SelectClipRgn_(memDC, 0)
    SGCE_ReplaceColors(memDC, width, height, @*GadgetsColors\Gadget_BackColor, @*GadgetsColors\Gadget_FrontColor)
    ;
    ; Copy memDC to hdc
    BitBlt_(hdc, gadgetRect\left, gadgetRect\top, width, height, memDC, 0, 0, #SRCCOPY)
    ;
    ; Clean up:
    SelectObject_(memDC, oldBitmap)
    DeleteObject_(hBitmap)
    DeleteDC_(memDC)
    ;           
    ; End the drawing process:
    EndPaint_(gHandle, ps)
    ;
    If *GadgetsColors\Gadget_Enabled <> IsWindowEnabled_(gHandle)
      ; Gadget has been enabled or desabled.
      ; Force a second repaint after this one:
      *GadgetsColors\Gadget_Enabled = IsWindowEnabled_(gHandle)
      InvalidateRect_(gHandle, 0, #True)
    EndIf
    ;
    If *GadgetsColors\Gadget_ParentOf
      ; Redraw child window if any:
      InvalidateRect_(*GadgetsColors\Gadget_ParentOf, 0, #True)
    EndIf
    ;
    ProcedureReturn 0
    ;
  EndIf
  ;
  ; Normal callback for all other messages::
  ProcedureReturn CallWindowProc_(OldREProc, gHandle, uMsg, wParam, lParam)
  ;
EndProcedure
;
Procedure SGCE_MainWindowCallbackCleaner(gHandle, uMsg, wParam, lParam)
  ;
  Shared GadgetsColors()
  ;
  Protected OldREProc = GetWindowLongPtr_(gHandle, #GWL_USERDATA)
  ;
  If uMsg = #WM_DESTROY
    ; Clean the memory when the main window is destroyed:
    ForEach GadgetsColors()
      If Not (IsGadget(GadgetsColors()\Gadget_ID)) Or GadgetsColors()\Gadget_MainWindow = gHandle
        If GadgetsColors()\Gadget_OldREProc
          SetWindowLongPtr_(GadgetsColors()\Gadget_Handle, #GWL_WNDPROC, GadgetsColors()\Gadget_OldREProc)
        EndIf
        DeleteElement(GadgetsColors())
      EndIf
    Next
  EndIf
  ;
  ProcedureReturn CallWindowProc_(OldREProc, gHandle, uMsg, wParam, lParam)
  ;
EndProcedure
;
Procedure GetGadgetWindow(Gadget)
  ;
  ; This function is by 'mk-soft', english forum.
  ;
  Protected ID, r1
  
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
;
Procedure SGCE_FindChildCallback(hWnd, val)
  ;
  Shared ChildHandle
  ChildHandle = hWnd
  ;
  Protected windowTitle.s = Space(256)
  ;
  GetWindowText_(hWnd, @windowTitle, 255)
  ;
  ProcedureReturn 0
EndProcedure
;
Procedure SGCE_EnumChildWindows(Gadget)
  ;
  Shared ChildHandle
  ;
  If IsGadget(Gadget)
    Gadget = GadgetID(Gadget)
  EndIf
  If Gadget
    EnumChildWindows_(Gadget, @SGCE_FindChildCallback(), 0)
  EndIf
  ;
  ProcedureReturn ChildHandle
EndProcedure
;
Procedure SGCE_IsCustomColorGadget(GadgetID)
  Protected CustomColor = 1
  Select GadgetType(GadgetID)
    ; All the gadgets listed here are natively managed by PureBasic and don't need any special process:
    Case #PB_GadgetType_Editor, #PB_GadgetType_String, #PB_GadgetType_ListView, #PB_GadgetType_ButtonImage
      CustomColor = -1
    Case #PB_GadgetType_Calendar, #PB_GadgetType_Container, #PB_GadgetType_Date, #PB_GadgetType_ExplorerList
      CustomColor = 0
    Case #PB_GadgetType_ExplorerTree, #PB_GadgetType_HyperLink, #PB_GadgetType_MDI;, #PB_GadgetType_ListIcon
      CustomColor = 0
    Case #PB_GadgetType_ProgressBar, #PB_GadgetType_ScrollArea, #PB_GadgetType_Spin, #PB_GadgetType_Text, #PB_GadgetType_Tree
      CustomColor = 0
  EndSelect
  ProcedureReturn CustomColor
EndProcedure
;
Procedure SetGadgetColorEx(GadgetID, ColorType, Color = 0, ThisColorOnly = 0)
  ;
  ; Extended SetGadgetColor function for all gadgets.
  ;
  ; If parameter 'ColorType' is set to '#SGCE_StandardColors',
  ; standard colors are restored for the gadget.
  ;
  ; The 'Color' parameter can be an RGB() number or can be '#PB_Default'.
  ; The 'ColorType' parameter can be #PB_Gadget_BackColor or #PB_Gadget_FrontColor.
  ;
  ; If 'ThisColorOnly' is set to zero and if #PB_Gadget_FrontColor is not allready set,
  ; the FrontColor will be adjusted in opposition from the BackColor.
  ;
  ; If 'ThisColorOnly' is set to zero and if #PB_Gadget_BackColor is not allready set,
  ; the BackColor with be set to the same color as the window background color.
  ;
  ; If 'ThisColorOnly' is set to 1, only the color specified by ColorType
  ; (#PB_Gadget_BackColor or #PB_Gadget_FrontColor) is adjusted.
  ;
  ;
  Shared GadgetsColors()
  ;
  Protected Luminosity, contener, ct, Found, OldREProc
  ;
  If IsGadget(GadgetID)
    ;
    Contener = GetGadgetWindow(GadgetID)
    If IsWindow(contener)
      contener = WindowID(contener)
      ; Set a callback procedure for the main window, in order to clean
      ; the memory when the window is closed:
      If GetWindowLongPtr_(contener, #GWL_USERDATA) = 0
        OldREProc = SetWindowLongPtr_(contener, #GWL_WNDPROC, @SGCE_MainWindowCallbackCleaner())
        SetWindowLongPtr_(contener, #GWL_USERDATA, OldREProc)
      EndIf
    EndIf
    ;
    ; Check is a GadgetsColors() element allready exists for the gadget:
    Found = 0
    If ListSize(GadgetsColors())
      ForEach GadgetsColors()
        If GadgetsColors()\Gadget_Handle = GadgetID(GadgetID)
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
        ; Clean memory --> Clean the element and its child element
        ForEach GadgetsColors()
          If GadgetsColors()\Gadget_ID = GadgetID
            If GadgetsColors()\Gadget_OldREProc
              ; Reset to standard management:
              SetWindowLongPtr_(GadgetsColors()\Gadget_Handle, #GWL_WNDPROC, GadgetsColors()\Gadget_OldREProc)
            EndIf
            DeleteElement(GadgetsColors())
          EndIf
        Next
        ;
        ; Force gadget to redraw:
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
      AddElement(GadgetsColors())
      GadgetsColors()\Gadget_ID = GadgetID
      GadgetsColors()\Gadget_Handle = GadgetID(GadgetID)
      GadgetsColors()\Gadget_FrontColor = #PB_Default
      GadgetsColors()\Gadget_BackColor = #PB_Default
      GadgetsColors()\Gadget_OldREProc = 0
      GadgetsColors()\Gadget_Enabled = IsWindowEnabled_(GadgetID(GadgetID))
      GadgetsColors()\Gadget_MainWindow = Contener
      GadgetsColors()\Gadget_ChildOf = 0
    EndIf
    ;
    If ColorType = #PB_Gadget_BackColor
      ;
      GadgetsColors()\Gadget_BackColor = Color
      ;
      ; From BackColor, compute an automatic color for FrontColor:
      ;
      If ThisColorOnly = 0 And GadgetsColors()\Gadget_FrontColor = #PB_Default And Color <> #PB_Default
        ; If gadget's front color is not allready set,
        ; setup automatic FrontColor from BackColor:
        Luminosity = (Red(Color) - $80) + (Green(Color) - $80) + (Blue(Color) - $80)
        If Abs(Luminosity) < 100
          ; Gadget color is quite middle grey
          FrontColor = RGB(255, 255, 255)
        ElseIf Luminosity > 0
          ; Bright theme
          FrontColor = 0
        Else
          ; Dark theme
          FrontColor = RGB(220, 220, 220)
        EndIf
        GadgetsColors()\Gadget_FrontColor = FrontColor
      EndIf
      ;
    ElseIf ColorType = #PB_Gadget_FrontColor
      ;
      GadgetsColors()\Gadget_FrontColor = Color
      ;
      If ThisColorOnly = 0 And GadgetsColors()\Gadget_BackColor = #PB_Default And Color <> #PB_Default
        ; If gadget's back color is not allready set,
        ; use the contener back color:
        contener = GetGadgetWindow(GadgetID)
        If IsWindow(contener)
          GadgetsColors()\Gadget_BackColor = GetWindowColor(contener)
        ElseIf IsGadget(contener)
          GadgetsColors()\Gadget_BackColor = GetGadgetColor(contener, #PB_Gadget_BackColor)
        EndIf
      EndIf
    EndIf
    ;
    If (GadgetsColors()\Gadget_FrontColor = #PB_Default Or GadgetsColors()\Gadget_FrontColor = GetSysColor_(#COLOR_WINDOWTEXT)) And (GadgetsColors()\Gadget_BackColor = #PB_Default Or GadgetsColors()\Gadget_BackColor = GetSysColor_(#COLOR_BTNFACE))
      ; Both Gadget_BackColor and Gadget_FrontColor are set to #PB_Default.
      ColorType = #SGCE_StandardColors
      Goto SetStandardColor
    EndIf
    ;
    ;
    ApplyDarkMode = 0
    Select GadgetType(GadgetID)
      Case #PB_GadgetType_ScrollArea, #PB_GadgetType_ScrollBar, #PB_GadgetType_Editor, #PB_GadgetType_ListIcon, #PB_GadgetType_ListView
        ApplyDarkMode = 1
      Case #PB_GadgetType_ListView, #PB_GadgetType_ComboBox, #PB_GadgetType_ExplorerCombo, #PB_GadgetType_ExplorerList
        ApplyDarkMode = 1
      Case #PB_GadgetType_ExplorerTree, #PB_GadgetType_Container, #PB_GadgetType_Web, #PB_GadgetType_WebView, #PB_GadgetType_Scintilla
        ApplyDarkMode = 1
    EndSelect
    If ApplyDarkMode
      ;
      ; If the background is darker than lighter, set the scrollbar to dark mode:
      If (Red(GadgetsColors()\Gadget_BackColor) + Green(GadgetsColors()\Gadget_BackColor) + Blue(GadgetsColors()\Gadget_BackColor)) > (127 * 3)
        SetWindowTheme_(GadgetID(GadgetID), @"Explorer", #Null)
      Else
        SetWindowTheme_(GadgetID(GadgetID), @"DarkMode_Explorer", #Null)
      EndIf
    EndIf
    ;
    Protected CustomColor = SGCE_IsCustomColorGadget(GadgetID)
    ;
    If CustomColor
      ;
      If GadgetsColors()\Gadget_FrontColor = #PB_Default
        GadgetsColors()\Gadget_FrontColor = GetSysColor_(#COLOR_WINDOWTEXT)
      EndIf
      If GadgetsColors()\Gadget_BackColor = #PB_Default
        GadgetsColors()\Gadget_BackColor = GetSysColor_(#COLOR_BTNFACE)
      EndIf
      ;
      ; Compute the color of the borders, for some gadgets, as EditorGadget, which need to repaint the border:
      Ratio.f = 0.25
      Red = Red(GadgetsColors()\Gadget_FrontColor) * Ratio + Red(GadgetsColors()\Gadget_BackColor) * (1 - Ratio)
      Green = Green(GadgetsColors()\Gadget_FrontColor) * Ratio + Green(GadgetsColors()\Gadget_BackColor) * (1 - Ratio)
      Blue = Blue(GadgetsColors()\Gadget_FrontColor) * Ratio + Blue(GadgetsColors()\Gadget_BackColor) * (1 - Ratio)
      GadgetsColors()\Gadget_BorderColor =  RGB(Red, Green, Blue)
      Ratio.f = 0.6
      Red = Red(GadgetsColors()\Gadget_FrontColor) * Ratio + Red(GadgetsColors()\Gadget_BackColor) * (1 - Ratio)
      Green = Green(GadgetsColors()\Gadget_FrontColor) * Ratio + Green(GadgetsColors()\Gadget_BackColor) * (1 - Ratio)
      Blue = Blue(GadgetsColors()\Gadget_FrontColor) * Ratio + Blue(GadgetsColors()\Gadget_BackColor) * (1 - Ratio)
      GadgetsColors()\Gadget_OverColor =  RGB(Red, Green, Blue)
      ;
      ; Set the gadget callback if not already done:
      ;
      If GadgetsColors()\Gadget_OldREProc = 0 
        GadgetsColors()\Gadget_OldREProc = SetWindowLongPtr_(GadgetsColors()\Gadget_Handle, #GWL_WNDPROC, @SGCE_ChangeColorsCallback())
      EndIf
      ; Register color data adress in #GWL_USERDATA
      SetWindowLongPtr_(GadgetsColors()\Gadget_Handle, #GWL_USERDATA, @GadgetsColors())
      ;
      ;
      If GadgetType(GadgetID) = #PB_GadgetType_ListIcon Or GadgetType(GadgetID) = #PB_GadgetType_ComboBox
        WChild = SGCE_EnumChildWindows(GadgetID)
        If WChild
          *Parent = @GadgetsColors()
          GadgetsColors()\Gadget_ParentOf = WChild
          SetWindowLong_(GadgetsColors()\Gadget_Handle, #GWL_EXSTYLE, GetWindowLong_(GadgetsColors()\Gadget_Handle, #GWL_EXSTYLE) & ~#WS_EX_CLIENTEDGE)
          SetWindowLong_(GadgetsColors()\Gadget_Handle, #GWL_STYLE, GetWindowLong_(GadgetsColors()\Gadget_Handle, #GWL_STYLE) | #WS_BORDER)
          Found = 0
          ForEach GadgetsColors()
            If GadgetsColors()\Gadget_Handle = WChild
              Found = 1
              Break
            EndIf
          Next
          If found = 0
            AddElement(GadgetsColors())
            CopyMemory(*Parent, GadgetsColors(), SizeOf(SGCE_GadgetsColors))
            GadgetsColors()\Gadget_Handle = WChild
            GadgetsColors()\Gadget_ChildOf = 1
            GadgetsColors()\Gadget_ParentOf = 0
            GadgetsColors()\Gadget_OldREProc = SetWindowLongPtr_(GadgetsColors()\Gadget_Handle, #GWL_WNDPROC, @SGCE_ChangeColorsCallback())
          EndIf
          SetWindowLongPtr_(WChild, #GWL_USERDATA, @GadgetsColors())
          ;
          ChangeCurrentElement(GadgetsColors(), *Parent)
          ;
        EndIf
      EndIf
      ;
      ; Force gadget to be redrawn:
      InvalidateRect_(GadgetID(GadgetID), 0, #True)
      For ct = 1 To 50 : WindowEvent() : Next
      ;
    EndIf
    If CustomColor  <>  1
      If ColorType = #PB_Gadget_FrontColor Or ColorType = #PB_Gadget_BackColor
        SetGadgetColor(GadgetID, #PB_Gadget_FrontColor, GadgetsColors()\Gadget_FrontColor)
        SetGadgetColor(GadgetID, #PB_Gadget_BackColor, GadgetsColors()\Gadget_BackColor)
      Else
        SetGadgetColor(GadgetID, ColorType, Color)
      EndIf
    EndIf
  EndIf
EndProcedure
;
Procedure GetGadgetColorEx(GadgetID, ColorType)
  ;
  ; Extended GetGadgetColor function for buttons and other
  ; gadgets with wich GetGadgetColor doesn't work.
  ;
  ;
  Shared GadgetsColors()
  ;
  If IsGadget(GadgetID)
    ;
    Protected Found = 0
    ;
    If SGCE_IsCustomColorGadget(GadgetID) = 1
      ;
      If ListSize(GadgetsColors())
        ForEach GadgetsColors()
          If GadgetsColors()\Gadget_Handle = GadgetID(GadgetID)
            Found = 1
            Break
          EndIf
        Next
      EndIf
      ;
      If Found
        If ColorType = #PB_Gadget_BackColor
          ;
          ProcedureReturn GadgetsColors()\Gadget_BackColor
          ;
        ElseIf ColorType = #PB_Gadget_FrontColor
          ;
          ProcedureReturn GadgetsColors()\Gadget_FrontColor
          ;
        Else
          ;
          ProcedureReturn GetGadgetColor(GadgetID, ColorType)
          ;
        EndIf
      EndIf
      ;
    Else
      ;
      ProcedureReturn GetGadgetColor(GadgetID, ColorType)
      ;
    EndIf
  EndIf
EndProcedure
;
;
; *************************************************************************************
;
;-                      2--- THIRD PART: DEMO PROCEDURE ---
;
; *************************************************************************************
;
Procedure MulticolorDemoWindow()
  ;
  If OpenWindow(1, 100, 100, 560, 370, "SetGadgetColorEx Demo", #PB_Window_ScreenCentered | #PB_Window_SystemMenu)
    ApplyDarkModeToWindow(1)
    ;
    BackColor = RGB(0, 40, 10)
    BorderColor = RGB(200, 200, 100)
    ;
    SetWindowColor(1, BackColor)
    ;
    Enumeration GadgetNum
      #Frame = 0
      #Option1
      #Option2
      #CheckBox
      #Button1
      #Text
      #String
      #Editor
      #ListView
      #Panel
      #Contener
      #ContenerButton1
      #ContenerButton2
      #ContenerButton3
      #ListIcon
      #Combo
      #ComboEdit
      #BQuit
    EndEnumeration
    ;
    ; Creating two option buttons
    FrameGadget(#Frame, 10, 10, 120, 70, "Options")
    ;
    SetGadgetColorEx(#Frame, #PB_Gadget_BackColor, RGB(0, 50, 15))
    SetGadgetColorEx(#Frame, #PB_Gadget_FrontColor, RGB($FF, $60, $00))
    ;
      OptionGadget(#Option1, 15, 25, 100, 25, "Option #1")
      SetGadgetState(#Option1, 1)
      ;
      SetGadgetColorEx(#Option1, #PB_Gadget_BackColor, RGB(0, 50, 15))
      SetGadgetColorEx(#Option1, #PB_Gadget_FrontColor, RGB($FF, $50, $FF))
      ;
      OptionGadget(#Option2, 15, 45, 100, 25, "Option #2")
      SetGadgetColorEx(#Option2, #PB_Gadget_BackColor, RGB(0, 50, 15))
      SetGadgetColorEx(#Option2, #PB_Gadget_FrontColor, RGB($60, $FF, $60))
    ;
    CheckBoxGadget(#CheckBox, 10, 90, 100, 25, "Desable")
    ;
    ; If #PB_Gadget_BackColor is not defined, SetGadgetColorEx() will automatically take
    ; the main window color.
    ;
    SetGadgetColorEx(#CheckBox, #PB_Gadget_FrontColor, BorderColor)
    SetGadgetState(#CheckBox, 1)
    ButtonGadget(#Button1, 10, 120, 100, 25, "Desabled Button")
    DisableGadget(#Button1, #True)
    SetGadgetColorEx(#Button1, #PB_Gadget_FrontColor, BorderColor)
    ;
    TextGadget(#Text, 10, 150, 100, 25, "Simple text")
    SetGadgetColorEx(#Text, #PB_Gadget_FrontColor, BorderColor)
    ;
    StringGadget(#String, 10, 170, 100, 25, "StringGadget")
    SetGadgetColorEx(#String, #PB_Gadget_FrontColor, BorderColor)
    ;
    EditorGadget(#Editor, 10, 200, 100, 50)
    SetGadgetText(#Editor, "EditorGadget avec un texte long")
    SetGadgetColorEx(#Editor, #PB_Gadget_FrontColor, BorderColor)
    ;
    ;
    ListViewGadget(#ListView, 140, 10, 200, 240)
    SetGadgetColorEx(#ListView, #PB_Gadget_FrontColor, BorderColor)
    ;
    For ct = 1 To 20
      AddGadgetItem(#ListView, ct - 1, "Line " + Str(ct))
    Next
    ;
    PanelGadget(#Panel, 350, 10, 200, 240)
    SetGadgetColorEx(#Panel, #PB_Gadget_FrontColor, RGB($60, $FF, $60))
    ;
    ; If #PB_Gadget_FrontColor is not defined, SetGadgetColorEx() will automatically choose
    ; a color with a strong contrast from #PB_Gadget_BackColor
    ;
    ;SetGadgetColorEx(#Panel, #PB_Gadget_FrontColor, BorderColor)
    ;
    ;
    ; For PanelGadget, SetWindowTheme_() doesn't work to get black arrows on top-right corner :(
    ;
    For ct = 1 To 10
      AddGadgetItem(#Panel, -1, "Tab " + Str(ct))
    Next
    CloseGadgetList()
    ;
    ContainerGadget(#Contener, 10, 260, 120, 100)
      SetGadgetColorEx(#Contener, #PB_Gadget_BackColor, RGB($90, 00, $90))
      ButtonGadget(#ContenerButton1, 10, 10, 100, 22, "Button #1")
      SetGadgetColorEx(#ContenerButton1, #PB_Gadget_BackColor, BorderColor)
      ButtonGadget(#ContenerButton2, 10, 40, 100, 22, "Button #2")
      ;
      ; For buttons and some other gadgets type,
      ; SetWindowTheme_(GadgetID(NoGadget), "DarkMode_Explorer", #Null)
      ; can be used to get a dark theme gadget.
      ; Then, if you want to go back to normal appearance, you can
      ; do SetWindowTheme_(GadgetID(NoGadget), "Explorer", #Null)
      ; or SetWindowTheme_(GadgetID(NoGadget), "", #Null)
      ;
      ; Notice that, for some types of gadget, SetWindowTheme_(GadgetID(NoGadget), "", #Null)
      ; set the gadget style to old fashion (old Windows versions) appearance.
      ;
      SetWindowTheme_(GadgetID(#ContenerButton2), "DarkMode_Explorer", #Null)
      ;
      ButtonGadget(#ContenerButton3, 10, 70, 100, 22, "Button #3")
      SetGadgetColorEx(#ContenerButton3, #PB_Gadget_FrontColor, RGB($FF, $FF, $0))
    CloseGadgetList()
    ;
    ListIconGadget(#ListIcon, 140, 260, 200, 90, "ListIcon", 120)
    AddGadgetColumn(#ListIcon, 1, "Column 2", 240)
    For ct = 1 To 5 : AddGadgetItem(#ListIcon, -1, "ListIcon Element " + Str(ct) + #LF$ + "Column 2 Element " + Str(ct))
    Next
    SetGadgetColorEx(#ListIcon, #PB_Gadget_FrontColor, RGB($C0, $FF, $0))
    
    ComboBoxGadget(#Combo, 350, 260, 200, 28)
    For ct = 1 To 5 : AddGadgetItem(#Combo, -1, "Combo Element " + Str(ct)) : Next
    SetGadgetState(#Combo, 0)
    SetGadgetColorEx(#Combo, #PB_Gadget_FrontColor, RGB($FF, $FF, $0))
    
    ComboBoxGadget(#ComboEdit, 350, 300, 200, 28, #PB_ComboBox_Editable)
    For ct = 1 To 5 : AddGadgetItem(#ComboEdit, -1, "Combo Editable Element " + Str(ct)) : Next
    SetGadgetState(#ComboEdit, 0)
    SetGadgetColorEx(#ComboEdit, #PB_Gadget_FrontColor, RGB($FF, $FF, $0))
    ;     
    ButtonGadget(#BQuit, WindowWidth(1) - 110, WindowHeight(1) - 35, 100, 25, "Exit")
    SetGadgetColorEx(#BQuit, #PB_Gadget_BackColor, RGB(255, 0, 0))
    ;       
    Repeat
      Event = WaitWindowEvent()
      
      If Event = #PB_Event_Gadget

        If EventGadget() = #CheckBox
          If GetGadgetState(#CheckBox)
            DisableGadget(#Button1, #True)
            SetGadgetText(#Button1, "Desabled Button")
          Else
            DisableGadget(#Button1, #False)
            SetGadgetText(#Button1, "Enabled Button")
          EndIf
        EndIf
    
        If EventGadget() = #BQuit
          Break
        EndIf
      ElseIf Event = #PB_Event_CloseWindow
          Break
      EndIf
    ForEver
    CloseWindow(1)
  EndIf
EndProcedure
;
CompilerIf #PB_Compiler_IsMainFile
  ; The following won't run when this file is used as 'Included'.
  ;
  MulticolorDemoWindow()

CompilerEndIf
; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; CursorPosition = 849
; FirstLine = 526
; Folding = 9h0
; EnableXP
; DPIAware