;
;**********************************************************************************
;
;                             SetMenuItemEx.pbi library
;                                   Windows only
;                             Zapman - March 2025 - 6.2
;
;          This file should be saved under the name "SetMenuItemEx.pbi".
;
;     This library offers functions to overcome the limitations of the current
;          PureBasic (6.20) functions dedicated to the menu management.
;
; • It is not possible to change the icon of a menu item after assigning it.
;   You have to destroy the menu and rebuild it entirely to be able to change one
;   of its icons.
; • It is not possible to retrieve the ImageID of a menu icon. You have to memorize
;   it in a variable at the time you assign it to be able to retrieve it later.
; • It is not possible to assign an icon to an item of the menu-bar of a window,
;   if this item has subitems.
; • It is not possible to change the menu items font and colors.
;
;  Many functions of this library (as SetMenuItemFont()) are simple added to the
;  original set of PureBasic functions.
;  Some existing functions (as MenuTitle()) are overrided to win more capacities
;  and flexibilities.
;
;                                 Added functions:
;
;  • SetMenuItemFont()   and GetMenuItemFont()
;  • SetMenuTitleFont()  and GetMenuTitleFont()
;  • SetMenuItemColor()  and GetMenuItemColor()
;  • SetMenuTitleColor() and GetMenuTitleColor()
;  • SetMenuItemImage()  and GetMenuItemImage()
;  • SetMenuTitleImage() and GetMenuTitleImage()
;  • CheckMenuItem()     and IsMenuItemChecked()
;  • SetMenuColor()      ; Apply a color to all existing items of a menu
;  • SetMenuFont()       ; Apply a font to all existing items of a menu
;
;                                Improved functions:
;
;  • MenuTitle() can now has an imageID as second parameter. It also returns
;    the handle of the submenu that this command creates.
;  • MenuBar() can now has an ID as parameter. This allows you to attribute
;    particular colors to menu-bars.
;  • You can now use dynamic allocation for items ID using #PB_Any as usual.
;    For example:   MyItem = MenuItem(#PB_Any, "Text of my item")
;  • You can use submenu's handles exactly the same way you would use IDs.
;    It means that you can change a menu title like this:
;        MyTitle = MenuTitle("Text of the title")
;        SetMenuItemText(MyMenu, MyTitle, "New text of the title")
;    This also means that you no longer need to use functions that start with
;    "SetMenuTitle..." (but they are still functional).
;  • The menu ID can be replaced by '#PB_Default' in all functions except
;    the ones beggining by "SetMenuTitle...". Consequently, you can do:
;        SetMenuItemText(#PB_Default, MyTitle, "New text of the title")
;    Item IDs are supposed to be unique and handles are also. Consequently,
;    the functions of this library will search in which menu or PopupMenu the
;    ID/handle provided as a parameter is located and they will replace
;    #PB_Default with the correct menu ID. This can be very usefull to retreive
;    the text of the command activated by the user in the main loop of the program.
;    You can do :  If WaitWindowEvent() = #PB_Event_Menu
;                    Define ItemText$ = GetMenuItemText(#PB_Default, EventMenu())
;                    MessageRequester("Info", "The user chose: " + ItemText$, 0)
;                  Endif
;    
;
;   This library is compatible with all other Zapman libraries, including
;   SetGadgetColorEx.pbi and ApplyColorThemes.pbi
; 
;***********************************************************************************
;
;                           SOME EXPLANATIONS ABOUT MENUS
;
; To understand what is done through the different functions of that library, it is
; important to know some things about how Windows has organized that part of its
; interface and how PureBasic manages it.
;
;           *---------------- About CreateImageMenu() ----------------*
;
; • The Windows API offers a complete set of functionnalities which allow you to add
;   menus and popup menus to an application without having to manage much.
;   The menu item text can be stored using the API SetMenuItemInfo_() function and
;   can be retreived using GetMenuItemInfo_().
; • The same functions can be used to store a menu item image pointer and the existence
;   of the CreateImageMenu() function, in addition to CreateMenu(), is not necessary
;   on Windows. At the very least, these two functions should be strictly equivalent
;   on this system. However, for reasons that are likely historical (Windows only
;   supports menu icons since Windows 2000), PureBasic handles menus very differently
;   depending on whether they are created with CreateMenu() (and CreatePopupMenu())
;   or with CreateImageMenu() (and CreatePopupImageMenu()). In the first case, menus
;   use the standard Windows functions for their entire functionality. In the second
;   case, the menu items are defined as "ownerdrawn" and PureBasic takes care of
;   drawing each of them.
; • When a menu item is defined as "ownerdrawn", it is common to associate it with 
;   a set of data that will define, in addition to the text and image linked to 
;   the item, the text color, for example, as well as the background color and, 
;   optionally, the font with which the item should be displayed.
; • As developers, when we want to add display options to the items of a menu 
;   created with CreateImageMenu(), we face a difficulty: these items are already 
;   defined as ownerdrawn by PureBasic, and a set of data is already associated 
;   with them. When attempting to add data to the existing data and manage the 
;   display of these items in a customized way, other issues arise: the functions 
;   GetMenuItemText() and SetMenuItemText() are not designed for such operations 
;   and lead to malfunctions. Worse yet: if a popup menu is created with the 
;   CreatePopupImageMenu() function, its management by PureBasic interferes with 
;   other menus managed as ownerdrawn by our application (which should absolutely 
;   not be the case).
; • In summary, the CreateImageMenu() function and the way its features are 
;   handled by PureBasic create unsolvable issues as soon as we want to add 
;   functionalities such as SetMenuItemColor(). I have therefore decided to 
;   override this function, along with many other menu-related functions, in 
;   order to create a reliable and functional library. One of the advantages of
;   this choice is that menus, even if they include images, are only ownerdrawn
;   when it is strictly necessary. (As we will see a bit later, owner-drawn menus
;   come with a series of drawbacks).
;   Notice the small arrow on the right that appears on items with a submenu:
;   —> Without using this library, it does not has the same appearance when
;      creating the menu with CreateMenu() and CreateImageMenu(). In the first
;      case, the Windows theme is applied. In the second case, it is not,
;      because the menu is then ownerdrawn.
;   —> When using this library, CreateMenu() and CreateImageMenu() are strictly
;      equivalent and you can add images to items in both cases.
;      Windows theme is applied until you set a particular color or font
;      for a menu item. In this case (and only in this case), the menu is
;      ownerdrawn and the Windows theme cannot be applied (but the menu behavior
;      comes as close as possible)
;
;          *---------------- About ownerdrawn menus ----------------*
;      All the following comments concern Windows 11 and older versions.
;
; • When a menu item is defined as ownerdrawn, its display must be managed by a 
;   callback procedure associated with the application's window. This procedure 
;   receives the #WM_INITMENU (or #WM_INITMENUPOPUP) message when the menu is 
;   about to be drawn for the first time. Then it receives the #WM_MEASUREITEM 
;   message, which allows it to calculate and record the dimensions of the rectangle 
;   in which the item will be drawn. Finally, it receives the #WM_DRAWITEM message 
;   when the item needs to be drawn.
; • To calculate the width of the menu item during the handling of the
;   #WM_MEASUREITEM message and to draw it during the handling of the #WM_DRAWITEM
;   message, the callback procedure needs data defining its text, image, colors,
;   and font. These data are stored in a structure, and the address of this structure
;   is associated with the item using the SetMenuItemInfo_() function.
;   This address can then be retrieved using GetMenuItemInfo_(). It is also available 
;   in measureItemStruct\itemData at the time of the #WM_MEASUREITEM message, and in
;   drawItemStruct\itemData, at the time of the #WM_DRAWITEM message.
; • When an image is associated with an ownerdrawn item using SetMenuItemInfo_()
;   and #MIIM_BITMAP, the #WM_MEASUREITEM  message is never sent to the callback
;   procedure. It is very unfortunate, and I personally consider it a mistake on 
;   Windows' part, but it is a fact that must be taken into account. If you associate
;   an image to an item using the classical way, Windows takes it upon itself 
;   to calculate the width required to display the item, but unfortunately, it 
;   does this incorrectly and only considers the image, without accounting for 
;   the width needed to display the text.
; • Consequently, when an item is defined as ownerdrawn, the image associated 
;   with it must NOT be registered in the standard way using SetMenuItemInfo_(). 
;   The 'hbmpItem' field of the MenuItemInfo structure must be set to '0' or
;   '#HBMMENU_CALLBACK' for the #WM_MEASUREITEM message to be sent and handled by
;   the callback procedure. However, since it will still be necessary to have
;   the address of the image during the handling of #WM_MEASUREITEM and #WM_DRAWITEM,
;   this address must be stored in the set of data associated with the item.
; • The 'ownerdrawn' mode of Windows menu items suffers from other bugs (also reported
;   by other developers on the web) that I had to discover, understand, and handle:
;   Since each menu item can be independently defined as ownerdrawn or not, one would
;   naturally expect everything to continue functioning normally if some items are
;   ownerdrawn while others are not. But this is not the case. As soon as a single
;   item in a menu is ownerdrawn, the display of all other items of this menu
;   encounters issues: they are no longer displayed using the current Windows theme
;   and are shifted to the right, with the offset oddly depending on the width assigned
;   to the ownerdrawn item. In short, it doesn't work (except if the ownerdrawn item is
;   the last of the menu). If you define one menu item as ownerdrawn, then you must
;   define ALL items in that menu as ownerdrawn, even if you don't need them to be.
; • When you change the text or the image of a menu-bar item that is not ownerdrawn,
;   the menu-bar is automatically redrawn. When you change the data of an ownerdrawn
;   menu-bar item, it is NOT automatically resized and redrawn (it should be!).
;   The #WM_MEASUREITEM and #WM_DRAWITEM are not automatically sent to it.
;   And because no explicit 'ResizeMenuItem()' function does exist, you have to use
;   some sort of 'hack' to get the equivalent: by calling SetMenuItemInfo_()
;   with fMask = #MIIM_BITMAP and hbmpItem = 0, #WM_MEASUREITEM is resent to the callback
;   procedure, even if hbmpItem was allready set to zero before the call of SetMenuItemInfo_().
;   This whole process has something absurd about it.
; • When you set a MenuBar item as ownerdrawn, the #WM_MEASUREITEM message is normally
;   sent to the callback procedure, but the #WM_DRAWITEM is not sent if you have not
;   attributed an ID to the MenuBar. Any positive number can be used, and you can even
;   attribute an ID allready used by another item (what shouldn't work!). It doesn't
;   matter. The system never check or use this ID but it needs one to send the #WM_DRAWITEM.
;   message. This is another absurd aspect of the ownerdrawing system's functions.
; • When an item is a submenu entry, it's ID should be -1. But the Windows management
;   of ownerdrawn menus has a very rare but very annoying bug (this is not systematic
;   as for MenuBars): when the ID of an item is equal to -1, approximately once in 200,
;   the #WM_DRAWITEM message is not sent to the callback procedure and the item is never
;   drawn. When an item falls victim to this bug, it's for good: you can send DrawMenuBar_()
;   repeatedly, the faulty item will never be drawn. Something in its data is apparently
;   corrupt without it being possible to know exactly why. After numerous and exhausting
;   tests, it appears that assigning an ID of 1 to all items that are submenu entries
;   corrects this problem. Notice that, in reality, the ID will not REALLY be assigned
;   to the item. If you test this with GetMenuItemID_(hmenu, ItemPos) after the operation,
;   you will see that the item still has an ID equal to -1. But this fixes the malfunction
;   and the bug will no longer occur.
; • In modern versions of Windows, menu bar items are activated when the mouse cursor
;   hovers over them, providing a smooth and intuitive user experience.
;   Unfortunately, this behavior does not apply to owner-drawn menu items, which are
;   managed in the "old-fashioned" Windows way. These items do not react to mouse hover
;   events, and no #WM_DRAWITEM message is sent to the window's callback procedure
;   in such cases. I address this limitation by explicitly triggering a #WM_DRAWITEM 
;   message when the mouse cursor hovers over a menu item. This is handled by the
;   callback procedure (Case #WM_NCMOUSEMOVE).
; • In modern versions of Windows, the menu bar is grayed out when the window becomes
;   inactive. However, this behavior typically does not occur with owner-drawn menus
;   because Windows does not send the #WM_DRAWITEM message to owner-drawn menus when
;   the window is activated or deactivated. This oversight has been addressed by adding
;   a snippet of code in the window's callback procedure (Case #WM_ACTIVATE).
; • These last two points may explain why PureBasic never defines the menu bar as
;   owner-drawn, even when the menu is created with CreateImageMenu().
;   This avoids the aforementioned drawbacks but this choice introduces a third one:
;   it becomes impossible to display an image in the menu bar, even when it is created 
;   with CreateImageMenu().
;   By rewriting PureBasic's native functions, this library resolves all these issues:
;   1- The menu bar can now display images without being owner-drawn.
;   2- If the menu bar is set as owner-drawn to allow its items to be colored or 
;      displayed with a specific font, it will still behave like a non-owner-drawn
;      menu bar. If a theme exists, colors of the theme will be used by default.
;
;   *---------------- About the 'MenuTitle()' PureBasic function ----------------*
;
; • The native MenuTitle() function in PureBasic is essentially a disguised version
;   of the OpenSubMenu() function. When you use it to add a title to the main menu,
;   it actually creates a submenu entry, into which you can then add items with
;   the MenuItem() function. From your perspective, MenuItem() adds items in the
;   same menu where you added a title with MenuTitle(). In reality, (i.e. for Window)
;   these are two separate menus (a menu and a submenu) that are nested. Windows
;   handle a submenu (or a popumenu) the same way as a menu.
; • Unlike the 'OpenSubmenu()' function which returns a handle to the created submenu,
;   PureBasic's native MenuTitle() function does not return a handle. This is why,
;   when you want to change the title of a menu with SetMenuTitleText(), you are forced
;   to designate the item by its position in the menu (which is not very practical).
;   This lack is corrected by this library. Now MenuTitle() returns you a handle and
;   you can use SetMenuItemText() with this handle to change the menu title. You can
;   also give it an imageID as second parameter.
; • The design of the native OpenSubMenu() function can be improved, because if the
;   user forget to call 'CloseSubmenu()' at the end of the menu creation, in case of
;   heavy menu usage, including creations and destructions (FreeMenu) of a menu with
;   an unclosed submenu, this can end up causing a memory error (this happens exactly
;   on the 64rd iteration of a create/destruct loop). This library fixes this problem
;   by forcing a call to CloseSubmenu() before each CreateMenu(), CreatePopupMenu()
;   and FreeMenu().
;
;   *---------------- About the 'MenuBar()' PureBasic function ----------------*
;
;   PureBasic's native MenuBar() function does not return a handle and you cannot
;   attribute an ID to it. You can with this library. So you may use this ID to
;   attribute a particular color to a menu-bar.
;
CompilerIf #PB_Compiler_IsMainFile
  EnableExplicit
CompilerEndIf
;
;**********************************************************************************
;
;-               1- STRUCTURES GLOBALS AND CONSTANTES DECLARATIONS
; 
Global SMI_MenuHeight = GetSystemMetrics_(#SM_CYMENU)
#SMI_MenuBulletSize = 4
;
EnumerationBinary SMI_ShowSelectionMethod
  #SMI_SSW_Bullet = 1
  #SMI_SSW_Borders
  #SMI_SSW_SystemColor
  #SMI_SSW_ShadeBackground
EndEnumeration
;
Global SMI_ShowSelectionMethod
If SMI_ShowSelectionMethod = 0
  ; By changing the value of SMI_ShowSelectionMethod,
  ; you change the manner of showing to the user that the mouse
  ; is over an item of the menu. Usualy, a simple background
  ; shading is enough to do that, but it can be inefficient
  ; in a dark mode environment or when each item has its own
  ; particular color (or both). So, you can choose four different
  ; methods (see the SMI_ShowSelectionMethod enumeration).
  ; Try each of them to decide.
  ;
  ; Note that your can eventually combine different methods.
  ; SMI_ShowSelectionMethod = #SMI_SSW_SystemColor | #SMI_SSW_Bullet
  ;
  SMI_ShowSelectionMethod =  #SMI_SSW_SystemColor
EndIf
;
#SMI_SubItemSearch = -1
;
Structure MENUITEMINFO_Fixed Align #PB_Structure_AlignC
  ; The MENUITEMINFO structure described in PureBasic 6.20
  ; and olders is uncomplete.
  ; Great thanks to 'idle' from english PureBasic forum,
  ; for the right form of this structure.
  cbSize.l
  fMask.l
  fType.l
  fState.l
  wID.l
  hSubMenu.i
  hbmpChecked.i
  hbmpUnchecked.i
  dwItemData.i
  *dwTypeData
  cch.l
  hbmpItem.i ; This field is missing in the PureBasic 6.20 (and older) description of MENUITEMINFO.
EndStructure
;

#MENU_BARITEM         = 8
#MENU_POPUPITEM       = 14
#MBI_HOT              = 2
#MPI_HOT              = 2
#MBI_NORMAL           = 1
#MPI_NORMAL           = 1
#MENU_BARBACKGROUND   = 7
#MENU_POPUPBACKGROUND = 9
#MENU_POPUPCHECK      = 11
#MC_CHECKMARKNORMAL   = 1
#TMT_FILLCOLOR        = 3802
#TMT_TEXTCOLOR        = 3803
;
#MIM_MENUDATA         = 8
#HBMMENU_CALLBACK     = -1
;
Structure MENUINFO Align #PB_Structure_AlignC
  cbSize.l           ; Size of the structure (in bytes).
  fMask.l            ; Mask specifying which members are being set or retrieved.
  dwStyle.l          ; Style of the menu.
  cyMax.l            ; Maximum height of the menu, in pixels.
  hbrBack.i          ; Handle to the brush used to paint the menu background.
  dwContextHelpID.l  ; Context help ID for the menu.
  dwMenuData.i       ; Custom data associated with the menu.
EndStructure
;
Structure MENUBARINFO Align #PB_Structure_AlignC
  cbSize.l
  rcBar.RECT
  hMenu.i
  hwndMenu.i
  fBarFocused.b
  fFocused.b
  fUnused.b
  Padding_1.b
EndStructure
;
Enumeration MustBeOwnerdrawn
  #SMI_NOT_Ownerdrawn
  #SMI_Ownerdrawn
EndEnumeration
;
Structure OwnerdrawnMenuItemDataStruct
  ;
  ; As the Windows documentation says : "The menu item data can represent
  ; any information that is meaningful to your application, and that will
  ; be available to your application when the item is to be displayed.
  ; For example, the value could contain a pointer to a structure;
  ; the structure, in turn, might contain a text string and a handle to the
  ; logical font your application will use to draw the string."
  ; The data of this structure can be accessed through the field
  ; *drawItem\itemData when the #WM_DRAWITEM message is sent to the main
  ; window of the program. It can also be accessed with the GetMenuItemInfo_()
  ; API function. See SMI_FillMenuItemData() as an example about how to do that.
  ;
  *MenuItemTextPtr      ; This field contains a pointer to the text of
                        ;    the menu item. To be compatible with the way
                        ;    PureBasic handles ownedrawn items, the text
                        ;    must be stored in memory with the
                        ;    SysAllocString_(ItemText$) function.
  *MenuItemImgHandle    ; This field contains a pointer to the image
                        ;   illustrating the menu item.
  MenuItemImgNum.i      ; This field is used to store the PureBasic
                        ;   imageID of the image.
  MustBeOwnerdrawn.i    ; Equal to #SMI_Ownerdrawn when the item must be ownerdrawn.
  IsItemOwnerdrawn.i    ; Equal to #SMI_Ownerdrawn when the item is ownerdrawn.
  IsItemSeparator.i     ; Equal to #True when the item is a separator.
  *ParentMenuHandle     ; Handle of the menu owning the item.
  MenuItemPos.i         ; Position of the item in the menu.
  MenuItemID.i          ; ID of the item, if it gets one.
                        ;   (submenu entries has not ID).
  *MenuItemFont         ; Font to use when drawing the item.
  MenuItemBackColor.l   ; BackColor to use when drawing the item.
  MenuItemTextColor.l   ; TextColor to use when drawing the item.
  MenuItemState.l       ; Check or UnCheck state of the item.
EndStructure
;
Global NewList MenuItemData.OwnerdrawnMenuItemDataStruct()
;
Define SMI_MemMenuNum = -1, SMI_MemMenuHandle = 0
;
;*******************************************************************************
;
;-                             2- GENERAL FUNCTIONS
;                     (possibly reusable in other programs)
;
;*******************************************************************************
;
Procedure LoadMenuFont()
  #SPI_GETNONCLIENTMETRICS = $0029
  Protected ncm.NONCLIENTMETRICS
  ncm\cbSize = SizeOf(NONCLIENTMETRICS)
  ;
  If SystemParametersInfo_(#SPI_GETNONCLIENTMETRICS, ncm\cbSize, @ncm, 0)
    Protected FontName$ = PeekS(@ncm\lfMenuFont\lfFaceName)
    Protected FontSize = ncm\lfMenuFont\lfHeight / DesktopResolutionY()
    ProcedureReturn LoadFont(#PB_Any, FontName$, FontSize)
  EndIf
EndProcedure
;
Global DefaultMenuFont = LoadMenuFont()
;
CompilerIf Not(Defined(GetImageFromShell32, #PB_Procedure))
  Procedure GetImageFromShell32(IconNum, ImgWidth, ImgHeight)
    ;
    Protected TransparentImage = CreateImage(#PB_Any, ImgWidth, ImgHeight, 32, #PB_Image_Transparent)
    Protected hIcon = ExtractIcon_(0, "shell32.dll", IconNum)
    ;
    If IsImage(TransparentImage) And hIcon
      Protected Dest_hDC = StartDrawing(ImageOutput(TransparentImage))
      If Dest_hDC
        DrawingMode(#PB_2DDrawing_AlphaBlend)
        Box(0, 0, ImgWidth, ImgHeight, RGBA(0, 0, 0, 0))
        DrawIconEx_(Dest_hDC, 0, 0, hIcon, ImgWidth, ImgHeight, 0, #Null, #DI_NORMAL)
        StopDrawing()
        DeleteDC_(Dest_hDC)
      EndIf
    EndIf
    DestroyIcon_(hIcon)
    ;
    ProcedureReturn TransparentImage
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(ResizeImageToIconSize, #PB_Procedure))
  Procedure ResizeImageToIconSize(SourceImage)
    ;
    If SourceImage
      If IsImage(SourceImage)
        SourceImage = ImageID(SourceImage)
      EndIf
      ;
      Protected ResizedImage = CreateImage(#PB_Any, GetSystemMetrics_(#SM_CXSMICON), GetSystemMetrics_(#SM_CYSMICON), 32, #PB_Image_Transparent)
      If ResizedImage
        If StartDrawing(ImageOutput(ResizedImage))
          DrawingMode(#PB_2DDrawing_AlphaBlend)
          DrawImage(SourceImage, 0, 0, GetSystemMetrics_(#SM_CXSMICON), GetSystemMetrics_(#SM_CYSMICON))
          StopDrawing()
          ProcedureReturn ResizedImage
        EndIf
      EndIf
    EndIf
    ProcedureReturn #Null
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(CreateIconFromImage, #PB_Procedure))
  Procedure CreateIconFromImage(hBitmap)
    ;
    Protected iconInfo.ICONINFO, bitmap.BITMAP
    ;
    If IsImage(hBitmap)
      hBitmap = ImageID(hBitmap)
    EndIf
    ;
    If GetObject_(hBitmap, SizeOf(BITMAP), @bitmap)
      ;
      ; Fill the ICONINFO Structure
      iconInfo\fIcon = #True
      iconInfo\xHotspot = 0
      iconInfo\yHotspot = 0
      iconInfo\hbmMask = CreateBitmap_(bitmap\bmWidth, bitmap\bmHeight, 1, 1, #Null)
      iconInfo\hbmColor = hBitmap
      Protected hIcon = CreateIconIndirect_(@iconInfo)
      ;
      ; Free the mask
      DeleteObject_(iconInfo\hbmMask)
      ;
      ProcedureReturn hIcon
    EndIf
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(DrawTransparentRectangle, #PB_Procedure))
  Procedure DrawTransparentRectangle(Dest_hDC, *rect.RECT, CoverColor, Opacity)
    ;
    Protected TempRect.Rect, hBrush
    Protected Srce_hDC = CreateCompatibleDC_(Dest_hDC)
    ;
    If CoverColor = 0 : CoverColor = 1 : EndIf ; There is a bug when color = 0
    ;
    If Srce_hDC
      Protected ImgWidth = *rect\right - *rect\left
      Protected ImgHeight = *rect\bottom - *rect\top
      Protected hbmTemp = CreateCompatibleBitmap_(Dest_hDC, ImgWidth, ImgHeight)
      If hbmTemp
        Protected oldBitmap = SelectObject_(Srce_hDC, hbmTemp)
        If oldBitmap
          Protected blend, *blend.BLENDFUNCTION = @blend
          ;
          If OpenLibrary(0, "Msimg32.dll")
            ;
            ; Fill hbmTemp with CoverColor
            TempRect\left = 0 : TempRect\top = 0
            TempRect\right = ImgWidth : TempRect\bottom = ImgHeight
            hBrush = CreateSolidBrush_(CoverColor)
            FillRect_(Srce_hDC, TempRect, hBrush)
            DeleteObject_(hBrush)
            ;
            ; Copy hbmTemp from Srce_hDC to Dest_hDC, respecting AlphaBlend:
            *blend\BlendOp = #AC_SRC_OVER
            *blend\BlendFlags = 0
            *blend\AlphaFormat = 0
            *blend\SourceConstantAlpha = Opacity
            CallFunction(0, "AlphaBlend", Dest_hDC, *rect\left, *rect\top, ImgWidth, ImgHeight, Srce_hDC, 0, 0, ImgWidth, ImgHeight, blend)
            ;
            CloseLibrary(0)
          EndIf
          ;
          SelectObject_(Srce_hDC, oldBitmap)
        EndIf
        DeleteObject_(hbmTemp)
      EndIf
      DeleteDC_(Srce_hDC)
    EndIf
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(DrawRightPointingTriangle, #PB_Procedure))
  Procedure DrawRightPointingTriangle(hDC, *rc.Rect, tSize, FrontColor, BackColor)
    ;
    Protected *points = AllocateMemory(3 * SizeOf(POINT))
    Protected vCenter = *rc\top + (*rc\bottom - *rc\top) / 2 - 1
    ;
    ; Calculate the coordinates of the triangle
    PokeL(*points + 0, *rc\left + tSize) ; Point for the tip of the triangle
    PokeL(*points + 4, vCenter)  ; Centered vertically within the triangle
    ;
    PokeL(*points + 8, *rc\left)
    PokeL(*points + 12, vCenter + tSize) ; Bottom of the triangle
    ;
    PokeL(*points + 16, *rc\left);
    PokeL(*points + 20, vCenter - tSize) ; Top of the triangle
    ;
    Protected hPen = CreatePen_(#PS_SOLID, 1, FrontColor) ; Pen with 1-pixel thickness
    Protected hBrush = CreateSolidBrush_(BackColor)
    SelectObject_(hDC, hPen)
    SelectObject_(hDC, hBrush)
    ;
    Polygon_(hDC, *points, 3) ; Draw the triangle
    ;
    ; CleanUp
    DeleteObject_(hPen)
    DeleteObject_(hBrush)
    FreeMemory(*points)
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(DrawCheckmark, #PB_Procedure))
  Procedure DrawCheckmark(hDC, *rc.Rect, Size, Color)
    Protected pt.POINT, hPen
    
    ; Calcul de la position du rectangle 10x10 à l'intérieur de rc
    pt\x = *rc\left  ; Aligné à gauche
    pt\y = *rc\top + (*rc\bottom - *rc\top - Size) / 2; Centrage vertical
    
    ; Création d'un stylo pour dessiner la checkmark
    hPen = CreatePen_(#PS_SOLID, 2, Color)
    SelectObject_(hDC, hPen)
    
    ; Dessiner la checkmark (?) avec MoveToEx_ et LineTo_
    MoveToEx_(hDC, pt\x + Size/5, pt\y + Size * 2 / 5, 0)
    LineTo_(hDC, pt\x + Size * 3 / 5, pt\y + Size * 4 / 5)
    LineTo_(hDC, pt\x + Size * 7 / 5, pt\y)
    
    ; Nettoyage
    DeleteObject_(hPen)
  EndProcedure
CompilerEndIf
;
CompilerIf Not(Defined(GetTextWidthInWindowContext, #PB_Procedure))
  Procedure GetTextWidthInWindowContext(hWnd, FontID, Text$)
    If IsWindow(hWnd) : hWnd = WindowID(hWnd) : EndIf
    If IsFont(FontID) : FontID = FontID(FontID) : EndIf
    Protected hDC = GetDC_(hWnd), hOldFont, Size.SIZE
    If hDC
      hOldFont = SelectObject_(hDC, FontID) ; Apply the specified font
      GetTextExtentPoint32_(hDC, Text$, Len(Text$), @Size) ; Get the text width
      SelectObject_(hDC, hOldFont) ; Restore the previous font
      ReleaseDC_(hWnd, hDC) ; Release the drawing context
      ProcedureReturn Size\cx
    EndIf
    ProcedureReturn 0
  EndProcedure
CompilerEndIf
; Prototype for the DwmSetWindowAttribute_ function
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
; *****************************************************************************
;
;-                 3. SPECIALIZED PROCEDURES OF THE LIBRARY
;
;      The following procedure are subroutines of the library functions.
;
; *****************************************************************************
;
Procedure   SMI_GetLastMenuItem(hMenu)
  ;
  ; Return the position of the last menu item of hMenu.
  ;
  Protected LastMenuItem = GetMenuItemCount_(hMenu) - 1
  ;
  If LastMenuItem < 0 Or LastMenuItem > 65535
    ProcedureReturn -1
  EndIf
  ProcedureReturn LastMenuItem
EndProcedure
;
Procedure   SMI_InitMenuItemInfoData(*MenuItemInfo.MENUITEMINFO_Fixed)
  FillMemory(*MenuItemInfo, SizeOf(MENUITEMINFO_Fixed), 0)
  *MenuItemInfo\cbSize     = SizeOf(MENUITEMINFO_Fixed)
EndProcedure
;
Procedure.s SMI_GetClassicMenuStringFromPosition(hMenu, Position)
  ;
  Protected MenuItemInfo.MENUITEMINFO_Fixed, ItemString$
  ;
  ; Get the size of the string (the number of chars):
  SMI_InitMenuItemInfoData(@MenuItemInfo)
  MenuItemInfo\fMask      = #MIIM_STRING
  MenuItemInfo\dwTypeData = 0
  ; The first call to GetMenuItemInfo_ with dwTypeData set to zero
  ; allows to get the string size (returned in cch):
  GetMenuItemInfo_(hMenu, Position, #MF_BYPOSITION, @MenuItemInfo)
  ;
  If MenuItemInfo\cch
    ; Allocate memory for a unicode string:
    ItemString$ = Space(MenuItemInfo\cch)
    MenuItemInfo\cch + 1 ; Add room for the ending character.
    ; Put the item string buffer address in MenuItemInfo\dwTypeData:
    MenuItemInfo\dwTypeData = @ItemString$
    ; Retreive the string:
    GetMenuItemInfo_(hMenu, Position, #MF_BYPOSITION, @MenuItemInfo)
  EndIf
  ProcedureReturn ItemString$
EndProcedure
;
Procedure   SMI_GetMenuItemPos(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, *ItemPos.Integer, RecursiveSearch = #True)
  ;
  ; Explore a menu and, eventually,  its submenus, looking for an item
  ; having an ID corresponding to ItemPosOrIdOrSubmenuHandle, or being
  ; a submenu entry for the submenu having the ItemPosOrIdOrSubmenuHandle
  ; handle.
  ;
  ; ByPosOrIdOrSubMenuHandle decides what type of search must be done.
  ;
  ; The menu handle of the menu or submenu containing the item is returned.
  ; *ItemPos is then set with the position of the item in hMenu.
  ;
  ; If RecursiveSearch = #True, all the submenus will be explored. Else
  ; only the menu designated by the hMenu entry parameter will be explored.
  ;
  Protected LastItem, Counter, hSubMenu, Result
  ;
  If IsMenu(hMenu)
    hMenu = MenuID(hMenu)
  EndIf
  ;
  If IsMenu_(ItemPosOrIdOrSubmenuHandle)
    ByPosOrIdOrSubMenuHandle = #SMI_SubItemSearch
    ; When using the Windows API, items that are submenu entries
    ; has no ID. So, in the callback procedure, when receiving
    ; an identifier for that type of item, it can't be an ID!
    ; Instead of that, the handle of the corresponding submenu
    ; is given. If IsMenu_(ItemPosOrIdOrSubmenuHandle) returns
    ; a non-zero value, it means that the ItemPosOrIdOrSubmenuHandle
    ; parameter is a submenu handle.
    ; We will exlore all the menu items and subitems until we
    ; found the submenu entry.
    ; You might be wondering "What happens if an Item ID is equal
    ; to a Submenu handle?" This theoretically cannot happen,
    ; because IDs are normally always less than #FFFF, and handles
    ; are always greater.
  EndIf
  ;
  If ByPosOrIdOrSubMenuHandle = #MF_BYPOSITION
    *ItemPos\i = ItemPosOrIdOrSubmenuHandle
    ProcedureReturn hMenu
  Else
    LastItem = SMI_GetLastMenuItem(hMenu)
    For Counter = 0 To LastItem
      hSubMenu = GetSubMenu_(hMenu, Counter)
      If hSubMenu = 0 And ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND And GetMenuItemID_(hMenu, Counter) = ItemPosOrIdOrSubmenuHandle
        *ItemPos\i = Counter
        ProcedureReturn hMenu
      Else
        If hSubMenu
          If ByPosOrIdOrSubMenuHandle = #SMI_SubItemSearch And hSubMenu = ItemPosOrIdOrSubmenuHandle
            *ItemPos\i = Counter
            ProcedureReturn hMenu
          ElseIf RecursiveSearch
            Result = SMI_GetMenuItemPos(hSubMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, *ItemPos, RecursiveSearch)
            If Result
              ProcedureReturn Result
            EndIf
          EndIf
        EndIf
      EndIf
    Next
  EndIf
EndProcedure
;
Procedure   SMI_FillMenuItemData(hMenu, ItemPosOrIdOrSubmenuHandle, *MyMenuItemData.OwnerdrawnMenuItemDataStruct, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND, RecursiveSearch = #True)
  ;
  ; Retreive the item designated by the entry parameters and copy its
  ; data into *MyMenuItemData, whether the item is ownerdrawn or not.
  ; This allows the program to deal with a unic type of data storing.
  ;
  Protected MenuItemInfo.MENUITEMINFO_Fixed
  Protected ItemString$, ItemPos
  Protected bitmap.bitmap
  ;
  MenuItemInfo\dwItemData = 0
  ;
  hMenu = SMI_GetMenuItemPos(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, @ItemPos, RecursiveSearch)
  ;  
  If hMenu And *MyMenuItemData
    ; The item has been found. hMenu will be the returned value.
    ;
    ; CAUTION: If the searched item is an ID or a SubMenu handle, the returned
    ; hMenu may be different from the hmenu value given as parameter because
    ; the item may be found in a submenu of hMenu.
    ;
    ; Check if the menu item is #MFT_OWNERDRAW:
    SMI_InitMenuItemInfoData(@MenuItemInfo)
  	MenuItemInfo\fMask      = #MIIM_FTYPE
  	GetMenuItemInfo_(hMenu, ItemPos, #MF_BYPOSITION, @MenuItemInfo)
  	;
  	If MenuItemInfo\fType & #MFT_SEPARATOR
  	  *MyMenuItemData\IsItemSeparator = #True
  	EndIf
  	;
  	If MenuItemInfo\fType & #MFT_OWNERDRAW
  	  ;
  	  ; MenuItem is ownerDraw. Get its data:
  	  SMI_InitMenuItemInfoData(@MenuItemInfo)
      MenuItemInfo\fMask      = #MIIM_DATA
      GetMenuItemInfo_(hMenu, ItemPos, #MF_BYPOSITION, @MenuItemInfo)
    EndIf
    ;
    If MenuItemInfo\dwItemData = 0
      ; MenuItem has no data. It is not ownerDraw.
      *MyMenuItemData\IsItemOwnerdrawn    = #SMI_NOT_Ownerdrawn
      ;
      ; Retreive the item data using the classical API functions:
      SysFreeString_(*MyMenuItemData\MenuItemTextPtr)
      ItemString$ = SMI_GetClassicMenuStringFromPosition(hMenu, ItemPos)
      *MyMenuItemData\MenuItemTextPtr = SysAllocString_(ItemString$)
      ;
      SMI_InitMenuItemInfoData(@MenuItemInfo)
      MenuItemInfo\fMask      = #MIIM_BITMAP | #MIIM_STATE
      GetMenuItemInfo_(hMenu, ItemPos, #MF_BYPOSITION, @MenuItemInfo)
      ; Check if the image is valid:
      If GetObject_(MenuItemInfo\hbmpItem, SizeOf(BITMAP), @bitmap)
        *MyMenuItemData\MenuItemImgHandle = MenuItemInfo\hbmpItem
      EndIf
      *MyMenuItemData\MenuItemState     = MenuItemInfo\fState 
      ;
    Else
      If MenuItemInfo\dwItemData <> *MyMenuItemData
        ; The data pointed by dwItemData is not the same that data
        ; pointed by the *MyMenuItemData parameter.
        CopyStructure(MenuItemInfo\dwItemData, *MyMenuItemData, OwnerdrawnMenuItemDataStruct)
      EndIf
    EndIf
    ;
    *MyMenuItemData\ParentMenuHandle = hMenu
    *MyMenuItemData\MenuItemPos      = ItemPos
    *MyMenuItemData\MenuItemID       = GetMenuItemID_(hMenu, ItemPos)
    If GetSubMenu_(hMenu, ItemPos)
      *MyMenuItemData\MenuItemID     = -1
    EndIf
    ;
  EndIf
  ;
  ProcedureReturn hMenu
  ;
EndProcedure
;
Procedure   SMI_SaveMenuWindow(hMenu, WindowID)
  Protected MenuInfo.MenuInfo
  ;
  MenuInfo\cbSize = SizeOf(MenuInfo)
  MenuInfo\fMask  = #MIM_MENUDATA
  MenuInfo\dwMenuData = WindowID
  SetMenuInfo_(hMenu, @MenuInfo)
EndProcedure
;
Procedure   SMI_RetreiveMenuWindow(hMenu)
  ;
  Protected MenuInfo.MenuInfo
  ;
  MenuInfo\cbSize = SizeOf(MenuInfo)
  MenuInfo\fMask  = #MIM_MENUDATA
  GetMenuInfo_(hMenu, @MenuInfo)
  ;
  ProcedureReturn MenuInfo\dwMenuData
EndProcedure
;
Procedure   SMI_CheckID(*ItemID.Integer)
  Static IDNum = $AFFF
  If *ItemID\i > $AFFF
    MessageRequester("Error", "ID upper that 45055 ($AFFF) are reserved for dynamic allocation.")
  EndIf
  If *ItemID\i = #PB_Any Or *ItemID\i > $AFFF
    IDNum + 1
    *ItemID\i = IDNum
  EndIf
  ProcedureReturn *ItemID\i
EndProcedure
;
Procedure   SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND, *ItemPos.Integer = 0, RecursiveSearch = #True)
  ;
  ; An element of MenuItemData() is created for each encountered menu item,
  ; regardless it is ownerdrawn or not.
  ; This allows the program to deal with a unic type of data storing.
  ;
  Protected Found = 0, ItemPos, SearchType$, ErrorDetails$
  ;
  If hMenu = #PB_Default And (ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND Or IsMenu_(ItemPosOrIdOrSubmenuHandle))
    ; No menu is specified. Explore all existing items to find the searched one
    ForEach MenuItemData()
      If MenuItemData()\MenuItemID = ItemPosOrIdOrSubmenuHandle Or GetSubMenu_(MenuItemData()\ParentMenuHandle, MenuItemData()\MenuItemPos) = ItemPosOrIdOrSubmenuHandle
        hMenu = MenuItemData()\ParentMenuHandle
        ItemPos = MenuItemData()\MenuItemPos
        Found = 1
        Break
      EndIf
    Next
  Else
    ; Retreive the exact menu handle and the position of the searched item:
    hMenu = SMI_GetMenuItemPos(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, @ItemPos, RecursiveSearch)
  EndIf
  If hMenu And hMenu <> #PB_Default
    If Found = 0
      ; Look for an existing element representing the item data:
      ForEach MenuItemData()
        If MenuItemData()\ParentMenuHandle = hMenu And MenuItemData()\MenuItemPos = ItemPos
          Found = 1
          Break
        EndIf
      Next
    EndIf
    If Found = 0
      ; Create a new element:
      AddElement(MenuItemData())
      MenuItemData()\MenuItemFont = #PB_Default
      MenuItemData()\MenuItemBackColor = #PB_Default
      MenuItemData()\MenuItemTextColor = #PB_Default
    EndIf
    ;
    ; Fill the element with the item data:
    SMI_FillMenuItemData(hMenu, ItemPos, @MenuItemData(), #MF_BYPOSITION)
    ;
  Else
    If ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND
      SearchType$ = "BY COMMAND"
    Else
      SearchType$ = "BY POSITION" + #CR$ + #CR$ + "You should try to call 'SetMenuItem...()' instead of 'SetMenuTitle...()'."
    EndIf
    If IsMenu_(ItemPosOrIdOrSubmenuHandle)
      ErrorDetails$ = #CR$ + " • " + Str(ItemPosOrIdOrSubmenuHandle) + " is a menu handle."
    ElseIf ItemPosOrIdOrSubmenuHandle > 20 And ItemPosOrIdOrSubmenuHandle < $FFFF
      ErrorDetails$ = #CR$ + " • " + Str(ItemPosOrIdOrSubmenuHandle) + " seems to be an item ID."
    EndIf
    MessageRequester("Error", "Unable to retreive the item " + Str(ItemPosOrIdOrSubmenuHandle) + #CR$ + ErrorDetails$ + #CR$ + " • The search was: " + SearchType$)
  EndIf
  ;
  If *ItemPos
    *ItemPos\i = ItemPos
  EndIf
  ;
  ProcedureReturn hMenu
EndProcedure
;
Procedure   SMI_HasMenuImage(hMenu)
  ; When drawing a menu item from a vertical parent menu
  ; (all menus except the window main one), it is necessary
  ; to know if one of the other items of the parent menu
  ; has an image, in order to reserve the left necessary room
  ; to draw the image.
  Protected LastItem, Counter, Result = #False
  ;
  PushListPosition(MenuItemData())
  ;
  LastItem = SMI_GetLastMenuItem(hMenu)
  For Counter = 0 To LastItem
    SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, Counter, #MF_BYPOSITION, 0, #False)
    If MenuItemData()\MenuItemImgHandle
      Result = #True
      Break
    EndIf
  Next
  ;
  PopListPosition(MenuItemData())
  ;
  ProcedureReturn Result
  ;
EndProcedure
;
Procedure   SMI_HasMenuSubmenu(hMenu)
  ; When drawing a menu item from a vertical parent menu
  ; (all menus except the window main one), it is necessary
  ; to know if one of the other items of the parent menu
  ; has a submenu, in order to reserve the right necessary room
  ; to draw the arrow.
  Protected LastItem, Counter
  ;
  LastItem = SMI_GetLastMenuItem(hMenu)
  For Counter = 0 To LastItem
    If GetSubMenu_(hMenu, Counter)
      ProcedureReturn #True
    EndIf
  Next
  ProcedureReturn #False
  ;
EndProcedure
;
Procedure   SMI_HasMenuOwnerdrawnItems(hMenu, RecursiveSearch = #False)
  ;
  Protected Counter, hSubMenu, Result
  ;
  If IsMenu(hMenu)
    hMenu = MenuID(hMenu)
  EndIf
  ;
  PushListPosition(MenuItemData())
  ;
  Protected LastItem = SMI_GetLastMenuItem(hMenu)
  For Counter = 0 To LastItem
    SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, Counter, #MF_BYPOSITION)
  	If MenuItemData()\MustBeOwnerdrawn = #SMI_Ownerdrawn
  	  Result = #True
  	  Break
  	EndIf
  	;
  	If RecursiveSearch
  	  hSubMenu = GetSubMenu_(hMenu, Counter)
  	  Protected *CurrentElement = @MenuItemData()
      If hSubMenu And SMI_HasMenuOwnerdrawnItems(hSubMenu)
        Result = #True
        Break
      Else
        ChangeCurrentElement(MenuItemData(), *CurrentElement)
      EndIf
    EndIf
  Next
  ;
  PopListPosition(MenuItemData())
  ;
  ProcedureReturn Result
EndProcedure
;
Procedure   SMI_HasMenuItemChecked(hMenu)
  ; When drawing a menu item from a vertical parent menu
  ; (all menus except the window main one), it is necessary
  ; to know if one of the other items of the parent menu
  ; is checked, in order to reserve the left necessary room
  ; to draw the check mark.
  Protected LastItem, Counter, Result = #False
  ;
  PushListPosition(MenuItemData())
  ;
  LastItem = SMI_GetLastMenuItem(hMenu)
  For Counter = 0 To LastItem
    SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, Counter, #MF_BYPOSITION, 0, #False)
    If MenuItemData()\MenuItemState & #MF_CHECKED
      Result = #True
      Break
    EndIf
  Next
  ;
  PopListPosition(MenuItemData())
  ;
  ProcedureReturn Result
  ;
EndProcedure
;
Declare     SMI_OwnerDrawnCallback(hWnd, uMsg, wParam, lParam)
;
Procedure   SMI_AttachCallbackToWindow(WindowID)
  If GetProp_(WindowID, "SMI_OldCallBack") = 0
    Protected SMI_OldCallBack = SetWindowLongPtr_(WindowID, #GWL_WNDPROC, @SMI_OwnerDrawnCallback())
    SetProp_(WindowID, "SMI_ActualCallBack", @SMI_OwnerDrawnCallback())
    SetProp_(WindowID, "SMI_OldCallBack", SMI_OldCallBack)
    SetProp_(WindowID, "SMI_MustBeinitialized", 1)
  EndIf
EndProcedure
;
Procedure   SMI_PrepareItemToBeOwnerDrawn(hMenu = #PB_Default, ItemPos = #PB_Default)
  ;
  ; If the program has not yet entered its main loop and has not started executing
  ; WaitWindowEvent() or WindowEvent(), the item will not really be converted to 
  ; ownerdrawn until the #WM_INITMENU or #WM_INITMENUPOPUP is sent to the callback
  ; procedure, but its data will be prepared to that.
  ; Once the program enters its main loop and starts executing WaitWindowEvent()
  ; or WindowEvent(), then the items are directly converted to ownerdrawing.
  ;
  Protected MenuItemInfo.MENUITEMINFO_Fixed
  Protected LastItem, Counter
  ;
  If hMenu <> #PB_Default And ItemPos <> #PB_Default
    SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPos, #MF_BYPOSITION)
  EndIf
  ;
  MenuItemData()\MustBeOwnerdrawn = #SMI_Ownerdrawn
  ;
  If MenuItemData()\IsItemOwnerdrawn <> #SMI_Ownerdrawn And IsWindow(EventWindow())
    ; The menu item is #SMI_NOT_Ownerdrawn.
    ; It must be set to #SMI_Ownerdrawn to handle colors and/or font.
    ; More than that: due to a Windows bug affecting the ownerdrawn system,
    ; ALL items in this menu must be set as ownerdrawn. To avoid display
    ; issues, a menu must either be entirely non-ownerdrawn or entirely
    ; ownerdrawn. Mixed menus do not work properly.
    PushListPosition(MenuItemData())
    hMenu = MenuItemData()\ParentMenuHandle
    LastItem = SMI_GetLastMenuItem(hMenu)
    For Counter = 0 To LastItem
      SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, Counter, #MF_BYPOSITION)
      ;
      MenuItemData()\IsItemOwnerdrawn = #SMI_Ownerdrawn
      SMI_InitMenuItemInfoData(@MenuItemInfo)
      ; When an item is ownerdrawn, the field hbmpItem of its MenuItemInfo
      ; MUST be set to zero or #HBMMENU_CALLBACK, otherwise, Windows does not send
      ; the #WM_MEASUREITEM message for this item, and it sizes it itself by taking
      ; into account only the size of the image (without taking into account the text).
      ; This field must be set to get a correct display.
      MenuItemInfo\fMask      = #MIIM_FTYPE | #MIIM_BITMAP | #MIIM_DATA
      MenuItemInfo\fType      = #MFT_OWNERDRAW
      MenuItemInfo\hbmpItem   = #HBMMENU_CALLBACK
      If MenuItemData()\IsItemSeparator
        MenuItemInfo\fType | #MFT_SEPARATOR
      EndIf
      If GetSubMenu_(MenuItemData()\ParentMenuHandle, MenuItemData()\MenuItemPos)
        ; When an item is a submenu entry, it's ID should be -1.
        ; But the Windows management of ownerdrawn menus has a very rare but very
        ; annoying bug: when the ID of an item is equal to -1, approximately once
        ; in 200, the #WM_DRAWITEM message is not sent to the callback procedure
        ; and the item is never drawn. When an item falls victim to this bug,
        ; it's for good: you can send DrawMenuBar_ repeatedly, it will never be
        ; drawn. Something in the item data is apparently malfunctioning.
        ; After numerous and exhausting tests, it appears that assigning an ID of 1
        ; to the item corrects this problem. This hack cancels the bug.
        ; Notice that, in reality, the ID will not really be assigned to the item.
        ; If you test this with 
        ; GetMenuItemID_(MenuItemData()\ParentMenuHandle, MenuItemData()\MenuItemPos)
        ; after the operation, you will see that the item still has an ID equal to -1.
        ; But the malfunction is then fixed and the bug will no longer occur.
        ; 
        MenuItemInfo\fMask | #MIIM_ID
        MenuItemInfo\wID  = 1
      EndIf
      MenuItemInfo\dwItemData = @MenuItemData()
      SetMenuItemInfo_(MenuItemData()\ParentMenuHandle, MenuItemData()\MenuItemPos, #MF_BYPOSITION, @MenuItemInfo)
      ;
    Next
    PopListPosition(MenuItemData())
    ;
  EndIf
  ;
  ; Retreive the windows handle of the menu (if any):
  Protected WindowID = SMI_RetreiveMenuWindow(MenuItemData()\ParentMenuHandle)
  If WindowID
    SMI_AttachCallbackToWindow(WindowID)
  EndIf
  ;
EndProcedure
;
Declare     SMI_FreeMenu(hMenu)
Declare     SMI_CloseSubMenu()
;
Procedure   SMI_OwnerDrawnCallback(hWnd, uMsg, wParam, lParam)
  ;
  ; Callback procedure for the main window hosting the gadgets or menu
  ; whose items must be colored or set with a particular font.
  ;
  Protected *drawItem.DRAWITEMSTRUCT
  Protected *measureItem.MEASUREITEMSTRUCT
  Protected hDC, rc.RECT, rc2.RECT, Text$, ImgAddress
  Protected BackColor, TextColor, hBrush, ImgVerticalMargin
  Protected hMenu, ItemNum, ODTType, ObjectTheme
  Protected SelectedFont, BackLuminosity, TextLuminosity
  Protected HasMenuImage, HasCheckedItem, Selected, Disabled, Opacity, CoverColor, hTheme
  Protected ItemThemePart, ItemHotState, ItemNormalState, ItemBackGround, ApplySelectionEffect
  Protected HasSubMenu = 0, ItemPos, menuItemRect.RECT, cont
  Protected MenuItemInfo.MENUITEMINFO_Fixed, TWidth
  Protected *MyMenuItemData.OwnerdrawnMenuItemDataStruct
  Static ItemPosOver = #PB_Default, ItemPosHOTLIGHT = #PB_Default, DisableMainMenu = #WA_ACTIVE
  ;
  If GetProp_(hWnd, "SMI_MustBeinitialized") = 1 And IsWindow(EventWindow()) And WindowID(EventWindow()) = hWnd
    ;
    ; While the command 'WindowEvent()' or 'WaitWindowEvent()' have not been
    ; executed, the value of 'EventWindow()' is not valid.
    ; Therefore, by testing 'IsWindow(EventWindow())', we have the opportunity to know
    ; the precise moment when the main loop of the program begins. This moment is
    ; the good one to make some check or to initialize something.
    ;
    ;
    If GetMenu_(hWnd)
      ; For the main menu of the window, PureBasic send the #WM_INITMENU message when
      ; the menu is created by CreateMenu(). At this moment, our callback procedure is
      ; generally not installed, so, it never receive the #WM_INITMENU for this menu.
      ; The following will resend this message at the moment the program enters into
      ; the main loop and begin to call 'WindowEvent()'.
      ;
      SetProp_(hWnd, "SMI_MustBeinitialized", 2)
      SendMessage_(hWnd, #WM_INITMENU, GetMenu_(hWnd), 0)
    EndIf
  EndIf
  ;
  Select uMsg
      ;
    Case #WM_DESTROY
      ; Clean the memory when the main window is destroyed:
      SMI_FreeMenu(GetMenu_(hWnd))
      ;
    Case #WM_INITMENU, #WM_INITMENUPOPUP
      ; Before #WM_INITMENU or  #WM_INITMENUPOPUP, the menus are not really
      ; set to Ownerdrawn. All their data is ready, but the ownerdrawn flags
      ; of their items is not set, to avoid unnecessary redrawing of items
      ; before all menu settings are completed.
      ; When the program enters in its main loop for the first time and then
      ; execute 'WindowEvent()' or 'WaitWindowEvent()', it's time to finish
      ; the work. SMI_PrepareItemToBeOwnerDrawn(hMenu, 0) will set the
      ; ownerdrawn flag of all items configured for ownerdrawing.
      hMenu = wParam
      If SMI_HasMenuOwnerdrawnItems(hMenu)
        SMI_PrepareItemToBeOwnerDrawn(hMenu, 0)
        BackColor = #PB_Default
        CompilerIf Defined(SetGadgetsColorsFromTheme, #PB_Procedure)
          ; Manage compatibility with the ApplyColorThemes.pbi library
          ; by painting the menu bar with the background color:
          If ListSize(InterfaceColorPresets()) > 0
            BackColor = GetRealColorFromType("BackgroundColor", InterfaceColorPresets()\BackgroundColor)
          EndIf
        CompilerElseIf Defined(ObjectTheme, #PB_Module)
          ; Manage compatibility with the ObjectTheme.pbi library
          ; by painting the menu bar with the background color:
          BackColor = ObjectTheme::GetObjectThemeAttribute(0, #PB_Gadget_BackColor)
        CompilerEndIf
        If BackColor <> #PB_Default
          #MIM_BACKGROUND = $2
          #MIM_APPLYTOSUBMENUS = $80000000
          Protected menuInfo.MENUINFO
          hBrush = CreateSolidBrush_(BackColor)
          menuInfo\cbSize = SizeOf(MENUINFO)
          menuInfo\hbrBack = hBrush
          menuInfo\fMask = #MIM_BACKGROUND ;| #MIM_APPLYTOSUBMENUS
          SetMenuInfo_(hMenu, @menuInfo)
        EndIf
        If hMenu = GetMenu_(hWnd) And GetProp_(hWnd, "SMI_MustBeinitialized")
          DrawMenuBar_(hWnd)
        EndIf
      EndIf
      SetProp_(hWnd, "SMI_MustBeinitialized", 0)
      ;
    Case #WM_ACTIVATE
      ; Check if the window is active or not and redraw menu
      ; to gray or degray it.
      If wParam <> DisableMainMenu
        DisableMainMenu = wParam
        DrawMenuBar_(hwnd)
      EndIf
      ;
    Case #WM_MEASUREITEM
      *measureItem = lParam
      ;
      ODTType = *measureItem\CtlType
      ;
      If ODTType = #ODT_MENU
        ;
        *MyMenuItemData = *measureItem\itemData
        hMenu = *MyMenuItemData\ParentMenuHandle
        ;
        If hMenu
          ;
          If *MyMenuItemData\IsItemSeparator = #False
            ; Compute the width of the text:
            Text$ = PeekS(*MyMenuItemData\MenuItemTextPtr)
            
            If IsFont(*MyMenuItemData\MenuItemFont)
              TWidth = GetTextWidthInWindowContext(hWnd, *MyMenuItemData\MenuItemFont, Text$)
            Else
              TWidth = GetTextWidthInWindowContext(hWnd, DefaultMenuFont, Text$)
            EndIf
            ;
            If SMI_ShowSelectionMethod & #SMI_SSW_Bullet
              TWidth + DesktopScaledX(#SMI_MenuBulletSize)
            EndIf
            ;
        		If hMenu = GetMenu_(hWnd)
              ; If the menu is the main window menu, look if the current
              ; item has an image (because the menu is horizontal and each
              ; item can be independently aligned).
        		  HasMenuImage = *MyMenuItemData\MenuItemImgHandle
        		  If *MyMenuItemData\MenuItemState & #MF_CHECKED
        		    HasCheckedItem = 1
        		  EndIf
        		  TWidth + DesktopScaledX(6) ; Left + Right margins for menu bar items
            Else
              ; If the menu is not the main window menu, look if any of
              ; its items has an image (because the menu is vertical and
              ; all items must be aligned the same way).
              HasMenuImage   = SMI_HasMenuImage(hMenu)
              HasSubMenu     = SMI_HasMenuSubmenu(hMenu)
              HasCheckedItem = SMI_HasMenuItemChecked(hMenu)
              TWidth + DesktopScaledX(18) ; Left + Right margins for normal items
            EndIf
            ;
            If HasMenuImage
              ; If the menu has an image, the width must be increased to draw the image.
              TWidth + GetSystemMetrics_(#SM_CXSMICON) + DesktopScaledX(5)
            EndIf
            ; If the menu has a submenu, it also necessary to increase the width
            ; in order to print the small right triangle indicating the submenu.
            If HasSubMenu  : TWidth + DesktopScaledX(25) : EndIf
           ; If the menu has a checked item, it also necessary to increase the width
           ; in order to print the check mark.
            If HasCheckedItem
              TWidth + DesktopScaledX(15)
              If SMI_ShowSelectionMethod & #SMI_SSW_Bullet
                TWidth + DesktopScaledX(#SMI_MenuBulletSize)
              EndIf
            EndIf
            ;
            If FindString(Text$, #TAB$)
              TWidth + DesktopScaledX(15)
            EndIf
            ;
            ; Save measures:
            *measureItem\itemWidth = TWidth
            *measureItem\itemHeight = GetSystemMetrics_(#SM_CYMENU) + 2
          Else
            ; Mesures for menu separator:
            *measureItem\itemWidth = DesktopScaledX(10)
            *measureItem\itemHeight = DesktopScaledY(8)
          EndIf
          ;
        EndIf
        ;
      EndIf
      ;
    Case #WM_DRAWITEM
      ;- Callback: #WM_DRAWITEM
      If lParam
        *drawItem = lParam
        ItemNum = *drawItem\itemID
        ODTType = *drawItem\CtlType
        If ODTType = #ODT_MENU
          ;
          ; _________________________________________
          ;
          ;              Layout settings
          ;
          hMenu = *drawItem\hwndItem
          *MyMenuItemData = *drawItem\itemData
          ;
          ImgVerticalMargin = (GetSystemMetrics_(#SM_CYMENU) - GetSystemMetrics_(#SM_CYSMICON)) / 2
          If *drawItem\itemState & #ODS_SELECTED Or *drawItem\itemState & #ODS_HOTLIGHT Or ItemPosHOTLIGHT = *MyMenuItemData\MenuItemPos
            If ItemPosHOTLIGHT <> *MyMenuItemData\MenuItemPos
              ItemPosHOTLIGHT = #PB_Default
            EndIf
            Selected = 1
          EndIf
          If DisableMainMenu = #WA_INACTIVE Or *drawItem\itemState & #ODS_GRAYED
            Disabled = 1
            Selected = 0
          EndIf
          ;
          hDC = *drawItem\hdc
          rc = *drawItem\rcItem
          hTheme = OpenThemeData_(#Null, @"MENU")
          If hTheme
            If hMenu = GetMenu_(hWnd)
              ItemThemePart = #MENU_BARITEM
              ItemBackGround = #MENU_BARBACKGROUND
              ItemHotState = #MBI_HOT
              ItemNormalState = #MBI_NORMAL
            Else
              ItemThemePart = #MENU_POPUPITEM
              ItemBackGround = #MENU_POPUPBACKGROUND
              ItemHotState = #MPI_HOT
              ItemNormalState = #MPI_NORMAL
            EndIf
          EndIf
          ;
          Text$ = PeekS(*MyMenuItemData\MenuItemTextPtr)
          ImgAddress = *MyMenuItemData\MenuItemImgHandle
          ;
          If *MyMenuItemData\IsItemSeparator Or Text$ = "" : Disabled = 0 : EndIf
          ;
          If IsFont(*MyMenuItemData\MenuItemFont)
            SelectedFont = FontID(*MyMenuItemData\MenuItemFont)
          Else
            SelectedFont = FontID(DefaultMenuFont)
          EndIf
          SelectObject_(hDC, SelectedFont)
          ;
          ; _________________________________________
          ;
          ; Define colors for text and background:
          ;
          Protected BackGroundMustBePainted = 1
          Protected TextMustBePainted = 1
          ;
          BackColor = *MyMenuItemData\MenuItemBackColor
          If BackColor = #PB_Default
            CompilerIf Defined(SetGadgetsColorsFromTheme, #PB_Procedure)
              ; Manage compatibility with the ApplyColorThemes.pbi library
              If ListSize(InterfaceColorPresets()) > 0
                BackColor = GetRealColorFromType("BackgroundColor", InterfaceColorPresets()\BackgroundColor)
              EndIf
            CompilerElseIf Defined(ObjectTheme, #PB_Module)
              ; Manage compatibility with the ObjectTheme.pbi library:
              BackColor = ObjectTheme::GetObjectThemeAttribute(0, #PB_Gadget_BackColor)
            CompilerEndIf
            If BackColor = #PB_Default
              If hTheme
                BackGroundMustBePainted = DrawThemeBackground_(hTheme, hDC, ItemBackGround, ItemNormalState, @rc, 0)
                ; BackGroundMustBePainted will be equal to zero (#S_OK) if DrawThemeBackground_ works.
                If GetThemeColor_(hTheme, ItemThemePart, ItemNormalState, #TMT_FILLCOLOR, @BackColor) <> #S_OK
                  ; If GetThemeColor_ doesn't work, mesure the color at the center of the drawn rectangle:
                  BackColor = GetPixel_(hDC, (rc\left + rc\right) / 2, (rc\top + rc\bottom) / 2)
                EndIf
              EndIf
              If BackColor = #PB_Default
                ; If the theme functions didn't work, use GetSysColor_:
                If hMenu = GetMenu_(hWnd)
                  BackColor = GetSysColor_(#COLOR_MENUBAR)
                Else
                  BackColor = GetSysColor_(#COLOR_MENU)
                EndIf
              EndIf
            EndIf
          EndIf
          ;
          TextColor = *MyMenuItemData\MenuItemTextColor
          Protected KeepTextColorDefault = 1
          If TextColor = #PB_Default
            CompilerIf Defined(SetGadgetsColorsFromTheme, #PB_Procedure)
              ; Manage compatibility with the ApplyColorThemes.pbi library:
              If ListSize(InterfaceColorPresets()) > 0
                TextColor = GetRealColorFromType("TextColor", InterfaceColorPresets()\TextColor)
                KeepTextColorDefault = 0
              EndIf
            CompilerElseIf Defined(ObjectTheme, #PB_Module)
              ; Manage compatibility with the ObjectTheme.pbi library:
              TextColor = ObjectTheme::GetObjectThemeAttribute(#PB_GadgetType_Button, #PB_Gadget_FrontColor)
            CompilerEndIf
            If TextColor = #PB_Default
              TextColor = GetSysColor_(#COLOR_MENUTEXT)
              If hTheme
                If GetThemeColor_(hTheme, ItemThemePart, ItemNormalState, #TMT_TEXTCOLOR, @TextColor) <> #S_OK
                  ; If GetThemeColor_ doesn't work, use GetSysColor_
                  TextColor = GetSysColor_(#COLOR_MENUTEXT)
                EndIf
              EndIf
            EndIf
          EndIf
          ;
          ; Based on Human perception of color
          TextLuminosity = Red(TextColor)*0.299 + Green(TextColor)*0.587 + Blue(TextColor)*0.114
          BackLuminosity = Red(BackColor)*0.299 + Green(BackColor)*0.587 + Blue(BackColor)*0.114
          ;
          ; _________________________________________
          ;
          ; Paint the background if its not allready painted by DrawThemeBackground_:
          ;
          If BackGroundMustBePainted <> #S_OK
            hBrush = CreateSolidBrush_(BackColor)
            FillRect_(hDC, @rc, hBrush)
            DeleteObject_(hBrush)
          EndIf
          ;
          ; _________________________________________
          ;
          ; Shade the background (eventually):
          ;
          If Selected And (SMI_ShowSelectionMethod & #SMI_SSW_ShadeBackground)
            ;
            If BackLuminosity > 128
              CoverColor = #Black
              Opacity = 40
            Else
              CoverColor = #White
              Opacity = 60
            EndIf
            ;
            DrawTransparentRectangle(hDC, rc, CoverColor, Opacity)
            ;
            If BackLuminosity < 128
              TextColor = #White
            Else
              TextColor = #Black
            EndIf
          EndIf
          ;
          ; _________________________________________
          ;
          ; Draw borders (eventually):
          ;
          If Selected And (SMI_ShowSelectionMethod & #SMI_SSW_Borders)
            hBrush   = CreateSolidBrush_(GetSysColor_(#COLOR_HIGHLIGHT))
            FrameRect_(hDC, @rc, hBrush)
            DeleteObject_(hBrush)
          EndIf
          ;
          ; _________________________________________
          ;
          ; Draw the separation line if the item is a separator:
          ;
          If *MyMenuItemData\IsItemSeparator
            If *MyMenuItemData\MenuItemTextColor <> #PB_Default
              hBrush = CreateSolidBrush_(*MyMenuItemData\MenuItemTextColor)
            Else
              If BackLuminosity > 128
                hBrush = CreateSolidBrush_($A0A0A0)
              Else
                hBrush = CreateSolidBrush_($505050)
              EndIf
            EndIf
            CopyMemory(@rc, @rc2, SizeOf(RECT))
            rc2\top = (rc2\top + rc2\bottom) / 2
            rc2\bottom = rc2\top + 1
            rc2\left + DesktopScaledX(7)
            rc2\right - DesktopScaledX(7)
            FillRect_(hDC, @rc2, hBrush)
            DeleteObject_(hBrush)
          EndIf
          ; _________________________________________
          ;
          ; Apply 'HotState' when SMI_ShowSelectionMethod & #SMI_SSW_SystemColor
          ;
          If Selected And (SMI_ShowSelectionMethod & #SMI_SSW_SystemColor)
            ApplySelectionEffect = 1
            If hTheme And (*MyMenuItemData\MenuItemBackColor = #PB_Default Or BackLuminosity > 128)
              ApplySelectionEffect = DrawThemeBackground_(hTheme, hDC, ItemThemePart, ItemHotState, @rc,0)
            EndIf
            If ApplySelectionEffect <> #S_OK
              If hMenu = GetMenu_(hWnd)
                Opacity = 40
              Else
                Opacity = 100
              EndIf
              CoverColor = GetSysColor_(#COLOR_HIGHLIGHT)
              If BackLuminosity < 128
                Opacity = 150
              EndIf
              ;
              DrawTransparentRectangle(hDC, rc, CoverColor, Opacity)
            EndIf
          EndIf
          ;
          ; _________________________________________
          ;
          ; Draw a checkmark (eventually):
          ;
          If *MyMenuItemData\MenuItemState & #MF_CHECKED
            CopyMemory(@rc, @rc2, SizeOf(RECT))
            rc2\left + DesktopScaledX(3)
            rc2\right = rc2\left + DesktopScaledX(16)
           ; DrawThemeBackground_(hTheme, hDC, #MENU_POPUPCHECK, #MC_CHECKMARKNORMAL, @rc2, 0)
            ;DrawFrameControl_(hDC, @rc2, #DFC_MENU, #DFCS_MENUCHECK)
            DrawCheckmark(hDC, @rc2, DesktopScaledX(10), TextColor)
          EndIf
          ;
          If *MyMenuItemData\MenuItemState & #MF_CHECKED Or (hMenu <> GetMenu_(hWnd) And SMI_HasMenuItemChecked(hMenu))
            rc\left + DesktopScaledX(18)
          EndIf
          ;
          ; _________________________________________
          ;
          ; Draw a bullet (eventually):
          ;
          If SMI_ShowSelectionMethod & #SMI_SSW_Bullet
            rc\Left + DesktopScaledX(5)
            If Selected
              DrawRightPointingTriangle(hDC, rc, DesktopScaledX( #SMI_MenuBulletSize), TextColor, BackColor)
              rc\left + DesktopScaledX(#SMI_MenuBulletSize + 5)
            Else
              rc\left + DesktopScaledX(#SMI_MenuBulletSize + 5) / 2
            EndIf
          Else
            If hMenu = GetMenu_(hWnd)
              rc\Left + DesktopScaledX(8)
            Else
              rc\Left + DesktopScaledX(6)
            EndIf
          EndIf
          ;
          ; _________________________________________
          ;
          ; Draw an icon (eventually):
          ;
          If ImgAddress ; Image illustrating the menu line
            ;
            Protected icone = CreateIconFromImage(ImgAddress)
            If icone <> 0
              DrawIconEx_(hDC, rc\left, rc\top + ImgVerticalMargin, icone, 0, 0, 0, #Null, #DI_NORMAL)
            EndIf
            ;
          EndIf
          ;
          If ImgAddress Or (hMenu <> GetMenu_(hWnd) And SMI_HasMenuImage(hMenu))
            rc\left + GetSystemMetrics_(#SM_CXSMICON) + DesktopScaledX(5)
          EndIf
          ;
          ; _________________________________________
          ;
          ; Draw the item text:
          ;
          rc\Left + DesktopScaledX(2)
          ;
          Protected PosTab = FindString(Text$, #TAB$)
          If PosTab
            Protected AfterTab$ = Mid(Text$, PosTab + 1)
            Text$ = Left(Text$, PosTab - 1)
            rc\right - DesktopScaledX(7)
          EndIf
          If hTheme And TextColor = #PB_Default And KeepTextColorDefault
            TextMustBePainted = DrawThemeText_(hTheme, hDC, ItemThemePart, ItemNormalState, @Text$, Len(Text$),#DT_LEFT | #DT_VCENTER | #DT_SINGLELINE,0,@rc)
            If TextMustBePainted = #S_OK And AfterTab$
              DrawThemeText_(hTheme, hDC, ItemThemePart, ItemNormalState, @AfterTab$, Len(AfterTab$),#DT_RIGHT | #DT_VCENTER | #DT_SINGLELINE,0,@rc)
            EndIf
          EndIf
          If TextMustBePainted <> #S_OK
            SetTextColor_(hDC, TextColor)
            SetBkMode_(hDC, #TRANSPARENT)
            DrawText_(hDC, Text$, Len(Text$), @rc, #DT_LEFT | #DT_VCENTER | #DT_SINGLELINE)
            If AfterTab$
              DrawText_(hDC, AfterTab$, Len(AfterTab$), @rc, #DT_RIGHT | #DT_VCENTER | #DT_SINGLELINE)
            EndIf
          EndIf
          ;
          ; _________________________________________
          ;
          ; Gray the item if it is Disabled:
          ;
          If Disabled
            ; Item is grayed. Recover the drawing with a semi-transparent rectangle:
            If BackLuminosity > 128
              CoverColor = #White
            Else
              CoverColor = #Black
            EndIf
            rc = *drawItem\rcItem
            DrawTransparentRectangle(hDC, rc, CoverColor, 128)
          EndIf
          ;
          If hTheme
            CloseThemeData_(hTheme)
          EndIf
          ;
        EndIf
        ;
      EndIf
    Case #WM_NCMOUSEMOVE
      ; In modern versions of Windows, menu bar items are activated when the mouse cursor
      ; hovers over them, providing a smooth and intuitive user experience.
      ; Unfortunately, this behavior does not apply to owner-drawn menu items, which are
      ; managed in the "old-fashioned" Windows way. These items do not react to mouse hover
      ; events, and no #WM_DRAWITEM message is sent to the window's callback procedure
      ; in such cases. The following code addresses this limitation by explicitly
      ; triggering a #WM_DRAWITEM message when the mouse cursor hovers over a menu item,
      ; ensuring consistent behavior for owner-drawn menus.
      hMenu = GetMenu_(hwnd)
      If hMenu ; Checks if the window has a menu
        ; Gets the coordinates of the menu bar:
        Protected mbi.MENUBARINFO\cbSize = SizeOf(MENUBARINFO)
        GetMenuBarInfo_(hwnd, $FFFFFFFD, 0, @mbi) ; #OBJID_MENU = $FFFFFFFD
        If mbi And PtInRect_(@mbi\rcBar, ((lParam & $FFFF0000) << 16) + (lParam & $FFFF))
          ; The mouse cursor is over the menu bar:
          Cont = 1
          If ItemPosOver <> 0 And GetMenuItemRect_(hwnd, hMenu, ItemPosOver, @menuItemRect)
            If PtInRect_(@menuItemRect, ((lParam & $FFFF0000) << 16) + (lParam & $FFFF))
              ; The cursor is on the same item as in the previous test.
              Cont = 0
            EndIf
          EndIf
          If Cont
            Protected LastMenuItem = SMI_GetLastMenuItem(hMenu)
            ; Iterates through the menu items to check which one is hovered over
            For ItemPos = 0 To LastMenuItem
              If ItemPos <> ItemPosOver
                GetMenuItemRect_(hwnd, hMenu, ItemPos, @menuItemRect)
                ; If the mouse is within the boundaries of this item
                If PtInRect_(@menuItemRect, ((lParam & $FFFF0000) << 16) + (lParam & $FFFF))
                  ItemPosOver = ItemPos
                  ; Checks if the item is ownerdrawn:
                  SMI_InitMenuItemInfoData(@MenuItemInfo.MENUITEMINFO_Fixed)
                  MenuItemInfo\fMask      = #MIIM_FTYPE
                	GetMenuItemInfo_(hMenu, ItemPos, #MF_BYPOSITION, @MenuItemInfo)
                	If MenuItemInfo\fType & #MFT_OWNERDRAW
                	  ItemPosHOTLIGHT = ItemPos
                	  DrawMenuBar_(hWnd)
                	ElseIf ItemPosHOTLIGHT > #PB_Default
                	  ; If the mouse cursor has left the area of an ownerdrawn item
                    ; and the item is still selected, deselect it:
                	  ItemPosHOTLIGHT = #PB_Default
                	  DrawMenuBar_(hWnd)
                  EndIf
                EndIf
              EndIf
            Next
          EndIf
        Else
          ItemPosOver = #PB_Default
          ; If the mouse cursor has left the menu bar area while
          ; leaving an item selected, deselect it:
          If ItemPosHOTLIGHT > #PB_Default
        	  ItemPosHOTLIGHT = #PB_Default
        	  DrawMenuBar_(hWnd)
        	EndIf
        EndIf
      EndIf
      ;
    Case #WM_MOUSEMOVE, #WM_NCLBUTTONDOWN
      ; If the mouse cursor has quickly left the menu bar area while
      ; leaving an item selected, or if the user click onto a menuitem,
      ; deselect the old one.
      ItemPosOver = #PB_Default
      If ItemPosHOTLIGHT > #PB_Default
        ItemPosHOTLIGHT = #PB_Default
    	  DrawMenuBar_(hWnd)
    	EndIf

  EndSelect
  ;
  ; Normal callback for all other messages:
  Protected SMI_OldCallBack = GetProp_(hWnd, "SMI_OldCallBack")
  ProcedureReturn CallWindowProc_(SMI_OldCallBack, hWnd, uMsg, wParam, lParam)
EndProcedure
;
Procedure   SMI_RedrawMenuIfMainMenu(hMenu)
  ;
  ; Update the menu, If it is the window main menu:
  ;
  If IsMenu(hMenu)
    hMenu = MenuID(hMenu)
  EndIf
  If IsWindow(EventWindow()) And GetMenu_(WindowID(EventWindow())) = hMenu
    DrawMenuBar_(WindowID(EventWindow()))
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure   SMI_ResizeMenu(hMenu, RecursiveSearch = #False)
  ;
  ; When an image is added to or removed from an item of the main menu,
  ; or when the font or the text of the item is modified, it is necessary
  ; to resize it, and not only to redraw it.
  ; By calling SetMenuItemInfo_ with fMask = #MIIM_BITMAP and
  ; hbmpItem = #HBMMENU_CALLBACK, we force a #WM_MEASUREITEM to be sent
  ; to the item.
  ;
  Protected MenuItemInfo.MENUITEMINFO_Fixed
  Protected ItemPos, LastMenuItem, Counter, hSubMenu
  ;
  If IsMenu(hMenu)
    hMenu = MenuID(hMenu)
  EndIf
  ;
  If hMenu
    If IsWindow(EventWindow())
      LastMenuItem  = SMI_GetLastMenuItem(hMenu)
      ; Force item(s) to be redrawn:
      For Counter = 0 To LastMenuItem
        If SMI_HasMenuOwnerdrawnItems(hMenu)
          SMI_InitMenuItemInfoData(@MenuItemInfo)
          MenuItemInfo\fMask    = #MIIM_BITMAP
          MenuItemInfo\hbmpItem = #HBMMENU_CALLBACK
          SetMenuItemInfo_(hMenu, Counter, #MF_BYPOSITION, @MenuItemInfo)
        EndIf
        hSubMenu = GetSubMenu_(hMenu, Counter)
        If RecursiveSearch And hSubMenu
          SMI_ResizeMenu(hSubMenu, RecursiveSearch)
        EndIf
      Next
      If GetMenu_(WindowID(EventWindow())) = hMenu
        DrawMenuBar_(WindowID(EventWindow()))
      EndIf
      ProcedureReturn #True
    EndIf
  EndIf
EndProcedure
;
; *****************************************************************************
;
;-                         4. LIBRARY NEW FUNCTIONS
;
;          Functions added to the native set of PureBasic functions
;
; *****************************************************************************
;
Procedure   ApplyThemesToMenu(hMenu, RecursiveSearch = #True)
  ;
  ; Will force a menu to be ownerdrawn, even if none of its items
  ; uses a particular font or color. This will allow the library
  ; to apply color themes from 'ObjectTheme.pbi' or 'ApplyColorThemes.pbi'
  ; to the menus.
  ;
  Protected ItemPos, LastMenuItem, Counter, hSubMenu
  ;
  If IsMenu(hMenu)
    hMenu = MenuID(hMenu)
  EndIf
  ;
  If hMenu
    LastMenuItem  = SMI_GetLastMenuItem(hMenu)
    ; Force item(s) to be redrawn:
    For Counter = 0 To LastMenuItem
      SMI_PrepareItemToBeOwnerDrawn(hMenu, Counter)
      hSubMenu = GetSubMenu_(hMenu, Counter)
      If RecursiveSearch And hSubMenu
        ApplyThemesToMenu(hSubMenu, RecursiveSearch)
      EndIf
    Next
    If IsWindow(EventWindow()) And GetMenu_(WindowID(EventWindow())) = hMenu
      DrawMenuBar_(WindowID(EventWindow()))
    EndIf
    ProcedureReturn #True
  EndIf
EndProcedure
;
Procedure SetMenuItemImage(hMenu, ItemPosOrIdOrSubmenuHandle, *ItemImagePtr, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  ;
  Protected MenuItemInfo.MENUITEMINFO_Fixed
  Protected ItemPos, Result
  ;
  If *ItemImagePtr <> #PB_Default
    PushListPosition(MenuItemData())
    ;
    hMenu = SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, @ItemPos)
    ;
    If IsImage(*ItemImagePtr)
      *ItemImagePtr = ImageID(*ItemImagePtr)
    EndIf
    ;
    If hMenu
      ;
      If IsImage(MenuItemData()\MenuItemImgNum)
        FreeImage(MenuItemData()\MenuItemImgNum)
      EndIf
      MenuItemData()\MenuItemImgHandle = 0
      ;
      If *ItemImagePtr
        MenuItemData()\MenuItemImgNum = ResizeImageToIconSize(*ItemImagePtr)
        If MenuItemData()\MenuItemImgNum
          MenuItemData()\MenuItemImgHandle = ImageID(MenuItemData()\MenuItemImgNum)
        EndIf
      EndIf
      ;
      If MenuItemData()\IsItemOwnerdrawn = #SMI_NOT_Ownerdrawn
        ; Records information using classic Windows API functions.
        ; The item is automatically redrawn.
        SMI_InitMenuItemInfoData(@MenuItemInfo)
        MenuItemInfo\fMask      = #MIIM_BITMAP
        MenuItemInfo\hbmpItem   = MenuItemData()\MenuItemImgHandle
        SetMenuItemInfo_(hMenu, ItemPos, #MF_BYPOSITION, @MenuItemInfo)
        ;
      ElseIf MenuItemData()\MustBeOwnerdrawn = #SMI_Ownerdrawn
        ; The item is NOT automatically resized and redrawn.
        SMI_ResizeMenu(hMenu)
      EndIf
      ;
      Result = #True
    EndIf
    ;
    PopListPosition(MenuItemData())
  EndIf
  ;
  ProcedureReturn Result
EndProcedure
;
Procedure SetMenuTitleImage(hMenu, ItemPos, *ItemImagePtr)
  ProcedureReturn SetMenuItemImage(hMenu, ItemPos, *ItemImagePtr, #MF_BYPOSITION)
EndProcedure
;
Procedure GetMenuItemImage(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  PushListPosition(MenuItemData())
  ;
  If SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle)
    Protected Result = MenuItemData()\MenuItemImgHandle
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
Procedure GetMenuTitleImage(hMenu, ItemPos)
  ProcedureReturn GetMenuItemImage(hMenu, ItemPos, #MF_BYPOSITION)
EndProcedure
;
Procedure SetMenuItemColor(hMenu, ItemPosOrIdOrSubmenuHandle, ColorType, ItemColor, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  ;
  ; This function is similar to the PureBasic native function SetMenuItemText()
  ; except it manages the case of program's ownerdrawn menu item texts.
  ;
  Protected ItemPos, Result
  ;
  PushListPosition(MenuItemData())
  ;
  hMenu = SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, @ItemPos)
  ;
  If hMenu
    ;
    SMI_PrepareItemToBeOwnerDrawn()
    ;
    If ColorType = #PB_Gadget_FrontColor
      MenuItemData()\MenuItemTextColor = ItemColor
    Else
      MenuItemData()\MenuItemBackColor = ItemColor
    EndIf
    ;
    ; Now, update the menu, if it is the window main menu:
    SMI_RedrawMenuIfMainMenu(hMenu)
    ;
    Result = #True
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
Procedure SetMenuTitleColor(hMenu, ItemPos, ColorType, ItemColor)
  ProcedureReturn SetMenuItemColor(hMenu, ItemPos, ColorType, ItemColor, #MF_BYPOSITION)
EndProcedure
;
Procedure GetMenuItemColor(hMenu, ItemPosOrIdOrSubmenuHandle, ColorType, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  PushListPosition(MenuItemData())
  ;
  If SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle)
    If ColorType = #PB_Gadget_FrontColor
      Protected Result = MenuItemData()\MenuItemTextColor
    Else
      Result = MenuItemData()\MenuItemBackColor
    EndIf
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
Procedure GetMenuTitleColor(hMenu, ItemPos, ColorType)
  ProcedureReturn GetMenuItemColor(hMenu, ItemPos, ColorType, #MF_BYPOSITION)
EndProcedure
;
Procedure SetMenuColor(hMenu, ColorType, ItemColor, RecursiveSearch = #True)
  ;
  ; Apply a color to all existing items of a menu
  ;
  Protected LastItem, Counter, hSubMenu
  If IsMenu(hMenu)
    hMenu = MenuID(hMenu)
  EndIf
  LastItem = SMI_GetLastMenuItem(hMenu)
  For Counter = 0 To LastItem
    SetMenuItemColor(hMenu, Counter, ColorType, ItemColor, #MF_BYPOSITION)
    hSubMenu = GetSubMenu_(hMenu, Counter)
    If RecursiveSearch And hSubMenu
      SetMenuColor(hSubMenu, ColorType, ItemColor, #True)
    EndIf
  Next
  If IsWindow(EventWindow())
    SetProp_(WindowID(EventWindow()), "SMI_MustBeinitialized", 1)
  EndIf
EndProcedure
;
Procedure SetMenuItemFont(hMenu, ItemPosOrIdOrSubmenuHandle, Font, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  ;
  ; This function is similar to the PureBasic native function SetMenuItemText()
  ; except it manages the case of program's ownerdrawn menu item texts.
  ;
  Protected ItemPos, Result
  ;
  PushListPosition(MenuItemData())
  ;
  hMenu = SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, @ItemPos)
  ;
  If hMenu
    ;
    SMI_PrepareItemToBeOwnerDrawn()
    ;
    MenuItemData()\MenuItemFont = Font
    ;
    ; Now, update the menu:
    SMI_ResizeMenu(hMenu)
    ;
    Result = #True
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
Procedure SetMenuTitleFont(hMenu, ItemPosOrIdOrSubmenuHandle, Font)
  ProcedureReturn SetMenuItemFont(hMenu, ItemPosOrIdOrSubmenuHandle, Font, #MF_BYPOSITION)
EndProcedure
;
Procedure GetMenuItemFont(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  PushListPosition(MenuItemData())
  ;
  If SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle)
    Protected Result = MenuItemData()\MenuItemFont
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
Procedure GetMenuTitleFont(hMenu, ItemPosOrIdOrSubmenuHandle)
  ProcedureReturn GetMenuItemFont(hMenu, ItemPosOrIdOrSubmenuHandle, #MF_BYPOSITION)
EndProcedure
;
Procedure SetMenuFont(hMenu, Font, RecursiveSearch = #True)
  ;
  ; Apply a font to all existing items of a menu
  ;
  Protected LastItem, Counter, hSubMenu
  If IsMenu(hMenu)
    hMenu = MenuID(hMenu)
  EndIf
  LastItem = SMI_GetLastMenuItem(hMenu)
  For Counter = 0 To LastItem
    SetMenuItemFont(hMenu, Counter, Font, #MF_BYPOSITION)
    hSubMenu = GetSubMenu_(hMenu, Counter)
    If RecursiveSearch And hSubMenu
      SetMenuColor(hSubMenu, Font, #True)
    EndIf
  Next
  If IsWindow(EventWindow())
    SetProp_(WindowID(EventWindow()), "SMI_MustBeinitialized", 1)
  EndIf
EndProcedure
;
Procedure CheckMenuItem(hMenu, ItemPosOrIdOrSubmenuHandle, State = #True, Style = #PB_Default, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  ;
  Protected MenuItemInfo.MENUITEMINFO_Fixed
  Protected ItemPos, Result
  ;
  PushListPosition(MenuItemData())
  ;
  hMenu = SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, @ItemPos)
  ;
  If hMenu
    ;
    If State
      MenuItemData()\MenuItemState | #MF_CHECKED
      CheckMenuItem_(hMenu, ItemPos, #MF_BYPOSITION | #MF_CHECKED)
    Else
      MenuItemData()\MenuItemState & ~#MF_CHECKED
      CheckMenuItem_(hMenu, ItemPos, #MF_BYPOSITION | #MF_UNCHECKED)
    EndIf
    ;
    SMI_ResizeMenu(hMenu)
    ;
    Result = #True
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
Procedure IsMenuItemChecked(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  PushListPosition(MenuItemData())
  ;
  If SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle)
    If MenuItemData()\MenuItemState & #MF_CHECKED
      Protected Result = #True
    EndIf
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
;
; *****************************************************************************
;
;-         5. FUNCTIONS THAT WILL OVERRIDE PUREBASIC NATIVE FUNCTIONS
;
; *****************************************************************************
;
Procedure.s SMI_GetMenuItemText(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  PushListPosition(MenuItemData())
  ;
  If SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle)
    Protected Result$ = PeekS(MenuItemData()\MenuItemTextPtr)
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result$
EndProcedure

Procedure.s SMI_GetMenuTitleText(hMenu, ItemPos)
  ProcedureReturn SMI_GetMenuItemText(hMenu, ItemPos, #MF_BYPOSITION)
EndProcedure
;
Procedure   SMI_SetMenuItemText(hMenu, ItemPosOrIdOrSubmenuHandle, ItemText$, ByPosOrIdOrSubMenuHandle = #MF_BYCOMMAND)
  ;
  Protected MenuItemInfo.MENUITEMINFO_Fixed
  Protected ItemPos, Result
  ;
  PushListPosition(MenuItemData())
  ;
  hMenu = SMI_RetreiveOrCreateMenuItemDataListElement(hMenu, ItemPosOrIdOrSubmenuHandle, ByPosOrIdOrSubMenuHandle, @ItemPos)
  ;
  If hMenu
    ;
    SysFreeString_(MenuItemData()\MenuItemTextPtr)
    MenuItemData()\MenuItemTextPtr = SysAllocString_(ItemText$)
    ;
    If MenuItemData()\IsItemOwnerdrawn = #SMI_NOT_Ownerdrawn
      ; Records information using classic Windows API functions.
      ; The item is automatically redrawn.
      SMI_InitMenuItemInfoData(@MenuItemInfo)
      MenuItemInfo\fMask      = #MIIM_STRING
      MenuItemInfo\dwTypeData = @ItemText$
      SetMenuItemInfo_(hMenu, ItemPos, #MF_BYPOSITION, @MenuItemInfo)
      ;
    ElseIf MenuItemData()\MustBeOwnerdrawn = #SMI_Ownerdrawn
      ; The item is NOT automatically redrawn.
      ; Update the menu.
      SMI_ResizeMenu(hMenu)
    EndIf
    ;
    Result = #True
  EndIf
  ;
  PopListPosition(MenuItemData())
  ProcedureReturn Result
EndProcedure
;
Procedure   SMI_SetMenuTitleText(hMenu, ItemPos, ItemText$)
  ProcedureReturn SMI_SetMenuItemText(hMenu, ItemPos, ItemText$, #MF_BYPOSITION)
EndProcedure
;
Procedure   SMI_CloseSubMenu()
  ;
  Shared SMI_MemMenuNum, SMI_MemMenuHandle
  Protected hMenu, ItemPos
  ;
  ; Finds the menu to which the submenu belongs:
  If IsMenu(SMI_MemMenuNum) And SMI_MemMenuHandle And MenuID(SMI_MemMenuNum) <> SMI_MemMenuHandle
    hMenu = SMI_GetMenuItemPos(SMI_MemMenuNum, SMI_MemMenuHandle, #MF_BYCOMMAND, @ItemPos)
    If hMenu
      ; Updates SMI_MemMenuHandle with the top menu handle:
      CloseSubMenu()
      SMI_MemMenuHandle = hMenu
    EndIf
    ProcedureReturn #True
  Else
    If IsMenu(SMI_MemMenuNum)
      SMI_MemMenuHandle = MenuID(SMI_MemMenuNum)
    EndIf
  EndIf
  ;
EndProcedure
;
Procedure   SMI_OpenSubMenu(ItemText$, ItemImage)
  ;
  Shared SMI_MemMenuHandle
  ;
  Protected Result = OpenSubMenu(ItemText$)
  If ItemImage <> #PB_Default
    SetMenuItemImage(SMI_MemMenuHandle, Result, ItemImage)
  Else
    SMI_RetreiveOrCreateMenuItemDataListElement(SMI_MemMenuHandle, Result)
  EndIf
  ;
  If SMI_HasMenuOwnerdrawnItems(SMI_MemMenuHandle)
    SMI_PrepareItemToBeOwnerDrawn(SMI_MemMenuHandle, SMI_GetLastMenuItem(SMI_MemMenuHandle))
  EndIf
  ;
  SMI_MemMenuHandle = Result
  ;
  ProcedureReturn Result
EndProcedure
;
Procedure   SMI_MenuTitle(ItemText$, ItemImage)
  ; The native PureBasic MenuTitle() function can't be used with PopupMenu,
  ; it cannot has an image as second parameter and it doesn't return
  ; the created submenu handle. This function cancels this limitations.
  ;
  Shared SMI_MemMenuNum, SMI_MemMenuHandle
  ;
  While SMI_CloseSubMenu() : Wend
  ProcedureReturn SMI_OpenSubMenu(ItemText$, ItemImage)
EndProcedure
;
Procedure   SMI_CreateMenu(MenuNum, WindowID)
  ;
  ; This function intercept the CreateMenu() and CreateImageMenu()
  ; calls to register the menu PureBasic number (MenuNum. It also simplify
  ; the menu management by replacing CreateImageMenu() by CreateMenu().
  ;
  Shared SMI_MemMenuNum, SMI_MemMenuHandle
  Protected Result
  ;
  While SMI_CloseSubMenu() : Wend
  ;
  If MenuNum = #PB_Any
    Result = CreateMenu(#PB_Any, WindowID)
    MenuNum = Result
  Else
    Result = CreateMenu(MenuNum, WindowID)
  EndIf
  ;
  SMI_MemMenuNum = MenuNum
  SMI_MemMenuHandle = MenuID(MenuNum)
  
  SMI_SaveMenuWindow(SMI_MemMenuHandle, WindowID)
  ;
  ProcedureReturn Result
EndProcedure
;
Procedure   SMI_CreatePopUpMenu(MenuNum)
  ;
  ; This function intercept the CreatePopupMenu() and CreatePopupImageMenu()
  ; calls to register the menu PureBasic number. It also simplify
  ; the menu management by replacing CreatePopupImageMenu() by CreatePopupMenu().
  ;
  Shared SMI_MemMenuNum, SMI_MemMenuHandle
  Protected Result
  ;
  While SMI_CloseSubMenu() : Wend
  ;
  If MenuNum = #PB_Any
    Result = CreatePopupMenu(#PB_Any)
    MenuNum = Result
  Else
    Result = CreatePopupMenu(MenuNum)
  EndIf
  ;
  SMI_MemMenuNum = MenuNum
  SMI_MemMenuHandle = MenuID(MenuNum)
  ;
  ProcedureReturn Result
EndProcedure
;
Procedure   SMI_MenuItem(ItemID, ItemText$, ItemImage)
  ;
  Shared SMI_MemMenuHandle
  ;
  SMI_CheckID(@ItemID)
  ;
  MenuItem(ItemID, ItemText$)
  ;
  If ItemImage <> #PB_Default
    SetMenuItemImage(SMI_MemMenuHandle, ItemID, ItemImage)
  Else
    SMI_RetreiveOrCreateMenuItemDataListElement(SMI_MemMenuHandle, ItemID)
  EndIf
  ;
  ; Set the item to ownerdrawn if its menu is allready ownerdrawn:
  If SMI_HasMenuOwnerdrawnItems(SMI_MemMenuHandle)
    SMI_PrepareItemToBeOwnerDrawn(SMI_MemMenuHandle, SMI_GetLastMenuItem(SMI_MemMenuHandle))
  EndIf
  ProcedureReturn ItemID
EndProcedure
;
Procedure   SMI_MenuBar(ItemID)
  ;
  Shared SMI_MemMenuHandle
  Protected MenuItemInfo.MENUITEMINFO_Fixed, ItemPos
  ;
  MenuBar()
  ;
  ItemPos = SMI_GetLastMenuItem(SMI_MemMenuHandle)
  ;
  SMI_CheckID(@ItemID)
  ;
  ; The ID must not be null if we want the separator-item to be drawn
  ; as an ownerdrawn items.
  SMI_InitMenuItemInfoData(@MenuItemInfo)
  MenuItemInfo\fMask = #MIIM_ID
  MenuItemInfo\wID = ItemID
  SetMenuItemInfo_(SMI_MemMenuHandle, ItemPos, #MF_BYPOSITION, @MenuItemInfo)
  ;
  ; Set the menubar item to ownerdrawn if its menu is allready ownerdrawn:
  If SMI_HasMenuOwnerdrawnItems(SMI_MemMenuHandle)
    SMI_PrepareItemToBeOwnerDrawn(SMI_MemMenuHandle, ItemPos)
  EndIf
  ProcedureReturn ItemID
EndProcedure
;
Procedure   SMI_FreeMenu(hMenu)
  ;
  Shared SMI_MemMenuNum, SMI_MemMenuHandle
  ;
  While SMI_CloseSubMenu() : Wend
  ;
  If IsMenu(hMenu)
    Protected MenuNum = hMenu
    hMenu = MenuID(hMenu)
  Else
    MenuNum = #PB_Default
  EndIf
  ;
  ForEach MenuItemData()
    If MenuItemData()\ParentMenuHandle = hMenu
      If IsImage(MenuItemData()\MenuItemImgNum)
        FreeImage(MenuItemData()\MenuItemImgNum)
      EndIf
      SysFreeString_(MenuItemData()\MenuItemTextPtr)
      Protected hSubMenu = GetSubMenu_(hMenu, MenuItemData()\MenuItemPos)
      If hSubMenu
        Protected *CurrentElement = @MenuItemData()
        SMI_FreeMenu(hSubMenu)
        ChangeCurrentElement(MenuItemData(), *CurrentElement)
      EndIf
      DeleteElement(MenuItemData())
    EndIf
  Next
  If MenuNum <> #PB_Default
    FreeMenu(MenuNum)
  EndIf
  ;
EndProcedure
;
Procedure   SMI_DisplayPopupMenu(MenuNum, WindowID, X, Y)
  ;
  SMI_SaveMenuWindow(MenuID(MenuNum), WindowID)
  ;
  If SMI_HasMenuOwnerdrawnItems(MenuNum, #True)
    ; Set the callback procedure of the window showing the menu:
    SMI_AttachCallbackToWindow(WindowID)
  EndIf
  ;
  If X = #PB_Ignore Or Y = #PB_Ignore
    ProcedureReturn DisplayPopupMenu(MenuNum, WindowID)
  Else
    ProcedureReturn DisplayPopupMenu(MenuNum, WindowID, X, Y)
  EndIf
EndProcedure
;
;
; *****************************************************************************
;
;-                   6. OVERRIDE PUREBASIC NATIVE FUNCTIONS
;
; The native PureBasic functions dedicated to menu management are not designed
; to work in a context where not all menu items are handled the same way
; (some being owner-drawn and others not).
; They also encounter several issues when an image is assigned to a menu bar item:
; • GetMenuItemText() And GetMenuTitleText() are then unable to retrieve the
;   item's text, while SetMenuItemText() and SetMenuTitleText() erase the image
;   assigned to the item.
; • Simply displaying a popup menu created With CreatePopupImageMenu() can
;   disrupt the main menu's data.
;
; Consequently, to allow the use of these functions in the usual way while
; benefiting from the additional features of this library, the following functions
; have been overridden (i.e., replaced) by the library's functions.
; For example, when you include GetMenuItemText() in your program, it is actually
; SMI_GetMenuItemText() that will be called.
; This replacement is transparent for the library user.
;
; *****************************************************************************
;
;
Macro GetMenuItemText(MenuNum, MenuPos)
  ; Override GetMenuItemText
  SMI_GetMenuItemText(MenuNum, MenuPos)
EndMacro
Macro SetMenuItemText(MenuNum, MenuPos, MenuItemText)
  ; Override SetMenuItemText
  SMI_SetMenuItemText(MenuNum, MenuPos, MenuItemText)
EndMacro
Macro GetMenuTitleText(MenuNum, MenuPos)
  ; Override GetMenuTitleText
  SMI_GetMenuTitleText(MenuNum, MenuPos)
EndMacro
Macro SetMenuTitleText(MenuNum, MenuPos, MenuItemText)
  ; Override SetMenuTitleText
  SMI_SetMenuTitleText(MenuNum, MenuPos, MenuItemText)
EndMacro
Macro OpenSubMenu(MenuNum, MenuItemImage = #PB_Default)
  ; Override OpenSubMenu
  SMI_OpenSubMenu(MenuNum, MenuItemImage)
EndMacro
Macro CloseSubMenu()
  ; Override CloseSubMenu
  SMI_CloseSubMenu()
EndMacro
Macro MenuTitle(MenuNum, MenuItemImage = #PB_Default)
  ; Override MenuTitle
  SMI_MenuTitle(MenuNum, MenuItemImage)
EndMacro
Macro MenuItem(MenuItemID, MenuItemText, MenuItemImage = #PB_Default)
  ; Override MenuItem
  SMI_MenuItem(MenuItemID, MenuItemText, MenuItemImage)
EndMacro
Macro MenuBar(MenuItemID = #PB_Any)
  ; Override MenuBar
  SMI_MenuBar(MenuItemID)
EndMacro
Macro CreateMenu(MenuNum, WindowID)
  ; Override CreateMenu
  SMI_CreateMenu(MenuNum, WindowID)
EndMacro
Macro CreateImageMenu(MenuNum, WindowID, Options = 0)
  ; Override CreateImageMenu
  SMI_CreateMenu(MenuNum, WindowID)
EndMacro
Macro CreatePopupMenu(MenuNum)
  ; Override CreatePopupMenu
  SMI_CreatePopupMenu(MenuNum)
EndMacro
Macro CreatePopupImageMenu(MenuNum, Options = 0)
  ; Override CreatePopupImageMenu
  SMI_CreatePopupMenu(MenuNum)
EndMacro
Macro FreeMenu(MenuNum)
  ; Override FreeMenu
  SMI_FreeMenu(MenuNum)
EndMacro
Macro DisplayPopupMenu(MenuNum, WindowID, X = #PB_Ignore, Y = #PB_Ignore)
  ; Override DisplayPopupMenu
  SMI_DisplayPopupMenu(MenuNum, WindowID, X, Y)
EndMacro
;
;
; ********************************************************************************
;
;-                                   7- DEMO
;
; ********************************************************************************
;
CompilerIf #PB_Compiler_IsMainFile
  Define MainWindowNum = OpenWindow(#PB_Any, 100, 100, 500, 350, "SetMenuItemEx library demo", #PB_Window_SystemMenu | #PB_Window_ScreenCentered)
  If MainWindowNum
    ApplyDarkModeToWindow(MainWindowNum)
    ;
    Define ArialFont    = LoadFont(#PB_Any, "Arial", 9)
    Define CourierFont  = LoadFont(#PB_Any, "Courier New", 9)
    Define SegoeFont    = LoadFont(#PB_Any, "Segoe UI", 9)
    Define ComicFont    = LoadFont(#PB_Any, "Comic Sans MS", 9)
    Define ImpactFont   = LoadFont(#PB_Any, "Impact", 10)
    Define JokermanFont = LoadFont(#PB_Any, "Jokerman", 9)
    Define CurlzMTFont  = LoadFont(#PB_Any, "Curlz MT", 9)
    Define GabriolaFont = LoadFont(#PB_Any, "Gabriola", 12)

    Define VioletColor        = $FF00FF
    Define DarkBlueColor      = $B00000
    Define DarkRedColor       = $0000D0
    Define DarkTurquoise      = $807000
    Define PaleOrangeColor    = $60A0FF
    Define PaleRedColor       = $8080FF
    Define PaleVioletColor    = $F685FF
    Define PaleBlueColor      = $FF9090
    Define PaleTurquoiseColor = $FFFF80
    Define PaleGreenColor     = $80FF80
    Define PaleYellowColor    = $80FFFF
    ;
    Define img1 = GetImageFromShell32(13, 32, 32)
    Define img2 = GetImageFromShell32(43, 32, 32)
    Define img3 = GetImageFromShell32(22, 32, 32)
    Define img4 = GetImageFromShell32(259, 32, 32)
    Define img5 = GetImageFromShell32(134, 32, 32)
    ;
    ;
    Define MainMenu = CreateMenu(#PB_Any, WindowID(MainWindowNum))
    ;
    ; You can now get a handle for a title and attribute an image to it:
    Define Menu1 = MenuTitle("Menu1", ImageID(img1))
    ; Before, you should select the title item by its position (still functionnal):
    ;     SetMenuTitleColor(MainMenu, 0, #PB_Gadget_FrontColor, #White)
    ;     SetMenuTitleColor(MainMenu, 0, #PB_Gadget_BackColor, 0)
    ; Now, you can use its handle (more simple):
    SetMenuTitleColor(MainMenu, Menu1, #PB_Gadget_FrontColor, #White)
    SetMenuTitleColor(MainMenu, Menu1, #PB_Gadget_BackColor, DarkTurquoise)
    ;
    ; Now, you can allocate a dynamic ID to an item with #PB_Any:
    Define Menu1_Command1 = MenuItem(#PB_Any, "Menu1_Command1", ImageID(img2))
    SetMenuItemColor(MainMenu, Menu1_Command1, #PB_Gadget_BackColor, DarkTurquoise)
    SetMenuItemColor(MainMenu, Menu1_Command1, #PB_Gadget_FrontColor, #White)
    ; Now, you can get a handle for a MenuBar
    Define Menu1_MenuBar = MenuBar()
    ; So, you can set colors to a MenuBar!
    SetMenuItemcolor(MainMenu, Menu1_MenuBar, #PB_Gadget_BackColor, DarkTurquoise)
    SetMenuItemcolor(MainMenu, Menu1_MenuBar, #PB_Gadget_FrontColor, #Red)
    ;   
    Define Menu1_Submenu = OpenSubMenu("Non-ownerdrawn Submenu", ImageID(img3))
    SetMenuItemColor(MainMenu, Menu1_Submenu, #PB_Gadget_BackColor, DarkTurquoise)
    SetMenuItemColor(MainMenu, Menu1_Submenu, #PB_Gadget_FrontColor, #White)
      Define Menu1_Submenu_Command1 = MenuItem(#PB_Any, "Menu1_Submenu_Command1")
      Define Menu1_SubSubmenu = OpenSubMenu("Ownerdrawn Sub-Submenu", ImageID(img3))
        Define Menu1_SubSubmenu_Command1 = MenuItem(#PB_Any, "Menu1_SubSubmenu_Command1")
        Define Menu1_SubSubmenu_Command2 = MenuItem(#PB_Any, "Menu1_SubSubmenu_Command2")
        ;
        ; You can omit specifying the menu number when using the SetMenuItem...
        ; and GetMenuItem.... commands. The item ID is sufficient to identify it:        
        SetMenuItemColor(#PB_Default, Menu1_SubSubmenu_Command2, #PB_Gadget_FrontColor, VioletColor)
      CloseSubMenu()
      Define Menu1_Submenu_Command2 = MenuItem(#PB_Any, "Menu1_Submenu_Command2")
    CloseSubMenu()
    ;
    Define Menu1_Command2 = MenuItem(#PB_Any, "Menu1_Command2")
    SetMenuItemImage(MainMenu, Menu1_Command2, ImageID(img5))
    SetMenuItemColor(MainMenu, Menu1_Command2, #PB_Gadget_BackColor, DarkTurquoise)
    SetMenuItemColor(MainMenu, Menu1_Command2, #PB_Gadget_FrontColor, #White)
    ;
    Define Colors = MenuTitle("Color")
    SetMenuTitleImage(MainMenu, Colors, ImageID(img4))
    SetMenuTitleColor(MainMenu, Colors, #PB_Gadget_FrontColor, DarkRedColor)
    SetMenuTitleColor(MainMenu, Colors, #PB_Gadget_BackColor, PaleYellowColor)
    Define Orange = MenuItem(#PB_Any, "Orange")
    SetMenuItemColor(MainMenu, Orange, #PB_Gadget_BackColor, PaleOrangeColor)
    Define Red = MenuItem(#PB_Any, "Red")
    SetMenuItemColor(MainMenu, Red, #PB_Gadget_BackColor, PaleRedColor)
    Define Violet = MenuItem(#PB_Any, "Violet")
    SetMenuItemColor(MainMenu, Violet, #PB_Gadget_BackColor, PaleVioletColor)
    Define Blue = MenuItem(#PB_Any, "Blue")
    SetMenuItemColor(MainMenu, Blue, #PB_Gadget_BackColor, PaleBlueColor)
    Define Turquoise = MenuItem(#PB_Any, "Turquoise")
    SetMenuItemColor(MainMenu, Turquoise, #PB_Gadget_BackColor, PaleTurquoiseColor)
    Define Green = MenuItem(#PB_Any, "Green")
    SetMenuItemColor(MainMenu, Green, #PB_Gadget_BackColor, PaleGreenColor)
    Define Yellow = MenuItem(#PB_Any, "Yellow")
    SetMenuItemColor(MainMenu, Yellow, #PB_Gadget_BackColor, PaleYellowColor)
    CheckMenuItem(MainMenu, Yellow, #True)
    Define OldCheckedColor = Yellow
    ;
    Define Fonts = MenuTitle("Choose your font")
    SetMenuTitleFont(MainMenu, Fonts, ImpactFont)
    Define Arial = MenuItem(#PB_Any, "Arial")
    SetMenuItemFont(MainMenu, Arial, ArialFont)
    Define Courier = MenuItem(#PB_Any, "Courier New")
    SetMenuItemFont(MainMenu, Courier, CourierFont)
    Define Comic = MenuItem(#PB_Any, "Comic Sans MS")
    SetMenuItemFont(MainMenu, Comic, ComicFont)
    Define Jokerman = MenuItem(#PB_Any, "Jokerman")
    SetMenuItemFont(MainMenu, Jokerman, JokermanFont)
    Define CurlzMT = MenuItem(#PB_Any, "Curlz MT")
    SetMenuItemFont(MainMenu, CurlzMT, CurlzMTFont)
    Define Gabriola = MenuItem(#PB_Any, "Gabriola")
    SetMenuItemFont(MainMenu, Gabriola, GabriolaFont)
    ;
    ; You can check or uncheck an item:
    MenuTitle("Check/Unchek my item")
    Define CheckMe = MenuItem(#PB_Any, "Click on me")
    CheckMenuItem(MainMenu, CheckMe)
    SetMenuItemColor(MainMenu, CheckMe, #PB_Gadget_FrontColor, 0)
    SetMenuItemImage(MainMenu, CheckMe, ImageID(img5))
    MenuItem(#PB_Any, "Alignement witness")
    ;
    ; Styled PopupMenu
    ;
    ; You can do exactly the same for PopupMenu items:
    Define StyledPopUpMenu = CreatePopupMenu(#PB_Any)
    MenuItem(#PB_Any, "PopUpMenu", ImageID(img5))
    Define PopupSubMenu = OpenSubMenu("PopupSubMenu", ImageID(img1))
    SetMenuItemColor(StyledPopUpMenu, PopupSubMenu, #PB_Gadget_BackColor, PaleOrangeColor)
      Define PopupSubMenu_Command1 = MenuItem(#PB_Any, "PopupSubMenu_Command1")
      Define PopupSubMenu_Command2 = MenuItem(#PB_Any, "PopupSubMenu_Command2")
      SetMenuItemColor(StyledPopUpMenu, PopupSubMenu_Command2, #PB_Gadget_FrontColor, PaleOrangeColor)
      SetMenuItemColor(StyledPopUpMenu, PopupSubMenu_Command2, #PB_Gadget_BackColor, DarkRedColor)
      SetMenuItemImage(StyledPopUpMenu, PopupSubMenu_Command2, ImageID(img5))
    CloseSubMenu()
    ;
    Define PUCourier = MenuItem(#PB_Any, "PopUpMenu Courier New")
    SetMenuItemFont(StyledPopUpMenu, PUCourier, CourierFont)
    Define PUImpact = MenuItem(#PB_Any, "PopUpMenu Impact")
    SetMenuItemFont(StyledPopUpMenu, PUImpact, ImpactFont)
    ;
    ;* ----------------------------- Buttons -----------------------------
    ;
    Define Msg$ = "By changing the value of SMI_ShowSelectionMethod, you change the manner"
    Msg$ + " of showing to the user that the mouse is over an item of the menu. Usualy, a "
    Msg$ + "simple background shading is enough to do that, but it can be inefficient in a "
    Msg$ + "dark mode environment or when each item has its own particular color (or both). "
    Msg$ + "So, you can choose four different methods (see the SMI_ShowSelectionMethod enumeration)."
    Msg$ + " Try each of them to decide. Note that your can eventually combine different methods:"
    ;
    TextGadget(#PB_Any, 10, 25, WindowWidth(MainWindowNum) - 20, 100, Msg$, #PB_Text_Center)
    Define VPos = 120
    Define HPos = 40
    Define BulletButton = CheckBoxGadget(#PB_Any, HPos, VPos, 60, 25, "Bullet")
    HPos + 85
    Define BordersButton = CheckBoxGadget(#PB_Any, HPos, VPos, 60, 25, "Framing")
    HPos + 100
    Define SystemColorButton = CheckBoxGadget(#PB_Any, HPos, VPos, 90, 25, "SystemColor")
    HPos + 120
    SetGadgetState(SystemColorButton, #True)
    SMI_ShowSelectionMethod = #SMI_SSW_SystemColor
    Define ShadeButton = CheckBoxGadget(#PB_Any, HPos, VPos, 120, 25, "Shade background")
    ;
    Msg$ = "This library requires no special skills to use. No initialization is needed. "
    Msg$ + "You don't need to install a callback procedure (this will be done automatically "
    Msg$ + "if necessary)." + #CR$
    Msg$ + "The only thing you need to do is include the library file in your program's code"
    Msg$ + " by using 'XIncludeFile (path)/SetMenuItemEx.pbi'." + #CR$ +"From then on, you can call"
    Msg$ + " functions like 'SetMenuItemColor()', 'SetMenuItemFont()', etc."
    ;
    TextGadget(#PB_Any, 10, 180, WindowWidth(MainWindowNum) - 20, 100, Msg$, #PB_Text_Center)
    ;
    VPos = WindowHeight(MainWindowNum) - 55
    Define PopUpMenuButton = ButtonGadget(#PB_Any, 10, VPos, 115, 25, "Open Styled PopUp")
    Define DisableButton = CheckBoxGadget(#PB_Any, 140, VPos, 165, 25, "Disable Menu1_Command2")
    Define HideMenu = CheckBoxGadget(#PB_Any, 320, VPos, 135, 25, "Hide menu")
    ;
    ; Main loop
    Repeat
      Select WaitWindowEvent()
        Case #PB_Event_Gadget
          Select EventGadget()
            Case PopUpMenuButton
              DisplayPopupMenu(StyledPopUpMenu, WindowID(MainWindowNum))
            Case DisableButton
              DisableMenuItem(MainMenu, Menu1_Command2, GetGadgetState(DisableButton))
            Case HideMenu
              HideMenu(MainMenu, GetGadgetState(HideMenu))
            Case BulletButton
              If GetGadgetState(BulletButton)
                SMI_ShowSelectionMethod | #SMI_SSW_Bullet
              Else
                SMI_ShowSelectionMethod & ~(#SMI_SSW_Bullet)
              EndIf
              ; CAUTION : If you modifiy SMI_ShowSelectionMethod
              ; on the fly, you need to resize menus:
              SMI_ResizeMenu(MainMenu, #True)
              SMI_ResizeMenu(StyledPopUpMenu, #True)
            Case BordersButton
              If GetGadgetState(BordersButton)
                SMI_ShowSelectionMethod | #SMI_SSW_Borders
              Else
                SMI_ShowSelectionMethod & ~(#SMI_SSW_Borders)
              EndIf
            Case SystemColorButton
              If GetGadgetState(SystemColorButton)
                SMI_ShowSelectionMethod | #SMI_SSW_SystemColor
              Else
                SMI_ShowSelectionMethod & ~(#SMI_SSW_SystemColor)
              EndIf
            Case ShadeButton
              If GetGadgetState(ShadeButton)
                SMI_ShowSelectionMethod | #SMI_SSW_ShadeBackground
              Else
                SMI_ShowSelectionMethod & ~(#SMI_SSW_ShadeBackground)
              EndIf
          EndSelect
          ;
        Case #PB_Event_Menu
          ; Now, you can find an item from its ID, without knowing the menu
          ; to which it belongs:
          Define ItemText$ = GetMenuItemText(#PB_Default, EventMenu())
          Msg$ = "The user chose: " + ItemText$
          Select EventMenu()
            Case CheckMe
              If IsMenuItemChecked(#PB_Default, CheckMe)
                CheckMenuItem(#PB_Default, CheckMe, #False)
                Msg$ + #CR$ + "The item is now unchecked"
              Else
                CheckMenuItem(#PB_Default, CheckMe, #True)
                Msg$ + #CR$ + "The item is now checked"
              EndIf
            Case Orange, Red, Violet, Blue, Turquoise, Green, Yellow
              CheckMenuItem(#PB_Default, OldCheckedColor, #False)
              CheckMenuItem(#PB_Default, EventMenu(), #True)
              OldCheckedColor = EventMenu()
              Define TitleColor = GetMenuItemColor(#PB_Default, EventMenu(), #PB_Gadget_BackColor)
              SetMenuTitleColor(#PB_Default, Colors, #PB_Gadget_BackColor, TitleColor)
          EndSelect
          MessageRequester("Info", Msg$, 0)
          ;
        Case #PB_Event_CloseWindow
          Break
      EndSelect
    ForEver
    CloseWindow(MainWindowNum)
  EndIf
  ;
CompilerEndIf
; IDE Options = PureBasic 6.20 (Windows - x64)
; CursorPosition = 5
; Folding = qqKvz1PAAAAA+--
; EnableXP
; DPIAware