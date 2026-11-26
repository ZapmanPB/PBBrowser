;***********************************************************************
;
;                          ChooseLanguage.pb
;                      Part of PBBrowser project
;                          Zapman - Nov 2026
;
;  Open a window allowing the user to choose the application language.
;
;***********************************************************************
;
Enumeration CL_Shortcuts
  #CL_Shortcut_Confirm = 45010
  #CL_Shortcut_Cancel
EndEnumeration
;
Structure CL_ChoiceEntry
  Dir$
  GadgetID.i
EndStructure
;
Procedure.s ChooseLanguage(folder$, LDefault$ = "")
  ;
  ; Role: Scans 'folder$' for language subdirectories, opens a modal selection dialog,
  ;       and returns the chosen language directory path (with a guaranteed trailing backslash).
  ;       - Confirming via OK or Enter returns the newly selected directory.
  ;       - Closing via [X] or Escape preserves 'LDefault$' when valid, or falls back to the
  ;         default selection on first launch so callers always receive a usable catalog path.
  ;       - Returns "" only if 'folder$' cannot be opened or contains no subdirectories.
  ;
  Protected NewList Choices.CL_ChoiceEntry()
  Protected BaseFolder$, DefaultDir$, EntryName$, SelectedFolder$, GadgetList$
  Protected DirID, WLang, BOK, ParentWindowID, Event
  Protected NumDirs, MaxChars, WWidth, WHeight, OX, OY, RadioY
  Protected SelectedGadget, DefaultMatched, Canceled
  ;
  BaseFolder$ = Trim(folder$)
  If BaseFolder$ = ""
    ProcedureReturn ""
  EndIf
  If Right(BaseFolder$, 1) <> "\" And Right(BaseFolder$, 1) <> "/"
    BaseFolder$ + "\"
  EndIf
  ;
  ; Extract the exact trailing folder name from LDefault$ (supports both full path and bare folder name)
  DefaultDir$ = RTrim(RTrim(Trim(LDefault$), "\"), "/")
  If DefaultDir$ <> ""
    DefaultDir$ = GetFilePart(DefaultDir$)
  EndIf
  ;
  DirID = ExamineDirectory(#PB_Any, BaseFolder$, "*.*")
  If DirID = 0
    MessageRequester("Error", "Unable to open the specified folder:" + #CRLF$ + BaseFolder$)
    ProcedureReturn ""
  EndIf
  ;
  While NextDirectoryEntry(DirID)
    If DirectoryEntryType(DirID) = #PB_DirectoryEntry_Directory
      EntryName$ = DirectoryEntryName(DirID)
      If EntryName$ <> "." And EntryName$ <> ".."
        AddElement(Choices())
        Choices()\Dir$ = EntryName$
        If Len(EntryName$) > MaxChars
          MaxChars = Len(EntryName$)
        EndIf
      EndIf
    EndIf
  Wend
  FinishDirectory(DirID)
  ;
  NumDirs = ListSize(Choices())
  If NumDirs = 0
    MessageRequester("Error", "No language folder found in:" + #CRLF$ + BaseFolder$)
    ProcedureReturn ""
  EndIf
  ;
  SortStructuredList(Choices(), #PB_Sort_Ascending | #PB_Sort_NoCase, OffsetOf(CL_ChoiceEntry\Dir$), TypeOf(CL_ChoiceEntry\Dir$))
  ;
  ; Adapt window width to long folder names while keeping all OptionGadgets strictly inside client bounds
  WWidth = 80 + (MaxChars * 8)
  If WWidth < 190
    WWidth = 190
  EndIf
  WHeight = (NumDirs * 22) + 50
  ;
  ParentWindowID = ComputeWinOrigins(@OX, @OY, WWidth, WHeight, GPBBGadgets\PBBWindow)
  WLang = OpenWindow(#PB_Any, OX, OY, WWidth, WHeight, "Choose your language", #PB_Window_SystemMenu | #PB_Window_Invisible, ParentWindowID)
  If WLang = 0
    ProcedureReturn ""
  EndIf
  ;
  ApplyDarkModeToWindow(WLang)
  StickyWindow(WLang, 1)
  ;
  RadioY = 8
  ForEach Choices()
    Choices()\GadgetID = OptionGadget(#PB_Any, 45, RadioY, WWidth - 55, 20, Choices()\Dir$)
    GadgetList$ + Str(Choices()\GadgetID) + ","
    ; Priority: 1) Exact match on LDefault$, 2) "English" when LDefault$ is empty, 3) First sorted folder
    If DefaultDir$ <> "" And LCase(Choices()\Dir$) = LCase(DefaultDir$)
      SelectedGadget = Choices()\GadgetID
      DefaultMatched = #True
    ElseIf DefaultMatched = #False
      If DefaultDir$ = "" And LCase(Choices()\Dir$) = "english"
        SelectedGadget = Choices()\GadgetID
      ElseIf SelectedGadget = 0
        SelectedGadget = Choices()\GadgetID
      EndIf
    EndIf
    RadioY + 22
  Next
  ;
  If SelectedGadget
    SetGadgetState(SelectedGadget, 1)
  EndIf
  ;
  BOK = ButtonGadget(#PB_Any, WWidth - 80, WHeight - 32, 70, 22, "OK", #PB_Button_Default)
  GadgetList$ + Str(BOK) + ","
  ;
  ; Keep shortcut IDs strictly below $AFFF (45055) for compatibility with SetMenuItemEx
  AddKeyboardShortcut(WLang, #PB_Shortcut_Return, #CL_Shortcut_Confirm)
  AddKeyboardShortcut(WLang, #PB_Shortcut_Escape, #CL_Shortcut_Cancel)
  ;
  SetFontAndGadgetsColors(WLang, InterfaceColorPresets(), GadgetList$)
  HideWindow(WLang, #False)
  SetActiveWindow(WLang)
  ;
  Repeat
    Event = WaitWindowEvent()
    If EventWindow() <> WLang
      SetActiveWindow(WLang)
      Continue
    EndIf
    ;
    Select Event
      Case #PB_Event_Gadget
        If EventGadget() = BOK
          Break
        EndIf
      Case #PB_Event_Menu
        If EventMenu() = #CL_Shortcut_Confirm
          Break
        ElseIf EventMenu() = #CL_Shortcut_Cancel
          Canceled = #True
          Break
        EndIf
      Case #PB_Event_CloseWindow
        Canceled = #True
        Break
    EndSelect
  ForEver
  ;
  If Canceled And DefaultMatched
    ; User aborted via [X] or Escape while a valid active language was already configured: keep it unchanged
    SelectedFolder$ = DefaultDir$
  Else
    ForEach Choices()
      If GetGadgetState(Choices()\GadgetID)
        SelectedFolder$ = Choices()\Dir$
        Break
      EndIf
    Next
  EndIf
  ;
  CloseWindow(WLang)
  ;
  If SelectedFolder$ = ""
    ProcedureReturn ""
  EndIf
  ProcedureReturn BaseFolder$ + SelectedFolder$ + "\"
EndProcedure
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 169
; Folding = -
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb