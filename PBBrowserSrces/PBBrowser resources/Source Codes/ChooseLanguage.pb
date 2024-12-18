;***********************************************************************
;
;                          ChooseLanguage.pb
;                      Part of PBBrowser project
;                          Zapman - Dec 2024
;
;  Open a window allowing the user to choose the application language.
;
;***********************************************************************
;
Procedure.s ChooseLanguage(folder$, LDefault$ = "")
  ;
  Protected selectedFolder$
  Protected numDirs, radioY, selectedGadget
  ;
  If ExamineDirectory(0, folder$, "*.*")
    ; Create a list to store the folder names
    NewList Dirs$()
    While NextDirectoryEntry(0)
      If DirectoryEntryType(0) = #PB_DirectoryEntry_Directory And DirectoryEntryName(0) <> "." And DirectoryEntryName(0) <> ".."
        AddElement(Dirs$())
        Dirs$() = DirectoryEntryName(0)
      EndIf
    Wend
    FinishDirectory(0)
    
    SortList(Dirs$(), #PB_Sort_Ascending)
    
    Structure ChoicesList
      Dir$
      GadgetID.i
    EndStructure
    NewList finalList.ChoicesList()
    
    ; Calculate the window size based on the number of folders
    numDirs = ListSize(Dirs$())
    WindowHeight = (numDirs * 22) + 50
    WindowWidth = 190
    
    If OpenWindow(0, 250, 200, WindowWidth, WindowHeight, "Choose your language", #PB_Window_SystemMenu | #PB_Window_ScreenCentered | #PB_Window_Invisible)
      ApplyDarkModeToWindow(0)
      StickyWindow(0, 1)
      ; Create radio buttons for each folder and select the first one by default
      radioY = 5
      selectedGadget = -1
      ForEach Dirs$()
        GadgetID = OptionGadget(#PB_Any, 60, radioY, WindowWidth - 20, 20, Dirs$())
        GadgetList$ + Str(GadgetID) + ","
        If selectedGadget = -1 Or FindString(LDefault$, Dirs$())
          SetGadgetState(gadgetID, 1) ; Select the first button
          selectedGadget = gadgetID
        EndIf
        AddElement(finalList())
        finalList()\Dir$ = dirs$()
        finalList()\GadgetID = GadgetID
        radioY + 22
      Next dirs$()
      
      ; Create the OK button at the bottom right
      ButtonGadget(1, WindowWidth - 80, WindowHeight - 32, 70, 22, "OK")
      GadgetList$ + "1,"
      ;
      SetFontAndGadgetsColors(0, InterfaceColorPresets(), GadgetList$)
      ;
      ; The window was invisible until now, because we created it with #PB_Window_Invisible.
      ; We make it visible now.
      HideWindow(0, #False)
      ;
      ; Main application loop
      Repeat
        Event = WaitWindowEvent()
        If Event = #PB_Event_Gadget
          If EventGadget() = 1 ; If OK is pressed
            Break
          EndIf
        EndIf
      Until Event = #PB_Event_CloseWindow ; Close the window via the close icon
    EndIf
  Else
    MessageRequester("Error", "Unable to open the specified folder.")
  EndIf
  ;
  ; Retrieve the name of the selected folder
  ForEach finalList()
    If GetGadgetState(finalList()\GadgetID)
      selectedFolder$ = finalList()\Dir$
      Break
    EndIf
    selectedGadget + 1
  Next finalList()
  CloseWindow(0)
  ProcedureReturn GetPathPart(folder$) + selectedFolder$ + "\"
EndProcedure

; Calling the procedure with the folder address to explore
;Debug ChooseLanguage("C:\MyFolder")

; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; Folding = -
; EnableXP
; DPIAware