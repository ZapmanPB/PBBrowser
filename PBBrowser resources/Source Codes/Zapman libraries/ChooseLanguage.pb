Procedure.s ChooseLanguage(folder$, LDefault$ = "")
  ;
  Protected selectedFolder$
  Protected numDirs, radioY, selectedGadget
  ;
  If ExamineDirectory(0, folder$, "*.*")
    ; Créer une liste pour stocker les noms de dossiers
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
    
    ; Calcul de la taille de la fenêtre en fonction du nombre de dossiers
    numDirs = ListSize(Dirs$())
    WindowHeight = (numDirs * 22) + 50
    WindowWidth = 190
    
    If OpenWindow(0, 200, 200, WindowWidth, WindowHeight, "Choose your language", #PB_Window_SystemMenu | #PB_Window_ScreenCentered)
      StickyWindow(0, 1)
      ; Créer les boutons radio pour chaque dossier et sélectionner le premier par défaut
      radioY = 5
      selectedGadget = -1
      ForEach Dirs$()
        GadgetID = OptionGadget(#PB_Any, 60, radioY, WindowWidth - 20, 20, Dirs$())
        If selectedGadget = -1 Or FindString(LDefault$, Dirs$())
          SetGadgetState(gadgetID, 1) ; Sélectionner le premier bouton
          selectedGadget = gadgetID
        EndIf
        AddElement(finalList())
        finalList()\Dir$ = dirs$()
        finalList()\GadgetID = GadgetID
        radioY + 22
      Next dirs$()
      
      ; Créer le bouton OK en bas à droite
      ButtonGadget(1, WindowWidth - 80, WindowHeight - 32, 70, 22, "OK")
      
      ; Boucle principale de l'application
      Repeat
        Event = WaitWindowEvent()
        If Event = #PB_Event_Gadget
          If EventGadget() = 1 ; Si on appuie sur OK
            Break
          EndIf
        EndIf
      Until Event = #PB_Event_CloseWindow ; Fermer la fenêtre via l'icône de fermeture
    EndIf
  Else
    MessageRequester("Error", "Unable to open the specified folder.")
  EndIf
  ;
  ; Récupérer le nom du dossier sélectionné
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

; Appel de la procédure en passant l'adresse du dossier à explorer
;Debug ChooseLanguage("D:\LUC\Dropbox\RankSpiritDev\CreateMyEBook\PBBrowser resources\Catalogs")

; IDE Options = PureBasic 6.11 LTS (Windows - x64)
; CursorPosition = 14
; Folding = -
; EnableXP
; DPIAware