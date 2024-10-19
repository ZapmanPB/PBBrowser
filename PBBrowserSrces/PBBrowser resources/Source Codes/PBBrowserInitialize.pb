; *******************************************************************************
;
;                         Initialisation de PBBrowser
;
; *******************************************************************************
;
;
; *******************************************************************************
;
;        Stockage et récupération des images pour les images-boutons
;           et du catalog contenant les messages de l'application.
;
Procedure ExtractData(dataStart, dataEnd, outputPath$)
  ;
  outputPath$ = ReplaceString(outputPath$, "/", "\")
  CreatePathFolders(outputPath$)
  If CreateFile(0, outputPath$)
    WriteData(0, dataStart, dataEnd - dataStart)
    CloseFile(0)
  Else
    MessageRequester("Ooops!", "Unable to create AppData file!")
    End
  EndIf
EndProcedure
;
; Le contenu du dossier des données (AppData) de PBBrowser est mis-à-jour au premier
; lancement de l'application 'PBBrowser.exe' avec les données contenues
; dans l'application elle-même (voir la DataSection de PBBrowserDeclaration.pb).
; Si l'application est lancée depuis l'IDE de PureBasic ('Compiler/Exécuter')
; les données sont systématiquement recopiées depuis le dossier 'PBBrowser resources'
; qui accompagne le fichier-source de l'application. Ainsi, en cas de mise-à-jour des
; fichiers de 'PBBrowser resources', l'application dispose toujours des dernières
; versions de ces fichiers.
;
If FileSize(MyAppDataFolder$ + "Next.jpg") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BNext, ?BNextEnd, MyAppDataFolder$ + "Next.jpg")
EndIf
If FileSize(MyAppDataFolder$ + "Previous.jpg") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BPrevious, ?BPreviousEnd, MyAppDataFolder$ + "Previous.jpg")
EndIf
If FileSize(MyAppDataFolder$ + "NoNext.jpg") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BNoNext, ?BNoNextEnd, MyAppDataFolder$ + "NoNext.jpg")
EndIf
If FileSize(MyAppDataFolder$ + "NoPrevious.jpg") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BNoPrevious, ?BNoPreviousEnd, MyAppDataFolder$ + "NoPrevious.jpg")
EndIf
If FileSize(MyAppDataFolder$ + "Refresh.jpg") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BRefresh, ?BRefreshEnd, MyAppDataFolder$ + "Refresh.jpg")
EndIf
If FileSize(MyAppDataFolder$ + "Catalogs\Français\PBBrowser.catalog") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BCatalogFR, ?BCatalogFREnd, MyAppDataFolder$ + "Catalogs\Français\PBBrowser.catalog")
EndIf
If FileSize(MyAppDataFolder$ + "Catalogs\Français\IntroPBBrowser.rtf") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BIntroFR, ?BIntroFREnd, MyAppDataFolder$ + "Catalogs\Français\IntroPBBrowser.rtf")
EndIf
If FileSize(MyAppDataFolder$ + "Catalogs\English\PBBrowser.catalog") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BCatalogEN, ?BCatalogENEnd, MyAppDataFolder$ + "Catalogs\English\PBBrowser.catalog")
EndIf
If FileSize(MyAppDataFolder$ + "Catalogs\English\IntroPBBrowser.rtf") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BIntroEN, ?BIntroENEnd, MyAppDataFolder$ + "Catalogs\English\IntroPBBrowser.rtf")
EndIf
If FileSize(MyAppDataFolder$ + "Catalogs\Italiano\PBBrowser.catalog") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BCatalogIT, ?BCatalogITEnd, MyAppDataFolder$ + "Catalogs\Italiano\PBBrowser.catalog")
EndIf
If FileSize(MyAppDataFolder$ + "Catalogs\Italiano\IntroPBBrowser.rtf") < 2 Or #PB_Compiler_Debugger
  ExtractData(?BIntroIT, ?BIntroITEnd, MyAppDataFolder$ + "Catalogs\Italiano\IntroPBBrowser.rtf")
EndIf

If FileSize(MyAppDataFolder$ + "PBFunctionList.Data") < 2 Or #PB_Compiler_Debugger
  ExtractData(?PBFunction, ?PBFunctionEnd, MyAppDataFolder$ + "PBFunctionList.Data")
EndIf

If FileSize(MyAppDataFolder$ + "APIFunctionListing.txt") < 2 Or #PB_Compiler_Debugger
  ExtractData(?APIFunction, ?APIFunctionEnd, MyAppDataFolder$ + "APIFunctionListing.txt")
EndIf
;
;
; *******************************************************************************
;
;           Lecture du fichier des préférences et choix de la langue

If OpenPreferencesWithPatience(PBBrowserPrefile$)
  PBBLanguageFolder$ = ReadPreferenceString("PBBLanguageFolder", "")
  ClosePreferences()
  PBBFirstLaunch = 0
Else ; Il n'y a pas de préférences. L'application est lancée pour la première fois
  CreatePathFolders(PBBrowserPrefile$)
  If CreatePreferences(PBBrowserPrefile$)
    ClosePreferences()
  EndIf
  PBBFirstLaunch = 1
EndIf
;
If PBBLanguageFolder$ = ""
  PBBLanguageFolder$ = ChooseLanguage(MyAppDataFolder$ + "Catalogs\")
  If OpenPreferencesWithPatience(PBBrowserPrefile$)
    WritePreferenceString("PBBLanguageFolder", PBBLanguageFolder$)
    ClosePreferences()
  Else
    MessageRequester("Ooops!", "Unable to create 'preference' file!")
    End
  EndIf
EndIf
;
; *******************************************************************************
;
;                            Initialisation du 'Catalog'
;
; Tous les messages sont stockés dans le fichier 'PBBrowser.catalog' afin de pouvoir
; être traduits dans différentes langues.
;
; La ligne ci-dessous initialise l'adresse du fichier de référence à utiliser par
; GetTextFromCatalog().
; Comme aucun texte n'est passé dans le premier argument, cet appel n'aura pas
; d'autre effet :
GetTextFromCatalog("", PBBLanguageFolder$ + "PBBrowser.catalog")
;
; *******************************************************************************
;
;      Récupération des arguments passés par PureBasic au moment de l'appel
;                 de l'application à l'aide du menu "Outils"
;
nl = 0
Repeat
  SAppArgument$ = ProgramParameter()
  nl + 1
  If SAppArgument$
    If nl = 1     ; En principe, le premier argument passé par l'appel d'outil est "%HOME"
                  ; qui contient le chemin complet vers l'application PureBasic.
      PureBasicProgAdr$ = SAppArgument$ + "PureBasic.exe"
    ElseIf nl = 2 ; En principe, le deuxième argument passé par l'appel d'outil est "%WORD"
                  ; qui contient le mot situé sous la souris dans l'éditeur de PureBasic.
      PBUnderCursor$ = SAppArgument$
    ElseIf nl = 3 ; En principe, le troisième argument passé par l'appel d'outil est "%SELECTION"
                  ; qui contient la position de la sélection dans le fichier.
      PBSelection$ = SAppArgument$
    ElseIf nl = 4 ; En principe, le cinquième argument passé par l'appel d'outil est "%CURSOR"
                  ; qui contient la position du curseur dans le fichier actuellement ouvert
                  ; dans l'éditeur de PureBasic.
      PBCursor$ = SAppArgument$
    ElseIf nl = 5 ; En principe, le cinquième argument passé par l'appel d'outil est "%FILE"
                  ; qui contient le fichier actuellement ouvert dans l'éditeur de PureBasic.
      FicPrincipalPB$ = SAppArgument$
      FicActualPB$ = SAppArgument$
    Else          ; En principe, le sixième argument passé par l'appel d'outil est "%TEMPFILE"
                  ; qui contient le fichier actuellement ouvert dans l'éditeur de PureBasic.
      TempFile$ = SAppArgument$
      If FicActualPB$ = ""
        FicActualPB$ = TempFile$
        FicPrincipalPB$ = TempFile$
      EndIf
    EndIf
  EndIf
Until nl > 5
;
;
; *******************************************************************************
;
;                             Message de présentation
;
; *******************************************************************************
;
If PBBFirstLaunch = 1
  AlertWithTitle(GetTextFromCatalog("file:IntroPBBrowser.rtf"), GetTextFromCatalog("Introduction"))
EndIf

;
; *******************************************************************************
;
;         Procédures diverses, utilisées dans l'ensemble du programme
;
; *******************************************************************************
;
Procedure.s GetPureBasicPrefAdresse()
  ; Détermine l'emplacement du fichier des préférences de PureBasic.
  ;
  ; PB Browser va puiser des données dans ce fichier, telles que les
  ; couleurs à attribuer aux mots-clés. C'est également dans ce dossier
  ; des préférences qu'on installe PB Browser en tant qu'outil de
  ; l'IDE PureBasic.
  Protected  itemid, PBPrefAdr$
  If SHGetSpecialFolderLocation_(0, #CSIDL_APPDATA, @itemid) = #NOERROR
    PBPrefAdr$ = Space(#MAX_PATH)
    SHGetPathFromIDList_(itemid, @PBPrefAdr$)
  EndIf
  PBPrefAdr$ + "\PureBasic\PureBasic.prefs"
  If FileSize(PBPrefAdr$) < 2
    PBPrefAdr$ = GetHomeDirectory() + "AppData\\Roaming\\PureBasic\\PureBasic.prefs"
  EndIf
  If FileSize(PBPrefAdr$) < 2
    PBPrefAdr$ = GetHomeDirectory() + "AppData\\PureBasic\\PureBasic.prefs"
  EndIf
  If FileSize(PBPrefAdr$) < 2
    ProcedureReturn ""
  Else
    ProcedureReturn PBPrefAdr$
  EndIf
EndProcedure
;
Procedure PBB_SuspendRedraw(gadget.i, YesNo = #True)
  ;
  ; Suspend gadget content's redrawing for a while,
  ; or retablish content's redrawing.
  ; This is usefull to change gadget's content without
  ; multiple scrollings.
  Protected Redraw
  ;
  If IsGadget(gadget)
    If YesNo = 0 : Redraw = #True : Else : Redraw = #False : EndIf
    SendMessage_(GadgetID(gadget), #WM_SETREDRAW, Redraw, 0)
    ;
    If Redraw = #False
      ; Because we suppose that redrawing is suspended for content modifications
      ; we set the gadget to READONLY/#False (allows modifications)
      SendMessage_(GadgetID(gadget), #EM_SETREADONLY, #False, 0)
    Else
      RedrawWindow_(GadgetID(gadget), 0, 0, #RDW_ERASE | #RDW_INVALIDATE | #RDW_UPDATENOW)
      SendMessage_(GadgetID(gadget), #EM_SETREADONLY, #True, 0)
    EndIf
  EndIf
EndProcedure
;
Procedure DisablePBBWindow()
  ;
  ; Désactive la fenêtre principale de PBBrowser afin de donner
  ; la priorité à une fenêtre de message ou à une fenêtre de recherche.
  ;
  If IsWindow(GPBBGadgets\PBBWindow) 
    If GPBBGadgets\Disabled = 0
      If Not (IsGadget(GPBBGadgets\IWhiteOver))
        GPBBGadgets\IWhiteOver = WhiteBoxOverWindow(GPBBGadgets\PBBWindow)
      EndIf
      HideGadget(GPBBGadgets\IWhiteOver, #False)
      DisableWindow(GPBBGadgets\PBBWindow, #True)
    EndIf
    GPBBGadgets\Disabled + 1
  EndIf
  ProcedureReturn
EndProcedure
;
Procedure EnablePBBWindow()
  ;
  ; Réactive la fenêtre principale de PBBrowser après un appel
  ; à 'DisablePBBWindow()'
  ;
  If IsWindow(GPBBGadgets\PBBWindow)
    GPBBGadgets\Disabled - 1
    If GPBBGadgets\Disabled < 1
      GPBBGadgets\Disabled = 0
      DisableWindow(GPBBGadgets\PBBWindow, #False)
      If IsGadget(GPBBGadgets\IWhiteOver)
        HideGadget(GPBBGadgets\IWhiteOver, #True)
      EndIf
    EndIf
    SetActiveWindow(GPBBGadgets\PBBWindow)
  EndIf
  ProcedureReturn
EndProcedure
;
Procedure PBBInitAlertTexts()
  ;
  ;         Initialisation des textes de boutons et du titre de fenêtre
  ;       pour la procédure 'Alert()', utilisée tout au long du programme
  ;
  Protected  AlertText.AlertWindowTitles
  ;
  AlertText.AlertWindowTitles\WTitle$  = GetTextFromCatalog("Attention")
  AlertText\BOK$     = GetTextFromCatalog("OK")
  AlertText\BCancel$ = GetTextFromCatalog("Cancel")
  AlertText\BYes$    = GetTextFromCatalog("Yes")
  AlertText\BNo$     = GetTextFromCatalog("No")
  AlertText\BCopy$   = GetTextFromCatalog("CopyAll")
  AlertText\BSearch$ = GetTextFromCatalog("Search")
  AlertText\BSave$   = GetTextFromCatalog("Save")
  AlertText\SaveAs$  = GetTextFromCatalog("SaveAs")
  AlertText\TextFile$ = GetTextFromCatalog("TextFile")
  AlertText\SearchGadgets\WTitle$ =       GetTextFromCatalog("SWTitle")
  AlertText\SearchGadgets\Search$ =       GetTextFromCatalog("SWSearch")
  AlertText\SearchGadgets\ReplaceTitle$ = GetTextFromCatalog("SWReplaceTitle")
  AlertText\SearchGadgets\Replace$ =      GetTextFromCatalog("SWReplace")
  AlertText\SearchGadgets\ReplaceAll$ =   GetTextFromCatalog("SWReplaceAll")
  AlertText\SearchGadgets\Quit$ =         GetTextFromCatalog("Quit")
  AlertText\SearchGadgets\CaseSensitive$ = GetTextFromCatalog("SWCaseSensitive")
  AlertText\SearchGadgets\WholeWord$ =    GetTextFromCatalog("SWWholeWord")
  AlertText\SearchGadgets\InAllDocument$ = GetTextFromCatalog("SWInAllDocument")
  AlertText\SearchGadgets\UnableToFind$ = GetTextFromCatalog("SWUnableToFind")
  AlertText\SearchGadgets\ReplacementMade$ = GetTextFromCatalog("SWReplacementMade")
  AlertText\SearchGadgets\SearchFromStart$ = GetTextFromCatalog("SWSearchFromStart")
  AlertText\SearchGadgets\SearchFromEnd$ =   GetTextFromCatalog("SWSearchFromEnd")
  ;
  Alert("", AlertText) ; On initialise les noms à utiliser
  ; Comme aucun texte ne figurait dans le premier argument, cet appel
  ; n'aura pas d'autre effet.
EndProcedure
;
PBBInitAlertTexts()
;
Procedure AlertInPBBWindow(AlertText$, Title$ = "", WSticky = 1, WhiteBackGround = 0, Txtleft = 0, FixedWidth = 0, TabList$ = "")
  ;
  ; Désactive la fenêtre principale de PBBrowser et affiche une fenêtre
  ; contenant un message. Puis réactive la fenêtre principale lorsque
  ; l'utilisateur ferme la fenêtre de message.
  ;
  DisablePBBWindow()
  If Title$ = ""
    Alert(AlertText$)
  Else
    AlertWithTitle(AlertText$, Title$, WSticky, WhiteBackGround, Txtleft, FixedWidth, TabList$)
  EndIf
  EnablePBBWindow()
EndProcedure
;
; *******************************************************************************
;
;           Procédures dédiées à la lecture des fichiers de code
;
; *******************************************************************************
;
Structure FileListingStruct ; Structure de la liste PBB_FileListing
  FileNameInList$
  FileContentInList.String
  FileLCaseContentInList.String
  FileLastModInList.i
EndStructure
;
NewList PBB_FileListing.FileListingStruct()
;
Procedure IsFileInMemory(FileName$)
  ;
  ; Cette procédure regarde si FileName$ figure en mémoire-cache et
  ; retourne #True ou #False selon le cas.
  ;
  Shared PBB_FileListing()
  ;
  If ListSize(PBB_FileListing())
    ; On ne parcourt la liste que si elle comporte des éléments
    ; et si elle n'est pas déjà positionnée sur le bon élément.
    If ListIndex(PBB_FileListing()) = -1 Or PBB_FileListing()\FileNameInList$ <> FileName$
      ForEach PBB_FileListing()
        If PBB_FileListing()\FileNameInList$ = FileName$
          Break
        EndIf
      Next
    EndIf
    If PBB_FileListing()\FileNameInList$ = FileName$
      ProcedureReturn #True
    EndIf
  EndIf
  ProcedureReturn #False
EndProcedure
;
Procedure.s GetCodeFromFile(FileName$)
  ;
  ; Retourne le contenu du fichier 'FileName$' sous la forme
  ; d'un simple chaîne.
  ;
  If FileSize(FileName$) = -1
    AlertInPBBWindow(GetTextFromCatalog("ErrorWithFile") + " " + FileName$)
  Else
    ; On simplifie les recherches à venir en remplaçant les tabulations par un double espace
    ; et en supprimant les éventuels Chr(10) (LineFeed).
    Protected Content$ = ReplaceString(ReplaceString(FileToText(FileName$, #PB_UTF8, 1), Chr(9), "  "), Chr(10), "")
    If Len(Content$) = 0 And FileSize(FileName$) > 2
      AlertInPBBWindow(GetTextFromCatalog("ErrorWithFile") + " " + FileName$)
    EndIf
    ProcedureReturn Content$
  EndIf
EndProcedure
;
Procedure GetPointedCodeFromFile(FileName$, ReturnLCase = #False)
  ;
  ; Cette procédure va stocker chaque fichier lu en mémoire, afin d'éviter de relire
  ; un fichier déjà lu. Quand un contenu est demandé une deuxième fois, c'est
  ; le contenu en mémoire qui est délivré.
  ;
  ; Quand le paramètre 'ReturnLCase' vaut '#True', c'est une version en minuscules (LCase()) du fichier
  ; qui est retournée. Cela permettra de faire des recherches sans tenir compte de la casse
  ; sur le contenu du fichier.
  ;
  ; La valeur de retour est un pointeur vers une chaîne String et ce pointeur
  ; demeure valide hors de la procédure parce que le 'String' est conservée en 'Shared' dans la liste PBB_FileListing().
  ; Cette façon de procéder évite de dupliquer en mémoire la chaîne retournée, ce qui permet de gagner
  ; du temps et de la mémoire.
  ;
  ;
  Shared PBB_FileListing()
  Static ErrorString.String
  ErrorString\s = "Error"
  ;
  If IsFileInMemory(FileName$) = #False And FileSize(FileName$) > 1
    AddElement(PBB_FileListing())
    PBB_FileListing()\FileNameInList$ = FileName$
    PBB_FileListing()\FileLastModInList = GetFileDate(FileName$, #PB_Date_Modified)
    PBB_FileListing()\FileContentInList\s = GetCodeFromFile(FileName$)
    PBB_FileListing()\FileLCaseContentInList\s = ""
  EndIf
  ;
  If ListSize(PBB_FileListing()) And ListIndex(PBB_FileListing()) <> -1
    If ReturnLCase = #True
      If PBB_FileListing()\FileLCaseContentInList\s = ""
        PBB_FileListing()\FileLCaseContentInList\s = LCase(PBB_FileListing()\FileContentInList\s)
      EndIf
      ProcedureReturn @PBB_FileListing()\FileLCaseContentInList
    Else
      ProcedureReturn @PBB_FileListing()\FileContentInList
    EndIf
  Else
    ProcedureReturn @ErrorString
  EndIf
EndProcedure
;
Procedure.s GetStringCodeFromFile(FileName$, ReturnLCase = #False)
  ; Cette version de GetCodeFromFile() retourne une chaîne
  ; au lieu de retourner un pointeur vers une structure 'String'.
  Protected *SPointer.String = GetPointedCodeFromFile(FileName$, ReturnLCase)
  ProcedureReturn *SPointer\s
EndProcedure
;
;
; *******************************************************************************
;
;   Mise à jour des paramètres passés par l'éditeur de PureBasic, dans le cas
;                où PBBrowser est lancé depuis l'éditeur.
;
;
If PBSelection$ And PBUnderCursor$ = "" ; Curieusement, quand une portion de texte est sélectionnée dans l'éditeur
                                        ; de PureBasic (même un simple mot), l'éditeur ne renvoie pas son contenu
                                        ; dans l'argument %Word. On va donc l'extraire à partir des positions
                                        ; de la sélection founies par l'argument %SELECTION.
  c1 = Val(StringField(PBSelection$, 2, "x"))
  c2 = Val(StringField(PBSelection$, 4, "x"))
  l1 = Val(StringField(PBSelection$, 1, "x"))
  l2 = Val(StringField(PBSelection$, 3, "x"))
  If l1 = l2 And c1 <> c2 And FicActualPB$    ; Cela ne fonctionnera pas si la sélection englobe plusieurs lignes
    *FileContent.String = GetPointedCodeFromFile(FicActualPB$)  ; mais c'est un cas de figure qui ne nous intéresse pas.
    Line$ = StringField(*FileContent\s, l2, Chr(13))
    PBUnderCursor$ = Mid(Line$, c1, c2 - c1)
  EndIf
EndIf
If PBCursor$                            ; Quand le curseur se trouve au-dessus d'un nom de constante,
                                        ; l'éditeur de PureBasic renvoie le nom de constante sans le '#'
                                        ; qui précède. C'est une information pourtant utile, que nous
                                        ; allons récupérer.
  l1 = Val(StringField(PBSelection$, 1, "x"))
  c1 = Val(StringField(PBSelection$, 2, "x"))
  If l1 And FicActualPB$
    *FileContent.String = GetPointedCodeFromFile(FicActualPB$)
    Line$ = StringField(*FileContent\s, l1, Chr(13))
    p = FindString(Line$, PBUnderCursor$, c1 - Len(PBUnderCursor$))
    If p And Mid(Line$, p - 1, 1) = "#"
      PBUnderCursor$ = "#" + PBUnderCursor$
    EndIf
  EndIf
EndIf
;
; *******************************************************************************
;
;      Quand PBBrowser.exe est lancé depuis l'éditeur à de multiples reprises,
;      on ne veut pas que cela ouvre plusieurs instances de l'application.
;      On veut qu'une seule instance récupère les paramètres fournis par
;      par l'éditeur et réagisse par rapport à ces paramètres.
;      Pour obtenir ce résultat, quand une nouvelle instance est crée, elle
;      récupère les paramètres de lancement, les communique à l'instance
;      précédante, puis met fin à son fonctionnement.
;      Il faut donc être en mesure de détecter si une instance de l'application
;      est déjà en cours de fonctionnemment, puis être en mesure de communiquer
;      avec elle. Tout cela se fait à travers la gestion d'un 'Pipe', un procédé
;      faisant partie de l'API Windows et dont le but est de permettre un dialogue
;      entre plusieurs threads d'une application ou entre plusieurs applications
;      différentes.
;
;
If CheckForOtherInstance(PureBasicProgAdr$ + Chr(13) + PBUnderCursor$ + Chr(13) + FicActualPB$ + Chr(13) + TempFile$ + Chr(13))
  ; Si une autre instance de cette application est déjà ouverte,
  ; on lui transmet les arguments que l'on vient de recevoir
  ; et on met fin au programme.
  If #PB_Compiler_Debugger
    ; Si nous sommes en mode 'Compilé', on signale à l'utilisateur
    ; l'éventuelle existence d'une instance précédente en mode
    ; 'StandAlone' (.exe).
    PlaySound_("SystemExclamation", 0, #SND_ALIAS | #SND_ASYNC)
    Debug "PBBrowser is allready launched as 'StandAlone'."
  EndIf
  End ; Une autre instance a été détectée. On met fin au programme.
  ;
Else ; Nous sommes dans la première instance.
  ; On bloque tout de suite les autres instances en initialisant le Pipe.
  ListenForPipeMessages()
EndIf
;
;
; *******************************************************************************
;                         Mise à jour des préférences
;
If PureBasicProgAdr$ = "" ; Apparemment, aucun argument n'a été reçu.
  ; On cherche l'adresse de l'application PureBasic dans les préférences.
  If OpenPreferencesWithPatience(PBBrowserPrefile$)
    FicPrincipalPB$ = ReadPreferenceString("FicPrincipalPB", "")
    PureBasicProgAdr$ = ReadPreferenceString("PureBasicProgAdr", "")
    ClosePreferences()
  EndIf
EndIf
;
; Et on enregistre l'adresse du fichier à examiner
If OpenPreferencesWithPatience(PBBrowserPrefile$)
  WritePreferenceString("FicPrincipalPB", FicPrincipalPB$)
  ClosePreferences()
EndIf
;
;
; *******************************************************************************
;
;          Vérification de l'adresse de l'application PureBasic.exe
;
;
Procedure.s FindInDirectory(directory.s, FindWhat$, Filters$ = "", Excluded$ = "")
  ;
  ; Recherche un fichier précis dans un répertoire donné
  ; ainsi que dans ses sous-répertoires.
  ;
  ; Cette procédure est récursive (elle s'appelle elle-même).
  ;
  ; Excluded$ peut contenir une liste de noms à exclure de la recherche, afin
  ; d'éviter une exploration trop longue.
  ;
  ; Filters$ peut contenir une liste de noms, auquel cas seuls les dossiers
  ; dont l'adresse comporte un de ces noms seront retenus.
  ;
  Protected PBFileListing$
  Protected directoryid, cont, DirType, file.s
  Protected exclude, nl, pExcluded$, include, pFilter$
  ;
  ;
  If directory
    If Right(directory, 1) <> "\" 
      directory + "\" 
    EndIf
    FindWhat$ = LCase(FindWhat$)
    directoryid = ExamineDirectory(#PB_Any, directory, "")
    If directoryid
      cont = 1
      Repeat
        If NextDirectoryEntry(directoryid)
          DirType = DirectoryEntryType(directoryid)
          file.s = DirectoryEntryName(directoryid)
          Select DirType 
            Case #PB_DirectoryEntry_File
              If LCase(file) = FindWhat$
                PBFileListing$ + directory + file + Chr(13)
              EndIf
            Case #PB_DirectoryEntry_Directory
              If file And file <> "." And file <> ".."
                exclude = 0
                If Excluded$
                  nl = 1
                  Repeat
                    pExcluded$ = StringField(Excluded$, nl, ",")
                    If pExcluded$ And FindString(LCase(directory + file), LCase(pExcluded$))
                      exclude = 1
                      Break
                    EndIf
                    nl + 1
                  Until pExcluded$ = ""
                EndIf
                If exclude = 0
                  If Filters$ = ""
                    include = 1
                  Else
                    include = 0
                    nl = 1
                    Repeat
                      pFilter$ = StringField(Filters$, nl, ",")
                      If pFilter$ And FindString(LCase(directory + file), LCase(pFilter$))
                        include = 1
                        Break
                      EndIf
                      nl + 1
                    Until pFilter$ = ""
                  EndIf
                  If include
                    PBFileListing$ + FindInDirectory(directory + file + "\", FindWhat$, Filters$, Excluded$)
                  EndIf
                EndIf
              EndIf 
          EndSelect 
        Else
          Cont = 0
        EndIf
      Until Cont = 0
      FinishDirectory(directoryid)
    EndIf
  EndIf
  ProcedureReturn PBFileListing$
EndProcedure
;
Procedure.s ChoosePBVersion(noms$)
  ;
  Protected.i SWindow, BList, nbNoms, button, event, i
  Protected.s choice$ ; Valeur de retour
  ;
  SWindow = OpenWindow(#PB_Any, 0, 0, 500, 280, GetTextFromCatalog("UpdatePBExeTitle"), #PB_Window_SystemMenu | #PB_Window_ScreenCentered)
  If SWindow
    StickyWindow(SWindow, 1)
    TextGadget(#PB_Any, 10, 10, WindowWidth(SWindow) - 20, 80, GetTextFromCatalog("ManyVersionsOfPB"), #PB_Text_Center)
    BList = ListViewGadget(#PB_Any, 10, 90, WindowWidth(SWindow) - 20, WindowHeight(SWindow) - 140)
    button = ButtonGadget(#PB_Any, WindowWidth(SWindow) / 2 - 50, WindowHeight(SWindow) - 35, 100, 25, GetTextFromCatalog("Choose"))

    ; Séparer les noms et ajouter à la BListe
    nbNoms = CountString(noms$, Chr(13)) + 1
    For i = 1 To nbNoms
      AddGadgetItem(BList, -1, StringField(noms$, i, Chr(13)))
    Next

    ; Sélectionner la dernière ligne par défaut
    SetGadgetState(BList, nbNoms - 1)

    ; Activer le bouton "Choisir" par défaut
    AddKeyboardShortcut(SWindow, #PB_Shortcut_Return, 1)
    ;
    ; On donne le focus à la liste
    SetActiveGadget(BList)

    Repeat
      event = WaitWindowEvent()
      If EventWindow() <> SWindow
        SetActiveWindow(SWindow)
      Else
        If event = #PB_Event_Gadget And EventGadget() = button Or event = #PB_Event_Menu Or event = #PB_Event_CloseWindow
          choice$ = GetGadgetItemText(BList, GetGadgetState(BList))
          Break
        EndIf
      EndIf
    ForEver

    CloseWindow(SWindow)
  EndIf

  ProcedureReturn choice$
EndProcedure
;
Procedure UpDatePureBasicExeAdr(Choose = 0)
  ;
  ; Recherche l'adresse du programme PureBasic.exe dans les dossiers "Program" de l'ordinateur
  ; PureBasicProgAdr$ doit avoir été précédemment déclarée comme une variable globale.
  ;
  Protected ProgFolder$, Excluded$, Filters$, PBFileListing$, ListSize, i, p
  ;
  ProgFolder$ = GetSystemFolder(#CSIDL_PROGRAM_FILES)
  ; Pour gagner du temps dans l'examen des répertoires des programmes, on exclut certains noms de dossier
  Excluded$ = "Adobe,OpenOffice,Explorer,Windows,Microsoft,Intel,Google,Chrome,DropBox,System,Languages,Catalogs,\SDK,\res\,Example,resources,Common"
  ; Seuls les dossiers comportant les mentions ci-dessous seront examinés
  Filters$ = "PB,Pure,Basic,Fantaisie,Dev,Soft,pers,home,vers"
  PBFileListing$ = FindInDirectory(ProgFolder$ + "\" , "PureBasic.exe", Filters$, Excluded$)
  If FindString(ProgFolder$, "(") ; #CSIDL_PROGRAM_FILES peut retourner "Program (x86)" alors que le ou les applications
                                 ; PureBasic se trouvent dans "Program".
                                 ; On cherche donc dans les deux dossiers.
    ProgFolder$ = Trim(Left(ProgFolder$, FindString(ProgFolder$, "(") - 1))
  Else                           ; Inversement.
    ProgFolder$ + " (x86)"
  EndIf
  PBFileListing$ + FindInDirectory(ProgFolder$ + "\" , "PureBasic.exe", Filters$, Excluded$)

  If CountString(PBFileListing$, Chr(13)) > 1
    ; On a trouvé plusieurs version de PureBasic dans les dossiers 'Program'.
    ; On va les trier par dates.
    Structure PBFilesAndDates
      ExeFName$
      ExeDate.q
    EndStructure
    ;
    Protected NewList PBFiles.PBFilesAndDates()
    ;
    ListSize = CountString(PBFileListing$, Chr(13))
    For i = 1 To ListSize
      AddElement(PBFiles())
      PBFiles()\ExeFName$ = StringField(PBFileListing$, i, Chr(13))
      ; On soustrait le longeur de l'adresse du fichier à sa date de modification
      ; Ainsi, les adresses le plus courtes seront considérées comme plus récentes
      ; et seront placées en dernier dans la liste.
      PBFiles()\ExeDate = GetFileDate(PBFiles()\ExeFName$, #PB_Date_Modified) - Len(PBFiles()\ExeFName$)
    Next
    ;
    SortStructuredList(PBFiles(), #PB_Sort_Ascending, OffsetOf(PBFilesAndDates\ExeDate), TypeOf(PBFilesAndDates\ExeDate))
    PBFileListing$ = ""
    ForEach(PBFiles())
      PBFileListing$ + PBFiles()\ExeFName$ + Chr(13)
    Next
    ; On enlève le dernier retour-chariot
    PBFileListing$ = Left(PBFileListing$, Len(PBFileListing$) - 1)
    ;
    If Choose
      ; On demande à l'utilisateur de choisir la version de PureBasic.exe à utiliser.
      PureBasicProgAdr$ = ChoosePBVersion(PBFileListing$)
    Else
      ; On prend la version la plus récente.
      p = ReverseFindString(PBFileListing$, Chr(13))
      PureBasicProgAdr$ = Mid(PBFileListing$, p + 1)
    EndIf
  Else
    PureBasicProgAdr$ = PBFileListing$
  EndIf
  ;
  If PureBasicProgAdr$ And OpenPreferencesWithPatience(PBBrowserPrefile$)
    WritePreferenceString("PureBasicProgAdr", PureBasicProgAdr$)
    ClosePreferences()
  EndIf
EndProcedure
;
Procedure ChoosePureBasicExeAdr()
  UpDatePureBasicExeAdr(1)
EndProcedure
;
If FileSize(PureBasicProgAdr$) < 2
  UpDatePureBasicExeAdr()
EndIf
;
If FileSize(PureBasicProgAdr$) < 2
  ; Si PBBrowser a été lancé depuis le fichier source (avec 'Compiler/Exécuter'),
  ; on peut récupérer l'adresse de 'PureBasic.exe' avec #PB_Compiler_Home.
  PureBasicProgAdr$ = #PB_Compiler_Home + "PureBasic.exe"
EndIf
;
If FileSize(PureBasicProgAdr$) < 2
  ; Si aucune de nos tentatives pour récupérer l'adresse de 'PureBasic.exe' n'a
  ; fonctionné, on demande à l'utilisateur de nous désigner l'adresse de l'application.
  Alert(GetTextFromCatalog("UnableToFindPB"), 1)
  PureBasicProgAdr$ = OpenFileRequester(GetTextFromCatalog("ShowPBPath"), GetSystemFolder(#CSIDL_PROGRAM_FILES) + "\PureBasic.exe", "PureBasic.exe", 0)
EndIf
;
If FileSize(PureBasicProgAdr$) > 2 And OpenPreferencesWithPatience(PBBrowserPrefile$)
  ; Si 'PureBasicProgAdr' n'existe pas encore dans les préférences, on l'enregistre.
  PrefPureBasicProgAdr$ = ReadPreferenceString("PureBasicProgAdr", "")
  If PrefPureBasicProgAdr$ = ""
    WritePreferenceString("PureBasicProgAdr", ReplaceString(PureBasicProgAdr$, Chr(34), ""))
  EndIf
  ClosePreferences()
EndIf
;
;
; *******************************************************************************
;
;             Récupération des mots clés du langage Basic de PureBasic
;
; Mots clé principaux (If, Endif, While, Wend, etc.
Global Dim PBBasicKeyword$(150), Dim PBBasicKeywordLCase$(150)
Restore PBBasicKeyWords
ne = 0
Repeat
  Read.s PBBasicKeyword$
  If PBBasicKeyword$ And PBBasicKeyword$ <> "EndPBBasicKeyWords"
    ne + 1
    PBBasicKeyword$(ne) = PBBasicKeyword$
    PBBasicKeywordLCase$(ne) = LCase(PBBasicKeyword$)
  EndIf
Until PBBasicKeyword$ = "EndPBBasicKeyWords"
;
; Liste des fonctions natives (FindString, Int, Mid, etc.)
;
Procedure UpDateNativeFunctionList(Confirm = 1)
  ;
  ; Met à jour la liste des fonctions du langage PureBasic à partir
  ; de l'index de "PureBasic.chm"
  ;
  ; MyAppDataFolder$ doit avoir été déclarée comme une variable globale
  ; et doit contenir le chemin (path) vers le dossier des données de la présente application (dans #CSIDL_COMMON_APPDATA)
  ;
  ; PureBasicProgAdr$ doit avoir été déclarée comme variable globale et contiendra l'adresse du programme PureBasic.exe
  ;
  ; PBFunctionList$ et PBFunctionListLCase$ sont des variables globales destinées à contenir la liste des fonctions
  ; natives de PureBasic.
  ;
  Protected FtKWSrce$, Line$, posf, posd, Entry$, FirstLetter$
  Protected PBFunctionListAdr$, Function$, noFile
  ;
  FtKWSrce$ = GetPathPart(PureBasicProgAdr$) + "PureBasic.chm" ; On va lire "PureBasic.chm" pour en extraire les noms des fonctions

  ;
  If FileSize(FtKWSrce$) > 2
    ;
    If ReadFile(0, FtKWSrce$)
      PBFunctionList$ = ""
      While Eof(noFile) = 0 And Line$ <> "Reference/reference.html"
        Line$ = ReadString(noFile)
        If Line$
          Posf = 0
          Repeat
            Posf = FindString(Line$, ".html", Posf + 1)
            If Posf
              Posd = Posf
              While Posd > 1 And FastMid(Line$, Posd - 1, 1) <> "/" : Posd - 1 : Wend
              Entry$ = FastMid(Line$, Posd, Posf - Posd)
              FirstLetter$ = Left(Entry$, 1)
              If FindString(Entry$, ".") = 0 And FindString(Entry$, "_") = 0 And LCase(FirstLetter$) <> FirstLetter$
                Posd - 1
                While Posd > 1 And Mid(Line$, Posd - 1, 1) <> "/" : Posd - 1 : Wend ; On cherche le dossier dans lequel se trouve la page
                Function$ = FastMid(Line$, Posd, Posf - Posd) + Chr(13)
                ; Si la fonction n'existe pas déjà dans PBFunctionList$, on l'ajoute :
                If FindString(PBFunctionList$, Function$ + Chr(13)) = 0
                  PBFunctionList$ + Function$ + Chr(13)
                EndIf
              EndIf
            EndIf
          Until Posf = 0
        EndIf
      Wend
      CloseFile(0)
      ;
      If OpenPreferencesWithPatience(PBBrowserPrefile$)
        PBFunctionListAdr$ = ReadPreferenceString("PBFunctionListAdr", MyAppDataFolder$ + "PBFunctionList.Data")
        ClosePreferences()
        TexteDansFichier(PBFunctionListAdr$, PBFunctionList$)
      EndIf
      PBFunctionListLCase$ = LCase(PBFunctionList$)
      ;
      If Confirm <> 0
        Alert(GetTextFromCatalog("FunctionsUpdateDone"))
      EndIf
    EndIf
  EndIf
EndProcedure
;
If OpenPreferencesWithPatience(PBBrowserPrefile$)
  ; Si la valeur "PBFunctionListAdr" n'existe pas déjà dans le fichier des préférences,
  ; on attribue par défaut la valeur MyAppDataFolder$ + "PBFunctionList.Data" à la
  ; variable globale 'PBFunctionListAdr$'.
  PBFunctionListAdr$ = ReadPreferenceString("PBFunctionListAdr", MyAppDataFolder$ + "PBFunctionList.Data")
  ; Et on l'enregistre dans le fichier des préférences.
  WritePreferenceString("PBFunctionListAdr", PBFunctionListAdr$)
  ClosePreferences()
Else
  ; Si on n'a pas réussi à ouvrir le fichier des préférences, on renseigne quand même
  ; la variable PBFunctionListAdr$.
  PBFunctionListAdr$ = MyAppDataFolder$ + "PBFunctionList.Data"
EndIf
;
If FileSize(PBFunctionListAdr$) > 2
  PBFunctionList$ = FileToText(PBFunctionListAdr$)
  PBFunctionListLCase$ = LCase(PBFunctionList$)
EndIf
;
Procedure UpDateAPIFunctionList()
  ;
  ; Met à jour la liste des fonctions de l'API Windows à partir
  ; du fichier "APIFunctionListing.txt" qui est normalement situé
  ; dans le dossier du compilateur de PureBasic.
  ;
  ; PureBasicProgAdr$ doit avoir été déclarée comme variable globale et contiendra l'adresse de "PureBasic.exe"
  ;
  ;
  Protected FtKWSrce$ = GetPathPart(PureBasicProgAdr$) + "Compilers\APIFunctionListing.txt"
  ;
  If FileSize(FtKWSrce$) > 2 And GetFileDate(FtKWSrce$, #PB_Date_Modified) > GetFileDate(MyAppDataFolder$ + "APIFunctionListing.txt", #PB_Date_Modified)
    ; Au début de 'PBBrowserInitialize.pb, nous avons déjà créé un fichier 'APIFunctionListing.txt'
    ; dans le dossier des données de PBBrowser. Cette version du fichier provient des données
    ; enregistrées dans le dossier 'PBBrowser resources', qui accompagne le fichier-source de
    ; 'PBBrowser'. Dans le cas où l'application fonctionne en mode 'StandAlone' (.exe), elle
    ; comporte dans son propre code une copie de 'APIFunctionListing.txt' qui lui servira de source.
    ; Cependant, si la version de 'APIFunctionListing.txt' contenue dans le dossier de PureBasic.exe est plus
    ; récente que celle que nous avions intégrée dans PBBrowser.exe, on la recopie dans le dossier
    ; des données.
    CopyFile(FtKWSrce$, MyAppDataFolder$ + "APIFunctionListing.txt")
  EndIf
EndProcedure
;
UpDateAPIFunctionList()
;
; *******************************************************************************
;
;    Récupération de certaines valeurs (les couleurs à utiliser dans le code)
;               dans le fichier des préférences de PureBasic,
;         ou dans une liste de secours, si le programme PureBasic.exe
;                        n'a pas pu être localisé.
;
Procedure GetValueFromBPPrefFile(ValueName$)
  ;
  ; Cette procédure récupère la valeur nommée 'ValueName$' dans le fichier des préférences
  ; de PureBasic.
  ;
  Static PBPrefile$
  ;
  Protected PosInText, pf, PBPrefAdr$
  ;
  If PBPrefile$ = ""
    ; On lit le fichier des préférences de PureBasic pour y récupérer des données
    PBPrefAdr$ = GetPureBasicPrefAdresse()
    If FileSize(PBPrefAdr$) > 2
      PBPrefile$ = FileToText(PBPRefAdr$, #PB_UTF8, 1)
    EndIf
    If PBPrefile$ = ""
      ; Impossible de lire le fichier préférence de PureBasic !!
      ; On va utiliser les valeurs de secours
      PBPrefile$ = PBPrefSecure$
    EndIf
  EndIf
  PosInText = FindString(PBPrefile$, ValueName$)
  If PosInText
    PosInText + Len(ValueName$)
    PosInText = FindString(PBPrefile$, "=", PosInText) + 1
    While Mid(PBPrefile$, PosInText, 1) = " " : PosInText + 1 : Wend
    pf = FindString(PBPrefile$, " ", PosInText)
    ProcedureReturn Val(Mid(PBPrefile$, PosInText, pf - PosInText))
  EndIf
EndProcedure
;
; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; CursorPosition = 476
; FirstLine = 241
; Folding = AAA-
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb