; *****************************************************************************
;
;                         Initializing PBBrowser
;
; *****************************************************************************
;
;
;        Storing and retrieving images for button images
;      and the catalog containing the application messages.
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
Procedure UpdateAppDataFolder(Force = 0)
  ; The contents of the PBBrowser AppData folder are updated the first time
  ; the 'PBBrowser.exe' application is launched with the data contained
  ; in the application itself (see the DataSection of PBBrowserDeclaration.pb).
  ; If the application is launched from the PureBasic IDE ('Compiler/Run')
  ; the data is always copied from the 'PBBrowser resources' folder
  ; that accompanies the source file of the application. Thus, when updating
  ; the 'PBBrowser resources' files, the application always has the latest
  ; versions of these files.
  ;
  If FileSize(MyAppDataFolder$ + "Next.jpg") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BNext, ?BNextEnd, MyAppDataFolder$ + "Next.jpg")
  EndIf
  If FileSize(MyAppDataFolder$ + "Previous.jpg") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BPrevious, ?BPreviousEnd, MyAppDataFolder$ + "Previous.jpg")
  EndIf
  If FileSize(MyAppDataFolder$ + "NoNext.jpg") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BNoNext, ?BNoNextEnd, MyAppDataFolder$ + "NoNext.jpg")
  EndIf
  If FileSize(MyAppDataFolder$ + "NoPrevious.jpg") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BNoPrevious, ?BNoPreviousEnd, MyAppDataFolder$ + "NoPrevious.jpg")
  EndIf
  If FileSize(MyAppDataFolder$ + "Refresh.jpg") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BRefresh, ?BRefreshEnd, MyAppDataFolder$ + "Refresh.jpg")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\Francais\PBBrowser.catalog") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BCatalogFR, ?BCatalogFREnd, MyAppDataFolder$ + "Catalogs\Francais\PBBrowser.catalog")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\Francais\IntroPBBrowser.rtf") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BIntroFR, ?BIntroFREnd, MyAppDataFolder$ + "Catalogs\Francais\IntroPBBrowser.rtf")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\English\PBBrowser.catalog") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BCatalogEN, ?BCatalogENEnd, MyAppDataFolder$ + "Catalogs\English\PBBrowser.catalog")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\English\IntroPBBrowser.rtf") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BIntroEN, ?BIntroENEnd, MyAppDataFolder$ + "Catalogs\English\IntroPBBrowser.rtf")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\Italiano\PBBrowser.catalog") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BCatalogIT, ?BCatalogITEnd, MyAppDataFolder$ + "Catalogs\Italiano\PBBrowser.catalog")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\Italiano\IntroPBBrowser.rtf") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BIntroIT, ?BIntroITEnd, MyAppDataFolder$ + "Catalogs\Italiano\IntroPBBrowser.rtf")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\Russian\PBBrowser.catalog") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BCatalogRU, ?BCatalogRUEnd, MyAppDataFolder$ + "Catalogs\Russian\PBBrowser.catalog")
  EndIf
  If FileSize(MyAppDataFolder$ + "Catalogs\Russian\IntroPBBrowser.rtf") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?BIntroRU, ?BIntroRUEnd, MyAppDataFolder$ + "Catalogs\Russian\IntroPBBrowser.rtf")
  EndIf
  
  If FileSize(MyAppDataFolder$ + "PBFunctionList.Data") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?PBFunction, ?PBFunctionEnd, MyAppDataFolder$ + "PBFunctionList.Data")
  EndIf
  
  If FileSize(MyAppDataFolder$ + "APIFunctionListing.txt") < 2 Or #PB_Compiler_Debugger Or Force
    ExtractData(?APIFunction, ?APIFunctionEnd, MyAppDataFolder$ + "APIFunctionListing.txt")
  EndIf
EndProcedure
;
UpdateAppDataFolder()
;
;
; *****************************************************************************
;
;           Reading the preferences file and choosing the language
;
If OpenPreferencesWithPatience(PBBrowserPrefile$)
  PBBLanguageFolder$ = ReadPreferenceString("PBBLanguageFolder", "")
  ClosePreferences()
  PBBFirstLaunch = 0
Else ; No preferences. The application is being launched for the first time
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
Procedure.s GetTextFromCatalogPB(SName$)
  ;
  Static AllreadyForced
  Shared PBBLanguageFolder$
  ;
  ; GTFCParam is equal to Not(AllreadyForced):
  GTFCParam = Abs(AllreadyForced - 1)
  ;    
  NText$ = GetTextFromCatalog(SName$, "", GTFCParam)
  ;
  If AllreadyForced = 0 And NText$ = "MissingMention"
    ; The first time a text is missing, try to force a copy of
    ; catalogs from binaries included 
    ;          to AppDataFolder.
    ; If the missing text is due to a new version of PBBrowser
    ; including new texts in its catalog, this is the only way to
    ; refresh the AppDataFolder.
    ;
    ; Force a copy from binaries to AppDataFolder:
    UpdateAppDataFolder(1)
    ;
    ; Now, force GetTextFromCatalog() to refresh its Static copy of the catalog:
    GetTextFromCatalog("", PBBLanguageFolder$ + "PBBrowser.catalog")
    ;
    ; Try again to get the text:
    NText$ = GetTextFromCatalog(SName$)
    ;
    ; Register that an attempt has allready been done:
    AllreadyForced = 1
  ElseIf NText$ = "MissingFile"
    PBBLanguageFolder$ = ChooseLanguage(MyAppDataFolder$ + "Catalogs\")
    If OpenPreferencesWithPatience(PBBrowserPrefile$)
      WritePreferenceString("PBBLanguageFolder", PBBLanguageFolder$)
      ClosePreferences()
    EndIf
    ;
    ; Now, force GetTextFromCatalog() to refresh its Static copy of the catalog:
    GetTextFromCatalog("", PBBLanguageFolder$ + "PBBrowser.catalog")
    ;
    ; Try again to get the text:
    NText$ = GetTextFromCatalog(SName$)
  EndIf
  ;
  ProcedureReturn NText$
EndProcedure
;
; *****************************************************************************
;
;                            Initializing the 'Catalog'
;
; All messages are stored in the 'PBBrowser.catalog' file so that they can be
; translated into different languages.
;
; The line below initializes the reference file address to be used by
; GetTextFromCatalogPB().
; Since no text is passed in the first argument, this call will have no
; other effect:
GetTextFromCatalog("", PBBLanguageFolder$ + "PBBrowser.catalog")
;
; *****************************************************************************
;
;      Retrieving the arguments passed by PureBasic when launching
;                 the application using the "Tools" menu
;
nl = 0
Repeat
  SAppArgument$ = ProgramParameter()
  nl + 1
  If SAppArgument$
    If nl = 1     ; In principle, the first argument passed by the tool call is "%HOME"
                  ; which contains the full path to the PureBasic application.
      PureBasicProgAddr$ = SAppArgument$ + "PureBasic.exe"
    ElseIf nl = 2 ; In principle, the second argument passed by the tool call is "%WORD"
                  ; which contains the word under the cursor in the PureBasic editor.
      PBUnderCursor$ = SAppArgument$
    ElseIf nl = 3 ; In principle, the third argument passed by the tool call is "%SELECTION"
                  ; which contains the position of the selection in the file.
      PBSelection$ = SAppArgument$
    ElseIf nl = 4 ; In principle, the fourth argument passed by the tool call is "%CURSOR"
                  ; which contains the position of the cursor in the currently opened
                  ; file in the PureBasic editor.
      PBCursor$ = SAppArgument$
    ElseIf nl = 5 ; In principle, the fifth argument passed by the tool call is "%FILE"
                  ; which contains the file currently opened in the PureBasic editor.
      PBBFicPrincipalPB$ = SAppArgument$
      FicActualPB$ = SAppArgument$
    Else          ; In principle, the sixth argument passed by the tool call is "%TEMPFILE"
                  ; which contains the file currently opened in the PureBasic editor.
      PBBTempFile$ = SAppArgument$
      If FicActualPB$ = ""
        FicActualPB$ = PBBTempFile$
        PBBFicPrincipalPB$ = PBBTempFile$
      EndIf
    EndIf
  EndIf
Until nl > 5
;
;
; *****************************************************************************
;
;                             Presentation Message
;
; *****************************************************************************
;
If PBBFirstLaunch = 1
  AlertWithTitle(GetTextFromCatalogPB("file:IntroPBBrowser.rtf"), GetTextFromCatalogPB("Introduction"))
EndIf

;
; *****************************************************************************
;
;         Various procedures, used throughout the program
;
; *****************************************************************************
;
Procedure.s GetPureBasicPrefAdresse()
  ; Determines the location of the PureBasic preferences file.
  ;
  ; PB Browser will retrieve data from this file, such as the
  ; colors to assign to keywords. It is also in this preferences folder
  ; that PB Browser is installed as a tool for the PureBasic IDE.
  Protected  itemid, PBPrefAddr$
  If SHGetSpecialFolderLocation_(0, #CSIDL_APPDATA, @itemid) = #NOERROR
    PBPrefAddr$ = Space(#MAX_PATH)
    SHGetPathFromIDList_(itemid, @PBPrefAddr$)
  EndIf
  PBPrefAddr$ + "\PureBasic\PureBasic.prefs"
  If FileSize(PBPrefAddr$) < 2
    PBPrefAddr$ = GetHomeDirectory() + "AppData\Roaming\PureBasic\PureBasic.prefs"
  EndIf
  If FileSize(PBPrefAddr$) < 2
    PBPrefAddr$ = GetHomeDirectory() + "AppData\PureBasic\PureBasic.prefs"
  EndIf
  If FileSize(PBPrefAddr$) < 2
    ProcedureReturn ""
  Else
    ProcedureReturn PBPrefAddr$
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
  ; Disables the main PBBrowser window to give priority
  ; to a message window or a search window.
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
  ; Reactivates the main PBBrowser window after a call
  ; to 'DisablePBBWindow()'
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
  ;         Initialization of button texts and window title
  ;     for the 'Alert()' procedure, used throughout the program
  ;
  Global  AlertText.AlertWindowTitles
  ;
  AlertText.AlertWindowTitles\WTitle$  = GetTextFromCatalogPB("Attention")
  AlertText\BOK$     = GetTextFromCatalogPB("OK")
  AlertText\BCancel$ = GetTextFromCatalogPB("Cancel")
  AlertText\BYes$    = GetTextFromCatalogPB("Yes")
  AlertText\BNo$     = GetTextFromCatalogPB("No")
  AlertText\BCopy$   = GetTextFromCatalogPB("CopyAll")
  AlertText\BSearch$ = GetTextFromCatalogPB("Search")
  AlertText\BSave$   = GetTextFromCatalogPB("Save")
  AlertText\SaveAs$  = GetTextFromCatalogPB("SaveAs")
  AlertText\TextFile$ = GetTextFromCatalogPB("TextFile")
  AlertText\SearchGadgets\WTitle$ =       GetTextFromCatalogPB("SWTitle")
  AlertText\SearchGadgets\Search$ =       GetTextFromCatalogPB("SWSearch")
  AlertText\SearchGadgets\ReplaceTitle$ = GetTextFromCatalogPB("SWReplaceTitle")
  AlertText\SearchGadgets\Replace$ =      GetTextFromCatalogPB("SWReplace")
  AlertText\SearchGadgets\ReplaceAll$ =   GetTextFromCatalogPB("SWReplaceAll")
  AlertText\SearchGadgets\Quit$ =         GetTextFromCatalogPB("Quit")
  AlertText\SearchGadgets\CaseSensitive$ = GetTextFromCatalogPB("SWCaseSensitive")
  AlertText\SearchGadgets\WholeWord$ =    GetTextFromCatalogPB("SWWholeWord")
  AlertText\SearchGadgets\InAllDocument$ = GetTextFromCatalogPB("SWInAllDocument")
  AlertText\SearchGadgets\UnableToFind$ = GetTextFromCatalogPB("SWUnableToFind")
  AlertText\SearchGadgets\ReplacementMade$ = GetTextFromCatalogPB("SWReplacementMade")
  AlertText\SearchGadgets\SearchFromStart$ = GetTextFromCatalogPB("SWSearchFromStart")
  AlertText\SearchGadgets\SearchFromEnd$ =   GetTextFromCatalogPB("SWSearchFromEnd")
  ;
  Alert("", AlertText) ; We initialize the names to be used
  ; Since no text was passed in the first argument, this call
  ; will have no other effect.
EndProcedure
;
PBBInitAlertTexts()
;
Procedure AlertInPBBWindow(AlertText$, Title$ = "", WSticky = 1, WhiteBackGround = 0, Txtleft = 0, FixedWidth = 0, TabList$ = "", YesNoCancel = #AW_AlertOnly)
  ;
  ; Désactive la fenêtre principale de PBBrowser et affiche une fenêtre
  ; contenant un message. Puis réactive la fenêtre principale lorsque
  ; l'utilisateur ferme la fenêtre de message.
  ;
  DisablePBBWindow()
  If Title$ = ""
    Protected Ret = Alert(AlertText$, 0, WSticky, WhiteBackGround, Txtleft, FixedWidth, TabList$, YesNoCancel)
  Else
    Ret = AlertWithTitle(AlertText$, Title$, WSticky, WhiteBackGround, Txtleft, FixedWidth, TabList$, YesNoCancel)
  EndIf
  EnablePBBWindow()
  ProcedureReturn Ret
EndProcedure
;
; *******************************************************************************
;
;           Procedures dedicated to reading code files
;
; *******************************************************************************
;
Structure FileListingStruct ; Structure for the PBB_FileListing
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
  ; This procedure checks if FileName$ is in the cache memory and
  ; returns #True or #False depending on the case.
  ;
  Shared PBB_FileListing()
  ;
  If ListSize(PBB_FileListing())
    ; We only iterate through the list if it has elements
    ; and if it's not already positioned on the correct item.
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
  ; Returns the content of the file 'FileName$' as a simple string.
  ;
  If FileSize(FileName$) = -1
    AlertInPBBWindow(GetTextFromCatalogPB("ErrorWithFile") + " " + FileName$)
  Else
    ; Simplifies future searches by replacing tabs with double spaces
    ; and removing any Chr(10) (LineFeed) characters.
    Protected Content$ = ReplaceString(ReplaceString(FileToText(FileName$, #PB_UTF8, 1), Chr(9), "  "), Chr(10), "")
    If Len(Content$) = 0 And FileSize(FileName$) > 2
      AlertInPBBWindow(GetTextFromCatalogPB("ErrorWithFile") + " " + FileName$)
    EndIf
    ProcedureReturn Content$
  EndIf
EndProcedure
;
Procedure GetPointedCodeFromFile(FileName$, ReturnLCase = #False)
  ;
  ; This procedure will store each read file in memory, to avoid re-reading
  ; a file that has already been read. When the content is requested a second time,
  ; the content in memory is returned.
  ;
  ; When the 'ReturnLCase' parameter is '#True', a lowercase version (LCase()) of the file
  ; will be returned. This allows case-insensitive searches
  ; on the content of the file.
  ;
  ; The return value is a pointer to a String, and this pointer
  ; remains valid outside the procedure because the 'String' is kept 'Shared' in the PBB_FileListing() list.
  ; This approach avoids duplicating the returned string in memory, which saves
  ; both time and memory.
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
  ; This version of GetCodeFromFile() returns a string
  ; instead of returning a pointer to a 'String' structure.
  Protected *SPointer.String = GetPointedCodeFromFile(FileName$, ReturnLCase)
  ProcedureReturn *SPointer\s
EndProcedure
;
; *******************************************************************************
;
;   Updating the parameters passed by the PureBasic editor, in case
;   PBBrowser is launched from the editor.
;
Procedure RecognizePointer(LineOfCode$, VarPos)
  ;
  ; When parsing the code, the "*" character poses a particular problem due to its dual function:
  ; it can represent the multiplication operator,
  ; or indicate that a variable is a pointer.
  ; We will try to identify its function by looking at what comes before it.
  ;
  Protected pTest, PreceedingWord$, dTest
  ;
  If VarPos > 1 And FastMid(LineOfCode$, VarPos - 1, 1) = "*"
    ; Skip "*" and move to the previous character:
    pTest = VarPos - 2
    ; Skip any spaces that may precede:
    While pTest And PeekC(@LineOfCode$ + (pTest - 1) * SizeOf(CHARACTER)) = 32
      pTest - 1
    Wend
    ;
    ; Look at the word that precedes:
    PreceedingWord$ = ""
    If pTest
      dTest = pTest
      While dTest And PeekC(@LineOfCode$ + (dTest - 1) * SizeOf(CHARACTER)) <> 32
        dTest - 1
      Wend
      dTest + 1
      PreceedingWord$ = "," + LCase(Fastmid(LineOfCode$, dTest, pTest - dTest + 1)) + ","
    EndIf
    ;
    If pTest = 0 Or FindString("(,@+-/*=&:~<>|[" + Chr(13), Chr(PeekC(@LineOfCode$ + (pTest - 1) * SizeOf(CHARACTER)))) Or (PreceedingWord$ And FindString(LCase(#OperatorAsWord$), PreceedingWord$))
      ; There is another operator before, or a separator indicating that the "*" sign
      ; should be considered as a pointer indicator. Thus integrate "*" into the variable name.
      ProcedureReturn #True
    EndIf
  EndIf
  ;
  ProcedureReturn #False
EndProcedure
;
If PBSelection$ And PBUnderCursor$ = "" ; Strangely, when a portion of text is selected in the editor
                                        ; of PureBasic (even a simple word), the editor doesn't return its content
                                        ; in the argument %Word. So, we will extract it from the positions
                                        ; of the selection provided by the %SELECTION argument.
  c1 = Val(StringField(PBSelection$, 2, "x"))
  c2 = Val(StringField(PBSelection$, 4, "x"))
  l1 = Val(StringField(PBSelection$, 1, "x"))
  l2 = Val(StringField(PBSelection$, 3, "x"))
  If l1 = l2 And c1 <> c2 And FicActualPB$    ; This will not work if the selection spans multiple lines
    *FileContent.String = GetPointedCodeFromFile(FicActualPB$)  ; but this is a scenario we are not interested in.
    Line$ = StringField(*FileContent\s, l2, Chr(13))
    PBUnderCursor$ = Mid(Line$, c1, c2 - c1)
  EndIf
EndIf
If PBCursor$                            ; When the cursor is over a constant name,
                                        ; the PureBasic editor returns the constant name without the '#'
                                        ; that precedes it. This is useful information that we
                                        ; will retrieve.
  l1 = Val(StringField(PBSelection$, 1, "x"))
  c1 = Val(StringField(PBSelection$, 2, "x"))
  If l1 And FicActualPB$
    *FileContent.String = GetPointedCodeFromFile(FicActualPB$)
    Line$ = StringField(*FileContent\s, l1, Chr(13))
    p = FindString(Line$, PBUnderCursor$, c1 - Len(PBUnderCursor$))
    If p And Mid(Line$, p - 1, 1) = "#"
      PBUnderCursor$ = "#" + PBUnderCursor$
    EndIf
    If RecognizePointer(Line$, p)
      PBUnderCursor$ = "*" + PBUnderCursor$
    EndIf
  EndIf
EndIf
;
; *******************************************************************************
;
;      When PBBrowser.exe is launched from the editor multiple times,
;      we do not want it to open several instances of the application.
;      We want a single instance to retrieve the parameters provided by
;      the editor and react to those parameters.
;      To achieve this, when a new instance is created, it retrieves the launch parameters,
;      communicates them to the previous instance, and then terminates its operation.
;      Therefore, we need to be able to detect if an instance of the application
;      is already running, and then communicate with it. This is done through the management
;      of a 'Pipe', a process part of the Windows API that allows dialogue
;      between multiple threads of an application or between different applications.
;
;
If CheckForOtherInstance(PureBasicProgAddr$ + Chr(13) + PBUnderCursor$ + Chr(13) + FicActualPB$ + Chr(13) + PBBTempFile$ + Chr(13))
  ; If another instance of this application is already open,
  ; we pass the arguments we just received to it
  ; and terminate the program.
  If #PB_Compiler_Debugger
    ; If we are in 'Compiled' mode, alert the user
    ; about the possible existence of a previous instance in 'StandAlone' (.exe) mode.
    PlaySound_("SystemExclamation", 0, #SND_ALIAS | #SND_ASYNC)
    Debug "PBBrowser is already launched as 'StandAlone'."
  EndIf
  End ; Another instance was detected. We terminate the program.
  ;
Else ; We are in the first instance.
  ; Block all other instances by initializing the Pipe.
  ListenForPipeMessages()
EndIf
;
;
; *******************************************************************************
;                         Update preferences
;
If PureBasicProgAddr$ = "" ; Apparently, no argument was received.
  ; We search for the address of the PureBasic application in the preferences.
  If OpenPreferencesWithPatience(PBBrowserPrefile$)
    PBBFicPrincipalPB$ = ReadPreferenceString("FicPrincipalPB", "")
    PureBasicProgAddr$ = ReadPreferenceString("PureBasicProgAdr", "")
    ClosePreferences()
  EndIf
EndIf
;
; And we save the address of the file to be examined
If OpenPreferencesWithPatience(PBBrowserPrefile$)
  WritePreferenceString("FicPrincipalPB", PBBFicPrincipalPB$)
  ClosePreferences()
EndIf
;
;
; *******************************************************************************
;
;          Verification of the address of the PureBasic.exe application
;
Procedure.s FindInDirectory(directory.s, FindWhat$, Filters$ = "", Excluded$ = "")
  ; 
  ; Searches for a specific file in a given directory
  ; as well as in its subdirectories.
  ;
  ; This procedure is recursive (it calls itself).
  ;
  ; Excluded$ can contain a list of names to exclude from the search, in order
  ; to avoid an overly long exploration.
  ;
  ; Filters$ can contain a list of names, in which case only the folders
  ; whose address contains one of these names will be considered.
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
  Protected.s choice$ ; Return Value
  ;
  SWindow = OpenWindow(#PB_Any, 0, 0, 500, 280, GetTextFromCatalogPB("UpdatePBExeTitle"), #PB_Window_SystemMenu | #PB_Window_ScreenCentered)
  If SWindow
    StickyWindow(SWindow, 1)
    TextGadget(#PB_Any, 10, 10, WindowWidth(SWindow) - 20, 80, GetTextFromCatalogPB("ManyVersionsOfPB"), #PB_Text_Center)
    BList = ListViewGadget(#PB_Any, 10, 90, WindowWidth(SWindow) - 20, WindowHeight(SWindow) - 140)
    button = ButtonGadget(#PB_Any, WindowWidth(SWindow) / 2 - 50, WindowHeight(SWindow) - 35, 100, 25, GetTextFromCatalogPB("Choose"))
    ;
    ; Separate the names and add to the BList
    nbNoms = CountString(noms$, Chr(13)) + 1
    For i = 1 To nbNoms
      AddGadgetItem(BList, -1, StringField(noms$, i, Chr(13)))
    Next
    
    ; Select the last line by default
    SetGadgetState(BList, nbNoms - 1)
    
    ; Activate the "Choose" button by default
    AddKeyboardShortcut(SWindow, #PB_Shortcut_Return, 1)
    ;
    ; Give focus to the list
    SetActiveGadget(BList)
    ;
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
  ; Searches for the address of the PureBasic.exe program in the "Program" folders of the computer
  ; PureBasicProgAddr$ must have been previously declared as a global variable.
  ;
  Protected ProgFolder$, Excluded$, Filters$, PBFileListing$, ListSize, i, p
  ;
  ProgFolder$ = GetSystemFolder(#CSIDL_PROGRAM_FILES)
  ; To save time in checking the program directories, we exclude certain folder names
  Excluded$ = "Adobe,OpenOffice,Explorer,Windows,Microsoft,Intel,Google,Chrome,DropBox,System,Languages,Catalogs,\SDK,\res\,Example,resources,Common"
  ; Only directories containing the following terms will be checked
  Filters$ = "PB,Pure,Basic,Fantaisie,Dev,Soft,pers,home,vers"
  PBFileListing$ = FindInDirectory(ProgFolder$ + "\" , "PureBasic.exe", Filters$, Excluded$)
  If FindString(ProgFolder$, "(") ; #CSIDL_PROGRAM_FILES may return "Program (x86)" while the PureBasic applications
                                 ; are in "Program". So we check both folders.
    ProgFolder$ = Trim(Left(ProgFolder$, FindString(ProgFolder$, "(") - 1))
  Else                           ; Otherwise, we check the "Program (x86)" folder.
    ProgFolder$ + " (x86)"
  EndIf
  PBFileListing$ + FindInDirectory(ProgFolder$ + "\" , "PureBasic.exe", Filters$, Excluded$)
  ;
  If CountString(PBFileListing$, Chr(13)) > 1
    ; We found multiple versions of PureBasic in the 'Program' folders.
    ; We will sort them by date.
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
      ; We subtract the length of the file path from its modification date
      ; This way, shorter file paths are considered more recent
      ; and will be placed last in the list.
      PBFiles()\ExeDate = GetFileDate(PBFiles()\ExeFName$, #PB_Date_Modified) - Len(PBFiles()\ExeFName$)
    Next
    ;
    SortStructuredList(PBFiles(), #PB_Sort_Ascending, OffsetOf(PBFilesAndDates\ExeDate), TypeOf(PBFilesAndDates\ExeDate))
    PBFileListing$ = ""
    ForEach(PBFiles())
      PBFileListing$ + PBFiles()\ExeFName$ + Chr(13)
    Next
    ; Remove the last carriage return
    PBFileListing$ = Left(PBFileListing$, Len(PBFileListing$) - 1)
    ;
    If Choose
      ; Ask the user to choose the version of PureBasic.exe to use.
      PureBasicProgAddr$ = ChoosePBVersion(PBFileListing$)
    Else
      ; Take the most recent version.
      p = ReverseFindString(PBFileListing$, Chr(13))
      PureBasicProgAddr$ = Mid(PBFileListing$, p + 1)
    EndIf
  Else
    PureBasicProgAddr$ = PBFileListing$
  EndIf
  ;
  If PureBasicProgAddr$ And OpenPreferencesWithPatience(PBBrowserPrefile$)
    WritePreferenceString("PureBasicProgAdr", PureBasicProgAddr$)
    ClosePreferences()
  EndIf
EndProcedure
;
Procedure ChoosePureBasicExeAdr()
  UpDatePureBasicExeAdr(1)
EndProcedure
;
If FileSize(PureBasicProgAddr$) < 2
  UpDatePureBasicExeAdr()
EndIf
;
If FileSize(PureBasicProgAddr$) < 2
  ; If PBBrowser was launched from the source file (using 'Compiler/Run'),
  ; we can get the address of 'PureBasic.exe' using #PB_Compiler_Home.
  PureBasicProgAddr$ = #PB_Compiler_Home + "PureBasic.exe"
EndIf
;
If FileSize(PureBasicProgAddr$) < 2
  ; If none of our attempts to get the address of 'PureBasic.exe' has
  ; worked, we ask the user to manually select the application address.
  Alert(GetTextFromCatalogPB("UnableToFindPB"), 1)
  PureBasicProgAddr$ = OpenFileRequester(GetTextFromCatalogPB("ShowPBPath"), GetSystemFolder(#CSIDL_PROGRAM_FILES) + "\PureBasic.exe", "PureBasic.exe", 0)
EndIf
;
If FileSize(PureBasicProgAddr$) > 2 And OpenPreferencesWithPatience(PBBrowserPrefile$)
  ; If 'PureBasicProgAdr' doesn't already exist in the preferences, we save it.
  PrefPureBasicProgAddr$ = ReadPreferenceString("PureBasicProgAdr", "")
  If PrefPureBasicProgAddr$ = ""
    WritePreferenceString("PureBasicProgAdr", ReplaceString(PureBasicProgAddr$, Chr(34), ""))  ; Removes quotes from the path before saving.
  EndIf
  ClosePreferences()
EndIf
;
;
; *******************************************************************************
;
;             Retrieving PureBasic language keywords
;
; Main keywords (If, Endif, While, Wend, etc.)
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
; List of native functions (FindString, Int, Mid, etc.)
;
Procedure UpDateNativeFunctionList(Confirm = 1)
  ;
  ; Updates the list of functions in the PureBasic language from
  ; the index of "PureBasic.chm"
  ;
  ; MyAppDataFolder$ must have been declared as a global variable
  ; and must contain the path to the data folder of the current application (in #CSIDL_COMMON_APPDATA)
  ;
  ; PureBasicProgAddr$ must have been declared as a global variable and will contain the address of the PureBasic.exe program
  ;
  ; PBFunctionList$ and PBFunctionListLCase$ are global variables intended to hold the list of
  ; native functions in PureBasic.
  ;
  Protected FtKWSrce$, Line$, posf, posd, Entry$, FirstLetter$
  Protected PBFunctionListAddr$, Function$, noFile
  ;
  FtKWSrce$ = GetPathPart(PureBasicProgAddr$) + "PureBasic.chm" ; "PureBasic.chm" will read to extract function names

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
                While Posd > 1 And Mid(Line$, Posd - 1, 1) <> "/" : Posd - 1 : Wend ; Looking for the folder in which the page is located
                Function$ = FastMid(Line$, Posd, Posf - Posd) + Chr(13)
                ; If the function doesn't already exist in PBFunctionList$, add it:
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
        PBFunctionListAddr$ = ReadPreferenceString("PBFunctionListAdr", MyAppDataFolder$ + "PBFunctionList.Data")
        ClosePreferences()
        TexteDansFichier(PBFunctionListAddr$, PBFunctionList$)
      EndIf
      PBFunctionListLCase$ = LCase(PBFunctionList$)
      ;
      If Confirm <> 0
        Alert(GetTextFromCatalogPB("FunctionsUpdateDone"))
      EndIf
    EndIf
  EndIf
EndProcedure
;
If OpenPreferencesWithPatience(PBBrowserPrefile$)
  ; If the value "PBFunctionListAdr" does not already exist in the preferences file,
  ; assign the default value MyAppDataFolder$ + "PBFunctionList.Data" to the
  ; global variable 'PBFunctionListAddr$'.
  PBFunctionListAddr$ = ReadPreferenceString("PBFunctionListAdr", MyAppDataFolder$ + "PBFunctionList.Data")
  ; And save it in the preferences file.
  WritePreferenceString("PBFunctionListAdr", PBFunctionListAddr$)
  ClosePreferences()
Else
  ; If failed to open the preferences file, still set the
  ; PBFunctionListAddr$ variable.
  PBFunctionListAddr$ = MyAppDataFolder$ + "PBFunctionList.Data"
EndIf
;
If FileSize(PBFunctionListAddr$) > 2
  PBFunctionList$ = FileToText(PBFunctionListAddr$)
  PBFunctionListLCase$ = LCase(PBFunctionList$)
EndIf
;
Procedure UpDateAPIFunctionList()
  ;
  ; Updates the list of Windows API functions from
  ; the "APIFunctionListing.txt" file, which is normally located
  ; in the PureBasic compiler folder.
  ;
  ; PureBasicProgAddr$ must have been declared as a global variable and will contain the address of "PureBasic.exe"
  ;
  ;
  Protected FtKWSrce$ = GetPathPart(PureBasicProgAddr$) + "Compilers\APIFunctionListing.txt"
  ;
  If FileSize(FtKWSrce$) > 2 And GetFileDate(FtKWSrce$, #PB_Date_Modified) > GetFileDate(MyAppDataFolder$ + "APIFunctionListing.txt", #PB_Date_Modified)
    ; At the beginning of 'PBBrowserInitialize.pb, we have already created an 'APIFunctionListing.txt' file
    ; in the PBBrowser data folder. This version of the file comes from data
    ; saved in the 'PBBrowser resources' folder, which accompanies the source file of
    ; 'PBBrowser'. In the case where the application is running in 'StandAlone' (.exe) mode, it
    ; includes a copy of 'APIFunctionListing.txt' within its own code to serve as its source.
    ; However, if the version of 'APIFunctionListing.txt' contained in the PureBasic.exe folder is more
    ; recent than the one we had integrated into PBBrowser.exe, it is copied into the data folder.
    CopyFile(FtKWSrce$, MyAppDataFolder$ + "APIFunctionListing.txt")
  EndIf
EndProcedure
;
UpDateAPIFunctionList()
;
; *******************************************************************************
;
;    Retrieval of certain values (the colors to use in the code)
;               from PureBasic's preferences file,
;         or from a backup list, if the PureBasic.exe program
;                        could not be located.
;
Procedure GetValueFromBPPrefFile(ValueName$)
  ;
  ; Cette procédure récupère la valeur nommée 'ValueName$' dans le fichier des préférences
  ; de PureBasic.
  ;
  Static PBPrefile$
  ;
  Protected PosInText, pf, PBPrefAddr$
  ;
  If PBPrefile$ = ""
    ; On lit le fichier des préférences de PureBasic pour y récupérer des données
    PBPrefAddr$ = GetPureBasicPrefAdresse()
    If FileSize(PBPrefAddr$) > 2
      PBPrefile$ = FileToText(PBPRefAddr$, #PB_UTF8, 1)
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
ListOfAllElementsColor(#PBBProcedure)       = GetValueFromBPPrefFile("PureKeywordColor")
ListOfAllElementsColor(#PBBStructure)       = GetValueFromBPPrefFile("StructureColor")
ListOfAllElementsColor(#PBBMacro)           = GetValueFromBPPrefFile("PureKeywordColor")
ListOfAllElementsColor(#PBBEnumeration)     = GetValueFromBPPrefFile("PureKeywordColor")
ListOfAllElementsColor(#PBBInterface)       = GetValueFromBPPrefFile("PureKeywordColor")
ListOfAllElementsColor(#PBBLabel)           = GetValueFromBPPrefFile("LabelColor")
ListOfAllElementsColor(#PBBConstante)       = GetValueFromBPPrefFile("ConstantColor")
ListOfAllElementsColor(#PBBVariable)        = GetValueFromBPPrefFile("NormalTextColor")

ListOfAllElementsColor(#PBBNativeFunction)  = GetValueFromBPPrefFile("PureKeywordColor")
ListOfAllElementsColor(#PBBBasicKeyword)    = GetValueFromBPPrefFile("BasicKeywordColor")
;
; IDE Options = PureBasic 6.10 LTS (Windows - x86)
; CursorPosition = 480
; FirstLine = 375
; Folding = uWY-
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb