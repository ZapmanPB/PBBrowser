; ***********************************************************************
;                         MakePBBZipFromSources.pb
;                        Projet PBBrowser - Zapman
;
;  Rôle : Génère les archives de distribution 'PBBrowserSrces.zip' et
;         'PBBrowserExe.zip' dans le dossier contenant 'PBBrowser.exe'.
;
;  Règles d'organisation des sources dans 'PBBrowserSrces.zip' :
;  - 'PBBrowser.pb' est placé à la racine de l'archive ;
;  - Le dossier 'PBBrowser resources\' (Catalogs, Images, fichiers .txt/.Data
;    et 'Source Codes\') conserve intégralement sa structure d'origine ;
;  - Tout fichier inclus situé en dehors de 'PBBrowser resources\' sur le
;    disque (tel que 'Zapman libraries\' ou tout autre dossier externe) est
;    automatiquement copié dans 'PBBrowser resources\Source Codes\<DossierParent>\'
;    et ses chemins d'inclusion relatifs sont mis à jour en conséquence.
; ***********************************************************************
;
#NomProg = "PBBrowserZipGenerator"
;
XIncludeFile "..\..\Zapman libraries\ZapmanCommon.pbi"
XIncludeFile "..\..\Zapman libraries\FastStringFunctions.pbi"
;
Global MyAppDataFolder$ = GetSystemFolder(#CSIDL_COMMON_APPDATA) + "\" + #NomProg + "\"
Global PBBrowserZipGeneratorPrefile$ = MyAppDataFolder$ + #NomProg + ".prefs"
Global ZipDir$
;
Define ExePath$, SourceFolder$, MainFile$, ZipFolderPath$
;
If OpenPreferencesWithPatience(PBBrowserZipGeneratorPrefile$)
  ExePath$ = ReadPreferenceString("ExePath", "")
  ClosePreferences()
Else
  ; Première exécution : création du dossier et du fichier de préférences
  CreatePathFolders(PBBrowserZipGeneratorPrefile$)
  If CreatePreferences(PBBrowserZipGeneratorPrefile$)
    ClosePreferences()
  EndIf
EndIf
;
ZipDir$ = GetCurrentApplicationPath() + "7-Zip\"
If FileSize(ZipDir$ + "7z.exe") < 10
  ZipDir$ = GetCurrentApplicationPath() + "ComponantPrograms\7-Zip\"
EndIf
If FileSize(ZipDir$ + "7z.exe") < 10
  ZipDir$ = "C:\Program Files\7-Zip\"
EndIf
If FileSize(ZipDir$ + "7z.exe") < 10
  Alert("Il faut que le dossier contenant l'application" + #CR$ + "contienne aussi le dossier ''7-Zip'' avec ''7z.exe'' à l'intérieur" + #CR$ + "(ou que 7-Zip soit installé dans ''C:\Program Files\7-Zip\'').")
  End
EndIf
;
Procedure DeleteDirectoryWithPatience(DirToDel$, Silence = 0)
  ;
  ; Rôle : Supprime récursivement un dossier en réessayant brièvement si un outil
  ;        de synchronisation (comme Dropbox) verrouille momentanément un fichier.
  ;
  Protected err, res
  ;
  If FileSize(DirToDel$) = -2
    If Silence = 0 And AlertYesNoCancel("OK pour effacer l'ancien contenu de" + #CR$ + DirToDel$ + " ?") <> #PB_MessageRequester_Yes
      End
    EndIf
    Repeat
      res = DeleteDirectory(DirToDel$, "*.*", #PB_FileSystem_Recursive)
      If res = 0
        Delay(20)
        err + 1
      EndIf
    Until res Or err = 30
    If res = 0
      Alert("L'un des fichiers du dossier " + DirToDel$ + " est ouvert. Impossible de supprimer le dossier.")
      End
    EndIf
  EndIf
EndProcedure
;
Procedure.s CorrectAddrSyntaxeSpecial(Adress$, IsFolder = 0)
  ;
  ; Rôle : Normalise la syntaxe d'une adresse de fichier ou de dossier selon les tolérances
  ;        du compilateur PureBasic, sans altérer les parenthèses valides des chemins Windows.
  ; Paramètres : IsFolder - Si non nul, garantit la présence d'un antislash final.
  ;
  Protected mAdress$, p
  ;
  p = FindString(Adress$, ";")
  If p
    Adress$ = Left(Adress$, p - 1)
  EndIf
  Adress$ = Trim(ReplaceString(Adress$, Chr(34), ""))
  If Adress$ = ""
    ProcedureReturn ""
  EndIf
  ;
  Adress$ = ReplaceString(Adress$, "/", "\")
  Adress$ = ReplaceString(Adress$, "\.\", "\")
  Repeat
    mAdress$ = Adress$
    Adress$ = ReplaceString(Adress$, "...\", "..\")
    If Len(Adress$) > 2
      ; Préserve un éventuel préfixe réseau UNC '\\' en tête de chaîne tout en réduisant les doublons internes
      Adress$ = Left(Adress$, 1) + ReplaceString(Mid(Adress$, 2), "\\", "\")
    EndIf
  Until mAdress$ = Adress$
  If Left(Adress$, 2) = ".\"
    Adress$ = Mid(Adress$, 3)
  EndIf
  If IsFolder And Right(Adress$, 1) <> "\"
    Adress$ + "\"
  EndIf
  ProcedureReturn Adress$
EndProcedure
;
Procedure.s FullAddressToRelativeAddress(Addr$, RefAddr$)
  ;
  ; Rôle : Convertit une adresse disque complète 'Addr$' en chemin relatif calculé
  ;        depuis le dossier de l'adresse de référence 'RefAddr$' (insensible à la casse).
  ;
  Protected PosInText
  Protected RefLow$, AddrLow$
  ;
  If LCase(Left(Addr$, 1)) <> LCase(Left(RefAddr$, 1))
    ; Les deux adresses sont sur des lecteurs différents : calcul relatif impossible
    ProcedureReturn Addr$
  EndIf
  RefAddr$ = ReplaceString(GetPathPart(ReplaceString(RefAddr$, "/", "\")), "\", "/")
  RefAddr$ = ReplaceString(RefAddr$, "//", "/")
  If Right(RefAddr$, 1) <> "/"
    RefAddr$ + "/"
  EndIf
  Addr$ = ReplaceString(ReplaceString(Addr$, "\", "/"), "//", "/")
  ;
  If FindString(FastLeft(Addr$, 3), ":")
    PosInText = 1
    RefLow$ = LCase(RefAddr$)
    AddrLow$ = LCase(Addr$)
    ; Détermine la longueur de la racine commune aux deux chemins
    While PosInText <= Len(RefLow$) And PosInText <= Len(AddrLow$) And FastMid(RefLow$, PosInText, 1) = FastMid(AddrLow$, PosInText, 1)
      PosInText + 1
    Wend
    ; Recule jusqu'au dernier '/' au cas où deux noms de dossiers commencent par les mêmes lettres
    While PosInText > 1 And FastMid(Addr$, PosInText - 1, 1) <> "/"
      PosInText - 1
    Wend
    Addr$ = Mid(Addr$, PosInText)
    RefAddr$ = Mid(RefAddr$, PosInText)
    ; Ajoute une remontée '../' pour chaque sous-dossier restant dans l'adresse de référence
    While RefAddr$
      PosInText = FindString(RefAddr$, "/")
      If PosInText
        Addr$ = "../" + Addr$
        RefAddr$ = Mid(RefAddr$, PosInText + 1)
      Else
        RefAddr$ = ""
      EndIf
    Wend
  EndIf
  ProcedureReturn ReplaceString(Addr$, "/", "\")
EndProcedure
;
Procedure.s RelativeAddressToFullAddress(RelAddr$, RefAddr$, IsFolder = 0)
  ;
  ; Rôle : Convertit une adresse relative 'RelAddr$' en adresse disque complète résolue
  ;        depuis le dossier de 'RefAddr$', en développant toutes les séquences '..\'.
  ;
  Protected NewAddr$, PosInText, pf
  ;
  RefAddr$ = CorrectAddrSyntaxeSpecial(RefAddr$)
  RelAddr$ = CorrectAddrSyntaxeSpecial(RelAddr$, IsFolder)
  If RelAddr$ = ""
    ProcedureReturn GetPathPart(RefAddr$)
  EndIf
  ;
  If FindString(Left(RelAddr$, 3), ":") Or Left(RelAddr$, 2) = "\\"
    ; L'adresse transmise est déjà absolue
    NewAddr$ = RelAddr$
  ElseIf Left(RelAddr$, 1) = "\"
    ; Adresse relative à la racine du lecteur de 'RefAddr$'
    NewAddr$ = Left(RefAddr$, 2) + RelAddr$
  Else
    NewAddr$ = CorrectAddrSyntaxeSpecial(GetPathPart(RefAddr$) + RelAddr$, IsFolder)
  EndIf
  ;
  ; Résout séquentiellement chaque remontée '..\' en retirant le dossier parent correspondant
  Repeat
    PosInText = FindString(NewAddr$, "\..\")
    If PosInText
      pf = PosInText + 4
      While PosInText > 1 And FastMid(NewAddr$, PosInText - 1, 1) <> "\"
        PosInText - 1
      Wend
      If PosInText < 4
        ; Empêche de remonter au-delà de la racine du lecteur (ex. 'D:\')
        PosInText = 4
      EndIf
      If pf <= Len(NewAddr$)
        NewAddr$ = FastLeft(NewAddr$, PosInText - 1) + FastMid(NewAddr$, pf)
      Else
        NewAddr$ = FastLeft(NewAddr$, PosInText - 1)
      EndIf
    EndIf
  Until PosInText = 0
  ProcedureReturn NewAddr$
EndProcedure
;
Procedure.s GetZipDestAddress(DiskAddr$, ZipRoot$, SourceRoot$, MainFile$, IsFolder = 0)
  ;
  ; Rôle : Détermine l'adresse de destination d'un fichier ou d'un dossier dans 'ZipRoot$' :
  ;        - 'MainFile$' (PBBrowser.pb) est placé à la racine du ZIP ;
  ;        - Tout élément déjà situé sous 'PBBrowser resources\' garde son chemin relatif intact ;
  ;        - Tout élément externe à 'PBBrowser resources\' est rangé dans
  ;          'PBBrowser resources\Source Codes\<NomDuDossierParent>\'.
  ;
  Protected ResRoot$, DestCodes$, ParentDir$, ParentName$
  ;
  DiskAddr$ = CorrectAddrSyntaxeSpecial(DiskAddr$, IsFolder)
  If IsFolder = 0 And LCase(DiskAddr$) = LCase(MainFile$)
    ProcedureReturn ZipRoot$ + GetFilePart(DiskAddr$)
  EndIf
  ;
  ResRoot$ = SourceRoot$ + "PBBrowser resources\"
  If LCase(Left(DiskAddr$, Len(ResRoot$))) = LCase(ResRoot$)
    ProcedureReturn ZipRoot$ + "PBBrowser resources\" + Mid(DiskAddr$, Len(ResRoot$) + 1)
  EndIf
  ;
  DestCodes$ = ZipRoot$ + "PBBrowser resources\Source Codes\"
  ParentDir$ = GetPathPart(DiskAddr$)
  If LCase(ParentDir$) <> LCase(SourceRoot$)
    ParentName$ = GetFilePart(RTrim(ParentDir$, "\"))
    If FindString(ParentName$, ":") = 0 And ParentName$ <> ""
      DestCodes$ + ParentName$ + "\"
    EndIf
  EndIf
  ProcedureReturn DestCodes$ + GetFilePart(DiskAddr$)
EndProcedure
;
Procedure CopyIncludedFiles(ZipRoot$, StartFile$, SourceRoot$, MainFile$)
  ;
  ; Rôle : Copie 'StartFile$' vers sa destination dans 'ZipRoot$', analyse ses directives
  ;        IncludePath, IncludeFile et XIncludeFile (avec un IncludePath local au fichier courant),
  ;        met à jour les chemins relatifs modifiés par la relocalisation et traite récursivement
  ;        chaque fichier inclus sans jamais dupliquer le traitement d'un fichier déjà visité.
  ;
  Static NewMap VisitedFiles.i()
  Protected FileEncoding, FileModified, HasBOM, p, dp, IsX, AfterKw, LineEnd, q1, q2, CommentPos, err, CopyOK
  Protected FileName$, NFileName$, FileContent$, EOL$, CurIncDir$, DestIncDir$
  Protected SubKw$, RawPath$, OldIncFile$, NIncFile$, NewRel$
  ;
  FileName$ = CorrectAddrSyntaxeSpecial(StartFile$)
  If FileName$ = ""
    Alert("Erreur : nom de fichier source vide dans CopyIncludedFiles !")
    End
  EndIf
  If LCase(FileName$) = LCase(MainFile$)
    ClearMap(VisitedFiles())
  EndIf
  If FindMapElement(VisitedFiles(), LCase(FileName$))
    ; Fichier déjà copié et analysé lors d'une inclusion antérieure
    ProcedureReturn
  EndIf
  VisitedFiles(LCase(FileName$)) = #True
  ;
  If FileSize(FileName$) < 0
    Alert("Fichier source introuvable :" + #CR$ + FileName$)
    End
  EndIf
  ;
  NFileName$ = GetZipDestAddress(FileName$, ZipRoot$, SourceRoot$, MainFile$)
  If LCase(Left(NFileName$, Len(ZipRoot$))) <> LCase(ZipRoot$)
    Alert("Sécurité : la destination calculée sort du dossier temporaire ZIP !" + #CR$ + NFileName$)
    End
  EndIf
  CreatePathFolders(NFileName$)
  For err = 1 To 20
    CopyOK = CopyFile(FileName$, NFileName$)
    If CopyOK
      Break
    EndIf
    Delay(20)
  Next
  If CopyOK = 0
    Alert("Copie de fichier impossible !" + #CR$ + "Source : " + FileName$ + #CR$ + "Destination : " + NFileName$)
    End
  EndIf
  ;
  ; Sauvegarde immédiatement l'encodage et la présence éventuelle d'un BOM avant tout appel récursif
  FileContent$ = FileToText(FileName$, -1, 2)
  FileEncoding = BOM
  HasBOM = FileHasBOM
  ;
  ; Mémorise le format de fin de ligne d'origine (#CRLF$, #LF$ ou #CR$) puis unifie en #CR$ pour l'analyse
  If FindString(FileContent$, #CRLF$)
    EOL$ = #CRLF$
  ElseIf FindString(FileContent$, #LF$)
    EOL$ = #LF$
  Else
    EOL$ = #CR$
  EndIf
  FileContent$ = ReplaceString(ReplaceString(FileContent$, #CRLF$, #CR$), #LF$, #CR$)
  ;
  ; Par défaut, le chemin relatif d'inclusion d'un fichier PureBasic est celui de sa propre adresse
  CurIncDir$ = GetPathPart(FileName$)
  DestIncDir$ = GetPathPart(NFileName$)
  ;
  Repeat
    p = FindString(FileContent$, "include", p + 1, #PB_String_NoCase)
    If p
      dp = p - 1
      IsX = 0
      If dp >= 1 And (FastMid(FileContent$, dp, 1) = "X" Or FastMid(FileContent$, dp, 1) = "x")
        IsX = 1
        dp - 1
      EndIf
      While dp >= 1 And (FastMid(FileContent$, dp, 1) = " " Or FastMid(FileContent$, dp, 1) = #TAB$)
        dp - 1
      Wend
      ;
      ; Vérifie que la directive figure bien en début de ligne (et non dans un commentaire ou une expression)
      If dp = 0 Or FastMid(FileContent$, dp, 1) = #CR$
        SubKw$ = LCase(Mid(FileContent$, p + 7, 4))
        AfterKw = p + 11
        If AfterKw <= Len(FileContent$) And ((SubKw$ = "path" And IsX = 0) Or SubKw$ = "file")
          If FindString(" " + #TAB$ + "(" + Chr(34), FastMid(FileContent$, AfterKw, 1))
            LineEnd = FindString(FileContent$, #CR$, AfterKw)
            If LineEnd = 0
              LineEnd = Len(FileContent$) + 1
            EndIf
            q1 = FindString(FileContent$, Chr(34), AfterKw)
            If q1 > 0 And q1 < LineEnd
              q2 = FindString(FileContent$, Chr(34), q1 + 1)
              CommentPos = FindString(FileContent$, ";", AfterKw)
              ;
              ; Cible uniquement la chaîne littérale située avant tout commentaire de fin de ligne
              If q2 > q1 And q2 < LineEnd And (CommentPos = 0 Or CommentPos > q1)
                RawPath$ = Mid(FileContent$, q1 + 1, q2 - q1 - 1)
                If SubKw$ = "path"
                  If Trim(RawPath$) = ""
                    CurIncDir$ = GetPathPart(FileName$)
                    DestIncDir$ = GetPathPart(NFileName$)
                    NewRel$ = ""
                  Else
                    CurIncDir$ = RelativeAddressToFullAddress(RawPath$, FileName$, 1)
                    DestIncDir$ = GetZipDestAddress(CurIncDir$, ZipRoot$, SourceRoot$, MainFile$, 1)
                    NewRel$ = RTrim(FullAddressToRelativeAddress(DestIncDir$, NFileName$), "\")
                  EndIf
                Else
                  OldIncFile$ = RelativeAddressToFullAddress(RawPath$, CurIncDir$)
                  If FileSize(OldIncFile$) < 0
                    Alert("Fichier inclus introuvable (référencé dans " + GetFilePart(FileName$) + ") :" + #CR$ + OldIncFile$)
                    End
                  EndIf
                  NIncFile$ = GetZipDestAddress(OldIncFile$, ZipRoot$, SourceRoot$, MainFile$)
                  NewRel$ = FullAddressToRelativeAddress(NIncFile$, DestIncDir$)
                  CopyIncludedFiles(ZipRoot$, OldIncFile$, SourceRoot$, MainFile$)
                EndIf
                ;
                ; Réécrit uniquement le contenu entre guillemets (préserve parenthèses et commentaires éventuels)
                If LCase(CorrectAddrSyntaxeSpecial(RawPath$)) <> LCase(CorrectAddrSyntaxeSpecial(NewRel$))
                  FileContent$ = Left(FileContent$, q1) + NewRel$ + Mid(FileContent$, q2)
                  p = q1 + Len(NewRel$) + 1
                  FileModified = #True
                Else
                  p = q2
                EndIf
              EndIf
            EndIf
          EndIf
        EndIf
      EndIf
    EndIf
  Until p = 0
  ;
  ; N'écrit sur le disque que si au moins un chemin d'inclusion a dû être adapté dans ce fichier
  If FileModified
    If EOL$ <> #CR$
      FileContent$ = ReplaceString(FileContent$, #CR$, EOL$)
    EndIf
    If TexteDansFichier(NFileName$, FileContent$, FileEncoding, 0, HasBOM) = 0
      Alert("Impossible d'enregistrer les chemins mis à jour dans :" + #CR$ + NFileName$)
      End
    EndIf
  EndIf
EndProcedure
;
Procedure CreateZipArchive(ZipFile$, TargetSpec$, WorkDir$)
  ;
  ; Rôle : Supprime toute version antérieure de 'ZipFile$' (afin que 7-Zip ne conserve jamais
  ;        d'entrées obsolètes), exécute la compression et contrôle le code de sortie.
  ;
  Protected err, DelOK, zipprog, ExitCode
  ;
  If FileSize(ZipFile$) >= 0
    For err = 1 To 20
      DelOK = DeleteFile(ZipFile$)
      If DelOK
        Break
      EndIf
      Delay(25)
    Next
    If DelOK = 0
      Alert("Impossible de remplacer l'archive existante (fichier ouvert ou verrouillé) :" + #CR$ + ZipFile$)
      ProcedureReturn #False
    EndIf
  EndIf
  ;
  zipprog = RunProgram(ZipDir$ + "7z.exe", "a -tzip " + Chr(34) + ZipFile$ + Chr(34) + " " + Chr(34) + TargetSpec$ + Chr(34), WorkDir$, #PB_Program_Open | #PB_Program_Hide)
  If zipprog = 0
    Alert("Impossible de lancer 7z.exe pour créer :" + #CR$ + ZipFile$)
    ProcedureReturn #False
  EndIf
  WaitProgram(zipprog)
  ExitCode = ProgramExitCode(zipprog)
  CloseProgram(zipprog)
  If ExitCode <> 0 Or FileSize(ZipFile$) <= 0
    Alert("Erreur lors de la compression 7-Zip (code " + Str(ExitCode) + ") :" + #CR$ + ZipFile$)
    ProcedureReturn #False
  EndIf
  ProcedureReturn #True
EndProcedure
;
; ***********************************************************************
;                           Programme principal
; ***********************************************************************
;
ExePath$ = PathRequester("Ouvrir le dossier contenant PBBrowser.pb (CreateMyEbook)", ExePath$)
If ExePath$
  ExePath$ = CorrectAddrSyntaxeSpecial(ExePath$, 1)
  If OpenPreferencesWithPatience(PBBrowserZipGeneratorPrefile$)
    WritePreferenceString("ExePath", ExePath$)
    ClosePreferences()
  EndIf
  ;
  If FileSize(ExePath$ + "PBBrowser.exe") < 2
    Alert("Je ne trouve pas PBBrowser.exe dans :" + #CR$ + ExePath$)
    End
  EndIf
  ;
  SourceFolder$ = CorrectAddrSyntaxeSpecial(GetCurrentApplicationPath(), 1)
  If FileSize(SourceFolder$ + "PBBrowser.pb") < 2
    SourceFolder$ = ExePath$
  EndIf
  MainFile$ = SourceFolder$ + "PBBrowser.pb"
  If FileSize(MainFile$) < 2
    Alert("Je ne trouve pas PBBrowser.pb dans :" + #CR$ + SourceFolder$)
    End
  EndIf
  If FileSize(SourceFolder$ + "PBBrowser resources") <> -2
    Alert("Je ne trouve pas le dossier 'PBBrowser resources' dans :" + #CR$ + SourceFolder$)
    End
  EndIf
  ;
  ZipFolderPath$ = ExePath$ + "PBBrowserZIP\"
  DeleteDirectoryWithPatience(ZipFolderPath$)
  ;
  ; Copie l'intégralité du dossier 'PBBrowser resources' (Catalogs, Images, Data et Source Codes)
  CreatePathFolders(ZipFolderPath$ + "PBBrowser resources\")
  If CopyDirectory(SourceFolder$ + "PBBrowser resources\", ZipFolderPath$ + "PBBrowser resources\", "*.*", #PB_FileSystem_Recursive) = 0
    Alert("La copie du dossier 'PBBrowser resources' vers le dossier temporaire a échoué !")
    End
  EndIf
  DeleteFile(ZipFolderPath$ + "PBBrowser resources\Source Codes\MakePBBZipFromSources.pb")
  DeleteFile(ZipFolderPath$ + "PBBrowser resources\Source Codes\MakePBBZipFromSources.exe")
  DeleteFile(ZipFolderPath$ + "PBBrowser resources\Source Codes\MakePBBZipFromSourcesOld.pb")
  ;
  ; Copie PBBrowser.pb et intègre récursivement tous les fichiers inclus (internes et externes)
  CopyIncludedFiles(ZipFolderPath$, MainFile$, SourceFolder$, MainFile$)
  ;
  ; Crée les deux archives .zip propres puis supprime le dossier de travail temporaire
  If CreateZipArchive(ExePath$ + "PBBrowserSrces.zip", "*", ZipFolderPath$)
    DeleteDirectoryWithPatience(ZipFolderPath$, 1)
    If CreateZipArchive(ExePath$ + "PBBrowserExe.zip", "PBBrowser.exe", ExePath$)
      Alert("Création des ZIPs terminée." + #CR$ + "'PBBrowserExe.zip' et 'PBBrowserSrces.zip' figurent à présent dans :" + #CR$ + ExePath$)
    EndIf
  EndIf
EndIf
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 374
; FirstLine = 368
; Folding = --
; EnableXP
; DPIAware