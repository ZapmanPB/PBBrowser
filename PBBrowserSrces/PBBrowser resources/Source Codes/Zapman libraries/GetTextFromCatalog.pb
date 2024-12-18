Procedure.s GetTextFromCatalog(SName$, FileName$ = "", DontPanic = 0)
  ;
  ; Get a text line from a '.catalog' language file.
  ;
  Static mCatalogContent$, mFileName$
  ;
  Protected fsize, noFile, CatalogContent$, FString$
  Protected pos, posf
  ;
  If FileName$ : mFileName$ = FileName$
  Else
    FileName$ = mFileName$
    CatalogContent$ = mCatalogContent$
  EndIf
  ;
  If FileSize(FileName$) < 2
    MessageRequester("Oops!", "GetTextFromCatalog(): Catalog FileName is wrong Or Catalog is missing!" + Chr(13) + FileName$)
    ProcedureReturn "MissingFile"
  EndIf
  ;
  If Left(SName$, 5) = "file:"
    ; Le paramètre SName$ contient un nom de fichier.
    ;
    FileName$ = GetPathPart(FileName$) + Mid(SName$, 6)
    If ReadFile(0, FileName$)
      fsize = Lof(0)
      If fsize > 0
        FString$ = Space(fsize)
        ReadData(0, @FString$, fsize)
      EndIf
      CloseFile(0)
      ProcedureReturn FString$
    Else
      MessageRequester("Oops!", "GetTextFromCatalog(): Unable to read the file!" + Chr(13) + FileName$)
      ProcedureReturn SName$
    EndIf
  Else
    ; Le paramètre SName$ contient un nom de chaîne.
    ;
    If CatalogContent$ = ""
      noFile = ReadFile(#PB_Any, FileName$, #PB_File_SharedRead | #PB_File_SharedWrite)
      If noFile
        While Eof(noFile) = 0
          CatalogContent$ + ReadString(noFile, #PB_UTF8) + Chr(13)
        Wend
        CloseFile(noFile)
        CatalogContent$ = ReplaceString(CatalogContent$, Chr(9), " ")
        mCatalogContent$ = CatalogContent$
      Else
        MessageRequester("Oops!", "GetTextFromCatalog(): ReadingError While reading Catalog For '" + SName$ + "'" + Chr(13) + "File exists, but can't be open." + Chr(13) + FileName$)
        ProcedureReturn SName$
      EndIf
    EndIf
    ;
    If SName$
      If CatalogContent$
        pos = FindString(CatalogContent$, Chr(13) + SName$ + " ")
        If pos = 0
          pos = FindString(CatalogContent$, Chr(13) + SName$ + "=")
        EndIf
        If pos = 0
          pos = FindString(CatalogContent$, Chr(13) + SName$ + " ", 0, #PB_String_NoCase)
        EndIf
        If pos = 0
          pos = FindString(CatalogContent$, Chr(13) + SName$ + "=", 0, #PB_String_NoCase)
        EndIf
        If pos = 0
          If DontPanic
            ProcedureReturn "MissingMention"
          Else
            MessageRequester("Oops!", "GetTextFromCatalog(): '" + SName$ + "' can't be found in catalog!" + Chr(13) + FileName$)
          EndIf
          ProcedureReturn SName$
        Else
          pos = FindString(CatalogContent$, "=", pos) + 2
          posf = FindString(CatalogContent$, Chr(13), pos)
          FString$ = Mid(CatalogContent$, pos, posf - pos)
          If Left(FString$, 5) = "file:"
            ; Le catalog nous redirige un nom de fichier.
            ;
            FileName$ = GetPathPart(FileName$) + Trim(Mid(FString$, 6))
            If ReadFile(0, FileName$)
              fsize = Lof(0)
              If fsize > 0
                FString$ = Space(fsize)
                ReadData(0, @FString$, fsize)
              EndIf
              CloseFile(0)
            Else
              MessageRequester("Oops!", "GetTextFromCatalog(): Unable to read the file!" + Chr(13) + FileName$)
              ProcedureReturn SName$
            EndIf
          Else
            FString$ = ReplaceString(FString$, "%newline%", Chr(13))
            FString$ = ReplaceString(FString$, "%quote%", Chr(34))
            FString$ = ReplaceString(FString$, "%equal%", "=")
            FString$ = ReplaceString(FString$, "%nonbreakingspace%", Chr(160))
            FString$ = ReplaceString(FString$, "£µ|", "%")
          EndIf
          ProcedureReturn FString$
        EndIf
      Else
        MessageRequester("Oops!", "Catalog is empty!")
        ProcedureReturn SName$
      EndIf
    Else
      ProcedureReturn SName$
    EndIf
  EndIf
EndProcedure
;
; IDE Options = PureBasic 6.12 LTS (Windows - x86)
; CursorPosition = 64
; FirstLine = 57
; Folding = -
; EnableXP
; DPIAware