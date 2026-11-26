;
; ****************************************************************************
;
;                              GetTextFromCatalog
;                              Zapman - Feb 2025
;
; This one-procedure library provides a method to search for a translation key
; in a file and return an expression matching that key.
;
; This file should be saved under the name "GetTextFromCatalog.pbi".
;
; ****************************************************************************
;
Procedure.s GetTextFromCatalog(SName$, FileName$ = "", DontPanic = 0)
  ;
  ; Get a text line from a '.catalog' language file. The name of this file can be specified
  ; in the 'FileName$' parameter. In this file, you'll store all the texts needed by the application
  ; in the following form:
  ;
  ; Introduction     = Introduction
  ; CopyAll          = Copy all
  ; Search           = Search
  ; Save             = Save
  ; SaveAs           = Save As...
  ; etc.
  ;
  ; The left part (before the '=' sign) is the expression key that allways remain the same regardless
  ; of the destination language. The right part (after '=') is the expression translated into the
  ; destination language. There is one file per language and you designate this file by changing the
  ; 'FileName$' parameter.
  ;
  ; If 'SName$' is an existing expression-key in the file 'FileName$', the procedure returns what has been found
  ; after the '=' sign.
  ;
  ; This Procedure attempts to handle several scenarios:
  ; 1- If the 'FileName$' parameter is empty, the procedure will use the last file name used in the previous call.
  ;    This avoids having to specify the file name on each call. You only need to do this once, on the first call,
  ;    or when the user decides to change the language of the application.
  ; 2- Into the file 'FileName$', instead of containing a translation, the right part after the '=' sign following
  ;    a translation-key can point to a .txt Or .rtf file as this:
  ;        Introduction     = file:Introduction.rtf
  ;    This allows to return a complete RTF file for some expression keys.
  ; 3- Instead of containing an expression key, the 'SName$' parameter can contain something like
  ;    'file:NameOfAfile.txt' or 'file:NameOfAfile.rtf'. In this case, the procedure will search for this file name
  ;    in the folder designated by 'FileName$'. This allows in particular to obtain a text in RTF format in cases
  ;    where this is necessary.
  ;
  ; If an expression-key is missing in the catalog, an error message is printed if the parameter 'DontPanic' is omitted
  ; or equal to zero. If 'DontPanic' is other than zero, the procedure simply returns "MissingMention".
  ;
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
    MessageRequester("Oops!", "GetTextFromCatalog(): Catalog FileName is wrong or 'Catalog' is missing!" + #CR$ + FileName$)
    ProcedureReturn "MissingFile"
  EndIf
  ;
  If Left(SName$, 5) = "file:"
    ; The SName$ parameter contains a file name.
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
      MessageRequester("Oops!", "GetTextFromCatalog(): Unable to read the file!" + #CR$ + FileName$)
      ProcedureReturn SName$
    EndIf
  Else
    ;  The SName$ parameter contains a string name.
    ;
    If CatalogContent$ = ""
      noFile = ReadFile(#PB_Any, FileName$, #PB_File_SharedRead | #PB_File_SharedWrite)
      If noFile
        While Eof(noFile) = 0
          CatalogContent$ + ReadString(noFile, #PB_UTF8) + #CR$
        Wend
        CloseFile(noFile)
        CatalogContent$ = ReplaceString(CatalogContent$, #TAB$, " ")
        mCatalogContent$ = CatalogContent$
      Else
        MessageRequester("Oops!", "GetTextFromCatalog(): ReadingError while reading Catalog for '" + SName$ + "'" + #CR$ + "File exists, but can't be open." + #CR$ + FileName$)
        ProcedureReturn SName$
      EndIf
    EndIf
    ;
    If SName$
      If CatalogContent$
        pos = FindString(CatalogContent$, #CR$ + SName$ + " ")
        If pos = 0
          pos = FindString(CatalogContent$, #CR$ + SName$ + "=")
        EndIf
        If pos = 0
          pos = FindString(CatalogContent$, #CR$ + SName$ + " ", 0, #PB_String_NoCase)
        EndIf
        If pos = 0
          pos = FindString(CatalogContent$, #CR$ + SName$ + "=", 0, #PB_String_NoCase)
        EndIf
        If pos = 0
          If DontPanic
            ProcedureReturn "MissingMention"
          Else
            MessageRequester("Oops!", "GetTextFromCatalog(): '" + SName$ + "' can't be found in catalog!" + #CR$ + FileName$)
          EndIf
          ProcedureReturn SName$
        Else
          pos = FindString(CatalogContent$, "=", pos) + 2
          posf = FindString(CatalogContent$, #CR$, pos)
          FString$ = Mid(CatalogContent$, pos, posf - pos)
          If Left(FString$, 5) = "file:"
            ; The catalog redirects us to a file name.
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
              MessageRequester("Oops!", "GetTextFromCatalog(): Unable to read the file!" + #CR$ + FileName$)
              ProcedureReturn SName$
            EndIf
          Else
            FString$ = ReplaceString(FString$, "%newline%", #CR$)
            FString$ = ReplaceString(FString$, "%quote%", #DOUBLEQUOTE$)
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
; IDE Options = PureBasic 6.20 Beta 4 (Windows - x64)
; CursorPosition = 19
; Folding = -
; EnableXP
; DPIAware