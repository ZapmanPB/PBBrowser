;
;***************************************************************************
;                           GetOpenProject.pb
;                         by Zapman - Nov. 2024
;
;        List of functions to check if a project is open by the IDE.
;
;    The PureBasic IDE has very few command lines to interract with it.
;       There's no one to ask it if any particular file is open.
;   The very complicated following set of functions has only one target:
;   -> Get the name of the project actually open by the IDE (if any).
;
;   To do so, the strategy is to:
;   • Get the PID (Process IDentifier) of a running version of PureBasic from its disk address with 'GetProcessIdFromFilePath(filePath.s)';
;   • Examine all windows of this process, looking for one containing "PureBasic" in it's name, with 'GetPureBasicMainWindowNumber(filePath.s)';
;   • Examine all child windows of the first one, looking for one containing ".pbp" in it's name, with 'FindChildFromPartialName(hWnd, *searchName.string)';
;   • Extract the full name of the found child window and then, extract the .pbp file name from it, with 'GetPBPOpenFile(PBAddress$)'.
;
;   It is clearly possible that the IDE evolutions may compromise the result of this strategy in the future.
;   But I didn't find a better way to do the job.
;
;***************************************************************************
;
ProcedureC EnumPureBasicWindowsProc(hWnd, lParam)
  ;
  ; Callback function looking for main PureBasic window.
  ;
  Shared MainPBWindowNumber
  ;
  Protected pid
  ;
  ; Get window PID
  GetWindowThreadProcessId_(hWnd, @pid)
  
  ; Compare the PID with PureBasic PID given by lParam
  If pid = lParam
    Protected windowTitle.s = Space(256)
    ; Get the window title
    GetWindowText_(hWnd, @windowTitle, 255)
    ;
    If FindString(windowTitle, "PureBasic")
      MainPBWindowNumber = hWnd
    EndIf
  EndIf
  
  ProcedureReturn 1
EndProcedure
;
Prototype protoQueryFullProcessImageName(hProcess, dwFlags, lpExeName, lpdwSize)
;
Procedure GetProcessIdFromFilePath(filePath.s)
  ;
  ; Get the processus PID from the application path.
  ; PID will only be returned if the application is running.
  ;
  Protected hSnapshot, proc.PROCESSENTRY32, pid = 0
  Protected QueryFullProcessImageName.protoQueryFullProcessImageName
  Protected hLib, modulehandle, lentgh, FullFileName$
  ;
  hLib = OpenLibrary(#PB_Any, "kernel32.dll")
  If IsLibrary(hLib)
    QueryFullProcessImageName = GetFunction(hLib, "QueryFullProcessImageNameW")
  Else
    ProcedureReturn 0
  EndIf 
  ;
  hSnapshot = CreateToolhelp32Snapshot_(#TH32CS_SNAPPROCESS, 0)
  If hSnapshot <> #INVALID_HANDLE_VALUE
    proc\dwSize = SizeOf(PROCESSENTRY32)
    If Process32First_(hSnapshot, @proc)
      Repeat
        If LCase(PeekS(@proc\szExeFile)) = LCase(GetFilePart(filePath)) ; Compare file name without path
          pid = proc\th32ProcessID
          modulehandle = OpenProcess_(#PROCESS_QUERY_INFORMATION, #False, pid)
          If modulehandle
            length = 1024
            FullFileName$ = Space(length)
            QueryFullProcessImageName(modulehandle, #Null, @FullFileName$, @length)
            CloseHandle_(modulehandle)
          EndIf
          If LCase(FullFileName$) = LCase(filePath) ; Compare file name with path
            Break
          Else
            pid = 0
          EndIf
        EndIf
      Until Process32Next_(hSnapshot, @proc) = 0
    EndIf
    CloseHandle_(hSnapshot)
  EndIf
  ProcedureReturn pid
EndProcedure
;
Procedure GetPureBasicMainWindowNumber(filePath.s)
  ;
  ; Function to get the main window of PureBasic for a given filePath
  ;
  Shared MainPBWindowNumber
  ; Find the PID based on the file path
  Protected pid = GetProcessIdFromFilePath(filePath)
  ;
  If pid
    EnumWindows_(@EnumPureBasicWindowsProc(), pid)
  EndIf
  ProcedureReturn MainPBWindowNumber
EndProcedure
;
Procedure FindChildFromPartialName(hWnd, *searchName.string)
  ;
  Protected windowTitle.s = Space(256)
  ;
  GetWindowText_(hWnd, @windowTitle, 255)
  ;
  If windowTitle And FindString(windowTitle, *searchName\s)
    *searchName\s = windowTitle
    ProcedureReturn 0 ; Stop
  EndIf
    
  ProcedureReturn 1 ; Continue
EndProcedure
;
Procedure.s GetPBPOpenFile(PBAddress$)
  ;
  Protected targetWindowTitle.string\s = ".pbp"
  Protected SFind$ = targetWindowTitle\s, Result$, pd, pf
  Protected Win = GetPureBasicMainWindowNumber(PBAddress$)
  ;
  If Win
    EnumChildWindows_(Win, @FindChildFromPartialName(), targetWindowTitle)
    If targetWindowTitle\s <> SFind$
      pd = FindString(LCase(targetWindowTitle\s), Chr(13)) + 1
      pf = FindString(LCase(targetWindowTitle\s), SFind$, pf) + Len(SFind$)
      Result$ = Mid(targetWindowTitle\s, pd, pf - pd)
      pd = FindString(Result$, ": ")
      Result$ = Mid(Result$, pd + 2)
      ProcedureReturn Result$
    EndIf
  EndIf
EndProcedure
  
;Debug GetPBPOpenFile("C:\Program Files (x86)\PureBasic ttVersions\PureBasic 6.1\PureBasic.exe")
; IDE Options = PureBasic 6.10 LTS (Windows - x86)
; CursorPosition = 12
; Folding = +
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb