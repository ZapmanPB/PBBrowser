; *******************************************************
;
;               PB Browser Native Tools
;
;
Define NCommand.CommandDetails\CommandName$ = "InCatalog:InstallPBBrowserTitle"
NCommand\Commandtype = 1
NCommand\CommandDontShow = 0
NCommand\CommandDescription$  = "InCatalog:InstallPBBrowserEx"
NCommand\CommandSimpleProcAddr = @InstallPBBTool()
AddCommandToList(NCommand)
;
NCommand\CommandName$ = "InCatalog:UninstallPBBrowserTitle"
NCommand\Commandtype = 1
NCommand\CommandDontShow = 0
NCommand\CommandDescription$  = "InCatalog:UnInstallPBBrowserEx"
NCommand\CommandSimpleProcAddr = @UnInstallPBBrowserAndPrintResult()
AddCommandToList(NCommand)
;
; NCommand\CommandName$ = "InCatalog:UpdateFunctionsTitle"
; NCommand\Commandtype = 1
; NCommand\CommandDontShow = 1
; NCommand\CommandDescription$  = "InCatalog:UpdateFunctionsEx"
; NCommand\CommandSimpleProcAddr = @UpdatePBNativeFunctionList()
; AddCommandToList(NCommand)
;
; NCommand\CommandName$ = "InCatalog:UpdatePBExeTitle"
; NCommand\Commandtype = 1
; NCommand\CommandDontShow = 1
; NCommand\CommandDescription$  = "InCatalog:UpdatePBExeEx"
; NCommand\CommandSimpleProcAddr = @ChoosePureBasicExeAdr()
; AddCommandToList(NCommand)


;
; IDE Options = PureBasic 6.41 (Windows - x64)
; CursorPosition = 24
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb