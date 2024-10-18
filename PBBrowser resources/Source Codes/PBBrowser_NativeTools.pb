; *******************************************************
;
;               PB Browser Native Tools
;
;
NCommand.CommandDetails\CommandName$ = "InCatalog:InstallPBBrowserTitle"
NCommand\Commandtype = 1
NCommand\CommandDescription$  = "InCatalog:InstallPBBrowserEx"
NCommand\CommandSimpleProcAdr = @InstallPBBrowserAndPrintResult()
AddCommandToList(NCommand)
;
NCommand\CommandName$ = "InCatalog:UninstallPBBrowserTitle"
NCommand\Commandtype = 1
NCommand\CommandDescription$  = "InCatalog:UnInstallPBBrowserEx"
NCommand\CommandSimpleProcAdr = @UnInstallPBBrowserAndPrintResult()
AddCommandToList(NCommand)
;
NCommand\CommandName$ = "InCatalog:UpdateFunctionsTitle"
NCommand\Commandtype = 1
NCommand\CommandDontShow = 1
NCommand\CommandDescription$  = "InCatalog:UpdateFunctionsEx"
NCommand\CommandSimpleProcAdr = @UpDateNativeFunctionList()
AddCommandToList(NCommand)
;
NCommand\CommandName$ = "InCatalog:UpdatePBExeTitle"
NCommand\Commandtype = 1
NCommand\CommandDontShow = 1
NCommand\CommandDescription$  = "InCatalog:UpdatePBExeEx"
NCommand\CommandSimpleProcAdr = @ChoosePureBasicExeAdr()
AddCommandToList(NCommand)
;
; IDE Options = PureBasic 6.11 LTS (Windows - x64)
; CursorPosition = 8
; EnableXP
; DPIAware
; UseMainFile = ..\..\PBBrowser.pb