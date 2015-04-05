!include UpgradeDLL.nsi
!include AddSharedDLL.nsi
!include un.RemoveSharedDLL.nsi

!define MUI_PRODUCT "ALC Touch" ;Define your own software name here
!define MUI_VERSION "1.02" ;Define your own software version here

!include "MUI.nsh"

;--------------------------------
;Configuration

	;Do A CRC Check
	CRCCheck On

	;Output File Name
	OutFile "ALCTouch-v1.02-Install.exe"

	;The Default Installation Directory
	InstallDir "$PROGRAMFILES\ALC Touch"

	;Remember install folder
	InstallDirRegKey HKCU "Software\ALC-WBC\${MUI_PRODUCT}" ""

;--------------------------------
;Modern UI Configuration

	!define MUI_WELCOMEPAGE
	!define MUI_LICENSEPAGE
	!define MUI_DIRECTORYPAGE
	!define MUI_FINISHPAGE
	!define MUI_FINISHPAGE_RUN "$INSTDIR\ALC Touch.exe"
	
	!define MUI_ABORTWARNING
 
	!define MUI_UNINSTALLER
	!define MUI_UNCONFIRMPAGE

;--------------------------------
;Languages
 
	!insertmacro MUI_LANGUAGE "French"

;--------------------------------
;Data

	;License Data
	LicenseData /LANG=${LANG_FRENCH} "Licence.txt"

;--------------------------------
;Installer Sections

Section "ALC Touch"
	;Install Files
	SetOutPath $INSTDIR
	File "ALC Touch.exe"
	File "Aide.html"
	File "${NSISDIR}\Contrib\Icons\modern-uninstall.ico"

	; Write the uninstall keys for Windows
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ALC Touch" "DisplayName" "ALC Touch (remove only)"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ALC Touch" "UninstallString" "$INSTDIR\Uninst.exe"
	WriteUninstaller "Uninst.exe"
SectionEnd

Section "VB Runtime DLLs"
	!insertmacro UpgradeDLL VBRun60sp5\Asycfilt.dll $SYSDIR\Asycfilt.dll
	!insertmacro UpgradeDLL VBRun60sp5\Comcat.dll $SYSDIR\Comcat.dll
	!insertmacro UpgradeDLL VBRun60sp5\Msvbvm60.dll $SYSDIR\Msvbvm60.dll
	!insertmacro UpgradeDLL VBRun60sp5\Oleaut32.dll $SYSDIR\Oleaut32.dll
	!insertmacro UpgradeDLL VBRun60sp5\Olepro32.dll $SYSDIR\Olepro32.dll
	!define UPGRADEDLL_NOREGISTER
		!insertmacro UpgradeDLL VBRun60sp5\Stdole2.tlb $SYSDIR\Stdole2.tlb
	!undef UPGRADEDLL_NOREGISTER
	!insertmacro UpgradeDLL VB6FR.dll $SYSDIR\VB6FR.dll
	!insertmacro UpgradeDLL MSCOMCTL.OCX $SYSDIR\MSCOMCTL.OCX
	;Skip shared count increasing if already done once for this application
	IfFileExists $INSTDIR\VIM32.exe skipAddShared
		Push $SYSDIR\Asycfilt.dll
		Call AddSharedDLL
		Push $SYSDIR\Comcat.dll
		Call AddSharedDLL
		Push $SYSDIR\Msvbvm60.dll
		Call AddSharedDLL
		Push $SYSDIR\Oleaut32.dll
		Call AddSharedDLL
		Push $SYSDIR\Olepro32.dll
		Call AddSharedDLL
		Push $SYSDIR\Stdole2.tlb
		Call AddSharedDLL
		Push $SYSDIR\VB6FR.dll
		Call AddSharedDLL
		Push $SYSDIR\MSCOMCTL.OCX
		Call AddSharedDLL
	skipAddShared:
SectionEnd

Section "Shortcuts"
	;Add Shortcuts
	CreateDirectory "$SMPROGRAMS\ALC Touch"
	CreateShortCut "$SMPROGRAMS\ALC Touch\ALC Touch.lnk" "$INSTDIR\ALC Touch.exe" "" "$INSTDIR\ALC Touch.exe" 0
	CreateShortCut "$SMPROGRAMS\ALC Touch\Help - Aide.lnk" "$INSTDIR\Aide.html" "" "$INSTDIR\Aide.html" 0
	CreateShortCut "$SMPROGRAMS\ALC Touch\Uninstall.lnk" "$INSTDIR\Uninst.exe" "" "$INSTDIR\modern-uninstall.ico" 0
SectionEnd

Section Uninstall
	;Delete Files
	Delete "$INSTDIR\ALC Touch.exe"
	Delete "$INSTDIR\Aide.html"
	Delete "$INSTDIR\modern-uninstall.ico"

	;Delete VB DLLs
	Push $SYSDIR\Asycfilt.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Comcat.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Msvbvm60.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Oleaut32.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Olepro32.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Stdole2.tlb
	Call un.RemoveSharedDLL
	Push $SYSDIR\VB6FR.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\MSCOMCTL.OCX
	Call un.RemoveSharedDLL

	;Delete Start Menu Shortcuts
	Delete "$SMPROGRAMS\ALC Touch\*.*"
	RmDir "$SMPROGRAMS\ALC Touch"

	;Delete Uninstaller And Unistall Registry Entries
	Delete "$INSTDIR\Uninst.exe"
	DeleteRegKey HKEY_LOCAL_MACHINE "SOFTWARE\ALC Touch"
	DeleteRegKey HKEY_LOCAL_MACHINE "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ALC Touch"
	RMDir "$INSTDIR"

	;Display the Finish header
	!insertmacro MUI_UNFINISHHEADER
SectionEnd