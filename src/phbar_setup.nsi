
; NSIS Excel Add-In Installer Script
; Include
!include MUI.nsh
!include LogicLib.nsh

; General
Name "NSIS Test"
OutFile "Setup.exe"
InstallDir "$PROGRAMFILES\NSIS Test"
InstallDirRegKey HKCU "Software\NSIS Test" "InstallDir" ;
Overrides InstallDir

; Interface Settings
!define MUI_ABORTWARNING

; Installer Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

; Uninstaller Pages
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Languages
!insertmacro MUI_LANGUAGE "English"

; Installer Section
Section "-Install"
SetOutPath "$INSTDIR"

; ADD FILES HERE
File "NSISTest.xla"
File "readme.txt"

; Check Installed Excel Version
ReadRegStr $1 HKCR "Excel.Application\CurVer" ""

${If} $1 == 'Excel.Application.8' ; Excel 95
StrCpy $2 "8.0"
${ElseIf} $1 == 'Excel.Application.9' ; Excel 2000
StrCpy $2 "9.0"
${ElseIf} $1 == 'Excel.Application.10' ; Excel XP
StrCpy $2 "10.0"
${ElseIf} $1 == 'Excel.Application.11' ; Excel 2003
StrCpy $2 "11.0"
${Else}
Abort "An appropriate version of Excel is not installed.
$\nNSIS Test setup will be canceled."
${EndIf}

; Find available "OPEN" key
StrCpy $3 ""
loop:
ReadRegStr $4 HKCU
"Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${If} $4 == ""
; Available OPEN key found
${Else}
IntOp $3 $3 + 1
Goto loop
${EndIf}

; Write install data to registry
WriteRegStr HKCU "Software\NSIS Test" "InstallDir" $INSTDIR
; Install Directory
WriteRegStr HKCU "Software\NSIS Test" "ExcelCurVer" $2
; Current Excel Version

; Write key to install AddIn in Excel Addin Manager
WriteRegStr HKCU "Software\Microsoft\Office\$2\Excel\Options"
"OPEN$3" '"$INSTDIR\NSISTest.xla"'

; Write keys to uninstall
WriteRegStr HKLM
"Software\Microsoft\Windows\CurrentVersion\Uninstall\NSIS Test"
"DisplayName" "NSIS Test"
WriteRegStr HKLM
"Software\Microsoft\Windows\CurrentVersion\Uninstall\NSIS Test"
"UninstallString" '"$INSTDIR\uninstall.exe"'
WriteRegDWORD HKLM
"Software\Microsoft\Windows\CurrentVersion\Uninstall\NSIS Test"
"NoModify" 1
WriteRegDWORD HKLM
"Software\Microsoft\Windows\CurrentVersion\Uninstall\NSIS Test"
"NoRepair" 1

; Create uninstaller
WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

; Uninstaller Section
Section "Uninstall"
; ADD FILES HERE...
Delete "$INSTDIR\NSISTest.xla"
;Delete "$INSTDIR\EngFunctHelp.chm"
Delete "$INSTDIR\readme.txt"
Delete "$INSTDIR\uninstall.exe"

RMDir "$INSTDIR"

; Find AddIn Manager Key and Delete
; AddIn Manager key name and location may have changed since
installation depending on actions taken by user in AddIn Manager.
; Need to search for the target AddIn key and delete if found.
ReadRegStr $2 HKCU "Software\NSIS Test" "ExcelCurVer"
StrCpy $3 ""

loop:
ReadRegStr $4 HKCU
"Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${If} $4 == '"$INSTDIR\NSISTest.xla"'
; Found Key
DeleteRegValue HKCU
"Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${ElseIf} $4 == ""
; Blank Key Found. Addin is no longer installed in
AddIn Manager.
; Need to delete Addin Manager Reference.
DeleteRegValue HKCU
"Software\Microsoft\Office\$2\Excel\Add-in Manager"
"$INSTDIR\NSISTest.xla"
${Else}
IntOp $3 $3 + 1
Goto loop
${EndIf}

DeleteRegKey HKCU "Software\NSIS Test"
DeleteRegKey HKLM
"Software\Microsoft\Windows\CurrentVersion\Uninstall\NSIS Test"
SectionEnd