;--------------------------------
;General

;Name and file
Name "MyVaultBrowser"
OutFile "MyVaultBrowser.exe"

Unicode true
SetCompressor /SOLID lzma

SetFont "Tahoma"  8

VIAddVersionKey "ProductName" "MyVaultBrowser"
VIAddVersionKey "LegalCopyright" "Copyright 2016 ${U+00A9} smilinger"
VIAddVersionKey "FileDescription" "Separate Vault Browser for Autodesk Inventor"
VIAddVersionKey "FileVersion" "0.9.9.0"
VIAddVersionKey "ProductVersion" "0.9.9.0"
VIProductVersion "0.9.9.0"

;--------------------------------
;Include Modern UI
!include "MUI2.nsh"

!define MULTIUSER_EXECUTIONLEVEL Highest
!define MULTIUSER_MUI

!include "MultiUser.nsh"

!include "LogicLib.nsh"
!include "WinVer.nsh"
!include "x64.nsh"
!include "nsProcess.nsh"

;--------------------------------
;Interface Settings

!define MUI_ABORTWARNING

;--------------------------------
;Language Selection Dialog Settings

;Remember the installer language
!define MUI_ICON "MyVaultBrowser.ico"
!define MUI_LANGDLL_REGISTRY_ROOT "HKCU"
!define MUI_LANGDLL_REGISTRY_KEY "Software\MyVaultBrowser"
!define MUI_LANGDLL_REGISTRY_VALUENAME "Installer Language"
!define MUI_COMPONENTSPAGE_NODESC
!define MUI_FINISHPAGE_SHOWREADME $INSTDIR\Readme.html
!define MUI_FINISHPAGE_NOAUTOCLOSE

;--------------------------------
;Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "License.txt"
!insertmacro MULTIUSER_PAGE_INSTALLMODE
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

;--------------------------------
;Languages

!insertmacro MUI_LANGUAGE "English" ;first language is the default language
!insertmacro MUI_LANGUAGE "SimpChinese"

;--------------------------------
;Reserve Files

;If you are using solid compression, files that are required before
;the actual installation should be stored first in the data block,
;because this will make your installer start faster.

!insertmacro MUI_RESERVEFILE_LANGDLL

;--------------------------------
;Descriptions

;USE A LANGUAGE STRING IF YOU WANT YOUR DESCRIPTIONS TO BE LANGAUGE SPECIFIC
LangString NAME_ADDIN ${LANG_ENGLISH} "Core add-in"
LangString NAME_LANG ${LANG_ENGLISH} "Language files"
LangString NAME_ZH-CN ${LANG_ENGLISH} "Simplified Chinese"
LangString NAME_ADDIN ${LANG_SIMPCHINESE} "插件主文件"
LangString NAME_LANG ${LANG_SIMPCHINESE} "语言文件"
LangString NAME_ZH-CN ${LANG_SIMPCHINESE} "简体中文"

LangString InventorIsRunning ${LANG_ENGLISH} "Inventor is running, please close Inventor and click OK, or click Cancel to install later."
LangString InventorIsRunning ${LANG_SIMPCHINESE} "Inventor 还在运行，请先关闭 Inventor，然后点确定继续安装，或者点取消，稍后再安装。"

;--------------------------------
;Installer Sections

Section

    Loop:
    ${nsProcess::FindProcess} "Inventor.exe" $R0
    StrCmp $R0 0 0 +3
    MessageBox MB_OKCANCEL|MB_ICONEXCLAMATION $(InventorIsRunning) IDOK Loop IDCANCEL +1
    Abort

    ${If} $MultiUser.InstallMode == "AllUsers"
        ExpandEnvStrings $INSTDIR "$APPDATA\Autodesk\Inventor Addins\MyVaultBrowser"
    ${Else}
        ExpandEnvStrings $INSTDIR "$APPDATA\Autodesk\ApplicationPlugins\MyVaultBrowser"
    ${EndIf}

SectionEnd

Section $(NAME_ADDIN) Addin

    SectionIn RO
    SetOutPath "$INSTDIR"

    File ..\MyVaultBrowser\bin\Release\MyVaultBrowser.dll
    File ..\MyVaultBrowser\bin\Release\Autodesk.MyVaultBrowser.Inventor.Addin
    File Readme.html
    File License.txt
    File MyVaultBrowser.ico

    ;Store installation folder
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyVaultBrowser" \
               "DisplayName" "MyVaultBrowser"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyVaultBrowser" \
               "DisplayVersion" "0.9.9"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyVaultBrowser" \
               "DisplayIcon" "$\"$INSTDIR\MyVaultBrowser.ico$\""
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyVaultBrowser" \
               "Publisher" "smilinger"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyVaultBrowser" \
               "UninstallString" "$\"$INSTDIR\uninstall.exe$\""

SectionEnd

SectionGroup $(NAME_LANG) Lang

    Section /o $(NAME_ZH-CN) zh-CN
        File /oname=Readme.html Readme.zh-CN.html
        File /r ..\MyVaultBrowser\bin\Release\zh-CN
    SectionEnd

SectionGroupEnd

Section
    ;Create uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd
;--------------------------------
;Uninstaller Section

Section "Uninstall"

    ${If} ${RunningX64}
        SetRegView 64
    ${EndIf}

    RMDir /r "$INSTDIR"
    DeleteRegKey HKCU "Software\MyVaultBrowser"
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MyVaultBrowser"

SectionEnd

;--------------------------------
;Installer Functions

Function .onInit
    ${If} ${RunningX64}
        SetRegView 64
    ${EndIf}

    ${If} $LANGUAGE == 2052
        SectionSetFlags ${zh-CN} 1
    ${EndIf}

    !insertmacro MULTIUSER_INIT
    !insertmacro MUI_LANGDLL_DISPLAY

FunctionEnd

;Uninstaller Functions

Function un.onInit

    !insertmacro MULTIUSER_UNINIT
    !insertmacro MUI_UNGETLANGUAGE

FunctionEnd
