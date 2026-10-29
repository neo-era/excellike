; Bộ cài ExcelLike Key Manager (NSIS 3, cài cho người dùng hiện tại)
Unicode true
Target amd64-unicode
ManifestDPIAware true
!include "MUI2.nsh"
!include "FileFunc.nsh"
!define APP "ExcelLike Key Manager"
!define EXE "ExcelLike Key Manager.exe"
!define UNKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExcelLikeKeyManager"
Name "${APP} ${VERSION}"
OutFile "${OUT}"
InstallDir "$LOCALAPPDATA\Programs\ExcelLikeKeyManager"
RequestExecutionLevel user
SetCompressor /SOLID lzma
BrandingText "${APP} ${VERSION}"
VIProductVersion "${VERSION}.0"
VIAddVersionKey /LANG=1066 "ProductName" "${APP}"
VIAddVersionKey /LANG=1066 "FileDescription" "Bộ cài ${APP}"
VIAddVersionKey /LANG=1066 "FileVersion" "${VERSION}"
VIAddVersionKey /LANG=1066 "ProductVersion" "${VERSION}"
VIAddVersionKey /LANG=1066 "LegalCopyright" "© 2026 Mai Vũ Lâm"
VIAddVersionKey /LANG=1066 "CompanyName" "LAVIPCO"
!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"
!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\${EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "Mở ${APP} ngay"
!define MUI_WELCOMEPAGE_TEXT "Trình cài đặt sẽ cài ${APP} ${VERSION} — công cụ cấp và quản lý key bản quyền ExcelLike.$\r$\n$\r$\nChỉ cài trên máy của người cấp key. Sau khi cài cần nhập tệp khóa ký.$\r$\n$\r$\nBấm Tiếp để tiếp tục."
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Vietnamese"
Section "Chương trình" SecMain
  SectionIn RO
  nsExec::Exec 'taskkill /F /T /IM "${EXE}"'
  Pop $0
  Sleep 800
  SetOutPath "$INSTDIR"
  File /r "${SRC}\*.*"
  WriteUninstaller "$INSTDIR\Gỡ cài đặt.exe"
  CreateDirectory "$SMPROGRAMS\${APP}"
  CreateShortcut "$SMPROGRAMS\${APP}\${APP}.lnk" "$INSTDIR\${EXE}" "" "$INSTDIR\${EXE}" 0
  CreateShortcut "$DESKTOP\${APP}.lnk" "$INSTDIR\${EXE}" "" "$INSTDIR\${EXE}" 0
  WriteRegStr HKCU "${UNKEY}" "DisplayName" "${APP}"
  WriteRegStr HKCU "${UNKEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "${UNKEY}" "Publisher" "LAVIPCO"
  WriteRegStr HKCU "${UNKEY}" "DisplayIcon" "$INSTDIR\${EXE}"
  WriteRegStr HKCU "${UNKEY}" "UninstallString" '"$INSTDIR\Gỡ cài đặt.exe"'
  WriteRegDWORD HKCU "${UNKEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNKEY}" "NoRepair" 1
SectionEnd
Section "Uninstall"
  Delete "$DESKTOP\${APP}.lnk"
  RMDir /r "$SMPROGRAMS\${APP}"
  DeleteRegKey HKCU "${UNKEY}"
  !include "uninst-files.nsh"
  Delete "$INSTDIR\Gỡ cài đặt.exe"
  RMDir "$INSTDIR"
  ; sổ key và khóa ký (AppData\Roaming\ExcelLike Key Manager) được GIỮ LẠI
SectionEnd
