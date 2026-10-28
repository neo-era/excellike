; Bộ cài ExcelLike cho Windows (NSIS 3, cài cho người dùng hiện tại, không cần quyền quản trị)
; makensis -DVERSION=0.4.0 -DSRC=release/win-unpacked -DOUT=release/ExcelLike-Setup-0.4.0.exe build/installer.nsi
Unicode true
Target amd64-unicode
ManifestDPIAware true
!include "MUI2.nsh"
!include "FileFunc.nsh"

!define APP "ExcelLike"
!define EXE "ExcelLike.exe"
!define UNKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}"

Name "${APP} ${VERSION}"
OutFile "${OUT}"
InstallDir "$LOCALAPPDATA\Programs\${APP}"
InstallDirRegKey HKCU "Software\${APP}" "InstallDir"
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
!define MUI_COMPONENTSPAGE_SMALLDESC
!define MUI_FINISHPAGE_RUN "$INSTDIR\${EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "Mở ${APP} ngay"
!define MUI_WELCOMEPAGE_TEXT "Trình cài đặt sẽ cài ${APP} ${VERSION} — phần mềm bảng tính kiểu Excel — cho tài khoản Windows của anh/chị.$\r$\n$\r$\nKhông cần quyền quản trị. Nếu ${APP} đang mở, hãy đóng trước khi cài.$\r$\n$\r$\nBấm Tiếp để tiếp tục."

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!define MUI_PAGE_CUSTOMFUNCTION_LEAVE DirLeave
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Vietnamese"

; luôn cài vào thư mục riêng “…\ExcelLike” để không ghi đè/xóa tệp của chương trình khác
Function DirLeave
  StrCpy $0 $INSTDIR "" -10
  StrCmp $0 "\${APP}" +2 0
    StrCpy $INSTDIR "$INSTDIR\${APP}"
FunctionEnd

Section "${APP} (bắt buộc)" SecMain
  SectionIn RO
  SetOutPath "$INSTDIR"
  File /r "${SRC}\*.*"
  WriteUninstaller "$INSTDIR\Uninstall ${APP}.exe"
  CreateDirectory "$SMPROGRAMS\${APP}"
  CreateShortcut "$SMPROGRAMS\${APP}\${APP}.lnk" "$INSTDIR\${EXE}" "" "$INSTDIR\${EXE}" 0
  CreateShortcut "$SMPROGRAMS\${APP}\Gỡ cài đặt ${APP}.lnk" "$INSTDIR\Uninstall ${APP}.exe"
  WriteRegStr HKCU "Software\${APP}" "InstallDir" "$INSTDIR"
  ; luôn có trong “Mở bằng…” cho .xlsx và .csv
  WriteRegStr HKCU "Software\Classes\${APP}.xlsx" "" "Sổ tính Excel"
  WriteRegStr HKCU "Software\Classes\${APP}.xlsx\DefaultIcon" "" "$INSTDIR\${EXE},0"
  WriteRegStr HKCU "Software\Classes\${APP}.xlsx\shell\open\command" "" '"$INSTDIR\${EXE}" "%1"'
  WriteRegStr HKCU "Software\Classes\${APP}.csv" "" "Tệp CSV"
  WriteRegStr HKCU "Software\Classes\${APP}.csv\DefaultIcon" "" "$INSTDIR\${EXE},0"
  WriteRegStr HKCU "Software\Classes\${APP}.csv\shell\open\command" "" '"$INSTDIR\${EXE}" "%1"'
  WriteRegStr HKCU "Software\Classes\.xlsx\OpenWithProgids" "${APP}.xlsx" ""
  WriteRegStr HKCU "Software\Classes\.csv\OpenWithProgids" "${APP}.csv" ""
  WriteRegStr HKCU "Software\Classes\Applications\${EXE}\SupportedTypes" ".xlsx" ""
  WriteRegStr HKCU "Software\Classes\Applications\${EXE}\SupportedTypes" ".csv" ""
  ; mục Gỡ cài đặt trong Cài đặt Windows
  WriteRegStr HKCU "${UNKEY}" "DisplayName" "${APP}"
  WriteRegStr HKCU "${UNKEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "${UNKEY}" "Publisher" "LAVIPCO"
  WriteRegStr HKCU "${UNKEY}" "DisplayIcon" "$INSTDIR\${EXE}"
  WriteRegStr HKCU "${UNKEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNKEY}" "UninstallString" '"$INSTDIR\Uninstall ${APP}.exe"'
  WriteRegDWORD HKCU "${UNKEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNKEY}" "NoRepair" 1
  ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
  WriteRegDWORD HKCU "${UNKEY}" "EstimatedSize" $0
SectionEnd

Section "Lối tắt trên màn hình" SecDesk
  CreateShortcut "$DESKTOP\${APP}.lnk" "$INSTDIR\${EXE}" "" "$INSTDIR\${EXE}" 0
SectionEnd

Section /o "Mở tệp .xlsx bằng ${APP} (mặc định)" SecXlsx
  WriteRegStr HKCU "Software\Classes\.xlsx" "" "${APP}.xlsx"
SectionEnd

Section /o "Mở tệp .csv bằng ${APP} (mặc định)" SecCsv
  WriteRegStr HKCU "Software\Classes\.csv" "" "${APP}.csv"
SectionEnd

Section "-thông báo Windows"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
SectionEnd

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecMain} "Chương trình ${APP}."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecDesk} "Tạo biểu tượng ${APP} trên màn hình."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecXlsx} "Nhấp đúp tệp .xlsx sẽ mở bằng ${APP} thay cho Excel. Không chọn thì vẫn mở được qua “Mở bằng…”."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecCsv} "Nhấp đúp tệp .csv sẽ mở bằng ${APP}."
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Section "Uninstall"
  Delete "$DESKTOP\${APP}.lnk"
  RMDir /r "$SMPROGRAMS\${APP}"
  ReadRegStr $0 HKCU "Software\Classes\.xlsx" ""
  StrCmp $0 "${APP}.xlsx" 0 +2
    DeleteRegValue HKCU "Software\Classes\.xlsx" ""
  ReadRegStr $0 HKCU "Software\Classes\.csv" ""
  StrCmp $0 "${APP}.csv" 0 +2
    DeleteRegValue HKCU "Software\Classes\.csv" ""
  DeleteRegValue HKCU "Software\Classes\.xlsx\OpenWithProgids" "${APP}.xlsx"
  DeleteRegValue HKCU "Software\Classes\.csv\OpenWithProgids" "${APP}.csv"
  DeleteRegKey HKCU "Software\Classes\${APP}.xlsx"
  DeleteRegKey HKCU "Software\Classes\${APP}.csv"
  DeleteRegKey HKCU "Software\Classes\Applications\${EXE}"
  DeleteRegKey HKCU "${UNKEY}"
  DeleteRegKey HKCU "Software\${APP}"
  ; chỉ xóa đúng các tệp đã cài (không xóa cả thư mục, phòng khi cài vào thư mục có tệp khác)
  !include "uninst-files.nsh"
  Delete "$INSTDIR\Uninstall ${APP}.exe"
  RMDir "$INSTDIR"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
  ; dữ liệu sổ tính (AppData\Roaming\ExcelLike) được giữ lại để cài lại không mất
SectionEnd
