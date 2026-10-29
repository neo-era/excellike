; Bộ cài ExcelLike cho Windows (NSIS 3, cài cho người dùng hiện tại, không cần quyền quản trị)
; makensis -DVERSION=0.4.0 -DSRC=release/win-unpacked -DOUT=release/ExcelLike-Setup-0.4.0.exe build/installer.nsi
Unicode true
Target amd64-unicode
ManifestDPIAware true
!include "MUI2.nsh"
!include "FileFunc.nsh"
!include "nsDialogs.nsh"
!include "LogicLib.nsh"
!include "WinMessages.nsh"

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
Page custom LicPage LicLeave
!insertmacro MUI_PAGE_COMPONENTS
!define MUI_PAGE_CUSTOMFUNCTION_LEAVE DirLeave
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Vietnamese"

; ---------- Kích hoạt bản quyền ----------
; Key ký số (Ed25519) do app Quản lý Key cấp. lic.exe kiểm chữ ký, Mã máy, hạn dùng ngay khi bấm Tiếp;
; key hợp lệ được lưu vào %APPDATA%\ExcelLike\license.key để ExcelLike tự kích hoạt khi mở.
; Cài im lặng (triển khai hàng loạt): ExcelLike-Setup.exe /S /KEY=XXXXX-XXXXX-…
Var LicKey
Var LicMachine
Var LicEdit
Var LicInfo

Function .onInit
  InitPluginsDir
  File "/oname=$PLUGINSDIR\lic.exe" "lic\lic.exe"
  nsExec::ExecToStack '"$PLUGINSDIR\lic.exe" machine'
  Pop $0
  Pop $LicMachine
  StrCpy $LicMachine $LicMachine 15
  ; key đã có từ lần cài trước → điền sẵn
  StrCpy $LicKey ""
  IfFileExists "$APPDATA\${APP}\license.key" 0 +5
    FileOpen $1 "$APPDATA\${APP}\license.key" r
    FileRead $1 $LicKey
    FileClose $1
    StrCpy $LicKey $LicKey 400
  ${GetParameters} $0
  ClearErrors
  ${GetOptions} $0 "/KEY=" $2
  ${IfNot} ${Errors}
    StrCpy $LicKey $2
  ${EndIf}
  ${If} ${Silent}
    Call LicCheck
    ${If} $0 != 0
      SetErrorLevel 3
      Quit
    ${EndIf}
  ${EndIf}
FunctionEnd

; kiểm $LicKey bằng lic.exe → $0 = mã kết quả, $1 = thông tin
Function LicCheck
  FileOpen $2 "$PLUGINSDIR\key.txt" w
  FileWrite $2 $LicKey
  FileClose $2
  nsExec::ExecToStack '"$PLUGINSDIR\lic.exe" verifyfile "$PLUGINSDIR\key.txt"'
  Pop $0
  Pop $1
FunctionEnd

Function LicPage
  !insertmacro MUI_HEADER_TEXT "Kích hoạt bản quyền" "Nhập key bản quyền ${APP} để tiếp tục cài đặt."
  nsDialogs::Create 1018
  Pop $0
  ${NSD_CreateLabel} 0 0 100% 24u "Dán key được cấp vào ô dưới (gồm các nhóm 5 ký tự, có hoặc không có dấu gạch đều được)."
  Pop $0
  nsDialogs::CreateControl EDIT ${DEFAULT_STYLES}|${WS_TABSTOP}|${ES_MULTILINE}|${ES_AUTOVSCROLL}|${WS_VSCROLL} ${WS_EX_CLIENTEDGE} 0 26u 100% 52u $LicKey
  Pop $LicEdit
  ${NSD_CreateLabel} 0 84u 100% 12u "Mã máy của máy này (gửi cho nơi cấp nếu cần key gắn với máy):"
  Pop $0
  ${NSD_CreateText} 0 98u 120u 13u $LicMachine
  Pop $0
  SendMessage $0 ${EM_SETREADONLY} 1 0
  ${NSD_CreateLabel} 0 118u 100% 22u "Chưa có key? Liên hệ LAVIPCO — email lamvietphat2022@gmail.com, điện thoại 0989 725 507."
  Pop $LicInfo
  ${NSD_SetFocus} $LicEdit
  nsDialogs::Show
FunctionEnd

Function LicLeave
  ${NSD_GetText} $LicEdit $LicKey
  Call LicCheck
  ${Switch} $0
    ${Case} 0
      Return
    ${Case} 10
      MessageBox MB_ICONEXCLAMATION "Key không đúng định dạng. Hãy dán nguyên key (gồm các nhóm 5 ký tự)."
      ${Break}
    ${Case} 11
      MessageBox MB_ICONEXCLAMATION "Key không hợp lệ (sai chữ ký). Kiểm tra lại từng ký tự hoặc liên hệ nơi cấp key."
      ${Break}
    ${Case} 12
      MessageBox MB_ICONEXCLAMATION "Key này không dành cho ${APP}."
      ${Break}
    ${Case} 13
      MessageBox MB_ICONEXCLAMATION "Key này cấp cho máy khác.$\r$\nMã máy của máy này: $LicMachine$\r$\nGửi Mã máy này cho nơi cấp để nhận key mới."
      ${Break}
    ${Case} 14
      MessageBox MB_ICONEXCLAMATION "Key đã hết hạn sử dụng."
      ${Break}
    ${Case} 15
      MessageBox MB_ICONEXCLAMATION "Key này đã bị thu hồi."
      ${Break}
    ${Default}
      MessageBox MB_ICONEXCLAMATION "Không kiểm được key (mã $0)."
  ${EndSwitch}
  Abort
FunctionEnd

; luôn cài vào thư mục riêng “…\ExcelLike” để không ghi đè/xóa tệp của chương trình khác
Function DirLeave
  StrCpy $0 $INSTDIR "" -10
  StrCmp $0 "\${APP}" +2 0
    StrCpy $INSTDIR "$INSTDIR\${APP}"
FunctionEnd

Section "${APP} (bắt buộc)" SecMain
  SectionIn RO
  ; đóng ExcelLike đang chạy (kể cả tiến trình chạy ngầm) để không bị khóa tệp khi chép đè bản mới
  nsExec::Exec 'taskkill /F /T /IM "${EXE}"'
  Pop $0
  Sleep 1000
  SetOutPath "$INSTDIR"
  File /r "${SRC}\*.*"
  WriteUninstaller "$INSTDIR\Uninstall ${APP}.exe"
  ; lưu key đã kiểm vào thư mục dữ liệu của ExcelLike (giữ lại khi gỡ cài đặt để cài lại không phải nhập)
  CreateDirectory "$APPDATA\${APP}"
  FileOpen $1 "$APPDATA\${APP}\license.key" w
  FileWrite $1 $LicKey
  FileClose $1
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
