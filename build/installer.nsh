; MyPDFKit takes over no file type when installed. It offers itself in "Open with" for PDFs and is listed in
; Windows' Default apps, where people can choose it for PDFs themselves.

!macro customInstall
  WriteRegStr SHCTX "Software\Classes\MyPDFKit.PDF" "" "PDF Document"
  WriteRegStr SHCTX "Software\Classes\MyPDFKit.PDF\DefaultIcon" "" "$INSTDIR\${APP_EXECUTABLE_FILENAME},0"
  WriteRegStr SHCTX "Software\Classes\MyPDFKit.PDF\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  WriteRegStr SHCTX "Software\Classes\.pdf\OpenWithProgids" "MyPDFKit.PDF" ""
  WriteRegStr SHCTX "Software\MyPDFKit\Capabilities" "ApplicationName" "MyPDFKit"
  WriteRegStr SHCTX "Software\MyPDFKit\Capabilities" "ApplicationDescription" "Edit, sign, convert and print PDFs"
  WriteRegStr SHCTX "Software\MyPDFKit\Capabilities\FileAssociations" ".pdf" "MyPDFKit.PDF"
  WriteRegStr SHCTX "Software\RegisteredApplications" "MyPDFKit" "Software\MyPDFKit\Capabilities"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
  ; The full version: the built-in AI moves out of the app folder into its own (C:\ProgramData\MyPDFKit\ai),
  ; so updates, which replace the app folder, keep it and stay small.
  SetShellVarContext all
  IfFileExists "$INSTDIR\resources\ai\*.gguf" 0 mypdfkit_noai
    CreateDirectory "$APPDATA\MyPDFKit\ai"
    FindFirst $0 $1 "$INSTDIR\resources\ai\*.*"
    mypdfkit_loop:
      StrCmp $1 "" mypdfkit_done
      StrCmp $1 "." mypdfkit_next
      StrCmp $1 ".." mypdfkit_next
      Delete "$APPDATA\MyPDFKit\ai\$1"
      Rename "$INSTDIR\resources\ai\$1" "$APPDATA\MyPDFKit\ai\$1"
      mypdfkit_next:
      FindNext $0 $1
      Goto mypdfkit_loop
    mypdfkit_done:
    FindClose $0
    RMDir /r "$INSTDIR\resources\ai"
  mypdfkit_noai:
!macroend

!macro customUnInstall
  DeleteRegValue SHCTX "Software\Classes\.pdf\OpenWithProgids" "MyPDFKit.PDF"
  DeleteRegKey SHCTX "Software\Classes\MyPDFKit.PDF"
  DeleteRegValue SHCTX "Software\RegisteredApplications" "MyPDFKit"
  DeleteRegKey SHCTX "Software\MyPDFKit\Capabilities"
  ; Removing MyPDFKit (not updating it): the built-in AI goes too.
  ${ifNot} ${isUpdated}
    SetShellVarContext all
    RMDir /r "$APPDATA\MyPDFKit\ai"
    RMDir "$APPDATA\MyPDFKit"
    SetShellVarContext current
    RMDir /r "$LOCALAPPDATA\MyPDFKit\ai"
    SetShellVarContext all
  ${endIf}
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
