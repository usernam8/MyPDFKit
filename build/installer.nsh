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
!macroend

!macro customUnInstall
  DeleteRegValue SHCTX "Software\Classes\.pdf\OpenWithProgids" "MyPDFKit.PDF"
  DeleteRegKey SHCTX "Software\Classes\MyPDFKit.PDF"
  DeleteRegValue SHCTX "Software\RegisteredApplications" "MyPDFKit"
  DeleteRegKey SHCTX "Software\MyPDFKit\Capabilities"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
