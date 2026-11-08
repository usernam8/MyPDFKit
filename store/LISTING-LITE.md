# Toast Docs Lite: Microsoft Store listing text (English)

Store ID 9N7QXTH8CBD3, identity Mel.ToastDocsLite. Package: `python store/make_msix.py release-store-lite --lite`
(after `npx electron-builder --win dir --publish never --config electron-builder.store-lite.js -c.directories.output=release-store-lite`).

## Description
Toast Docs Lite is the simple version of Toast Docs: it opens your files and turns them into PDFs, and that's it.

Drop in Word, Excel, PowerPoint files or pictures and save them as PDFs, even if you don't have Office. Put several files together into one PDF, split a PDF into separate pages, or turn a PDF back into Word, Excel or PowerPoint.

It also comes with a little desktop widget (a bear holding a page): drag files onto it and it converts, merges or splits them right there, without opening anything.

It's free, there are no ads, and you don't need an account. Want to edit, sign or print labels too? Get Toast Docs, the full version.

## What's new in this version
First version in the Microsoft Store.

## Product features (one per line)
Turn Word, Excel, PowerPoint and pictures into PDFs
Turn PDFs back into Word, Excel or PowerPoint
Combine several files into one PDF
Split a PDF into separate pages
A desktop drop box: drag files on it to convert, merge or split
Works without Microsoft Office
No account, no ads, free

## Short description
Turn your files into PDFs, combine them or split them. Free and simple.

## Search terms
PDF, PDF converter, Word to PDF, merge PDF, split PDF, JPG to PDF, combine PDF

## Screenshots
store/screenshots/lite-1-viewer.png, store/screenshots/lite-2-tools.png

## Notes for certification
Desktop app (Electron). It views PDFs and converts, merges and splits files with its own built-in engine. No account and no sign-in; the AI features of the full version are not in Lite.

## runFullTrust justification
It's a desktop app that opens and saves the user's files and runs its own PDF conversion engine.
