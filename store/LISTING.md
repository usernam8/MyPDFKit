# Toast Docs: Microsoft Store listing text

## Description
I made Toast Docs because my friends kept asking me how to sign a PDF, print a shipping label, or turn a Word file into a PDF without paying for a subscription.

So here it is. You can open PDFs, Word, Excel, PowerPoint files and pictures, even if you don't have Office. Fill in forms and sign them. Fix a typo in a PDF. Put pages together or pull them apart. Print labels on a 4x6 label printer and have them come out the right size.

There's also an AI helper that can summarize a long document or answer questions about it. It runs on your own computer, so your files don't go anywhere and it works without internet.

It's free, there are no ads, and you don't need an account. If something doesn't work, open an issue on the GitHub page and I'll take a look.

## What's new in this version
First version in the Microsoft Store.

## Product features (one per line)
Fill in and sign PDF forms
Fix text in a PDF
Print labels on 4x6 label printers
Turn Word, Excel, PowerPoint and pictures into PDFs and back
Summaries and answers from an AI that runs on your PC
Combine, split, shrink and rotate pages
Read the text in scanned pages
Password-protect files and black out private info
English, Spanish, German, French, Portuguese, Chinese, Japanese and Korean

## Short description
A free PDF app for signing, fixing, printing and converting your files.

## Search terms
PDF, PDF editor, sign PDF, label printer, PDF converter

## Screenshots
store/screenshots/1-document.png, store/screenshots/2-all-tools.png (1920x1080)

## Notes for certification
Desktop app built with Electron. The AI runs locally with llama.cpp and a bundled model, so testers don't need to sign in to anything.
The online AI options are optional and need the user's own key.

## runFullTrust justification
Toast Docs is a desktop app. It opens and saves the user's own files and runs its own PDF engine next to the window.
