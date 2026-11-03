# Uploads the installer built by electron-builder to the public MyPDFKit-releases repo on GitHub.
# Installed copies of the app check that repo and update themselves.
# Usage: bump "version" in package.json, then `npm run release` (builds, then runs this).
$ErrorActionPreference = 'Stop'
$version = (Get-Content package.json -Raw | ConvertFrom-Json).version
$files = @("release\MyPDFKit-Setup-$version.exe", "release\MyPDFKit-Setup-$version.exe.blockmap", "release\latest.yml")
foreach ($f in $files) { if (-not (Test-Path $f)) { throw "Missing $f - run the build first." } }

# The release page is what people get sent to, so it says how to install, in plain words.
$notes = @"
## How to install

1. Under **Assets** below, click **MyPDFKit-Setup-$version.exe** to download it.
2. Open the downloaded file.
3. If a blue box says **"Windows protected your PC"**, click **More info**, then **Run anyway**.
   (MyPDFKit isn't signed with a paid certificate, so Windows doesn't recognize it yet. It's safe.)
4. Click **Yes** when Windows asks to allow changes, then follow the steps and click **Install**.
5. Open **MyPDFKit** from the desktop or the Start menu and drop in any file.

Everything it needs comes with it: it opens PDFs, pictures, web pages, and Word, Excel and PowerPoint files,
even on a computer without Microsoft Office. Updates install themselves.
"@
gh release create "v$version" @files --repo usernam8/MyPDFKit-releases --title "MyPDFKit $version" --notes $notes
Write-Host "Published MyPDFKit $version"
