# Uploads both installers built by electron-builder to the public MyPDFKit-releases repo on GitHub.
# Installed copies of the app check that repo and update themselves (latest.yml points at the small Lite
# installer: the built-in AI lives in its own folder, so updates never download it again).
# Usage: bump "version" in package.json, build the full version (electron-builder --config electron-builder.full.js),
# then the Lite one (electron-builder), then run this.
$ErrorActionPreference = 'Stop'
$version = (Get-Content package.json -Raw | ConvertFrom-Json).version
$files = @("release\Toast-Docs-Setup-$version.exe", "release\Toast-Docs-Lite-Setup-$version.exe", "release\Toast-Docs-Lite-Setup-$version.exe.blockmap", "release\latest.yml")
foreach ($f in $files) { if (-not (Test-Path $f)) { throw "Missing $f - run both builds first." } }
if (-not (Select-String -Path release\latest.yml -Pattern "Toast-Docs-Lite-Setup-$version.exe" -Quiet)) { throw "latest.yml must point at the Lite installer: build the full version first, then the Lite one." }

# The built-in AI for the Lite version's one-time download: kept in a release of its own that is never
# replaced, so the download keeps working for years. Made once (it's older than every app release, so it's
# never mistaken for the newest version).
$model = 'vendor\ai\Qwen3.5-2B-Q4_K_M.gguf'
gh release view ai-model-1 --repo usernam8/MyPDFKit-releases *> $null
if ($LASTEXITCODE -ne 0) {
  $aiNotes = Join-Path $env:TEMP 'mypdfkit-ai-notes.md'
  [IO.File]::WriteAllText($aiNotes, "The built-in AI for MyPDFKit Lite (downloaded by the app when someone turns the AI on). You don't need to download this yourself; the full MyPDFKit installer already includes it.`n`nModel: Qwen3.5-2B (Apache 2.0 license), run by llama.cpp (MIT license).", (New-Object Text.UTF8Encoding $false))
  gh release create ai-model-1 $model --repo usernam8/MyPDFKit-releases --title "Built-in AI for MyPDFKit Lite" --notes-file $aiNotes --latest=false
  if ($LASTEXITCODE -ne 0) { throw "Uploading the AI failed (gh exited with $LASTEXITCODE)" }
  Remove-Item $aiNotes
}

# The release page is what people get sent to, so it says how to install, in plain words.
$notes = @"
## How to install

1. Under **Assets** below, click **Toast-Docs-Setup-$version.exe** to download it (about 1.5 GB, with the built-in AI).
   On a slow connection or a small disk, take **Toast-Docs-Lite-Setup-$version.exe** instead (about 200 MB); it can add the AI later.
2. Open the downloaded file.
3. If a blue box says **"Windows protected your PC"**, click **More info**, then **Run anyway**.
   (Toast Docs isn't signed with a paid certificate, so Windows doesn't recognize it yet. It's safe.)
4. Click **Yes** when Windows asks to allow changes, then follow the steps and click **Install**.
5. Open **Toast Docs** from the desktop or the Start menu and drop in any file.

Everything it needs comes with it: it opens PDFs, pictures, web pages, and Word, Excel and PowerPoint files,
even on a computer without Microsoft Office. The **AI** button summarizes and answers questions about your files
right on your computer: no account, no internet, nothing sent anywhere, free. Updates install themselves.
"@
# (from a file: Windows PowerShell breaks quotes inside a long argument)
$notesFile = Join-Path $env:TEMP "mypdfkit-notes-$version.md"
[IO.File]::WriteAllText($notesFile, $notes, (New-Object Text.UTF8Encoding $false))
gh release create "v$version" @files --repo usernam8/MyPDFKit-releases --title "Toast Docs $version" --notes-file $notesFile --latest
if ($LASTEXITCODE -ne 0) { throw "Publishing failed (gh exited with $LASTEXITCODE)" }
Remove-Item $notesFile
Write-Host "Published Toast Docs $version"
