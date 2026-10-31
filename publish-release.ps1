# Uploads the installer built by electron-builder to the public MyPDFKit-releases repo on GitHub.
# Installed copies of the app check that repo and update themselves.
# Usage: bump "version" in package.json, then `npm run release` (builds, then runs this).
$ErrorActionPreference = 'Stop'
$version = (Get-Content package.json -Raw | ConvertFrom-Json).version
$files = @("release\MyPDFKit-Setup-$version.exe", "release\MyPDFKit-Setup-$version.exe.blockmap", "release\latest.yml")
foreach ($f in $files) { if (-not (Test-Path $f)) { throw "Missing $f - run the build first." } }
gh release create "v$version" @files --repo usernam8/MyPDFKit-releases --title "MyPDFKit $version" --notes "MyPDFKit $version"
Write-Host "Published MyPDFKit $version"
