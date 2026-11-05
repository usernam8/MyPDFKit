"""Makes the Microsoft Store package (MSIX) from the built app, with Windows' own packaging tool (makeappx
from the Windows SDK). Microsoft signs the package when it's accepted in the Store, so no certificate is needed.

  1. npx electron-builder --win dir --publish never --config electron-builder.store.js -c.directories.output=release-store
  2. python store/make_msix.py
     -> release-store/MyPDFKit-Store-<version>.msix  (upload it in Partner Center: Packages)

The three Store names (identityName, publisher, publisherDisplayName) come from Partner Center, Product
identity, after reserving the app's name; put them in store-identity.json (see electron-builder.store.js)."""
import os, sys, json, glob, subprocess, shutil
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
out_dir = os.path.join(ROOT, sys.argv[1] if len(sys.argv) > 1 else 'release-store')
app = os.path.join(out_dir, 'win-unpacked')
version = json.load(open(os.path.join(ROOT, 'package.json'), encoding='utf-8'))['version']
ident = {'identityName': 'MyPDFKit.Test', 'publisher': 'CN=MyPDFKitTest', 'publisherDisplayName': 'MyPDFKit', 'displayName': 'MyPDFKit'}
try: ident.update(json.load(open(os.path.join(ROOT, 'store-identity.json'), encoding='utf-8')))
except FileNotFoundError: print('store-identity.json not found: using test names (fine for a test, not for the Store)')

kits = sorted(glob.glob(r'C:\Program Files (x86)\Windows Kits\10\bin\10.*\x64\makeappx.exe'))
if not kits: sys.exit('makeappx.exe not found: install the Windows SDK')
makeappx = kits[-1]

# the Store's tiles, from the app's icon
assets = os.path.join(out_dir, 'msix-assets'); os.makedirs(assets, exist_ok=True)
ico = Image.open(os.path.join(ROOT, 'build', 'icon.ico'))
ico.size = max(ico.ico.sizes()); src = ico.convert('RGBA')
def tile(name, w, h, pad=0.0):
    im = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    s = int(min(w, h) * (1 - pad)); icon = src.resize((s, s), Image.LANCZOS)
    im.paste(icon, ((w - s) // 2, (h - s) // 2), icon); im.save(os.path.join(assets, name))
tile('StoreLogo.png', 50, 50); tile('Square44x44Logo.png', 44, 44); tile('Square150x150Logo.png', 150, 150, 0.3)
tile('Wide310x150Logo.png', 310, 150, 0.25)

v4 = '.'.join((version.split('.') + ['0', '0', '0'])[:3]) + '.0'
manifest = f'''<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
         xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
         xmlns:uap3="http://schemas.microsoft.com/appx/manifest/uap/windows10/3"
         xmlns:desktop="http://schemas.microsoft.com/appx/manifest/desktop/windows10"
         xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
         IgnorableNamespaces="uap uap3 desktop rescap">
  <Identity Name="{ident['identityName']}" Publisher="{ident['publisher']}" Version="{v4}" ProcessorArchitecture="x64"/>
  <Properties>
    <DisplayName>{ident['displayName']}</DisplayName>
    <PublisherDisplayName>{ident['publisherDisplayName']}</PublisherDisplayName>
    <Logo>assets\\StoreLogo.png</Logo>
  </Properties>
  <Dependencies><TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.17763.0" MaxVersionTested="10.0.26100.0"/></Dependencies>
  <Resources><Resource Language="en-US"/><Resource Language="es-ES"/><Resource Language="de-DE"/><Resource Language="fr-FR"/><Resource Language="pt-BR"/><Resource Language="zh-CN"/><Resource Language="ja-JP"/><Resource Language="ko-KR"/></Resources>
  <Capabilities><rescap:Capability Name="runFullTrust"/><Capability Name="internetClient"/></Capabilities>
  <Applications>
    <Application Id="MyPDFKit" Executable="app\\MyPDFKit.exe" EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements DisplayName="{ident['displayName']}" Description="Edit, sign, convert, print and summarize PDFs"
          BackgroundColor="transparent" Square150x150Logo="assets\\Square150x150Logo.png" Square44x44Logo="assets\\Square44x44Logo.png">
        <uap:DefaultTile Wide310x150Logo="assets\\Wide310x150Logo.png"/>
      </uap:VisualElements>
      <Extensions>
        <uap3:Extension Category="windows.fileTypeAssociation">
          <uap3:FileTypeAssociation Name="pdf">
            <uap:SupportedFileTypes><uap:FileType>.pdf</uap:FileType></uap:SupportedFileTypes>
            <uap:DisplayName>PDF Document</uap:DisplayName>
          </uap3:FileTypeAssociation>
        </uap3:Extension>
      </Extensions>
    </Application>
  </Applications>
</Package>
'''
mf = os.path.join(out_dir, 'AppxManifest.xml'); open(mf, 'w', encoding='utf-8').write(manifest)
# what goes in the package: the app under app\, the tiles under assets\
mapping = os.path.join(out_dir, 'msix-files.txt')
with open(mapping, 'w', encoding='utf-8') as f:
    f.write('[Files]\n')
    f.write(f'"{mf}" "AppxManifest.xml"\n')
    for n in os.listdir(assets): f.write(f'"{os.path.join(assets, n)}" "assets\\{n}"\n')
    for r, _, fs in os.walk(app):
        for n in fs:
            # "[Content_Types].xml" and "_rels" folders are names MSIX keeps for itself; the only ones in the app are in
            # python-docx's unused template source folder (it uses templates\default.docx), so that folder is left out
            if n.lower() == '[content_types].xml' or '_rels' in os.path.relpath(r, app).lower().split(os.sep) or 'default-docx-template' in r: continue
            full = os.path.join(r, n); f.write(f'"{full}" "app\\{os.path.relpath(full, app)}"\n')
out = os.path.join(out_dir, f'MyPDFKit-Store-{version}.msix')
if os.path.exists(out): os.remove(out)
r = subprocess.run([makeappx, 'pack', '/f', mapping, '/p', out, '/o'], capture_output=True, text=True)
print((r.stdout or '')[-600:], (r.stderr or '')[-600:])
if r.returncode: sys.exit('makeappx failed')
print(f'{out}  ({os.path.getsize(out) / 1e9:.2f} GB)')
