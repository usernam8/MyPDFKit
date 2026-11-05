// The Microsoft Store version: the full app (with the built-in AI) as an MSIX package. Microsoft signs it and
// updates it, so it needs no certificate and the app's own updater stays off there (process.windowsStore).
// The three names below come from Partner Center (Product > Product identity) after reserving "MyPDFKit";
// put them in store-identity.json:  { "identityName": "...", "publisher": "CN=...", "publisherDisplayName": "..." }
// Build:  npx electron-builder --win appx --publish never --config electron-builder.store.js
const fs = require('fs');
const base = require('./package.json').build;
let id = { identityName: 'MyPDFKit.Test', publisher: 'CN=MyPDFKitTest', publisherDisplayName: 'MyPDFKit' };   // placeholders until then
try { id = { ...id, ...JSON.parse(fs.readFileSync(__dirname + '/store-identity.json', 'utf8')) }; } catch (e) {}
module.exports = {
  ...base,
  extraResources: [...base.extraResources, { from: 'vendor/ai', to: 'ai' }],
  win: { ...base.win, target: ['appx'] },
  appx: {
    ...id,
    applicationId: 'MyPDFKit',
    displayName: 'MyPDFKit',
    backgroundColor: '#1c1d20',
    languages: ['en-US', 'es-ES', 'de-DE', 'fr-FR', 'pt-BR', 'zh-CN', 'ja-JP', 'ko-KR'],
    artifactName: 'MyPDFKit-Store-${version}.${ext}',
  },
};
