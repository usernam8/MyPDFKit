// The full version: the same app with the built-in AI packed in (it works offline from the first click).
// Build it before the Lite one: both share one update feed (latest.yml), which must point at the small Lite
// installer, so updates stay small. The model is moved out of the app folder at install (build/installer.nsh),
// so updating with the Lite installer keeps it.
const base = require('./package.json').build;
module.exports = {
  ...base,
  extraResources: [...base.extraResources, { from: 'vendor/ai', to: 'ai' }],
  nsis: { ...base.nsis, artifactName: 'MyPDFKit-Setup-${version}.${ext}' },
};
