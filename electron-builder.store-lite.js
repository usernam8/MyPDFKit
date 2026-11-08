// Toast Docs Lite, its own Microsoft Store app: the same program with "edition": "lite" in its package.json, so it
// only views, converts, merges and splits (plus the desktop widget), and without the built-in AI (no 1.3 GB inside).
// Its Store names come from Partner Center (Toast Docs Lite > Product identity), in store-identity-lite.json.
// Build:  npx electron-builder --win dir --publish never --config electron-builder.store-lite.js -c.directories.output=release-store-lite
//         python store/make_msix.py release-store-lite --lite
const base = require('./package.json').build;
module.exports = {
  ...base,
  extraMetadata: { edition: 'lite' },
  win: { ...base.win, target: ['dir'] },
};
