const { app, BrowserWindow, dialog, Menu, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const http = require('http');

const SERVER_PORT = 5000;
let serverProc = null;
let win = null;

// Only allow one running copy; a second launch (e.g. double-clicking another file)
// hands its file to the first window.
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
}

// everything the backend can turn into a PDF (keep in step with OPENABLE_EXT in server.py)
const OPENABLE = new RegExp('\\.(' +['pdf', 'doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'odt', 'wps', 'wpd',
  'xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'csv', 'tsv', 'ods', 'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'pot', 'potx', 'odp',
  'jpg', 'jpeg', 'jfif', 'png', 'bmp', 'gif', 'tif', 'tiff', 'webp', 'ico', 'pnm', 'jxr', 'xps', 'oxps', 'epub', 'fb2', 'cbz', 'mobi', 'svg',
  'html', 'htm', 'mhtml', 'mht', 'xhtml', 'txt', 'md', 'markdown', 'log', 'json', 'xml', 'yaml', 'yml', 'ini', 'cfg', 'conf'].join('|') + ')$', 'i');
function filesFromArgv(argv) {
  return argv.slice(app.isPackaged ? 1 : 2).filter(a => OPENABLE.test(a) && fs.existsSync(a));
}

function startServer() {
  if (app.isPackaged) {
    // a PyInstaller folder build: starts at once, unlike a one-file exe that unpacks itself every launch
    const exe = path.join(process.resourcesPath, 'server', 'server.exe');
    serverProc = spawn(exe, [], { windowsHide: true });
  } else {
    // Dev mode: run the Python script directly
    serverProc = spawn('python', [path.join(__dirname, 'server.py')], { windowsHide: true });
  }
  serverProc.on('error', err => dialog.showErrorBox('Backend failed to start', String(err)));
  serverProc.stderr && serverProc.stderr.on('data', d => console.error('[server]', d.toString()));
}

function waitForServer(retries = 200) {
  return new Promise((resolve, reject) => {
    const attempt = n => {
      http.get(`http://127.0.0.1:${SERVER_PORT}/ping`, res => { res.resume(); resolve(); })
        .on('error', () => (n <= 0 ? reject(new Error('Backend did not start')) : setTimeout(() => attempt(n - 1), 150)));
    };
    attempt(retries);
  });
}

let pageLoaded = false, serverUp = false;
function flushOpenFiles() {
  if (!pageLoaded || !serverUp || !win) return;
  const files = filesFromArgv(process.argv);
  if (files.length) win.webContents.send('open-files', files);
}

function createWindow() {
  win = new BrowserWindow({
    width: 1400, height: 900, minWidth: 980, minHeight: 620, title: 'MyPDFKit',
    backgroundColor: '#25262a',
    show: false,
    autoHideMenuBar: true,
    // The app's own top bar acts as the title bar; Windows draws the min/max/close buttons over it.
    titleBarStyle: 'hidden',
    titleBarOverlay: { color: '#1c1d20', symbolColor: '#e6e5e1', height: 44 },
    webPreferences: {
      // index.html uses require('fs') / require('electron') directly.
      // Safe here because the page now loads only local files (pdf.js is bundled).
      nodeIntegration: true,
      contextIsolation: false,
    },
  });
  if (app.isPackaged) Menu.setApplicationMenu(null);
  win.once('ready-to-show', () => win.show());
  // closing the main window ends the app, even if a hidden helper window (printing) is still open
  win.on('closed', () => { win = null; app.quit(); });
  // Safety net: never leave the window hidden if the first paint is delayed or fails
  setTimeout(() => { if (win && !win.isDestroyed() && !win.isVisible()) win.show(); }, 2500);

  // Print renderer problems in the terminal so they can be diagnosed
  win.webContents.on('did-fail-load', (_e, code, desc, url) => console.error('[load failed]', code, desc, url));
  win.webContents.on('render-process-gone', (_e, d) => console.error('[renderer gone]', d.reason));
  win.webContents.on('console-message', (...a) => {
    const e = a[0];
    const msg = e && e.message !== undefined ? e.message : a[2];
    console.log('[page]', msg);
  });
  win.loadFile('index.html');
  win.webContents.on('did-finish-load', () => { pageLoaded = true; flushOpenFiles(); });
}

app.on('second-instance', (_e, argv) => {
  if (!win || win.isDestroyed()) return;
  if (win.isMinimized()) win.restore();
  win.focus();
  const files = filesFromArgv(argv);
  if (files.length) win.webContents.send('open-files', files);
});

if (gotLock) app.whenReady().then(async () => {
  startServer();
  createWindow();   // show the window immediately; the backend finishes starting in parallel
  try { await waitForServer(); serverUp = true; flushOpenFiles(); }
  catch (e) { dialog.showErrorBox('MyPDFKit', e.message); }
  setupAutoUpdate();
});

// ---------- updates ----------
// The installed app checks GitHub (the public MyPDFKit-releases repo) at startup and every few hours,
// downloads a newer version in the background, then offers a restart. If the user ignores it, the
// update is installed the next time the app closes.
function setupAutoUpdate() {
  if (!app.isPackaged) return;
  let autoUpdater;
  try { ({ autoUpdater } = require('electron-updater')); } catch (e) { console.error('[update] unavailable', e); return; }
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.on('update-downloaded', info => {
    if (win && !win.isDestroyed()) win.webContents.send('update-ready', info.version);
  });
  autoUpdater.on('error', err => console.error('[update]', err && err.message));
  const check = () => autoUpdater.checkForUpdates().catch(err => console.error('[update]', err && err.message));
  setTimeout(check, 4000);   // after startup settles
  setInterval(check, 4 * 60 * 60 * 1000);
  ipcMain.on('install-update', () => { killServer(); autoUpdater.quitAndInstall(true, true); });
}
ipcMain.handle('app-version', () => app.getVersion());

// ---------- printing ----------
// The renderer sends page pictures (made by the backend); they go into a hidden window, one per sheet,
// and Windows' own print dialog takes it from there (printer, copies, pages, paper).
ipcMain.handle('print-pages', async (_e, pages) => {
  if (!pages || !pages.length) return { error: 'Nothing to print' };
  const landscape = pages.filter(p => p.w > p.h).length > pages.length / 2;
  const html = `<!doctype html><meta charset="utf-8"><style>
    @page { margin: 0; }
    html, body { margin: 0; background: #fff; }
    .pg { width: 100vw; height: 100vh; display: flex; align-items: center; justify-content: center; page-break-after: always; break-after: page; overflow: hidden; }
    .pg:last-child { page-break-after: auto; break-after: auto; }
    img { max-width: 100%; max-height: 100%; object-fit: contain; }
  </style>` + pages.map(p => `<div class="pg"><img src="${require('url').pathToFileURL(p.path).href}"></div>`).join('');
  const file = path.join(path.dirname(pages[0].path), 'print.html');
  fs.writeFileSync(file, html);
  const pw = new BrowserWindow({ show: false, parent: win, webPreferences: { sandbox: true } });
  try {
    await pw.loadFile(file);
    await pw.webContents.executeJavaScript('Promise.all([...document.images].map(i => i.decode().catch(() => {})))');
    // pageSize must be given: Electron 34 otherwise sends an empty page size and no dialog ever opens
    return await new Promise(resolve => pw.webContents.print({ silent: false, printBackground: true, landscape, pageSize: 'Letter', margins: { marginType: 'none' } },
      (ok, reason) => resolve(ok || reason === 'cancelled' || reason === 'Print job canceled' ? { ok } : { error: "Couldn't print: " + reason })));
  } finally {
    pw.destroy();
    fs.rm(path.dirname(file), { recursive: true, force: true }, () => {});
  }
});

function killServer() {
  if (serverProc && !serverProc.killed) {
    // /T kills the whole process tree
    spawn('taskkill', ['/pid', String(serverProc.pid), '/f', '/t'], { windowsHide: true });
  }
}
app.on('before-quit', killServer);
app.on('window-all-closed', () => { killServer(); app.quit(); });
