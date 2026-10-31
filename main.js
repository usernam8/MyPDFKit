const { app, BrowserWindow, dialog, Menu } = require('electron');
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

const OPENABLE = /\.(pdf|docx?|xlsx?|pptx?|jpe?g|png)$/i;
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
  if (!win) return;
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
});

function killServer() {
  if (serverProc && !serverProc.killed) {
    // /T kills the whole process tree
    spawn('taskkill', ['/pid', String(serverProc.pid), '/f', '/t'], { windowsHide: true });
  }
}
app.on('before-quit', killServer);
app.on('window-all-closed', () => { killServer(); app.quit(); });
